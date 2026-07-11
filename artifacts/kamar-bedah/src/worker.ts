// src/worker.ts
// Worker entry untuk project "sistem-informasi-kamar-bedah".
// Menangani route /api/ocr-scan secara manual, lalu fallback ke static
// assets (env.ASSETS) untuk semua request lain — sesuai mode "Workers with
// static assets" yang dipakai project ini (BUKAN Cloudflare Pages).
//
// wrangler.toml WAJIB punya:
//   main = "src/worker.ts"
//   [assets]
//   directory = "./dist/public"
//   binding = "ASSETS"
//   [ai]
//   binding = "AI"

export interface Env {
  AI: Ai;
  ASSETS: Fetcher;
  // Secret — diisi lewat: npx wrangler secret put GROQ_API_KEY
  GROQ_API_KEY?: string;
}

const OCR_MODEL = "@cf/meta/llama-3.2-11b-vision-instruct";
// Model ini milik Meta — Cloudflare mewajibkan approve lisensi dulu lewat
// Cloudflare Dashboard → Workers AI → cari model ini → setujui, sebelum
// panggilan API berhasil.

// Model teks Groq — SATU-SATUNYA parser untuk jalur paste-teks (Tab
// Daftar). Dipilih dibanding CF karena Groq punya hardware inferensi
// khusus (LPU) yang jauh lebih cepat, dan model 120B-nya lebih akurat
// untuk ekstraksi data terstruktur dibanding model "flash" kecil.
// TRADE-OFF YANG DISADARI: karena hanya satu model, tidak ada lagi
// cross-check/fail-safe kalau model salah baca — dan kalau Groq down atau
// GROQ_API_KEY bermasalah, jalur ini gagal total tanpa fallback.
// "gpt-oss-120b" adalah pengganti resmi Groq untuk llama-3.3-70b-versatile
// (yang akan di-shutdown 16 Agustus 2026). Kalau nanti ingin ganti model
// lain, cek daftar model terkini di https://console.groq.com/docs/models.
const TEXT_MODEL_GROQ = "openai/gpt-oss-120b";
const GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions";

const EXTRACTION_PROMPT = `Baca gambar ini. Gambar berisi data permintaan operasi pasien kamar bedah rumah sakit — bisa berupa formulir, tulisan tangan, atau daftar dari chat WhatsApp.

LANGKAH KERJA:
1. Hitung dulu berapa jumlah pasien yang disebutkan (biasanya ditandai nomor urut 1, 2, 3, dst, atau nama "Ny./Tn./An." yang berbeda-beda).
2. Untuk SETIAP pasien yang kamu temukan, buat SATU objek JSON terpisah. Jumlah objek dalam array HARUS SAMA DENGAN jumlah pasien yang kamu hitung di langkah 1.
3. Untuk tiap pasien, isi field berikut apa adanya sesuai teks yang tertulis untuk pasien itu:
   - "patient": nama pasien
   - "age": HANYA angka umur, buang kata "th"/"tahun" dari teksnya. Contoh: teks "63th rencana ORIF" → age="63", procedure="ORIF pinning distal radius Sinistra" (umur TIDAK ikut ke procedure).
   - "procedure": nama tindakan/rencana operasi untuk pasien itu, TANPA angka umur di dalamnya.
   - "diagnosis": kosongkan jika tidak ada info diagnosis terpisah dari prosedur.
   - "surgeon": jika judul/header gambar menyebutkan nama dokter (misal "dr. Adhinanda dan dr. Timor"), isi nama dokter PERTAMA yang disebut ke "surgeon".
   - "anesthesiologist": isi nama dokter KEDUA yang disebut di judul/header ke field ini. Jika hanya ada satu nama dokter disebutkan, isi ke "surgeon" saja dan kosongkan "anesthesiologist".
   - "date", "time": HANYA isi jika benar-benar tertulis di gambar, format tanggal "YYYY-MM-DD" dan jam "HH:MM". JANGAN mengisi tanggal/jam hari ini atau tebakan apa pun jika tidak tertulis — kosongkan saja.
   - "opType": isi "elektif", "semi", atau "cyto" HANYA jika jelas disebutkan, selain itu kosongkan
   - Field lain ("rm", "room", "allergy", "bloodType") kosongkan jika tidak ada informasinya di gambar
   - "confidence": "high" jika teks jelas terbaca, "low" jika kabur/tulisan tangan sulit dibaca

CONTOH: jika judul gambar bertuliskan "Jam 12.00 dr. Adhinanda dan dr. Timor" dan salah satu baris berisi "Ny. Temu Rahayu, 63th rencana ORIF pinning distal radius Sinistra", maka objek untuk pasien itu adalah:
{"patient":"Ny. Temu Rahayu","age":"63","ageMonths":"","rm":"","diagnosis":"","procedure":"ORIF pinning distal radius Sinistra","opType":"","date":"","time":"12:00","surgeon":"dr. Adhinanda","anesthesiologist":"dr. Timor","room":"","allergy":"","bloodType":"","confidence":"high"}

ATURAN OUTPUT:
Balas HANYA dengan array JSON, dimulai dengan "[" dan diakhiri dengan "]". Tidak ada teks lain sebelum atau sesudahnya — tidak ada judul, penjelasan, atau kesimpulan.`;

// Prompt untuk jalur TEKS (Tab Daftar → paste teks). Sengaja dibuat mirip
// EXTRACTION_PROMPT gambar (field & aturan sama) supaya kedua model teks
// (CF & Groq) bisa dibandingkan hasilnya secara adil (apple-to-apple).
const TEXT_EXTRACTION_PROMPT = `Baca teks ini. Teks berisi data permintaan operasi pasien kamar bedah rumah sakit — bisa berupa memo, instruksi dokter, atau daftar hasil salin-tempel dari chat WhatsApp.

LANGKAH KERJA:
1. Hitung dulu berapa jumlah pasien yang disebutkan (biasanya ditandai nomor urut 1, 2, 3, dst, atau nama "Ny./Tn./An." yang berbeda-beda).
2. Untuk SETIAP pasien yang kamu temukan, buat SATU objek JSON terpisah. Jumlah objek dalam array HARUS SAMA DENGAN jumlah pasien yang kamu hitung di langkah 1.
3. Untuk tiap pasien, isi field berikut apa adanya sesuai teks yang tertulis untuk pasien itu:
   - "patient": nama pasien
   - "age": HANYA angka umur, buang kata "th"/"tahun" dari teksnya.
   - "procedure": nama tindakan/rencana operasi untuk pasien itu, TANPA angka umur di dalamnya.
   - "diagnosis": kosongkan jika tidak ada info diagnosis terpisah dari prosedur.
   - "surgeon": jika teks menyebutkan nama dokter, isi nama dokter PERTAMA yang disebut ke "surgeon".
   - "anesthesiologist": isi nama dokter KEDUA (jika ada) ke field ini.
   - "date", "time": HANYA isi jika benar-benar tertulis, format tanggal "YYYY-MM-DD" dan jam "HH:MM". JANGAN menebak tanggal/jam hari ini — kosongkan saja jika tidak tertulis.
   - "opType": isi "elektif", "semi", atau "cyto" HANYA jika jelas disebutkan, selain itu kosongkan.
   - Field lain ("rm", "room", "allergy", "bloodType") kosongkan jika tidak ada informasinya di teks.
   - "confidence": "high" jika teks jelas & tidak ambigu, "low" jika ambigu/tidak lengkap.

ATURAN OUTPUT:
Balas HANYA dengan array JSON, dimulai dengan "[" dan diakhiri dengan "]". Tidak ada teks lain sebelum atau sesudahnya — tidak ada judul, penjelasan, atau kesimpulan.

TEKS:
`;

// ── Sanitasi & validasi server-side (defense-in-depth, konsisten dengan
//    toISODateStrict dkk di client) ──────────────────────────────────────

function strField(v: unknown): string {
  return typeof v === "string" ? v.trim() : "";
}

function validDate(v: unknown): string {
  const s = strField(v);
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return "";
  const [, y, mo, da] = m;
  const d = new Date(`${y}-${mo}-${da}T00:00:00Z`);
  if (isNaN(d.getTime())) return "";
  if (
    d.getUTCFullYear() !== Number(y) ||
    d.getUTCMonth() + 1 !== Number(mo) ||
    d.getUTCDate() !== Number(da)
  ) {
    return "";
  }
  return s;
}

function validTime(v: unknown): string {
  const s = strField(v);
  return /^([01]\d|2[0-3]):([0-5]\d)$/.test(s) ? s : "";
}

function validAge(v: unknown, max: number): string {
  const s = strField(v);
  if (!s) return "";
  const n = Number(s);
  if (isNaN(n) || n < 0 || n > max) return "";
  return String(Math.round(n));
}

interface OcrResult {
  patient: string;
  age: string;
  ageMonths: string;
  rm: string;
  diagnosis: string;
  procedure: string;
  opType: "" | "elektif" | "semi" | "cyto";
  date: string;
  time: string;
  surgeon: string;
  anesthesiologist: string;
  room: string;
  allergy: string;
  bloodType: string;
  confidence: "high" | "low";
}

function sanitizeResult(r: unknown): OcrResult {
  const o = (r ?? {}) as Record<string, unknown>;
  const opType = o.opType;
  return {
    patient: strField(o.patient),
    age: validAge(o.age, 150),
    ageMonths: validAge(o.ageMonths, 11),
    rm: strField(o.rm),
    diagnosis: strField(o.diagnosis),
    procedure: strField(o.procedure),
    opType:
      opType === "elektif" || opType === "semi" || opType === "cyto"
        ? opType
        : "",
    date: validDate(o.date),
    time: validTime(o.time),
    surgeon: strField(o.surgeon),
    anesthesiologist: strField(o.anesthesiologist),
    room: strField(o.room),
    allergy: strField(o.allergy),
    bloodType: strField(o.bloodType),
    confidence: o.confidence === "high" ? "high" : "low",
  };
}

function extractJsonArrayFromText(text: string): unknown | null {
  try {
    return JSON.parse(text);
  } catch {
    /* lanjut ke fallback ekstraksi */
  }

  const start = text.indexOf("[");
  if (start === -1) return null;

  let depth = 0;
  let candidate = "";
  for (let i = start; i < text.length; i++) {
    if (text[i] === "[") depth++;
    else if (text[i] === "]") {
      depth--;
      if (depth === 0) {
        candidate = text.slice(start, i + 1);
        break;
      }
    }
  }
  if (!candidate) return null;

  try {
    return JSON.parse(candidate);
  } catch {
    /* fallback: model kadang salah pakai [ ] alih-alih { } untuk tiap
       objek pasien, menghasilkan [["key":"val",...],["key":"val",...]]
       alih-alih [{"key":"val",...},{"key":"val",...}]. Perbaiki dengan
       mengganti "[" dan "]" yang membungkus pasangan key:value menjadi
       "{" dan "}", tapi biarkan "[" dan "]" terluar (pembungkus array). */
    const inner = candidate.slice(1, -1).trim();
    const fixed =
      "[" +
      inner
        .replace(/\[(\s*"[^"]+"\s*:)/g, "{$1")
        .replace(/(:\s*(?:"[^"]*"|true|false|null|-?\d+(?:\.\d+)?))\s*\]/g, "$1}") +
      "]";
    try {
      return JSON.parse(fixed);
    } catch {
      return null;
    }
  }
}

// ── Jalur TEKS: Groq (satu-satunya parser, demi kecepatan) ─────────────

/** Panggil Groq (OpenAI-compatible chat completions) — SATU-SATUNYA
 *  parser untuk jalur paste-teks. Melempar Error kalau gagal — pemanggil
 *  (handleOcrScanText) memperlakukan ini sebagai FATAL, karena tidak ada
 *  model lain sebagai fallback (trade-off kecepatan vs redundansi, lihat
 *  komentar di TEXT_MODEL_GROQ). */
async function callGroqTextModel(env: Env, text: string): Promise<unknown[]> {
  if (!env.GROQ_API_KEY) throw new Error("GROQ_API_KEY belum diset di secret Worker");

  const res = await fetch(GROQ_API_URL, {
    method: "POST",
    headers: {
      "content-type": "application/json",
      authorization: `Bearer ${env.GROQ_API_KEY}`,
    },
    body: JSON.stringify({
      model: TEXT_MODEL_GROQ,
      messages: [
        { role: "user", content: TEXT_EXTRACTION_PROMPT + text },
      ],
      max_tokens: 4096,
      temperature: 0,
    }),
  });

  if (!res.ok) {
    const errText = await res.text().catch(() => "");
    throw new Error(`Groq API error (${res.status}): ${errText.slice(0, 300)}`);
  }

  const json = (await res.json()) as any;
  const responseText: string | undefined = json?.choices?.[0]?.message?.content;
  if (typeof responseText !== "string") {
    throw new Error("Format respons Groq tidak dikenali");
  }
  const extracted = extractJsonArrayFromText(
    responseText.replace(/```json\s*|```/g, "").trim()
  );
  if (extracted === null) throw new Error("Gagal parse hasil Groq");
  return Array.isArray(extracted) ? extracted : [extracted];
}

/** Handler untuk payload `{ text: string }` — jalur paste-teks di Tab
 *  Daftar. Groq = satu-satunya parser (demi kecepatan). Kalau Groq gagal
 *  (key belum diset, network error, format respons tidak dikenali, dsb),
 *  request GAGAL — tidak ada fallback ke model lain. */
async function handleOcrScanText(text: string, env: Env): Promise<Response> {
  let groqRaw: unknown[];
  try {
    groqRaw = await callGroqTextModel(env, text);
  } catch (e) {
    return Response.json(
      {
        ok: false,
        msg: "Gagal memproses teks dengan Groq: " + (e instanceof Error ? e.message : "unknown error"),
      },
      { status: 502 }
    );
  }

  const results = groqRaw.map(sanitizeResult);

  return Response.json({
    ok: true,
    data: results,
    msg: `✓ ${results.length} data operasi terbaca`,
  });
}



async function handleOcrScan(request: Request, env: Env): Promise<Response> {
  try {
    const body = (await request.json().catch(() => null)) as
      | { image_base64?: string; media_type?: string; text?: string }
      | null;

    // Jalur paste-teks (Tab Daftar) → Groq saja (satu model, demi kecepatan).
    // Dicek duluan sebelum validasi image_base64 supaya tidak mewajibkan
    // gambar saat payload memang berupa teks.
    if (typeof body?.text === "string" && body.text.trim().length > 0) {
      return handleOcrScanText(body.text.trim(), env);
    }

    if (!body?.image_base64) {
      return Response.json(
        { ok: false, msg: "image_base64 atau text wajib diisi" },
        { status: 400 }
      );
    }

    const approxBytes = (body.image_base64.length * 3) / 4;
    if (approxBytes > 6 * 1024 * 1024) {
      return Response.json(
        { ok: false, msg: "Ukuran gambar terlalu besar (maks ~6MB)" },
        { status: 413 }
      );
    }

    const mediaType = body.media_type || "image/jpeg";
    const dataUri = `data:${mediaType};base64,${body.image_base64}`;

    let aiResponse: unknown;
    try {
      aiResponse = await env.AI.run(OCR_MODEL, {
        messages: [
          {
            role: "user",
            content: [
              { type: "text", text: EXTRACTION_PROMPT },
              { type: "image_url", image_url: { url: dataUri } },
            ],
          },
        ],
        max_tokens: 4096,
      } as any);
    } catch (aiErr) {
      return Response.json(
        {
          ok: false,
          msg:
            "Gagal memanggil model AI: " +
            (aiErr instanceof Error ? aiErr.message : "unknown error"),
        },
        { status: 502 }
      );
    }

    let parsed: unknown;

    if (typeof aiResponse === "string") {
      const extracted = extractJsonArrayFromText(
        aiResponse.replace(/```json\s*|```/g, "").trim()
      );
      if (extracted === null) {
        return Response.json(
          {
            ok: false,
            msg: "Gagal parse hasil ekstraksi — model tidak mengembalikan JSON valid",
            raw: aiResponse,
          },
          { status: 502 }
        );
      }
      parsed = extracted;
    } else if (
      aiResponse &&
      typeof (aiResponse as { response?: unknown }).response === "string"
    ) {
      const responseText = (aiResponse as { response: string }).response;
      const extracted = extractJsonArrayFromText(
        responseText.replace(/```json\s*|```/g, "").trim()
      );
      if (extracted === null) {
        return Response.json(
          {
            ok: false,
            msg: "Gagal parse hasil ekstraksi — model tidak mengembalikan JSON valid",
            raw: responseText,
          },
          { status: 502 }
        );
      }
      parsed = extracted;
    } else if (
      aiResponse &&
      Array.isArray((aiResponse as { response?: unknown }).response)
    ) {
      parsed = (aiResponse as { response: unknown }).response;
    } else if (Array.isArray(aiResponse)) {
      parsed = aiResponse;
    } else {
      return Response.json(
        {
          ok: false,
          msg: "Format respons model tidak dikenali",
          raw_response_type: typeof aiResponse,
          raw_response: JSON.stringify(aiResponse).slice(0, 2000),
        },
        { status: 502 }
      );
    }

    const results = Array.isArray(parsed) ? parsed : [parsed];
    const safeResults = results.map(sanitizeResult);

    return Response.json({
      ok: true,
      data: safeResults,
      msg: `\u2713 ${safeResults.length} data operasi terbaca`,
    });
  } catch (err) {
    return Response.json(
      {
        ok: false,
        msg:
          "Gagal memproses OCR: " +
          (err instanceof Error ? err.message : "unknown error"),
      },
      { status: 500 }
    );
  }
}

// ── Entry point Worker ───────────────────────────────────────────────

export default {
  async fetch(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
    const url = new URL(request.url);

    if (url.pathname === "/api/ocr-scan") {
      if (request.method === "POST") {
        return handleOcrScan(request, env);
      }
      return Response.json(
        { ok: false, msg: "Gunakan method POST untuk endpoint ini" },
        { status: 405 }
      );
    }

    // Semua request lain diserve sebagai static asset (SPA React kamu).
    return env.ASSETS.fetch(request);
  },
};
