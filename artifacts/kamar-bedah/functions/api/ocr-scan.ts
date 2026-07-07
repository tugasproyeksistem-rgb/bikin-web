// functions/api/ocr-scan.ts
// Cloudflare Pages Function — otomatis ter-route ke /api/ocr-scan
// Env binding "AI" harus sudah ditambahkan di wrangler.toml:
//   [ai]
//   binding = "AI"
// Tidak perlu API key/token apa pun — binding AI tersedia otomatis saat deploy.

interface Env {
  AI: Ai;
}

const OCR_MODEL = "@cf/meta/llama-3.2-11b-vision-instruct";
// Model ini milik Meta — Cloudflare mewajibkan approve lisensi dulu lewat
// Cloudflare Dashboard → Workers AI → cari model ini → setujui, sebelum
// panggilan API berhasil.

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
    // Ganti setiap "[" yang diikuti oleh `"key":` jadi "{", dan setiap
    // "]" yang didahului oleh nilai (bukan diikuti koma ke elemen array
    // lain di level yang sama) jadi "}".
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

// ── Handler utama ─────────────────────────────────────────────────────

export const onRequestPost: PagesFunction<Env> = async (context) => {
  try {
    const body = await context.request
      .json<{ image_base64?: string; media_type?: string }>()
      .catch(() => null);

    if (!body?.image_base64) {
      return Response.json(
        { ok: false, msg: "image_base64 wajib diisi" },
        { status: 400 }
      );
    }

    // Batas ukuran wajar sebelum diproses (safety net kedua setelah
    // kompresi di sisi client). ~6MB base64 encoded.
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
      aiResponse = await context.env.AI.run(OCR_MODEL, {
        messages: [
          {
            role: "user",
            content: [
              { type: "text", text: EXTRACTION_PROMPT },
              { type: "image_url", image_url: { url: dataUri } },
            ],
          },
        ],
        max_tokens: 3000,
      });
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
      // Model langsung mengembalikan array/object JSON (bukan string) —
      // kasus ini terjadi ketika model taat instruksi tapi API
      // membungkusnya sebagai structured data, bukan teks mentah.
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
        msg: "Gagal memproses OCR: " + (err instanceof Error ? err.message : "unknown error"),
      },
      { status: 500 }
    );
  }
};

// Tolak method selain POST dengan jelas
export const onRequestGet: PagesFunction<Env> = async () => {
  return Response.json(
    { ok: false, msg: "Gunakan method POST untuk endpoint ini" },
    { status: 405 }
  );
};
