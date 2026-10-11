// src/worker.ts
// Worker entry untuk project "sisteminformasi-kamar-bedah".
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
//
// OPSIONAL (untuk rate limiting — lihat checkRateLimit()):
//   [[kv_namespaces]]
//   binding = "RATE_LIMIT_KV"
//   id = "<id dari: npx wrangler kv namespace create RATE_LIMIT_KV>"
// Kalau tidak di-setup, rate limiting otomatis di-skip (fail-open).

export interface Env {
  AI: Ai;
  ASSETS: Fetcher;
  // Secret — diisi lewat: npx wrangler secret put GROQ_API_KEY
  GROQ_API_KEY?: string;
  // KV Namespace untuk rate limiting — diisi lewat wrangler.toml:
  //   [[kv_namespaces]]
  //   binding = "RATE_LIMIT_KV"
  //   id = "<id dari: npx wrangler kv namespace create RATE_LIMIT_KV>"
  // Kalau binding ini tidak ada (belum disetup), rate limiting otomatis
  // di-skip (fail-open) — lihat checkRateLimit().
  RATE_LIMIT_KV?: KVNamespace;
}

// UPDATE (Juli 2026): diganti dari "@cf/moondream/moondream3.1-9B-A2B"
// ke Llama 4 Scout — model vision multimodal Meta (17B aktif / 109B total,
// MoE 16 experts, context window 131K token) yang tersedia native lewat
// binding env.AI.run() di Workers AI.
//
// CATATAN JUJUR (beda karakteristik dari Moondream):
// - Moondream 3.1 SECARA EKSPLISIT dioptimalkan Cloudflare untuk
//   "object detection, pointing, OCR, and structured output".
// - Llama 4 Scout adalah vision model GENERAL-PURPOSE (image reasoning,
//   captioning, visual QA) — bukan spesialis ekstraksi data terstruktur.
//   Akurasi untuk formulir/tulisan tangan medis Bahasa Indonesia BELUM
//   tentu lebih baik — WAJIB dites dulu dengan contoh nyata sebelum
//   dianggap upgrade.
// - Pricing berbasis token ($0.27/M input, $0.85/M output), bukan
//   flat/neuron seperti model vision kecil — gambar resolusi tinggi bisa
//   memakan token input cukup banyak, cek biaya aktualnya di dashboard.
const OCR_MODEL = "@cf/meta/llama-4-scout-17b-16e-instruct";

// Model teks Groq — parser UTAMA untuk jalur paste-teks (Tab Daftar),
// dipilih karena hardware inferensi khusus (LPU) yang jauh lebih cepat
// dan model 120B yang lebih akurat untuk ekstraksi data terstruktur.
// "gpt-oss-120b" adalah pengganti resmi Groq untuk llama-3.3-70b-versatile
// (yang akan di-shutdown 16 Agustus 2026). Kalau nanti ingin ganti model
// lain, cek daftar model terkini di https://console.groq.com/docs/models.
const TEXT_MODEL_GROQ = "openai/gpt-oss-120b";
const GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions";

// Model teks Cloudflare Workers AI — FALLBACK OTOMATIS kalau Groq gagal
// (down, GROQ_API_KEY belum diset, timeout, dsb). Ini BUKAN cross-check
// (tidak dipanggil dua-duanya setiap saat demi kecepatan) — CF hanya
// dipanggil KALAU Groq gagal, supaya jalur paste-teks tidak buntu total
// hanya karena satu provider bermasalah.
const TEXT_MODEL_CF = "@cf/zai-org/glm-4.7-flash";

const EXTRACTION_PROMPT = `Baca gambar ini. Gambar berisi data permintaan operasi pasien kamar bedah rumah sakit — bisa berupa formulir, tulisan tangan, atau daftar dari chat WhatsApp.

LANGKAH KERJA:
1. Hitung dulu berapa jumlah pasien yang disebutkan (biasanya ditandai nomor urut 1, 2, 3, dst, atau nama "Ny./Tn./An." yang berbeda-beda).
2. Untuk SETIAP pasien yang kamu temukan, buat SATU objek JSON terpisah. Jumlah objek dalam array HARUS SAMA DENGAN jumlah pasien yang kamu hitung di langkah 1.
3. Untuk tiap pasien, isi field berikut apa adanya sesuai teks yang tertulis untuk pasien itu:
   - "patient": nama pasien
   - "age": HANYA angka umur, buang kata "th"/"tahun" dari teksnya. Contoh: teks "63th rencana ORIF" → age="63", procedure="ORIF pinning distal radius Sinistra" (umur TIDAK ikut ke procedure).
   - "procedure": nama tindakan/rencana operasi untuk pasien itu, TANPA angka umur di dalamnya.
   - "diagnosis": kosongkan jika tidak ada info diagnosis terpisah dari prosedur.
   - "surgeon": dokter yang melakukan tindakan/operasi. Cari kata kunci peran DULU di dekat nama dokter — contoh: "operator", "bedah", "dr. bedah", "(bedah)", "op:", atau nama yang muncul bersama nama prosedur/tindakan. Kalau ADA kata kunci peran yang jelas, pakai itu untuk menentukan siapa "surgeon" — JANGAN asal pakai urutan sebutan.
   - "anesthesiologist": dokter anestesi/pembius. Cari kata kunci "anestesi", "anastesi", "dr. anestesi", "(anestesi)", "spAn" di dekat nama dokter.
   - HANYA kalau TIDAK ADA kata kunci peran sama sekali (cuma disebutkan "dr. X dan dr. Y" tanpa keterangan apa-apa), BARU gunakan urutan: dokter PERTAMA yang disebut → "surgeon", dokter KEDUA → "anesthesiologist".
   - Jika hanya ada satu nama dokter disebutkan (tanpa kata kunci peran), isi ke "surgeon" saja dan kosongkan "anesthesiologist".
   - "date", "time": HANYA isi jika benar-benar tertulis di gambar, format tanggal "YYYY-MM-DD" dan jam "HH:MM". JANGAN mengisi tanggal/jam hari ini atau tebakan apa pun jika tidak tertulis — kosongkan saja.
   - "opType": isi "elektif", "semi", atau "cyto" HANYA jika jelas disebutkan, selain itu kosongkan
   - Field lain ("rm", "room", "allergy", "bloodType") kosongkan jika tidak ada informasinya di gambar
   - "confidence": "high" jika teks jelas terbaca, "low" jika kabur/tulisan tangan sulit dibaca
   - "lowConfidenceFields": array berisi NAMA FIELD (persis seperti nama key JSON di atas, mis. "date", "surgeon") yang kamu SENDIRI tidak yakin/ragu-ragu saat membaca — supaya perawat tahu bagian mana yang perlu dicek ulang manual. Kosongkan array ini ([]) kalau semua field yang terisi kamu yakin benar.

CONTOH: jika judul gambar bertuliskan "Jam 12.00 dr. Adhinanda dan dr. Timor" dan salah satu baris berisi "Ny. Temu Rahayu, 63th rencana ORIF pinning distal radius Sinistra", maka objek untuk pasien itu adalah:
{"patient":"Ny. Temu Rahayu","age":"63","ageMonths":"","rm":"","diagnosis":"","procedure":"ORIF pinning distal radius Sinistra","opType":"","date":"","time":"12:00","surgeon":"dr. Adhinanda","anesthesiologist":"dr. Timor","room":"","allergy":"","bloodType":"","confidence":"high","lowConfidenceFields":[]}

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
   - "surgeon": dokter yang melakukan tindakan/operasi. Cari kata kunci peran DULU di dekat nama dokter — contoh: "operator", "bedah", "dr. bedah", "(bedah)", "op:", atau nama yang muncul bersama nama prosedur/tindakan. Kalau ADA kata kunci peran yang jelas, pakai itu untuk menentukan siapa "surgeon" — JANGAN asal pakai urutan sebutan.
   - "anesthesiologist": dokter anestesi/pembius. Cari kata kunci "anestesi", "anastesi", "dr. anestesi", "(anestesi)", "spAn" di dekat nama dokter.
   - HANYA kalau TIDAK ADA kata kunci peran sama sekali (cuma disebutkan "dr. X dan dr. Y" tanpa keterangan apa-apa), BARU gunakan urutan: dokter PERTAMA yang disebut → "surgeon", dokter KEDUA → "anesthesiologist".
   - "date", "time": HANYA isi jika benar-benar tertulis, format tanggal "YYYY-MM-DD" dan jam "HH:MM". JANGAN menebak tanggal/jam hari ini — kosongkan saja jika tidak tertulis.
   - "opType": isi "elektif", "semi", atau "cyto" HANYA jika jelas disebutkan, selain itu kosongkan.
   - Field lain ("rm", "room", "allergy", "bloodType") kosongkan jika tidak ada informasinya di teks.
   - "confidence": "high" jika teks jelas & tidak ambigu, "low" jika ambigu/tidak lengkap.
   - "lowConfidenceFields": array berisi NAMA FIELD (persis seperti nama key JSON di atas) yang kamu SENDIRI tidak yakin/ragu-ragu — kosongkan ([]) kalau semua field yang terisi kamu yakin benar.

ATURAN OUTPUT:
Balas HANYA dengan array JSON, dimulai dengan "[" dan diakhiri dengan "]". Tidak ada teks lain sebelum atau sesudahnya — tidak ada judul, penjelasan, atau kesimpulan.
TEKS:`;

// ── Sanitasi & validasi server-side (defense-in-depth, konsisten dengan
//    toISODateStrict dkk di client) ──────────────────────────────────

function strField(v: unknown): string {
  return typeof v === "string" ? v.trim() : "";
}

/* FIX BUG (penyebab: "kadang tanggal/jam/umur tidak terisi" — dilaporkan
   Andri): SEBELUMNYA validDate/validTime/validAge sangat KETAT — kalau
   model AI mengembalikan data yang BENAR tapi formatnya sedikit beda dari
   yang diharapkan PERSIS (mis. "11/07/2026" bukan "2026-07-11", atau
   "12.00" bukan "12:00" — yang mana "12.00" itu JUSTRU format yang dipakai
   di contoh prompt kita sendiri!), field itu DIKOSONGKAN begitu saja,
   padahal datanya sebenarnya ada & benar, cuma beda notasi.
   Sekarang: coba NORMALISASI dulu ke format standar (dengan aturan
   eksplisit, bukan tebakan), baru kosongkan kalau BENAR-BENAR tidak bisa
   dipahami sebagai tanggal/jam/umur yang valid. */
function validDate(v: unknown): string {
  let s = strField(v);
  if (!s) return "";
  // Format standar: YYYY-MM-DD
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) {
    const [, y, mo, da] = m;
    return normalizeDateParts(y, mo, da);
  }
  // Format umum Indonesia: DD/MM/YYYY atau DD-MM-YYYY
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) {
    const [, da, mo, y] = m;
    return normalizeDateParts(y, mo, da);
  }
  // Format YYYY/MM/DD (garis miring, urutan ISO)
  m = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if (m) {
    const [, y, mo, da] = m;
    return normalizeDateParts(y, mo, da);
  }
  return "";
}

function normalizeDateParts(y: string, mo: string, da: string): string {
  const yn = Number(y), mon = Number(mo), dan = Number(da);
  const yPad = String(yn).padStart(4, "0");
  const moPad = String(mon).padStart(2, "0");
  const daPad = String(dan).padStart(2, "0");
  const d = new Date(`${yPad}-${moPad}-${daPad}T00:00:00Z`);
  if (isNaN(d.getTime())) return "";
  if (
    d.getUTCFullYear() !== yn ||
    d.getUTCMonth() + 1 !== mon ||
    d.getUTCDate() !== dan
  ) {
    return ""; // tanggal tidak masuk akal (mis. 31 Februari) — tetap ditolak
  }
  return `${yPad}-${moPad}-${daPad}`;
}

function validTime(v: unknown): string {
  let s = strField(v);
  if (!s) return "";
  // Buang embel-embel umum: " WIB", " WITA", " WIT", spasi berlebih, detik.
  s = s.replace(/\s*(WIB|WITA|WIT)\s*$/i, "").trim();
  // Terima titik ATAU titik-dua sebagai pemisah jam:menit (mis. "12.00").
  const m = s.match(/^([01]?\d|2[0-3])[.:]([0-5]\d)(?:[.:]\d{2})?$/);
  if (!m) return "";
  const hh = m[1].padStart(2, "0");
  const mm = m[2];
  return `${hh}:${mm}`;
}

function validAge(v: unknown, max: number): string {
  const raw = strField(v);
  if (!raw) return "";
  // Coba angka langsung dulu; kalau gagal, ambil digit pertama dari teks
  // (menangani kasus model belum sepenuhnya membuang "th"/"tahun"/spasi
  // sesuai instruksi prompt — mis. "63 th" → tetap ambil "63").
  let n = Number(raw);
  if (isNaN(n)) {
    const digitMatch = raw.match(/\d+/);
    if (!digitMatch) return "";
    n = Number(digitMatch[0]);
  }
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
  /* PENAMBAHAN #5 (confidence per-field): daftar nama field yang model
     SENDIRI tandai sebagai tidak yakin — supaya perawat tahu PERSIS
     bagian mana yang perlu dicek ulang, bukan curiga ke SEMUA field
     hanya karena confidence keseluruhan "low". Contoh isi:
     ["date","surgeon"] — berarti tanggal & nama dokter operator perlu
     diperiksa lebih teliti, field lain relatif aman. */
  lowConfidenceFields: string[];
}

const VALID_FIELD_NAMES = [
  "patient", "age", "ageMonths", "rm", "diagnosis", "procedure", "opType",
  "date", "time", "surgeon", "anesthesiologist", "room", "allergy", "bloodType",
];

function sanitizeResult(r: unknown): OcrResult {
  const o = (r ?? {}) as Record<string, unknown>;
  const opType = o.opType;
  const rawLowConf = Array.isArray(o.lowConfidenceFields) ? o.lowConfidenceFields : [];
  const lowConfidenceFields = rawLowConf
    .filter((f): f is string => typeof f === "string" && VALID_FIELD_NAMES.includes(f));
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
    lowConfidenceFields,
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

// ── Jalur TEKS: Groq (utama) dengan fallback otomatis ke CF ────────────

/** Panggil Groq (OpenAI-compatible chat completions) — parser UTAMA
 *  untuk jalur paste-teks. Melempar Error kalau gagal — pemanggil
 *  (handleOcrScanText) akan otomatis coba CF sebagai fallback. */
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

/** Panggil model teks Cloudflare Workers AI — FALLBACK OTOMATIS, hanya
 *  dipanggil kalau callGroqTextModel gagal (lihat handleOcrScanText).
 *  Ini BUKAN cross-check: CF tidak dipanggil setiap saat, cuma sebagai
 *  cadangan supaya jalur paste-teks tidak buntu total kalau Groq
 *  bermasalah (down, key salah, rate limit Groq, dsb). */
async function callCfTextModel(env: Env, text: string): Promise<unknown[]> {
  const raw = await env.AI.run(TEXT_MODEL_CF as any, {
    messages: [
      { role: "user", content: TEXT_EXTRACTION_PROMPT + text },
    ],
    max_tokens: 4096,
  } as any);
  const responseText =
    typeof raw === "string"
      ? raw
      : (raw && typeof (raw as any).response === "string")
      ? (raw as any).response
      : null;
  if (responseText === null) {
    throw new Error("Format respons CF text model tidak dikenali");
  }
  const extracted = extractJsonArrayFromText(
    responseText.replace(/```json\s*|```/g, "").trim()
  );
  if (extracted === null) throw new Error("Gagal parse hasil CF text model");
  return Array.isArray(extracted) ? extracted : [extracted];
}

/** Handler untuk payload `{ text: string }` — jalur paste-teks di Tab
 *  Daftar. Coba Groq dulu (utama, tercepat). Kalau Groq gagal (network,
 *  key belum diset, timeout, dsb), OTOMATIS fallback ke CF text model
 *  sebagai cadangan — supaya jalur ini tidak buntu total hanya karena
 *  satu provider bermasalah. Kalau KEDUANYA gagal, baru request gagal. */
async function handleOcrScanText(text: string, env: Env): Promise<Response> {
  let results: unknown[];
  let usedFallback = false;
  let groqError: string | null = null;
  try {
    results = await callGroqTextModel(env, text);
  } catch (e) {
    groqError = e instanceof Error ? e.message : "unknown error";
    try {
      results = await callCfTextModel(env, text);
      usedFallback = true;
    } catch (e2) {
      return Response.json(
        {
          ok: false,
          msg: `Gagal memproses teks — Groq: ${groqError}; CF (fallback): ${
            e2 instanceof Error ? e2.message : "unknown error"
          }`,
        },
        { status: 502 }
      );
    }
  }

  let sanitized = results.map(sanitizeResult);

  /* PENAMBAHAN #3 (retry otomatis kalau confidence rendah): mengulang
     dengan MODEL & PARAMETER YANG SAMA tidak akan mengubah apa pun karena
     temperature:0 (deterministik) — hasilnya pasti identik, buang-buang
     waktu & biaya. Yang benar-benar berguna: minta MODEL LAIN (CF, kalau
     tadi pakai Groq) sebagai "opini kedua" HANYA kalau ada field yang
     ditandai ragu-ragu (lowConfidenceFields) — lalu untuk field itu SAJA,
     kalau model kedua ternyata yakin (confidence "high" & field tsb TIDAK
     ada di lowConfidenceFields versi dia), pakai nilai dari model kedua.
     Field yang tidak bermasalah dari hasil pertama TETAP dipakai apa
     adanya — ini penggantian di level field, bukan ganti seluruh objek. */
  const anyLowConfidence = sanitized.some(
    (r) => r.confidence === "low" || r.lowConfidenceFields.length > 0
  );
  if (anyLowConfidence && !usedFallback) {
    try {
      /* FIX BUG (konsisten dengan proteksi di jalur OCR foto): batasi
         waktu opini kedua supaya tidak berisiko menambah total waktu
         proses mendekati timeout client. Model teks (CF) umumnya jauh
         lebih cepat dari model vision, tapi tetap diberi batas wajar
         untuk jaga-jaga (network lambat dsb). */
      const secondOpinionRaw = await Promise.race([
        callCfTextModel(env, text),
        new Promise<never>((_, reject) =>
          setTimeout(() => reject(new Error("retry timeout budget habis")), 15000)
        ),
      ]);
      const secondOpinion = secondOpinionRaw.map(sanitizeResult);
      // Cocokkan berdasarkan index — hanya aman kalau jumlah pasien SAMA
      // antara kedua percobaan (kalau beda, jangan dicampur, terlalu
      // berisiko salah pasang seperti kasus cross-check sebelumnya).
      if (secondOpinion.length === sanitized.length) {
        sanitized = sanitized.map((primary, i) => {
          if (primary.confidence !== "low" && primary.lowConfidenceFields.length === 0) {
            return primary; // sudah yakin, tidak perlu opini kedua
          }
          const second = secondOpinion[i];
          const merged = { ...primary };
          const stillLowConf: string[] = [];
          for (const field of primary.lowConfidenceFields) {
            const secondIsConfidentForField = !second.lowConfidenceFields.includes(field);
            const secondValue = (second as any)[field];
            if (secondIsConfidentForField && secondValue) {
              (merged as any)[field] = secondValue; // pakai opini kedua untuk field ini saja
            } else {
              stillLowConf.push(field); // tetap ragu setelah opini kedua — biarkan perawat yang cek
            }
          }
          merged.lowConfidenceFields = stillLowConf;
          merged.confidence = stillLowConf.length === 0 ? "high" : "low";
          return merged;
        });
      }
    } catch {
      // Opini kedua gagal dipanggil (network dsb) — TIDAK FATAL, tetap
      // pakai hasil pertama apa adanya (lebih baik ada hasil dengan
      // catatan ragu-ragu, daripada gagal total).
    }
  }

  return Response.json({
    ok: true,
    data: sanitized,
    msg: usedFallback
      ? `✓ ${sanitized.length} data operasi terbaca (Groq gagal, pakai model cadangan — periksa lebih teliti)`
      : `✓ ${sanitized.length} data operasi terbaca`,
  });
}

/** Rate limiting sederhana berbasis Cloudflare KV: batasi jumlah panggilan
 *  /api/ocr-scan per-IP dalam jendela waktu tertentu, supaya endpoint ini
 *  tidak bisa di-spam habis-habisan (menghabiskan kuota Neuron & Groq API
 *  kalau di-exploit atau ada bug retry yang looping tak terkendali).
 *
 *  FAIL-OPEN BY DESIGN: kalau RATE_LIMIT_KV belum di-setup (binding tidak
 *  ada) atau KV sedang error, rate limit di-SKIP (request tetap diproses)
 *  — supaya kesalahan konfigurasi KV tidak sampai membuat SELURUH fitur
 *  OCR mati. Ini trade-off sadar: keamanan vs keandalan, kita pilih
 *  keandalan di sini karena kegagalan KV jauh lebih jarang daripada
 *  manfaat mencegah baru bisa didapat kalau di-setup dengan benar. */
async function checkRateLimit(
  env: Env,
  clientId: string,
  limit = 20,
  windowSeconds = 60
): Promise<{ allowed: boolean; remaining: number }> {
  if (!env.RATE_LIMIT_KV) return { allowed: true, remaining: limit };
  const key = `ratelimit:ocrscan:${clientId}`;
  try {
    const raw = await env.RATE_LIMIT_KV.get(key);
    const count = raw ? parseInt(raw, 10) || 0 : 0;
    if (count >= limit) {
      return { allowed: false, remaining: 0 };
    }
    // expirationTtl otomatis membersihkan key setelah jendela waktu habis
    // — tidak perlu job pembersihan terpisah.
    await env.RATE_LIMIT_KV.put(key, String(count + 1), { expirationTtl: windowSeconds });
    return { allowed: true, remaining: limit - count - 1 };
  } catch {
    // KV error → fail-open, lihat catatan di atas.
    return { allowed: true, remaining: limit };
  }
}

/** Panggil Llama 4 Scout SEKALI dengan parameter temperature yang
 *  diberikan, urai responsnya, dan kembalikan hasil yang sudah
 *  disanitasi. Dipisah jadi fungsi sendiri (bukan inline di
 *  handleOcrScan) supaya bisa dipanggil ULANG dengan temperature
 *  berbeda untuk retry — lihat PENAMBAHAN #3 di handleOcrScan.
 *  Tidak pernah throw — selalu mengembalikan objek {ok, ...} supaya
 *  pemanggil tidak perlu try/catch bertingkat.
 *
 *  PERBEDAAN PENTING vs Moondream (callMoondreamOnce lama):
 *  - Input pakai `messages` dengan `content` array gaya OpenAI:
 *    { type: "text", text } + { type: "image_url", image_url: { url } }.
 *    BUKAN { task, image, question } seperti Moondream.
 *  - Output ada di field `response` (string), BUKAN `answer`.
 *  - Tidak ada parameter `reasoning` seperti Moondream. Sebagai
 *    pengganti "opini kedua yang benar-benar beda" (bukan retry
 *    identik yang percuma), retry dipanggil dengan `temperature`
 *    lebih tinggi (0 → 0.4) supaya sampling-nya berbeda dari
 *    percobaan pertama. */
async function callLlamaScoutOnce(
  env: Env,
  dataUri: string,
  temperature: number
): Promise<{ ok: true; results: OcrResult[] } | { ok: false; msg: string; raw?: string }> {
  let aiResponse: unknown;
  try {
    aiResponse = await env.AI.run(OCR_MODEL as any, {
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
      temperature,
    } as any);
  } catch (aiErr) {
    return {
      ok: false,
      msg: "Gagal memanggil model AI: " + (aiErr instanceof Error ? aiErr.message : "unknown error"),
    };
  }

  let parsed: unknown;
  /* Output Llama 4 Scout ada di field `response` (mengikuti skema teks-
     generation Workers AI standar) — beda dari Moondream yang pakai
     `answer`. Fallback ke string polos / array tetap dipertahankan
     sebagai jaga-jaga kalau perilaku binding berubah. */
  if (aiResponse && typeof (aiResponse as { response?: unknown }).response === "string") {
    const responseText = (aiResponse as { response: string }).response;
    const extracted = extractJsonArrayFromText(
      responseText.replace(/```json\s*|```/g, "").trim()
    );
    if (extracted === null) {
      return { ok: false, msg: "Gagal parse hasil ekstraksi — model tidak mengembalikan JSON valid", raw: responseText };
    }
    parsed = extracted;
  } else if (typeof aiResponse === "string") {
    const extracted = extractJsonArrayFromText(
      aiResponse.replace(/```json\s*|```/g, "").trim()
    );
    if (extracted === null) {
      return { ok: false, msg: "Gagal parse hasil ekstraksi — model tidak mengembalikan JSON valid", raw: aiResponse };
    }
    parsed = extracted;
  } else if (aiResponse && Array.isArray((aiResponse as { response?: unknown }).response)) {
    parsed = (aiResponse as { response: unknown }).response;
  } else if (Array.isArray(aiResponse)) {
    parsed = aiResponse;
  } else {
    return {
      ok: false,
      msg: "Format respons model tidak dikenali",
      raw: JSON.stringify(aiResponse).slice(0, 2000),
    };
  }

  const results = Array.isArray(parsed) ? parsed : [parsed];
  return { ok: true, results: results.map(sanitizeResult) };
}

async function handleOcrScan(request: Request, env: Env): Promise<Response> {
  try {
    // Identifikasi klien lewat header Cloudflare (selalu ada di Workers,
    // tidak bisa dipalsukan dari luar karena di-inject Cloudflare edge).
    const clientId = request.headers.get("cf-connecting-ip") || "unknown";
    const rl = await checkRateLimit(env, clientId, 20, 60);
    if (!rl.allowed) {
      return Response.json(
        {
          ok: false,
          msg: "Terlalu banyak permintaan OCR dalam waktu singkat — tunggu sebentar lalu coba lagi (maks 20 scan per menit per perangkat).",
        },
        { status: 429 }
      );
    }

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

    let safeResults: OcrResult[];
    {
      const first = await callLlamaScoutOnce(env, dataUri, 0);
      if (!first.ok) {
        return Response.json({ ok: false, msg: first.msg, raw: (first as any).raw }, { status: 502 });
      }
      safeResults = first.results;

      /* PENAMBAHAN #3 (retry otomatis kalau confidence rendah, KHUSUS
         OCR foto): Llama 4 Scout dipanggil PERTAMA dengan temperature 0
         (paling deterministik/cepat). Kalau ada field yang di-flag
         ragu-ragu, retry SEKALI dengan temperature lebih tinggi (0.4)
         supaya sampling-nya benar-benar beda dari percobaan pertama —
         bukan retry identik yang percuma. Merge di level field yang sama
         seperti jalur paste-teks. */
      const anyLowConfidence = safeResults.some(
        (r) => r.confidence === "low" || r.lowConfidenceFields.length > 0
      );
      if (anyLowConfidence) {
        /* FIX BUG (risiko latency bertumpuk): batasi retry ke 20 detik —
           kalau lewat itu, batalkan retry & tetap pakai hasil PERTAMA apa
           adanya (dengan catatan ragu-ragu), supaya total waktu proses
           tidak mendekati/melebihi timeout client (lihat attemptWithRetry
           di App.tsx) yang bisa memicu client mengulang SELURUH request
           dari awal. */
        const RETRY_BUDGET_MS = 20000;
        const second = await Promise.race([
          callLlamaScoutOnce(env, dataUri, 0.4),
          new Promise<{ ok: false; msg: string }>((resolve) =>
            setTimeout(() => resolve({ ok: false, msg: "retry timeout budget habis" }), RETRY_BUDGET_MS)
          ),
        ]);
        if (second.ok && second.results.length === safeResults.length) {
          safeResults = safeResults.map((primary, i) => {
            if (primary.confidence !== "low" && primary.lowConfidenceFields.length === 0) return primary;
            const alt = second.results[i];
            const merged = { ...primary };
            const stillLowConf: string[] = [];
            for (const field of primary.lowConfidenceFields) {
              const altIsConfidentForField = !alt.lowConfidenceFields.includes(field);
              const altValue = (alt as any)[field];
              if (altIsConfidentForField && altValue) {
                (merged as any)[field] = altValue;
              } else {
                stillLowConf.push(field);
              }
            }
            merged.lowConfidenceFields = stillLowConf;
            merged.confidence = stillLowConf.length === 0 ? "high" : "low";
            return merged;
          });
        }
        // Kalau retry gagal (network dsb) atau jumlah pasien beda antara
        // percobaan 1 & 2 — TIDAK FATAL, tetap pakai hasil pertama apa
        // adanya (lebih baik ada hasil dengan catatan ragu-ragu).
      }
    }

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
}

// ── Entry point Worker ──────────────────────────────────────────────

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
