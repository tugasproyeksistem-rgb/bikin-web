/* src/utils/ocr-matching.ts
   Fungsi utilitas murni untuk memvalidasi & mencocokkan hasil ekstraksi
   OCR (foto/teks/Asisten Virtual) — koersi tipe, pembatasan panjang teks,
   normalisasi & fuzzy-match nama dokter ke master data staf, validasi
   field semi-terstruktur, dan aturan kehati-hatian (CYTO, gol.darah).
   Dipindah dari App.tsx (langkah 3 refactor — pisah utils, batch 4),
   TIDAK ADA perubahan logika, termasuk seluruh FIX AUDIT & FIX BUG. */

import type { OcrExtractResult } from "../types";
import { ROOMS, BT } from "../constants";
import { toISODateStrict } from "./format";

/* FIX AUDIT #34 (HIGH — Validasi Runtime pada Respons Eksternal):
   Sebelumnya `json.data as OcrExtractResult[]` langsung di-cast paksa tanpa
   validasi bentuk/tipe. Kalau model vision/Worker mengembalikan field
   dengan tipe tak terduga (mis. `room` berupa angka), pemanggilan method
   khusus-string (`.match`, `.toUpperCase`) di applyOcrResult/autoSaveOcrBatch
   akan throw dan bisa membuat komponen crash di tengah alur kerja perawat.
   coerceOcrResult() memaksa SEMUA field jadi string yang aman dipakai,
   apa pun tipe asli dari respons server. */
export function coerceOcrResult(raw: any): OcrExtractResult {
  const s = (v: any): string => typeof v === "string" ? v : (v==null ? "" : String(v));
  const ot = raw?.opType;
  return {
    patient: s(raw?.patient), age: s(raw?.age), ageMonths: s(raw?.ageMonths), rm: s(raw?.rm),
    diagnosis: s(raw?.diagnosis), procedure: s(raw?.procedure),
    opType: (ot==="elektif"||ot==="semi"||ot==="cyto") ? ot : "",
    date: s(raw?.date), time: s(raw?.time), surgeon: s(raw?.surgeon),
    anesthesiologist: s(raw?.anesthesiologist), room: s(raw?.room),
    allergy: s(raw?.allergy), bloodType: s(raw?.bloodType),
    confidence: raw?.confidence==="high" ? "high" : "low",
    lowConfidenceFields: Array.isArray(raw?.lowConfidenceFields) ? raw.lowConfidenceFields.filter((f:any)=>typeof f==="string") : [],
    notAFormError: raw?.notAFormError === true, // saran #6 — lihat catatan di interface
  };
}

/* FIX AUDIT #35 (MEDIUM — Tidak Ada Batas Panjang Teks Bebas): model vision
   yang berhalusinasi pada gambar buram/kompleks bisa mengembalikan teks
   sangat panjang. capOcrText() memotong ke panjang wajar untuk field
   nama/diagnosa/tindakan/dll sebelum disimpan ke form atau database. */
/* FIX BUG #41 (audit lanjutan): `.slice(0,max)` beroperasi per unit UTF-16,
   bukan per karakter — untuk emoji/karakter di luar BMP (dibentuk dari
   SEPASANG unit UTF-16/"surrogate pair"), potongan bisa jatuh di TENGAH
   pasangan, menghasilkan karakter rusak (mis. tampil sebagai "�"). Array.from
   memecah string per code point (aman terhadap surrogate pair), jadi
   pemotongan tidak akan pernah membelah satu karakter jadi dua. */
export function capOcrText(s: string, max = 200): string {
  if (s.length <= max) return s; // fast path: .length (unit UTF-16) sudah <= max, pasti aman dibiarkan utuh
  const chars = Array.from(s);
  return chars.length > max ? chars.slice(0, max).join("") + "…" : s;
}

/* PENAMBAHAN #1 (fuzzy-match dokter ke Master Data Staf): hasil OCR untuk
   nama dokter sering beda tipis dari yang terdaftar (typo kecil, gelar
   tidak lengkap, dsb — mis. "dr. Adhinanda" hasil OCR vs "dr. Adhinanda,
   Sp.B" yang terdaftar). Fungsi ini mencari nama staf TERDAFTAR yang
   paling mirip dengan hasil OCR, dan HANYA mengganti kalau similarity-nya
   sangat tinggi (>=0.72) — supaya tidak salah cocokkan ke dokter lain yang
   namanya kebetulan agak mirip. Kalau tidak ada yang cukup mirip, nilai
   OCR asli tetap dipakai apa adanya (fail-safe: tidak menebak). */
/* FIX BUG #45 (audit lanjutan): SEBELUMNYA regex `\bsp\.?\s*[a-z]{1,4}\b`
   cuma menangkap 1-4 huruf tepat setelah "sp" — untuk gelar subspesialis
   dengan akhiran tanda kurung/strip (mis. "Sp.B(K)Onk", "Sp.An-KIC"), sisa
   akhirannya ("(k)onk", "-kic") LOLOS tidak terhapus dan ikut masuk ke
   nama yang dinormalisasi, menurunkan akurasi fuzzy-match (skor similarity
   turun karena ada "sampah" karakter tambahan). Sekarang gelar spesialis
   dipotong dari kemunculan pertama "sp" SAMPAI AKHIR STRING — karena gelar
   selalu berada di BELAKANG nama, ini jauh lebih tahan terhadap variasi
   format gelar apa pun (termasuk yang belum pernah dilihat sebelumnya)
   dibanding mencoba menangkap tiap kemungkinan pola gelar satu-satu. */
export function normalizeDoctorName(s: string): string {
  let n = s
    .toLowerCase()
    .replace(/\b(dr|drg|prof)\b\.?/g, "")
    .trim();
  n = n.replace(/\bsp\.?.*$/, ""); // buang gelar spesialis & semua akhirannya, apa pun bentuknya
  return n
    .replace(/[.,()\-]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

export function levenshteinRatio(a: string, b: string): number {
  if (a === b) return 1;
  if (!a.length || !b.length) return 0;
  const dp: number[][] = Array.from({length: a.length+1}, () => new Array(b.length+1).fill(0));
  for (let i=0; i<=a.length; i++) dp[i][0] = i;
  for (let j=0; j<=b.length; j++) dp[0][j] = j;
  for (let i=1; i<=a.length; i++) {
    for (let j=1; j<=b.length; j++) {
      dp[i][j] = a[i-1] === b[j-1]
        ? dp[i-1][j-1]
        : 1 + Math.min(dp[i-1][j], dp[i][j-1], dp[i-1][j-1]);
    }
  }
  const dist = dp[a.length][b.length];
  return 1 - dist / Math.max(a.length, b.length);
}

/** Nilai anestesi yang BUKAN nama dokter (opsi valid di dropdown — lihat
 *  SF label="Dokter Anestesi" di ViewDaftar) — kalau OCR mengembalikan
 *  salah satu dari ini, JANGAN coba fuzzy-match ke nama dokter, karena
 *  ini memang bukan nama orang dan bisa salah "dikoreksi" jadi nama
 *  dokter yang kebetulan mirip hurufnya. */
export function looksLikeNonDoctorAnesthesiaValue(raw: string): boolean {
  const n = raw.toLowerCase().replace(/[.,]/g, "").trim();
  return /^(lokal anestesi|anestesi lokal|lokal|local|tanpa anestesi|-)$/i.test(n);
}

/** Cari nama staf paling mirip dari `candidates` (sudah difilter role-nya
 *  di pemanggil). Return nama TERDAFTAR (bentuk resmi) kalau similarity
 *  cukup tinggi, atau `null` kalau tidak ada yang cukup mirip (pemanggil
 *  harus tetap pakai nilai OCR asli dalam kasus ini). */
export function findBestStaffMatch(ocrName: string, candidates: string[], threshold = 0.72): string | null {
  const normOcr = normalizeDoctorName(ocrName);
  if (!normOcr) return null;
  let best: {name: string; score: number; distance: number; candidateLen: number} | null = null;
  for (const candidate of candidates) {
    const normCandidate = normalizeDoctorName(candidate);
    if (!normCandidate) continue;
    const score = levenshteinRatio(normOcr, normCandidate);
    const distance = Math.round((1 - score) * Math.max(normOcr.length, normCandidate.length));
    if (!best || score > best.score) best = {name: candidate, score, distance, candidateLen: normCandidate.length};
  }
  if (!best) return null;

  /* FIX BUG (safety-net nama pendek): untuk string PENDEK (<6 karakter,
     kira-kira 1 nama depan tanpa marga — mis. "Budi"), satu huruf beda
     SECARA MATEMATIS TIDAK BISA dibedakan antara "typo" vs "memang orang
     lain" (mis. "Budi" vs "Rudi" — skor 0.75, lolos threshold 0.72,
     padahal jelas 2 nama BERBEDA). Untuk kasus ini, WAJIB exact match
     (case/tanda baca diabaikan) — TIDAK ada toleransi fuzzy sama sekali.
     Untuk nama yang cukup panjang (>=6 karakter, biasanya nama lengkap
     dengan marga), toleransi fuzzy tetap dipakai seperti biasa karena
     kemungkinan tabrakan tak sengaja jauh lebih kecil. */
  const shorterLen = Math.min(normOcr.length, best.candidateLen);
  if (shorterLen < 6) {
    return normOcr === normalizeDoctorName(best.name) ? best.name : null;
  }
  if (best.score < threshold || best.distance > 2) return null;
  return best.name;
}

/* FIX AUDIT #30 (CRITICAL, bagian OCR): validasi terpusat untuk field
   semi-terstruktur (umur, tanggal, jam, kamar, gol. darah) hasil OCR.
   SEBELUMNYA hanya diterapkan di applyOcrResult (jalur 1-kandidat) — jalur
   autoSaveOcrBatch (≥2 kandidat) menulis field ini MENTAH langsung ke
   Supabase tanpa validasi apa pun. Sekarang kedua jalur memakai validator
   yang SAMA, supaya tidak ada lagi celah data tidak valid yang lolos ke
   database produksi hanya karena lewat jalur auto-save batch. */
export function validateOcrScalarFields(r: OcrExtractResult): {
  age: string; ageMonths: string; date: string; time: string; room: string; bloodType: string;
} {
  const ageNum = Number(r.age);
  const age = (r.age && !isNaN(ageNum) && ageNum>=0 && ageNum<=150) ? String(ageNum) : "";
  const moNum = Number(r.ageMonths);
  const ageMonths = (r.ageMonths && !isNaN(moNum) && moNum>=0 && moNum<=11) ? String(moNum) : "";
  const date = r.date ? (toISODateStrict(r.date) || "") : "";
  const time = (r.time && /^([01]\d|2[0-3]):([0-5]\d)$/.test(r.time)) ? r.time : "";
  const roomMatch = r.room ? r.room.match(/\d+/) : null;
  const room = roomMatch ? (ROOMS.find(rm2 => rm2.includes(roomMatch[0])) || "") : "";
  const bloodType = r.bloodType ? (BT.find(bt => bt.toUpperCase()===r.bloodType.toUpperCase()) || "") : "";
  return { age, ageMonths, date, time, room, bloodType };
}

/* SARAN PENINGKATAN #10 + #2 (SATU sumber kebenaran untuk aturan
   kehati-hatian OCR — sebelumnya applyOcrResult & autoSaveOcrBatch
   masing-masing punya logika "jangan auto-percaya CYTO" sendiri-sendiri,
   yang terbukti bisa DRIFT: satu jalur diperbaiki (FIX AUDIT #30 untuk
   batch), satu lagi tertinggal berbulan-bulan (FIX AUDIT #47 untuk
   1-kandidat) sampai ketahuan lewat audit manual. Supaya kelas bug ini
   tidak terulang, SEMUA keputusan "field mana yang butuh perhatian
   ekstra & kenapa" sekarang hidup di SATU fungsi ini. Kalau nanti perlu
   menambah field berisiko lain (mis. nama obat kemoterapi), cukup ubah
   di sini — otomatis berlaku ke jalur 1-kandidat maupun batch.

   Field yang diberi perlakuan berbeda:
   - opType "cyto": TIDAK di-auto-fill sama sekali (beda dari field lain)
     karena statusnya bukan cuma "data salah" tapi memicu EFEK SAMPING
     otomatis di seluruh app (banner darurat, lompat ke urutan teratas
     jadwal untuk SEMUA staf) — risiko salah-baca di sini jauh lebih
     mahal daripada field naratif biasa.
   - bloodType & allergy: TETAP di-auto-fill (beda dari cyto, karena
     keduanya tidak memicu efek samping otomatis apa pun di app — cuma
     data yang tersimpan), TAPI diberi styling "kritis" (merah, sama
     seperti RM & tanggal) supaya tidak tenggelam di antara field
     naratif lain yang risikonya lebih rendah kalau salah baca (nama
     dokter, ruang asal, dll — salah di situ gampang ketahuan & gampang
     diperbaiki; salah gol.darah/alergi baru ketahuan saat tindakan
     medis berlangsung). */
export function buildOcrSafeApply(r: OcrExtractResult): {
  scalars: ReturnType<typeof validateOcrScalarFields>;
  suspectedCyto: boolean;
  cautionNote: string;
  criticalFieldKeys: string[];
} {
  const scalars = validateOcrScalarFields(r);
  const suspectedCyto = r.opType === "cyto";
  const notes: string[] = [];
  if(suspectedCyto){
    notes.push("⚠ OCR mendeteksi kemungkinan CYTO/EMERGENCY — WAJIB diverifikasi & dipilih manual di \"Jenis Operasi\" jika benar.");
  }
  if(r.bloodType || r.allergy){
    notes.push("⚠ Golongan darah & riwayat alergi hasil OCR WAJIB dicek ulang terhadap formulir asli sebelum tindakan — jangan diandalkan sepenuhnya dari hasil scan.");
  }
  const criticalFieldKeys = ["rm","date", ...((r.bloodType||r.allergy) ? ["bloodType","allergy"] : [])];
  return { scalars, suspectedCyto, cautionNote: notes.join("\n"), criticalFieldKeys };
}

/* SARAN PENINGKATAN #11 (unit test untuk validasi OCR) — CATATAN JUJUR:
   App.tsx ini adalah satu file komponen React, BUKAN proyek dengan test
   runner (Vitest/Jest) terpasang — menambah test framework sungguhan
   perlu perubahan di luar file ini (package.json, config, folder test
   terpisah) yang tidak bisa dilakukan hanya dengan mengedit App.tsx.
   Sebagai pengganti sementara yang tetap berguna: fungsi ini menjalankan
   beberapa kasus tepi terhadap coerceOcrResult/validateOcrScalarFields/
   buildOcrSafeApply secara manual, dan mencetak hasilnya ke console.
   Cara pakai: buka DevTools Console di browser lalu panggil
   `window.__ocrSelfTest()`. Kalau nanti proyek ini benar-benar dipecah
   jadi beberapa file dengan Vitest, kasus-kasus di bawah ini tinggal
   dipindah jadi `it(...)` beneran. */
export function __ocrSelfTest(): void {
  const results: {name: string; pass: boolean; detail?: string}[] = [];
  const check = (name: string, cond: boolean, detail?: string) => results.push({name, pass: cond, detail});

  // 1) coerceOcrResult harus tahan terhadap tipe tak terduga dari server.
  const c1 = coerceOcrResult({ patient: 123, age: null, opType: "TIDAK_VALID", confidence: "medium" });
  check("coerce: patient number→string", c1.patient === "123");
  check("coerce: age null→empty string", c1.age === "");
  check("coerce: opType tak dikenal→empty", c1.opType === "");
  check("coerce: confidence tak dikenal→low (fail-safe)", c1.confidence === "low");

  // 2) Umur 0 tahun (bayi baru lahir) harus tetap valid, bukan dianggap kosong.
  const v1 = validateOcrScalarFields(coerceOcrResult({ age: "0", ageMonths: "6" }));
  check("umur 0 tahun valid (bukan falsy)", v1.age === "0", `dapat: "${v1.age}"`);
  check("umur bulan 6 valid", v1.ageMonths === "6");

  // 3) Umur di luar batas wajar harus ditolak (dikosongkan), bukan disimpan mentah.
  const v2 = validateOcrScalarFields(coerceOcrResult({ age: "999" }));
  check("umur 999 (tidak wajar) ditolak", v2.age === "", `dapat: "${v2.age}"`);

  // 4) Tanggal ambigu/tidak valid → dikosongkan, bukan disimpan mentah.
  const v3 = validateOcrScalarFields(coerceOcrResult({ date: "35/13/2026" }));
  check("tanggal mustahil ditolak", v3.date === "", `dapat: "${v3.date}"`);

  // 5) Jam di luar format HH:MM 24-jam → ditolak.
  const v4 = validateOcrScalarFields(coerceOcrResult({ time: "25:99" }));
  check("jam tidak valid ditolak", v4.time === "", `dapat: "${v4.time}"`);

  // 6) buildOcrSafeApply: CYTO harus terdeteksi & TIDAK auto-fill (lihat FIX AUDIT #47).
  const safeCyto = buildOcrSafeApply(coerceOcrResult({ opType: "cyto" }));
  check("cyto terdeteksi sebagai suspectedCyto", safeCyto.suspectedCyto === true);
  check("cyto menghasilkan cautionNote", safeCyto.cautionNote.length > 0);

  // 7) buildOcrSafeApply: golongan darah harus masuk criticalFieldKeys (saran #2).
  const safeBT = buildOcrSafeApply(coerceOcrResult({ bloodType: "A" }));
  check("bloodType terdaftar sebagai field kritis", safeBT.criticalFieldKeys.includes("bloodType"));
  check("allergy ikut kritis walau kosong (satu paket dgn bloodType)", safeBT.criticalFieldKeys.includes("allergy"));

  // 8) Kasus normal (bukan cyto, tanpa gol.darah/alergi) → tidak ada catatan ekstra.
  const safeNormal = buildOcrSafeApply(coerceOcrResult({ opType: "elektif", patient: "Tn. Test" }));
  check("kasus normal tidak memicu cautionNote", safeNormal.cautionNote === "");
  check("kasus normal criticalFieldKeys hanya rm+date", safeNormal.criticalFieldKeys.length === 2);

  const passed = results.filter(r=>r.pass).length;
  const failed = results.filter(r=>!r.pass);
  // eslint-disable-next-line no-console
  console.log(`[__ocrSelfTest] ${passed}/${results.length} lulus.`);
  results.forEach(r => console.log(`  ${r.pass ? "✓" : "✗"} ${r.name}${r.detail ? " — " + r.detail : ""}`));
  if(failed.length>0){
    // eslint-disable-next-line no-console
    console.warn(`[__ocrSelfTest] ${failed.length} kasus GAGAL — cek validateOcrScalarFields/buildOcrSafeApply.`);
  }
}
if(typeof window !== "undefined"){ (window as any).__ocrSelfTest = __ocrSelfTest; }
