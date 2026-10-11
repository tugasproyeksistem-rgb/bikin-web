/* src/services/ocr.ts
   Service pemanggilan endpoint OCR (/api/ocr-scan di Worker — worker.ts)
   untuk jalur foto DAN paste-teks (dipakai juga oleh Asisten Virtual
   chat). Dipindah dari App.tsx (langkah 4 refactor — pisah services),
   TIDAK ADA perubahan logika. */

import type { OcrExtractResult } from "../types";
import { OCR_SCAN_ENDPOINT, MAX_OCR_PASTE_CHARS } from "../constants";
import { coerceOcrResult, compressImageToBase64 } from "../utils";

/** Kirim gambar ke route OCR di Worker sendiri (same-origin — tidak
 *  perlu CORS khusus, tidak perlu API key/secret di client karena
 *  request ini berjalan di domain yang sama seperti App di-serve). */
/** Fungsi inti bersama: kirim body apa pun ke /api/ocr-scan lalu urai
 *  responsnya dengan cara yang SAMA PERSIS baik untuk jalur gambar maupun
 *  jalur paste-teks — supaya keduanya tunduk pada validasi & coerce yang
 *  identik (coerceOcrResult dkk) dan tidak bisa drift satu sama lain,
 *  sama semangatnya dengan buildOcrSafeApply() untuk validasi di sisi
 *  form (lihat SARAN PENINGKATAN #10). */
export async function callOcrScanEndpoint(
  body: Record<string, any>
): Promise<{ ok: boolean; data?: OcrExtractResult[]; msg: string; notAFormError?: boolean }> {
  try {
    const res = await fetch(OCR_SCAN_ENDPOINT, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(body),
    });
    if (!res.ok) {
      const errText = await res.text().catch(() => "");
      return { ok: false, msg: `Server OCR error (${res.status}): ${errText.slice(0, 200)}` };
    }
    const json = await res.json();
    /* SARAN PENINGKATAN #6 (forward-compat, lihat catatan di interface
       OcrExtractResult & OCR_SCAN_ENDPOINT di atas): kalau Worker SUATU
       SAAT diubah untuk mengembalikan `notAFormError:true` (baik di level
       respons top-level maupun disisipkan ke json.msg), diteruskan ke
       sini apa adanya. Sebelum Worker diubah, field ini akan selalu
       undefined/false dan tidak mengubah perilaku yang sudah ada. */
    const notAFormError = json?.notAFormError === true;
    if (!json?.ok) return { ok: false, msg: json?.msg || "Gagal — server mengembalikan respons tidak sukses", notAFormError };
    /* FIX AUDIT #34: coerce semua item ke tipe string yang aman sebelum
       diteruskan ke pemanggil (applyOcrResult / autoSaveOcrBatch). */
    const rawData = Array.isArray(json.data) ? json.data : [];
    return { ok: true, data: rawData.map(coerceOcrResult), msg: json.msg || "✓ Berhasil", notAFormError };
  } catch (e: any) {
    return { ok: false, msg: "Gagal menghubungi server OCR: " + (e?.message || "network error") };
  }
}

export async function runWorkerOcrScan(
  file: File
): Promise<{ ok: boolean; data?: OcrExtractResult[]; msg: string; notAFormError?: boolean }> {
  const { base64, mediaType } = await compressImageToBase64(file);
  return callOcrScanEndpoint({ image_base64: base64, media_type: mediaType });
}

/** Jalur 2 (Paste Teks): kirim TEKS MENTAH (mis. hasil copy-paste instruksi
 *  WhatsApp dokter) ke Worker yang SAMA — bukan endpoint terpisah, bukan
 *  pipeline terpisah. Worker perlu diperluas untuk membedakan payload
 *  `text` vs `image_base64` (lihat catatan lengkap & contoh kode di
 *  komentar OCR_SCAN_ENDPOINT di atas — bagian server TIDAK ada di
 *  App.tsx ini, jadi tidak bisa diubah langsung dari sini). Karena
 *  responsnya diurai lewat callOcrScanEndpoint yang SAMA dengan jalur
 *  gambar, hasilnya otomatis lewat coerceOcrResult juga — lalu di sisi
 *  form tetap lewat buildOcrSafeApply() sebelum boleh mengisi/menyimpan
 *  apa pun (gerbang CYTO/gol.darah/alergi tetap berlaku, TIDAK di-bypass
 *  seperti pada blueprint pipeline Gemini terpisah yang sempat diusulkan). */
export async function runWorkerOcrScanText(
  text: string
): Promise<{ ok: boolean; data?: OcrExtractResult[]; msg: string; notAFormError?: boolean }> {
  const trimmed = (text || "").trim();
  if (!trimmed) return { ok: false, msg: "Teks kosong — tempel dulu instruksi/memo sebelum diproses." };
  return callOcrScanEndpoint({ text: trimmed.slice(0, MAX_OCR_PASTE_CHARS) });
}
