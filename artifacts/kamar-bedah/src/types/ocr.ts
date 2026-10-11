/* src/types/ocr.ts
   Tipe hasil ekstraksi OCR (dipakai jalur foto, paste-teks, DAN Asisten
   Virtual chat — ketiganya reuse bentuk data yang sama) serta tipe pesan
   percakapan Asisten Virtual. Dipindah dari App.tsx, tidak ada perubahan
   logika. */

export interface OcrExtractResult {
  patient: string;
  age: string;
  ageMonths: string;
  rm: string;
  diagnosis: string;
  procedure: string;
  opType: "elektif" | "semi" | "cyto" | "";
  date: string;
  time: string;
  surgeon: string;
  anesthesiologist: string;
  room: string;
  allergy: string;
  bloodType: string;
  confidence: "high" | "low";
  lowConfidenceFields: string[];
  /* Flag opsional untuk kasus foto yang BUKAN formulir operasi sama sekali
     (salah foto, foto KTP, dsb) — begitu Worker mengirim flag
     `notAFormError:true`, UI menampilkan pesan yang lebih tepat ("sepertinya
     ini bukan formulir operasi") daripada pesan gagal generik. */
  notAFormError?: boolean;
}

export interface AssistantMsg {
  role: "user" | "assistant";
  text: string;
}
