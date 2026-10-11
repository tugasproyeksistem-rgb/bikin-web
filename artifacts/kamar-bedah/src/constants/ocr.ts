/* src/constants/ocr.ts
   Konstanta batasan & endpoint untuk fitur OCR (foto, paste-teks, dan
   Asisten Virtual chat — ketiganya reuse endpoint & batas yang sama).
   Dipindah dari App.tsx, tidak ada perubahan nilai. */

export const OCR_SCAN_ENDPOINT = "/api/ocr-scan"; // route di Worker — lihat worker.ts
export const MAX_OCR_BATCH = 20;
export const MAX_OCR_FILE_MB = 15;
export const MAX_OCR_PASTE_CHARS = 4000;
export const LS_OFFLINE_QUEUE_KEY = "kb_operasi_offline_queue_v1";
