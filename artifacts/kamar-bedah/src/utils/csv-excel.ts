/* src/utils/csv-excel.ts
   Fungsi utilitas murni untuk export CSV & Excel — styling sel XLSX,
   trigger unduh file browser, sanitasi anti formula-injection, dan parser
   CSV yang aman terhadap field ber-quote. Dipindah dari App.tsx (langkah 3
   refactor — pisah utils, batch 3), TIDAK ADA perubahan logika, termasuk
   seluruh FIX AUDIT & FIX BUG di dalamnya. */

import * as XLSX from "xlsx";

/* FIX AUDIT #21 (HIGH — Code Duplication): xlsxCell/xlsxSetCell diekstrak
   jadi SATU fungsi global yang menggabungkan SEMUA behavior styling sel
   (termasuk warna teal "16685F") yang tadinya terduplikasi di beberapa
   fungsi export Excel berbeda — dipakai oleh SEMUA fungsi export Excel di
   file ini. Murni stateless, tidak bergantung closure apa pun. */
export function xlsxCell(v:any, bold?:boolean, center?:boolean, border?:boolean, bg?:string, fontSize?:number): XLSX.CellObject {
  const s:any = {
    font: { bold: !!bold, sz: fontSize||10, color: bg==="16685F" ? {rgb:"FFFFFF"} : undefined },
    alignment: { horizontal: center?"center":"left", vertical:"center", wrapText:true },
  };
  if (border) s.border = { top:{style:"thin"}, bottom:{style:"thin"}, left:{style:"thin"}, right:{style:"thin"} };
  if (bg) s.fill = { fgColor:{rgb:bg}, patternType:"solid" };
  return { v, t:"s", s } as XLSX.CellObject;
}
/* xlsxSetCell: helper untuk menulis cell langsung ke worksheet pada
   alamat (r,c), memakai xlsxCell() di atas untuk styling. */
export function xlsxSetCell(ws: XLSX.WorkSheet, r:number, c:number, v:any, bold?:boolean, center?:boolean, border?:boolean, bg?:string, fs?:number) {
  ws[XLSX.utils.encode_cell({r,c})] = xlsxCell(v,bold,center,border,bg,fs);
}

/* ─── REAL FILE HELPERS ─────────────────────────────────────────────── */
/* FIX BUG #32 (audit lanjutan): SEBELUMNYA a.click() tidak dibungkus
   try/catch — di sebagian browser (Safari/iOS lama, atau saat popup/
   download diblokir kebijakan browser) pemanggilan ini bisa melempar
   exception yang TIDAK tertangani, membuat elemen <a> & Object URL
   tertinggal (memory leak kecil) dan pengguna tidak tahu unduhan gagal.
   Sekarang dibungkus try/finally: cleanup (hapus elemen, revoke URL)
   SELALU jalan baik sukses maupun gagal, dan kegagalan di-log + dikembalikan
   sebagai boolean supaya pemanggil BISA (opsional) menampilkan pesan error
   ke pengguna alih-alih gagal senyap. */
export const downloadBlob = (content: string, filename: string, mime="text/csv"): boolean => {
  const blob = new Blob(["\uFEFF"+content], {type:mime});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href=url; a.download=filename;
  document.body.appendChild(a);
  try {
    a.click();
    return true;
  } catch (e) {
    console.error("[downloadBlob] Gagal memicu unduhan:", e);
    return false;
  } finally {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }
};
/* FIX AUDIT #7 (HIGH — sanitizeForCSV Incomplete / CSV Injection Bypass):
   Implementasi lama hanya menangkap karakter berbahaya di AWAL string.
   Penyerang bisa menyisipkan "\n=cmd|..." di dalam field multi-line —
   Excel memproses baris setelah newline sebagai sel baru dengan formula aktif.
   Fix:
   1. Hapus semua \r\n dari dalam field (normalize ke spasi)
   2. Strip karakter kontrol lain (U+0000–U+001F kecuali tab yang sudah dihandle)
   3. Neutralisasi formula prefix (=, +, -, @, |, \t, \r) di awal hasil bersih
   Dengan urutan ini, injeksi via embedded newline tidak bisa lolos. */
export const sanitizeForCSV = (val: any): string => {
  /* FIX BUG #46 (audit lanjutan): U+2028 (LINE SEPARATOR) & U+2029
     (PARAGRAPH SEPARATOR) adalah karakter Unicode yang secara visual mirip
     spasi tapi diperlakukan sebagai pemisah baris oleh sebagian
     parser/aplikasi — bisa jadi celah row-injection yang sama seperti
     \r\n kalau tidak dinormalisasi juga. */
  // 1. Normalize newline (termasuk U+2028/U+2029) ke spasi agar tidak membentuk baris baru di CSV
  let s = String(val ?? "").replace(/\r?\n/g, " ").replace(/\r/g, " ").replace(/[\u2028\u2029]/g, " ");
  // 2. Escape tanda kutip ganda
  s = s.replace(/"/g, '""');
  // 3. Strip karakter kontrol non-printable (kecuali spasi/tab)
  s = s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "");
  // 4. Neutralisasi formula prefix di awal string hasil bersih
  if (/^[=+\-@\t\r|]/.test(s)) s = "'" + s;
  return `"${s}"`;
};

/* parseCSV: parser CSV minimal tapi AMAN terhadap field yang dibungkus tanda
   kutip (mis. `"Doe, John"` atau keterangan yang mengandung koma) — kasus ini
   sangat umum pada CSV hasil "Save As" dari Excel. Implementasi lama
   (`text.split("\n").map(r=>r.split(","))`) memecah SETIAP koma tanpa
   peduli tanda kutip, sehingga 1 baris bisa pecah menjadi kolom yang salah
   (nama/telepon/jabatan bergeser) TANPA error apa pun yang terlihat —
   silent data corruption saat import staf/roster. Parser di bawah
   menangani: koma di dalam tanda kutip, tanda kutip ganda escaped (""),
   serta akhir baris CRLF/LF. */
export const parseCSV = (text: string): string[][] => {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let inQuotes = false;
  const normalized = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  for (let i = 0; i < normalized.length; i++) {
    const ch = normalized[i];
    if (inQuotes) {
      if (ch === '"') {
        if (normalized[i + 1] === '"') { field += '"'; i++; }
        else inQuotes = false;
      } else field += ch;
    } else if (ch === '"') {
      inQuotes = true;
    } else if (ch === ",") {
      row.push(field.trim()); field = "";
    } else if (ch === "\n") {
      row.push(field.trim()); rows.push(row); row = []; field = "";
    } else {
      field += ch;
    }
  }
  if (field.length || row.length) { row.push(field.trim()); rows.push(row); }
  return rows.filter(r => r.some(c => c !== ""));
};
