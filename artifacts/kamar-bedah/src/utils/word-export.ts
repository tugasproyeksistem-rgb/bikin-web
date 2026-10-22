/* src/utils/word-export.ts
   Export tabel data ke file Word (.doc via HTML table) — dipakai di
   berbagai laporan (Arsip, Lembur, dll). Dipindah dari App.tsx (diekstrak
   bersamaan dengan ViewArsip yang memakainya), TIDAK ADA perubahan logika. */

import { HOSPITAL } from "../constants";
import { escHtml } from "./security";
import { fNow } from "./format";

export function downloadAsWord(title: string, rows: string[][], filename: string) {
  const thead = rows[0].map(h=>`<th style="background:#16685F;color:#fff;padding:6px 10px;font-size:11px">${escHtml(h)}</th>`).join("");
  const tbody = rows.slice(1).map((r,i)=>`<tr style="background:${i%2===0?"#F0FFF8":"#fff"}">${r.map(c=>`<td style="padding:5px 10px;border-bottom:1px solid #e0e0e0;font-size:11px">${escHtml(c)}</td>`).join("")}</tr>`).join("");
  const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"><title>${escHtml(title)}</title></head><body><h2 style="color:#16685F;font-family:Arial">${escHtml(title)}</h2><p style="color:#555;font-size:12px;font-family:Arial">Dicetak: ${escHtml(fNow())} — ${escHtml(HOSPITAL)}</p><table style="border-collapse:collapse;width:100%;font-family:Arial"><thead><tr>${thead}</tr></thead><tbody>${tbody}</tbody></table></body></html>`;
  const blob = new Blob([html], {type:"application/msword"});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href=url; a.download=filename;
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}


/* Unduh lembar (elemen HTML) sebagai file Word (.doc). Lebar piksel besar diubah ke 100%. */
export function downloadElementAsWord(el: HTMLElement, filename: string, landscape = false) {
  const klon = el.cloneNode(true) as HTMLElement;
  [klon, ...Array.from(klon.querySelectorAll<HTMLElement>("*"))].forEach((n) => {
    if (n.style && n.style.width.endsWith("px") && parseFloat(n.style.width) >= 600) n.style.width = "100%";
  });
  const ukuran = landscape ? "size:29.7cm 21cm;mso-page-orientation:landscape" : "size:21cm 29.7cm";
  const html = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"><title>' + escHtml(filename) + '</title><style>@page Section1{' + ukuran + ';margin:1.2cm}div.Section1{page:Section1}body{font-family:"Times New Roman",Times,serif}table{border-collapse:collapse}</style></head><body><div class="Section1">' + klon.innerHTML + '</div></body></html>';
  const blob = new Blob([html], { type: "application/msword" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

/* Sama seperti downloadElementAsWord, tetapi mengembalikan Blob (untuk dibagikan / dikirim WA). */
export function elementToWordBlob(el: HTMLElement, landscape = false): Blob {
  const klon = el.cloneNode(true) as HTMLElement;
  [klon, ...Array.from(klon.querySelectorAll<HTMLElement>("*"))].forEach((n) => {
    if (n.style && n.style.width.endsWith("px") && parseFloat(n.style.width) >= 600) n.style.width = "100%";
  });
  const ukuran = landscape ? "size:29.7cm 21cm;mso-page-orientation:landscape" : "size:21cm 29.7cm";
  const html = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"><title>Laporan</title><style>@page Section1{' + ukuran + ';margin:1.2cm}div.Section1{page:Section1}body{font-family:"Times New Roman",Times,serif}table{border-collapse:collapse}</style></head><body><div class="Section1">' + klon.innerHTML + '</div></body></html>';
  return new Blob([html], { type: "application/msword" });
}
