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
