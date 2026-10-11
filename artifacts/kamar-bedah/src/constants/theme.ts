/* src/constants/theme.ts
   Palet warna (C) & style input dasar (iS) yang dipakai lintas komponen.
   Dipindah dari App.tsx, tidak ada perubahan nilai. */

export const C = {
  /* ── Brand Blues (core gradient palette) ── */
  p:"#0284c7",   pL:"#38bdf8",  pBg:"#e0f2fe",
  /* ── Danger ── */
  d:"#dc2626",   dBg:"#fef2f2", dL:"#fca5a5",
  /* ── Warning ── */
  w:"#d97706",   wBg:"#fffbeb",
  /* ── Info / secondary blue ── */
  i:"#1d4ed8",   iBg:"#eff6ff",
  /* ── Success green ── */
  s:"#16a34a",   sBg:"#f0fdf4",
  /* ── WhatsApp ── */
  wa:"#25D366",  waBg:"#dcfce7",
  /* ── Neutrals ── */
  g:"#475569",   gBg:"#f1f5f9",
  t:"#0f172a",   tL:"#64748b",
  b:"#e2e8f0",   white:"#ffffff",
  bg:"#f8fafc",  gold:"#C9A961",
  /* ── PENAMBAHAN: aksen "mewah" (navy gelap + emas) untuk permukaan
     hero/signature (header, login, banner) — dipakai TERBATAS supaya
     tidak mengubah seluruh sistem warna fungsional yang sudah teruji. ── */
  navy:"#0B1220", navyDeep:"#060A12", gold2:"#E8CE8F",
};

export const iS: React.CSSProperties = { width:"100%", padding:"11px 14px", borderRadius:10, border:`1.5px solid ${C.b}`, fontSize:14, lineHeight:1.6, background:"#FAFBFD", boxSizing:"border-box", fontFamily:"inherit", color:C.t, outline:"none" };
