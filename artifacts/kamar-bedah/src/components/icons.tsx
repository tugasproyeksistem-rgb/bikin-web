/* src/components/icons.tsx
   Komponen ikon SVG line-style untuk tab navigasi (TabIcon) dan tombol
   aksi (ActionIcon) — pengganti emoji, konsisten stroke-based di seluruh
   app. Dipindah dari App.tsx (langkah 5 refactor — pisah components,
   batch 3), TIDAK ADA perubahan logika. Tanpa dependency ke state/types/
   constants aplikasi. */

export const TabIcon = ({k}: {k:string}) => {
  const common = {width:18,height:18,viewBox:"0 0 24 24",fill:"none",stroke:"currentColor",strokeWidth:2,strokeLinecap:"round" as const,strokeLinejoin:"round" as const};
  switch(k){
    case "home": return <svg {...common}><path d="M3 11.5 12 4l9 7.5"/><path d="M5 10v9a1 1 0 0 0 1 1h4v-6h4v6h4a1 1 0 0 0 1-1v-9"/></svg>;
    case "jadwal": return <svg {...common}><rect x="3" y="4" width="18" height="17" rx="2"/><path d="M8 2v4M16 2v4M3 9h18"/></svg>;
    case "daftar": return <svg {...common}><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>;
    case "jadwal_ok": case "laporan_harian": case "laporan": return <svg {...common}><path d="M14 3v4a1 1 0 0 0 1 1h4"/><path d="M6 3h8l5 5v12a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1Z"/><path d="M9 13h6M9 17h6"/></svg>;
    case "wa": return <svg {...common}><path d="M21 11.5a8.5 8.5 0 0 1-12.5 7.5L3 20l1.1-5.3A8.5 8.5 0 1 1 21 11.5Z"/></svg>;
    case "statistik": return <svg {...common}><path d="M3 20h18"/><rect x="6" y="10" width="3" height="8"/><rect x="11" y="6" width="3" height="12"/><rect x="16" y="13" width="3" height="5"/></svg>;
    case "staf": return <svg {...common}><circle cx="9" cy="8" r="3.2"/><path d="M2.5 20a6.5 6.5 0 0 1 13 0"/><circle cx="17.5" cy="9" r="2.6"/><path d="M15.8 13.2A5.2 5.2 0 0 1 21.5 18"/></svg>;
    case "lembur": return <svg {...common}><circle cx="12" cy="13" r="8"/><path d="M12 9v4l3 2M9 2h6"/></svg>;
    case "monitoring": return <svg {...common}><path d="M14 4v11a2 2 0 1 1-4 0V4a2 2 0 1 1 4 0Z"/><path d="M10 14h4M8 19h8"/></svg>;
    case "roster_gen": return <svg {...common}><rect x="3" y="4" width="18" height="17" rx="2"/><path d="M8 2v4M16 2v4M3 9h18M8 14h.01M12 14h.01M16 14h.01M8 17h.01M12 17h.01"/></svg>;
    case "arsip": return <svg {...common}><path d="M21 8H3v11a1 1 0 0 0 1 1h16a1 1 0 0 0 1-1V8Z"/><path d="M2 5a1 1 0 0 1 1-1h18a1 1 0 0 1 1 1v3H2Z"/><path d="M10 12h4"/></svg>;
    default: return null;
  }
};

/* PENAMBAHAN #6 (konsistensi ikon, perluasan): set ikon aksi SVG
   line-style — pengganti emoji di tombol aksi (Edit, Hapus, Batal, dsb)
   di seluruh card Jadwal. Sama gaya dengan TabIcon (stroke-based,
   strokeWidth konsisten), TANPA dependency baru. */
export const ActionIcon = ({k, size=14}: {k:string; size?: number}) => {
  const common = {width:size,height:size,viewBox:"0 0 24 24",fill:"none",stroke:"currentColor",strokeWidth:2,strokeLinecap:"round" as const,strokeLinejoin:"round" as const};
  switch(k){
    case "edit": return <svg {...common}><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>;
    case "trash": return <svg {...common}><path d="M3 6h18M8 6V4a1 1 0 0 1 1-1h6a1 1 0 0 1 1 1v2m2 0-.7 13.1A2 2 0 0 1 15.3 21H8.7a2 2 0 0 1-2-1.9L6 6"/></svg>;
    case "bell": return <svg {...common}><path d="M6 9a6 6 0 1 1 12 0c0 5 2 6 2 6H4s2-1 2-6Z"/><path d="M10 19a2 2 0 0 0 4 0"/></svg>;
    case "x-circle": return <svg {...common}><circle cx="12" cy="12" r="9"/><path d="m9 9 6 6m0-6-6 6"/></svg>;
    case "plus-circle": return <svg {...common}><circle cx="12" cy="12" r="9"/><path d="M12 8v8M8 12h8"/></svg>;
    case "refresh": return <svg {...common}><path d="M21 12a9 9 0 1 1-3-6.7"/><path d="M21 3v6h-6"/></svg>;
    case "play": return <svg {...common} fill="currentColor" stroke="none"><path d="M7 4.5v15l13-7.5Z"/></svg>;
    case "check": return <svg {...common}><path d="M20 6 9 17l-5-5"/></svg>;
    default: return null;
  }
};
