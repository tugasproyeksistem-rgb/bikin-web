/* src/components/header-clock.tsx
   Jam digital di header (tanggal + waktu WIB, update tiap detik). Dipindah
   dari App.tsx (langkah 5 refactor — pisah components, batch 3), TIDAK ADA
   perubahan logika. Tanpa dependency ke state/types/constants aplikasi. */

import { useState, useEffect } from "react";

export function HeaderClock() {
  const [now, setNow] = useState(() => new Date());
  useEffect(() => {
    const iv = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(iv);
  }, []);
  return (
    <>
      <div style={{color:"rgba(255,255,255,.55)",fontSize:10,marginTop:3,letterSpacing:.2}}>{now.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"})}</div>
      <div style={{color:"rgba(255,255,255,.9)",fontSize:13,fontWeight:700,fontVariantNumeric:"tabular-nums",letterSpacing:1,marginTop:1}}>{now.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"})} WIB</div>
    </>
  );
}
