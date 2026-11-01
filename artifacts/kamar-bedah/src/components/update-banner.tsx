/* src/components/update-banner.tsx
   Pita "Ada pembaruan": membandingkan versi build di kode ini dengan /version.json di server.
   Tidak memuat ulang otomatis supaya isian yang sedang diketik tidak hilang. */

import { useEffect, useState } from "react";

declare const __APP_VERSI__: string;

export function UpdateBanner() {
  const [baru, setBaru] = useState(false);

  useEffect(() => {
    if ((import.meta as any).env?.DEV) return;
    let batal = false;
    const cek = async () => {
      try {
        const r = await fetch("/version.json?t=" + Date.now(), { cache: "no-store" });
        if (!r.ok) return;
        const j = await r.json();
        if (!batal && j && typeof j.versi === "string" && j.versi !== __APP_VERSI__) setBaru(true);
      } catch { /* offline: abaikan */ }
    };
    cek();
    const t = setInterval(cek, 180000);
    const saatAktif = () => { if (document.visibilityState === "visible") cek(); };
    document.addEventListener("visibilitychange", saatAktif);
    window.addEventListener("focus", cek);
    return () => { batal = true; clearInterval(t); document.removeEventListener("visibilitychange", saatAktif); window.removeEventListener("focus", cek); };
  }, []);

  if (!baru) return null;

  const muatUlang = async () => {
    try { const reg = await navigator.serviceWorker?.getRegistration(); await reg?.update(); } catch { /* abaikan */ }
    window.location.reload();
  };

  return (
    <div style={{ position: "fixed", top: 0, left: 0, right: 0, zIndex: 100000, display: "flex", justifyContent: "center", alignItems: "center", gap: 12, flexWrap: "wrap", padding: "10px 14px", background: "linear-gradient(90deg,#4C1D95,#7C3AED)", color: "#fff", fontSize: 14, fontWeight: 700, boxShadow: "0 4px 16px rgba(76,29,149,.45)" }}>
      <span>🔔 Ada pembaruan aplikasi.</span>
      <button type="button" onClick={muatUlang} style={{ padding: "7px 16px", borderRadius: 10, border: "none", background: "#fff", color: "#5B21B6", fontSize: 14, fontWeight: 800, cursor: "pointer" }}>Muat ulang sekarang</button>
      <button type="button" onClick={() => setBaru(false)} aria-label="Tutup" style={{ border: "none", background: "transparent", color: "#fff", fontSize: 18, cursor: "pointer", lineHeight: 1 }}>✕</button>
    </div>
  );
}
