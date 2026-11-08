/* src/components/mutu-notif.tsx
   Peringatan "isi data pasien dengan lengkap" untuk Pemantauan Mutu Kamar Bedah. Tampil di paling atas tab Daftar
   dan di Home mulai 1 Januari 2027 (atau saat uji implant diaktifkan). Admin bisa menonaktifkan/mengaktifkannya;
   statusnya tersimpan di kamar_bedah_config (key "mutu_notif") sehingga berlaku untuk semua pengguna. */
import { useState, useEffect } from "react";
import { C } from "../constants";
import { SUPA_CLIENT } from "../services/supabase";
import { implantAktif } from "../constants/fitur-tanggal";
import { ClipboardCheck } from "lucide-react";

const KUNCI = "mutu_notif";

export function MutuNotif({ role, setTab, showToast }: any) {
  const [aktif, setAktif] = useState<boolean>(true);
  const [siap, setSiap] = useState<boolean>(false);
  const [sibuk, setSibuk] = useState<boolean>(false);
  const [, setTick] = useState<number>(0);
  const admin = role === "admin";

  useEffect(() => {
    let batal = false;
    const muat = async () => {
      try {
        const r: any = await SUPA_CLIENT.from("kamar_bedah_config").select("value").eq("key", KUNCI);
        if (!batal && !r.error) {
          const v = (r.data || [])[0]?.value;
          if (v) { try { setAktif(JSON.parse(v).aktif !== false); } catch { setAktif(true); } } else setAktif(true);
        }
      } catch { /* offline: pakai bawaan (tampil) */ }
      if (!batal) setSiap(true);
    };
    muat();
    const segarkan = () => { muat(); setTick((x) => x + 1); };
    const t = setInterval(segarkan, 60000);
    window.addEventListener("focus", segarkan);
    return () => { batal = true; clearInterval(t); window.removeEventListener("focus", segarkan); };
  }, []);

  const ubah = async (nilai: boolean) => {
    if (!nilai && !window.confirm("Nonaktifkan notifikasi ini untuk SEMUA pengguna?\nAdmin bisa mengaktifkannya lagi kapan saja.")) return;
    setSibuk(true);
    try {
      const { error }: any = await SUPA_CLIENT.from("kamar_bedah_config").upsert({ key: KUNCI, value: JSON.stringify({ aktif: nilai }) }, { onConflict: "key" });
      if (error) { showToast && showToast("Gagal menyimpan pengaturan notifikasi: " + (error.message || "error"), C.d); }
      else { setAktif(nilai); showToast && showToast(nilai ? "Notifikasi diaktifkan" : "Notifikasi dinonaktifkan", C.s); }
    } catch { showToast && showToast("Gagal menyimpan pengaturan notifikasi", C.d); }
    setSibuk(false);
  };

  if (!implantAktif() || !siap) return null;

  if (!aktif) {
    if (!admin) return null;
    return (
      <div style={{ display: "flex", flexWrap: "wrap", gap: 8, alignItems: "center", justifyContent: "space-between", padding: "8px 12px", marginBottom: 12, border: "1px dashed #D1D5DB", borderRadius: 10, background: "#F9FAFB", fontSize: 12, color: "#6B7280" }}>
        <span>Notifikasi pengisian data lengkap sedang <b>nonaktif</b> (hanya terlihat oleh admin).</span>
        <button type="button" disabled={sibuk} onClick={() => ubah(true)} style={{ border: "1px solid #D1D5DB", background: "#fff", borderRadius: 8, padding: "5px 10px", fontSize: 12, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>Aktifkan</button>
      </div>
    );
  }

  return (
    <div role="alert" style={{ border: "2px solid #F59E0B", borderRadius: 14, background: "#FFFBEB", padding: "12px 14px", marginBottom: 14 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 15, fontWeight: 800, color: "#92400E" }}>
        <ClipboardCheck size={20} strokeWidth={2.4} />
        <span>Mohon isi data pasien dengan lengkap</span>
      </div>
      <div style={{ fontSize: 13, color: "#78350F", marginTop: 6, lineHeight: 1.5 }}>
        Mulai 1 Januari 2027, data yang diisi di tab Daftar (nomor RM dan pilihan implant) dipakai untuk Pemantauan Mutu Kamar Bedah. Pastikan terisi dengan benar agar laporan mutu tidak kosong.
      </div>
      {(setTab || admin) && (
        <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginTop: 10 }}>
          {setTab && <button type="button" onClick={() => setTab("daftar")} style={{ border: "none", background: "#F59E0B", color: "#fff", borderRadius: 8, padding: "6px 12px", fontSize: 12, fontWeight: 800, cursor: "pointer", fontFamily: "inherit" }}>Buka tab Daftar →</button>}
          {admin && <button type="button" disabled={sibuk} onClick={() => ubah(false)} style={{ border: "1px solid #D97706", background: "#fff", color: "#92400E", borderRadius: 8, padding: "6px 12px", fontSize: 12, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>Nonaktifkan notifikasi</button>}
        </div>
      )}
    </div>
  );
}
