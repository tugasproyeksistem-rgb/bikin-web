/* src/components/mutu-notif.tsx
   Peringatan "input semua pasien ke sistem" untuk Pemantauan Mutu Kamar Bedah (kotak cerah, teks singkat).
   Tampil di paling atas tab Daftar dan di Home mulai 1 Januari 2027 (atau saat uji implant diaktifkan).
   Admin bisa menonaktifkan/mengaktifkan dan MENGEDIT pesannya; semuanya tersimpan di kamar_bedah_config
   (key "mutu_notif": {aktif, judul, isi}) sehingga berlaku untuk semua pengguna. */
import { useState, useEffect } from "react";
import { C } from "../constants";
import { SUPA_CLIENT } from "../services/supabase";
import { implantAktif } from "../constants/fitur-tanggal";
import { ClipboardCheck } from "lucide-react";

const KUNCI = "mutu_notif";
const JUDUL_BAWAAN = "Input semua pasien ke sistem";
const ISI_BAWAAN = "Mulai 1 Januari 2027, mohon semua pasien diinput dengan lengkap (nomor RM dan pilihan implant). Datanya dipakai untuk Pemantauan Mutu Kamar Bedah.";
const MAKS_JUDUL = 80, MAKS_ISI = 600;

export function MutuNotif({ role, setTab, showToast }: any) {
  const [aktif, setAktif] = useState<boolean>(true);
  const [judul, setJudul] = useState<string>("");
  const [isi, setIsi] = useState<string>("");
  const [siap, setSiap] = useState<boolean>(false);
  const [sibuk, setSibuk] = useState<boolean>(false);
  const [edit, setEdit] = useState<boolean>(false);
  const [fJudul, setFJudul] = useState<string>("");
  const [fIsi, setFIsi] = useState<string>("");
  const [, setTick] = useState<number>(0);
  const admin = role === "admin";

  useEffect(() => {
    let batal = false;
    const muat = async () => {
      try {
        const r: any = await SUPA_CLIENT.from("kamar_bedah_config").select("value").eq("key", KUNCI);
        if (!batal && !r.error) {
          const v = (r.data || [])[0]?.value;
          let j: any = {};
          if (v) { try { j = JSON.parse(v) || {}; } catch { j = {}; } }
          setAktif(j.aktif !== false);
          setJudul(typeof j.judul === "string" ? j.judul : "");
          setIsi(typeof j.isi === "string" ? j.isi : "");
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

  /* simpan SELURUH pengaturan sekaligus supaya mengubah satu bagian tidak menghapus bagian lain */
  const simpanCfg = async (p: { aktif?: boolean; judul?: string; isi?: string }): Promise<boolean> => {
    const nilai = { aktif: p.aktif !== undefined ? p.aktif : aktif, judul: p.judul !== undefined ? p.judul : judul, isi: p.isi !== undefined ? p.isi : isi };
    setSibuk(true);
    try {
      const { error }: any = await SUPA_CLIENT.from("kamar_bedah_config").upsert({ key: KUNCI, value: JSON.stringify(nilai) }, { onConflict: "key" });
      if (error) { showToast && showToast("Gagal menyimpan pengaturan notifikasi: " + (error.message || "error"), C.d); setSibuk(false); return false; }
      setAktif(nilai.aktif); setJudul(nilai.judul); setIsi(nilai.isi);
      setSibuk(false); return true;
    } catch { showToast && showToast("Gagal menyimpan pengaturan notifikasi", C.d); setSibuk(false); return false; }
  };

  const ubahAktif = async (nilai: boolean) => {
    if (!nilai && !window.confirm("Nonaktifkan notifikasi ini untuk SEMUA pengguna?\nAdmin bisa mengaktifkannya lagi kapan saja.")) return;
    if (await simpanCfg({ aktif: nilai }) && showToast) showToast(nilai ? "Notifikasi diaktifkan" : "Notifikasi dinonaktifkan", C.s);
  };
  const bukaEdit = () => { setFJudul(judul.trim() || JUDUL_BAWAAN); setFIsi(isi.trim() || ISI_BAWAAN); setEdit(true); };
  const simpanPesan = async () => {
    const j = fJudul.trim().slice(0, MAKS_JUDUL), s = fIsi.trim().slice(0, MAKS_ISI);
    if (!j || !s) { showToast && showToast("Judul dan isi pesan tidak boleh kosong", C.d); return; }
    if (await simpanCfg({ judul: j === JUDUL_BAWAAN ? "" : j, isi: s === ISI_BAWAAN ? "" : s })) { setEdit(false); showToast && showToast("Pesan notifikasi disimpan", C.s); }
  };
  const pakaiBawaan = async () => {
    if (!window.confirm("Kembalikan judul dan isi pesan ke teks bawaan?")) return;
    if (await simpanCfg({ judul: "", isi: "" })) { setEdit(false); showToast && showToast("Pesan dikembalikan ke teks bawaan", C.s); }
  };

  if (!implantAktif() || !siap) return null;

  const gayaIsian: any = { width: "100%", boxSizing: "border-box", padding: "8px 10px", borderRadius: 8, border: "1px solid #D97706", fontSize: 14, fontFamily: "inherit", background: "#fff", color: "#431407" };
  const formEdit = admin && edit ? (
    <div style={{ border: "2px solid #F97316", borderRadius: 14, background: "#FFF7ED", padding: "12px 14px", marginBottom: 14 }}>
      <div style={{ fontSize: 13, fontWeight: 800, color: "#7C2D12", marginBottom: 8 }}>Edit pesan notifikasi (tampil untuk semua pengguna)</div>
      <label style={{ display: "block", fontSize: 11, fontWeight: 700, color: "#9A3412", marginBottom: 4 }}>JUDUL</label>
      <input value={fJudul} maxLength={MAKS_JUDUL} onChange={(e) => setFJudul(e.target.value)} style={gayaIsian} />
      <label style={{ display: "block", fontSize: 11, fontWeight: 700, color: "#9A3412", margin: "10px 0 4px" }}>ISI PESAN (boleh beberapa baris)</label>
      <textarea value={fIsi} maxLength={MAKS_ISI} rows={5} onChange={(e) => setFIsi(e.target.value)} style={{ ...gayaIsian, resize: "vertical", lineHeight: 1.5 }} />
      <div style={{ fontSize: 11, color: "#9A3412", marginTop: 4 }}>{fIsi.length}/{MAKS_ISI} karakter</div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginTop: 10 }}>
        <button type="button" disabled={sibuk} onClick={simpanPesan} style={{ border: "none", background: "#16A34A", color: "#fff", borderRadius: 8, padding: "7px 16px", fontSize: 13, fontWeight: 800, cursor: "pointer", fontFamily: "inherit" }}>Simpan pesan</button>
        <button type="button" disabled={sibuk} onClick={() => setEdit(false)} style={{ border: "1px solid #D1D5DB", background: "#fff", color: "#374151", borderRadius: 8, padding: "7px 14px", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>Batal</button>
        <button type="button" disabled={sibuk} onClick={pakaiBawaan} style={{ border: "1px solid #D1D5DB", background: "#fff", color: "#6B7280", borderRadius: 8, padding: "7px 14px", fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>Kembalikan teks bawaan</button>
      </div>
    </div>
  ) : null;

  if (!aktif) {
    if (!admin) return null;
    return (
      <div>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 8, alignItems: "center", justifyContent: "space-between", padding: "8px 12px", marginBottom: 12, border: "1px dashed #D1D5DB", borderRadius: 10, background: "#F9FAFB", fontSize: 12, color: "#6B7280" }}>
          <span>Notifikasi pengisian data lengkap sedang <b>nonaktif</b> (hanya terlihat oleh admin).</span>
          <span style={{ display: "flex", gap: 8 }}>
            <button type="button" disabled={sibuk} onClick={bukaEdit} style={{ border: "1px solid #D1D5DB", background: "#fff", borderRadius: 8, padding: "5px 10px", fontSize: 12, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>Edit pesan</button>
            <button type="button" disabled={sibuk} onClick={() => ubahAktif(true)} style={{ border: "1px solid #D1D5DB", background: "#fff", borderRadius: 8, padding: "5px 10px", fontSize: 12, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>Aktifkan</button>
          </span>
        </div>
        {formEdit}
      </div>
    );
  }

  return (
    <div>
      <style>{`
        @keyframes kbMutuGlow{0%,100%{box-shadow:0 0 0 0 rgba(249,115,22,.7)}50%{box-shadow:0 0 18px 6px rgba(249,115,22,.35)}}
        .kbMutuGlow{animation:kbMutuGlow 1.6s ease-in-out infinite}
        @media (prefers-reduced-motion: reduce){.kbMutuGlow{animation:none}}
      `}</style>
      <div role="alert" className="kbMutuGlow" style={{ border: "3px solid #F97316", borderRadius: 14, background: "linear-gradient(135deg,#FEF08A,#FDBA74)", padding: "14px 16px", marginBottom: 14 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 18, fontWeight: 900, color: "#7C2D12" }}>
          <ClipboardCheck size={24} strokeWidth={2.6} />
          <span>{judul.trim() || JUDUL_BAWAAN}</span>
        </div>
        <div style={{ fontSize: 14, fontWeight: 600, color: "#431407", marginTop: 6, lineHeight: 1.5, whiteSpace: "pre-wrap" }}>
          {isi.trim() || ISI_BAWAAN}
        </div>
        {(setTab || admin) && (
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginTop: 10 }}>
            {setTab && <button type="button" onClick={() => setTab("daftar")} style={{ border: "none", background: "#C2410C", color: "#fff", borderRadius: 8, padding: "7px 14px", fontSize: 13, fontWeight: 800, cursor: "pointer", fontFamily: "inherit" }}>Buka tab Daftar →</button>}
            {admin && <button type="button" disabled={sibuk} onClick={bukaEdit} style={{ border: "1px solid #C2410C", background: "#fff", color: "#7C2D12", borderRadius: 8, padding: "7px 14px", fontSize: 13, fontWeight: 800, cursor: "pointer", fontFamily: "inherit" }}>Edit pesan</button>}
            {admin && <button type="button" disabled={sibuk} onClick={() => ubahAktif(false)} style={{ border: "1px solid #C2410C", background: "rgba(255,255,255,.7)", color: "#7C2D12", borderRadius: 8, padding: "7px 14px", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>Nonaktifkan notifikasi</button>}
          </div>
        )}
      </div>
      {formEdit}
    </div>
  );
}
