/* src/components/view-pengumuman.tsx
   Tab "Pengumuman" (khusus admin): kirim pesan ke staf lewat WhatsApp (wa.me), satu klik per orang.
   Tidak ada yang disimpan ke server (password hanya ada di layar dan di pesan yang dikirim). */

import { useState, useMemo } from "react";
import { C, iS } from "../constants";
import { Card, SH, Btn } from "./primitives";
import { normalizePhone62 } from "../utils";

const FOOTER = "\n\n🤖 Pesan ini dikirim otomatis, mohon tidak dibalas di chat ini.";
const TEMPLATE: { k: string; l: string; teks: string }[] = [
  { k: "password", l: "🔑 Password baru",
    teks: "📢 INFO PENTING\n\nHalo {NAMA}! 👋\n\nAda update sistem di aplikasi Kamar Bedah, jadi password untuk masuk sudah diganti ya.\n\n🔑 Password baru: {PASSWORD}\n🌐 Website: {WEBSITE}\n\nMulai sekarang pakai password yang baru. Password lama sudah tidak bisa dipakai.\n\nKalau belum bisa masuk, coba tutup aplikasinya lalu buka lagi, atau tekan Ctrl+Shift+R kalau lewat komputer.\n\nMohon password ini tidak dibagikan ke orang di luar tim kamar bedah ya. Kalau ada kendala, hubungi admin langsung. Terima kasih! 🙏" },
  { k: "update", l: "🛠 Info update sistem",
    teks: "📢 INFO UPDATE\n\nHalo {NAMA}! 👋\n\nAplikasi Kamar Bedah baru saja diperbarui, ada beberapa fitur baru dan perbaikan.\n\n🌐 Website: {WEBSITE}\n\nBiar tampilannya ikut terbaru, tutup aplikasinya lalu buka lagi, atau tekan Ctrl+Shift+R kalau lewat komputer. Kalau ada yang aneh, hubungi admin ya. Terima kasih! 🙏" },
  { k: "beset", l: "📅 Beset dibuka",
    teks: "📢 BESETAN DIBUKA\n\nHalo {NAMA}! 👋\n\nBesetan jadwal perawat sudah dibuka. Yuk segera diisi sebelum ditutup ya.\n\n🌐 Website: {WEBSITE}\n\nBuka tab Beset Jadwal, lalu isi tanggal yang kamu mau. Terima kasih! 🙏" },
  { k: "kosong", l: "✏️ Tulis sendiri", teks: "Halo {NAMA}! 👋\n\n" },
];
type Grup = "semua" | "perawat" | "bedah" | "anestesi";
const GRUP_LABEL: [Grup, string][] = [["perawat", "Perawat"], ["bedah", "Dokter bedah"], ["anestesi", "Dokter anestesi"], ["semua", "Semua"]];
const grupDari = (s: any): Grup => (s.type === "surgeon" ? "bedah" : s.type === "anesthesiologist" ? "anestesi" : "perawat");
const ganti = (t: string, kunci: string, nilai: string) => t.split(kunci).join(nilai);

export function ViewPengumuman({ staff, showToast, role }: any) {
  const [pesan, setPesan] = useState(TEMPLATE[0].teks);
  const [tpl, setTpl] = useState("password");
  const [website, setWebsite] = useState("https://kamarbedahrsrn.my.id");
  const [password, setPassword] = useState("");
  const [otomatis, setOtomatis] = useState(true);
  const [grup, setGrup] = useState<Grup>("perawat");
  const [cari, setCari] = useState("");
  const [dipilih, setDipilih] = useState<string[]>([]);
  const [terkirim, setTerkirim] = useState<string[]>([]);

  const semua = useMemo(() => (staff || [])
    .filter((s: any) => s && s.name && String(s.phone || "").trim())
    .sort((a: any, b: any) => String(a.name).localeCompare(String(b.name))), [staff]);
  const tampil = useMemo(() => semua.filter((s: any) =>
    (grup === "semua" || grupDari(s) === grup) && (!cari.trim() || String(s.name).toLowerCase().includes(cari.trim().toLowerCase()))), [semua, grup, cari]);
  const tanpaNomor = useMemo(() => (staff || []).filter((s: any) =>
    s && s.name && !String(s.phone || "").trim() && (grup === "semua" || grupDari(s) === grup)).length, [staff, grup]);

  if (role !== "admin") {
    return <Card><SH label="Pengumuman" /><div style={{ color: C.tL }}>Fitur ini hanya untuk admin.</div></Card>;
  }

  const butuhPassword = pesan.includes("{PASSWORD}");
  const isi = (nama: string): string => {
    let t = ganti(pesan, "{NAMA}", nama);
    t = ganti(t, "{WEBSITE}", website.trim());
    t = ganti(t, "{PASSWORD}", password);
    return t.trimEnd() + (otomatis ? FOOTER : "");
  };
  const pilihTemplate = (k: string) => {
    const t = TEMPLATE.find((x) => x.k === k);
    if (!t) return;
    if (pesan !== TEMPLATE.find((x) => x.k === tpl)?.teks && !window.confirm("Teks yang sudah Anda ubah akan diganti template. Lanjut?")) return;
    setTpl(k); setPesan(t.teks);
  };
  const antrean = semua.filter((s: any) => dipilih.includes(String(s.id)) && !terkirim.includes(String(s.id)));
  const kirim = (s: any) => {
    if (butuhPassword && !password.trim()) { showToast("⚠ Isi dulu kolom Password baru", C.w); return; }
    if (!pesan.trim()) { showToast("⚠ Pesan masih kosong", C.w); return; }
    window.open("https://wa.me/" + normalizePhone62(String(s.phone)) + "?text=" + encodeURIComponent(isi(String(s.name))), "_blank");
    setTerkirim((x) => (x.includes(String(s.id)) ? x : [...x, String(s.id)]));
  };
  const salin = async () => {
    if (butuhPassword && !password.trim()) { showToast("⚠ Isi dulu kolom Password baru", C.w); return; }
    try { await navigator.clipboard.writeText(isi("semuanya")); showToast("✓ Teks disalin (untuk grup)", C.s); }
    catch { showToast("Gagal menyalin teks", C.w); }
  };
  const toggle = (id: string) => setDipilih((x) => (x.includes(id) ? x.filter((y) => y !== id) : [...x, id]));
  const pilihSemuaTampil = () => setDipilih((x) => Array.from(new Set([...x, ...tampil.map((s: any) => String(s.id))])));

  return (
    <div style={{ paddingBottom: 24 }}>
      <Card>
        <SH label="Pengumuman ke staf" />
        <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 4 }}>Template</div>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 8 }}>
          {TEMPLATE.map((t) => (
            <button key={t.k} type="button" onClick={() => pilihTemplate(t.k)}
              style={{ padding: "8px 12px", borderRadius: 10, border: "2px solid #16A34A", background: tpl === t.k ? "#16A34A" : "#fff", color: tpl === t.k ? "#fff" : "#166534", fontSize: 14, fontWeight: 800, cursor: "pointer" }}>{t.l}</button>
          ))}
        </div>
        <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 2 }}>Isi pesan</div>
        <textarea value={pesan} onChange={(e) => setPesan(e.target.value)} rows={12}
          style={{ ...iS, width: "100%", boxSizing: "border-box", fontSize: 14, lineHeight: 1.4, resize: "vertical" }} />
        <div style={{ fontSize: 12, color: C.tL, marginTop: 2 }}>Kode otomatis: {"{NAMA}"} = nama penerima, {"{WEBSITE}"} = alamat web, {"{PASSWORD}"} = password baru.</div>
        <div style={{ fontWeight: 700, fontSize: 13, margin: "8px 0 2px" }}>Alamat website</div>
        <input style={iS} value={website} onChange={(e) => setWebsite(e.target.value)} />
        {butuhPassword && (
          <>
            <div style={{ fontWeight: 700, fontSize: 13, margin: "8px 0 2px" }}>Password baru (tidak disimpan di mana pun)</div>
            <input style={iS} value={password} onChange={(e) => setPassword(e.target.value)} placeholder="Ketik password baru" autoComplete="off" />
          </>
        )}
        <label style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 10, fontSize: 14 }}>
          <input type="checkbox" checked={otomatis} onChange={(e) => setOtomatis(e.target.checked)} />
          Tambah keterangan “pesan dikirim otomatis”
        </label>
        <div style={{ marginTop: 8, border: "1px dashed " + C.b, borderRadius: 10, padding: 8, background: "#F9FAFB", whiteSpace: "pre-wrap", fontSize: 13 }}>
          <div style={{ fontWeight: 700, marginBottom: 4 }}>Pratinjau</div>
          {isi("Nama Penerima")}
        </div>
        <div style={{ marginTop: 8 }}><Btn sm outline onClick={salin}>📋 Salin teks (untuk grup WA)</Btn></div>
      </Card>

      <Card style={{ marginTop: 12 }}>
        <SH label="Penerima" />
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 8 }}>
          {GRUP_LABEL.map(([k, l]) => (
            <button key={k} type="button" onClick={() => setGrup(k)}
              style={{ padding: "6px 12px", borderRadius: 10, border: "1.5px solid " + C.p, background: grup === k ? C.p : "#fff", color: grup === k ? "#fff" : C.p, fontSize: 13, fontWeight: 700, cursor: "pointer" }}>{l}</button>
          ))}
        </div>
        <input style={iS} value={cari} onChange={(e) => setCari(e.target.value)} placeholder="Cari nama..." />
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", margin: "8px 0" }}>
          <Btn sm outline onClick={pilihSemuaTampil}>Pilih semua yang tampil</Btn>
          <Btn sm outline onClick={() => setDipilih([])}>Kosongkan pilihan</Btn>
          {terkirim.length > 0 && <Btn sm outline onClick={() => setTerkirim([])}>↺ Reset tanda terkirim</Btn>}
        </div>
        {tanpaNomor > 0 && <div style={{ fontSize: 12, color: C.w, marginBottom: 6 }}>⚠ {tanpaNomor} orang di kelompok ini belum punya nomor WA di tab Staf, jadi tidak tampil.</div>}
        <div style={{ maxHeight: 320, overflowY: "auto" }}>
          {tampil.map((s: any) => {
            const id = String(s.id);
            return (
              <label key={id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "6px 0", borderBottom: "1px solid " + C.b, fontSize: 14 }}>
                <input type="checkbox" checked={dipilih.includes(id)} onChange={() => toggle(id)} />
                <span style={{ flex: 1 }}><b>{s.name}</b></span>
                {terkirim.includes(id) && <span style={{ color: "#166534", fontWeight: 800 }}>✓ terkirim</span>}
              </label>
            );
          })}
          {!tampil.length && <div style={{ color: C.tL, padding: "6px 0" }}>Tidak ada staf bernomor WA yang cocok.</div>}
        </div>
      </Card>

      <Card style={{ marginTop: 12 }}>
        <SH label="Kirim" />
        <div style={{ fontSize: 14, marginBottom: 8 }}>Dipilih: <b>{dipilih.length}</b> • Sudah terkirim: <b>{terkirim.filter((x) => dipilih.includes(x)).length}</b> • Sisa: <b>{antrean.length}</b></div>
        {antrean.length > 0
          ? <button type="button" onClick={() => kirim(antrean[0])}
              style={{ width: "100%", padding: "16px 12px", borderRadius: 12, border: "none", background: "#16A34A", color: "#fff", fontSize: 18, fontWeight: 800, cursor: "pointer", boxShadow: "0 2px 8px rgba(22,163,74,.5)" }}>
              📲 Kirim ke {String(antrean[0].name)} ({antrean.length} lagi)
            </button>
          : <div style={{ fontSize: 14, color: C.tL }}>{dipilih.length ? "✓ Semua penerima yang dipilih sudah dikirimi." : "Pilih penerima dulu di atas."}</div>}
        <div style={{ fontSize: 12, color: C.tL, marginTop: 8 }}>Setiap klik membuka WhatsApp dengan pesan yang sudah terisi. Tekan Kirim di WhatsApp, lalu kembali ke sini dan klik tombol hijau lagi untuk orang berikutnya.</div>
      </Card>
    </div>
  );
}
