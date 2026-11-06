/* Pembungkus KolomSaran: menambah tombol Simpan kecil di kanan isian kalau
   teks belum ada di daftar. Nama baru disimpan di kamar_bedah_config (key
   sendiri), daftar bawaan di kode tidak diubah. */
import { useState, useEffect, useMemo } from "react";
import type { CSSProperties } from "react";
import { KolomSaran } from "./kolom-saran";
import { SUPA_CLIENT } from "../services/supabase";

type Props = {
  kunci: string;
  daftar: string[];
  value?: string;
  onChange: (v: string) => void;
  style?: CSSProperties;
  placeholder?: string;
  onFocus?: (e: any) => void;
  onBlur?: (e: any) => void;
};

const rapi = (s: string): string => (s || "").replace(/\s+/g, " ").trim();
const norm = (s: string): string => rapi(s).toLowerCase();

function bacaDaftar(raw: any): string[] {
  try {
    const a = typeof raw === "string" ? JSON.parse(raw) : raw;
    return Array.isArray(a) ? a.filter((x: any) => typeof x === "string" && x.trim()) : [];
  } catch {
    return [];
  }
}

export function KolomSaranSimpan({ kunci, daftar, value, onChange, style, placeholder, onFocus, onBlur }: Props) {
  const [tambahan, setTambahan] = useState<string[]>([]);
  const [status, setStatus] = useState<"" | "proses" | "gagal">("");

  useEffect(() => {
    let batal = false;
    (async () => {
      try {
        const r: any = await SUPA_CLIENT.from("kamar_bedah_config").select("value").eq("key", kunci);
        const v = r && r.data && r.data[0] ? r.data[0].value : null;
        if (!batal && v) setTambahan(bacaDaftar(v));
      } catch {
        /* abaikan: daftar bawaan tetap jalan */
      }
    })();
    return () => { batal = true; };
  }, [kunci]);

  const gabung = useMemo<string[]>(
    () => [...daftar, ...tambahan.filter((t) => !daftar.some((d) => norm(d) === norm(t)))],
    [daftar, tambahan]
  );
  const teks = rapi(value || "");
  const sudahAda = useMemo(() => { const n = norm(teks); return gabung.some((d) => norm(d) === n); }, [gabung, teks]);
  const tampilTombol = teks.length >= 3 && teks.length <= 200 && !sudahAda;

  const simpan = async () => {
    if (!tampilTombol || status === "proses") return;
    setStatus("proses");
    try {
      const r: any = await SUPA_CLIENT.from("kamar_bedah_config").select("value").eq("key", kunci);
      if (r && r.error) throw r.error;
      const terbaru = r && r.data && r.data[0] ? bacaDaftar(r.data[0].value) : [];
      const n = norm(teks);
      const baru = terbaru.some((t) => norm(t) === n) ? terbaru : [...terbaru, teks];
      const { error } = await SUPA_CLIENT.from("kamar_bedah_config").upsert({ key: kunci, value: JSON.stringify(baru) }, { onConflict: "key" });
      if (error) throw error;
      setTambahan(baru);
      setStatus("");
    } catch {
      setStatus("gagal");
    }
  };

  return (
    <div style={{ display: "flex", alignItems: "stretch", gap: 6 }}>
      <div style={{ flex: 1, minWidth: 0 }}>
        <KolomSaran
          daftar={gabung}
          style={style}
          placeholder={placeholder}
          value={value}
          onChange={(v: string) => { if (status === "gagal") setStatus(""); onChange(v); }}
          onFocus={onFocus}
          onBlur={onBlur}
        />
      </div>
      {tampilTombol && (
        <button
          type="button"
          onClick={simpan}
          disabled={status === "proses"}
          style={{ flexShrink: 0, border: "none", borderRadius: 10, padding: "0 12px", background: status === "gagal" ? "#DC2626" : "#0369A1", color: "#fff", fontSize: 12, fontWeight: 700, cursor: "pointer", whiteSpace: "nowrap" }}
        >
          {status === "proses" ? "..." : status === "gagal" ? "Ulangi" : "Simpan"}
        </button>
      )}
    </div>
  );
}
