/* Isian teks dengan saran otomatis (autocomplete) dan tombol dropdown.
   Tetap bisa diketik bebas; daftar hanya membantu. */
import { useState, useMemo, useRef, useEffect } from "react";
import type { CSSProperties } from "react";

type Props = {
  value?: string;
  onChange: (v: string) => void;
  daftar: string[];
  style?: CSSProperties;
  placeholder?: string;
  onFocus?: (e: any) => void;
  onBlur?: (e: any) => void;
};

const norm = (s: string): string => s.toLowerCase().replace(/\s+/g, " ").trim();

function sorot(d: string, q: string) {
  if (!q) return d;
  const i = d.toLowerCase().indexOf(q);
  if (i < 0) return d;
  return (
    <span>
      {d.slice(0, i)}
      <b>{d.slice(i, i + q.length)}</b>
      {d.slice(i + q.length)}
    </span>
  );
}

export function KolomSaran({ value, onChange, daftar, style, placeholder, onFocus, onBlur }: Props) {
  const [buka, setBuka] = useState(false);
  const [semua, setSemua] = useState(false);
  const [aktif, setAktif] = useState(-1);
  const inputRef = useRef<HTMLInputElement>(null);
  const listRef = useRef<HTMLDivElement>(null);

  const q = norm(value || "");
  const saran = useMemo<string[]>(() => {
    if (semua) return daftar;
    if (!q) return [];
    const awal: string[] = [];
    const isi: string[] = [];
    for (const d of daftar) {
      const n = norm(d);
      if (n.startsWith(q)) awal.push(d);
      else if (n.includes(q)) isi.push(d);
    }
    return [...awal, ...isi];
  }, [q, daftar, semua]);

  const tepat = !semua && saran.length === 1 && norm(saran[0]) === q;
  const tampil = buka && saran.length > 0 && !tepat;

  useEffect(() => {
    if (!tampil || aktif < 0) return;
    const el = listRef.current?.children[aktif] as HTMLElement | undefined;
    if (el) el.scrollIntoView({ block: "nearest" });
  }, [aktif, tampil]);

  const tutup = () => { setBuka(false); setSemua(false); setAktif(-1); };
  const pilih = (d: string) => { onChange(d); tutup(); };

  return (
    <div style={{ position: "relative" }}>
      <input
        ref={inputRef}
        style={{ ...(style || {}), boxSizing: "border-box", paddingRight: 36 }}
        placeholder={placeholder}
        value={value ?? ""}
        autoComplete="off"
        spellCheck={false}
        role="combobox"
        aria-expanded={tampil}
        onChange={(ev) => { onChange(ev.target.value); setSemua(false); setBuka(true); setAktif(-1); }}
        onKeyDown={(ev) => {
          if (ev.key === "ArrowDown") {
            ev.preventDefault();
            if (!tampil) { setBuka(true); if (!q) setSemua(true); setAktif(0); }
            else setAktif((i) => Math.min(i + 1, saran.length - 1));
          } else if (ev.key === "ArrowUp") {
            if (tampil) { ev.preventDefault(); setAktif((i) => Math.max(i - 1, 0)); }
          } else if (ev.key === "Enter") {
            if (tampil && aktif >= 0 && saran[aktif]) { ev.preventDefault(); pilih(saran[aktif]); }
          } else if (ev.key === "Escape") {
            tutup();
          }
        }}
        onFocus={(e) => { if (onFocus) onFocus(e); }}
        onBlur={(e) => { tutup(); if (onBlur) onBlur(e); }}
      />
      <button
        type="button"
        tabIndex={-1}
        aria-label="Buka daftar"
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => {
          if (tampil) { tutup(); return; }
          if (inputRef.current) inputRef.current.focus();
          setSemua(true); setBuka(true); setAktif(-1);
        }}
        style={{ position: "absolute", right: 4, top: "50%", transform: "translateY(-50%)", width: 30, height: 30, border: "none", background: "transparent", cursor: "pointer", fontSize: 12, color: "#6B7280", padding: 0 }}
      >
        {tampil ? "▲" : "▼"}
      </button>
      {tampil && (
        <div
          ref={listRef}
          role="listbox"
          style={{ position: "absolute", left: 0, right: 0, top: "100%", marginTop: 4, background: "#fff", border: "1px solid #D1D5DB", borderRadius: 10, boxShadow: "0 8px 24px rgba(0,0,0,.15)", maxHeight: 220, overflowY: "auto", zIndex: 60 }}
        >
          {saran.map((d, i) => (
            <div
              key={d}
              role="option"
              aria-selected={i === aktif}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => pilih(d)}
              style={{ padding: "9px 12px", fontSize: 14, color: "#111827", cursor: "pointer", background: i === aktif ? "#EEF2FF" : "#fff", borderBottom: "1px solid #F3F4F6" }}
            >
              {sorot(d, q)}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
