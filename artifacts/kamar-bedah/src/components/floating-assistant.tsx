/* src/components/floating-assistant.tsx
   Widget chat "Asisten Virtual" melayang untuk pendaftaran pasien lewat
   teks natural. Reuse penuh runWorkerOcrScanText() (service OCR) &
   autoSaveOcrBatch() (tetap di App.tsx, dikirim lewat props) — pola
   "antrian ACC" yang sama dengan OCR foto/teks, TIDAK ada auto-commit
   tanpa review manusia. Dipindah dari App.tsx (langkah 5 refactor —
   pisah components, batch 4), TIDAK ADA perubahan logika, termasuk
   seluruh FIX BUG di dalamnya. */

import { useState, useEffect, useRef } from "react";
import type { OcrExtractResult, AssistantMsg, ShowToastFn } from "../types";
import { C, MAX_OCR_BATCH } from "../constants";
import { withTimeout } from "../utils";
import { runWorkerOcrScanText } from "../services/ocr";
import { AmbilLogMedis } from "../services/supabase";

export function FloatingAssistant({autoSaveOcrBatch, showToast}: {
  autoSaveOcrBatch: (results: OcrExtractResult[]) => void;
  showToast: ShowToastFn;
}) {
  const [open, setOpen] = useState(false);
  const [input, setInput] = useState("");
  const [busy, setBusy] = useState(false);
  const [msgs, setMsgs] = useState<AssistantMsg[]>([
    {role:"assistant", text:"Hai! Ceritakan data pasien yang mau didaftarkan (nama, umur, tindakan, tanggal/jam, dokter). Saya catat dulu — nanti kamu tinggal cek & ACC lewat badge biru di Tab Jadwal."}
  ]);
  const scrollRef = useRef<HTMLDivElement>(null);

  useEffect(()=>{
    scrollRef.current?.scrollTo({top: scrollRef.current.scrollHeight, behavior:"smooth"});
  },[msgs, open]);

  const send = async () => {
    const text = input.trim();
    if(!text || busy) return;
    setMsgs(m=>[...m,{role:"user",text}]);
    setInput("");
    setBusy(true);
    try {
      /* FIX BUG (ditemukan saat review): SEBELUMNYA runWorkerOcrScanText()
         dipanggil TANPA batas waktu sama sekali — beda dari Tab Daftar
         yang selalu lewat attemptWithRetry (timeout 60 detik + 1x retry
         otomatis kalau timeout). Kalau koneksi lambat/macet, widget ini
         bisa menggantung di "Sedang memproses…" TANPA BATAS, tidak seperti
         bagian app lain. Sekarang dibuatkan retry-on-timeout yang sama
         semangatnya (disederhanakan, tidak perlu state stage seperti
         attemptWithRetry karena UI-nya cuma satu baris status). */
      let res;
      try {
        res = await withTimeout(runWorkerOcrScanText(text), 60000, "Asisten pendaftaran");
      } catch(errFirst:any) {
        const isTimeout = typeof errFirst?.message === "string" && errFirst.message.includes("waktu tunggu habis");
        if(!isTimeout) throw errFirst;
        setMsgs(m=>[...m,{role:"assistant",text:"Koneksi lambat, mencoba sekali lagi…"}]);
        res = await withTimeout(runWorkerOcrScanText(text), 60000, "Asisten pendaftaran (percobaan ke-2)");
      }

      if(!res.ok || !res.data || res.data.length===0){
        /* FIX BUG: sebelumnya notAFormError tidak dibedakan — perawat yang
           kirim chat biasa (bukan data pasien, mis. salah ketik/nanya hal
           lain) akan dapat pesan generik yang membingungkan. Disamakan
           dengan pesan di handleOcrResponse (Tab Daftar). */
        const msg = res.notAFormError
          ? "Sepertinya itu bukan data pendaftaran pasien — pastikan sebutkan nama, umur, tindakan, dan dokter."
          : "Maaf, saya belum bisa memahami itu sebagai data pasien. " + (res.msg||"Coba tulis lebih lengkap ya — nama, umur, tindakan, dan dokter.");
        setMsgs(m=>[...m,{role:"assistant",text: msg}]);
        return;
      }

      const dataCapped = res.data.length > MAX_OCR_BATCH ? res.data.slice(0, MAX_OCR_BATCH) : res.data;
      autoSaveOcrBatch(dataCapped);
      /* FIX BUG: sebelumnya tidak ada audit log sama sekali dari jalur ini
         — beda dari OCR foto/paste-teks yang SELALU tercatat lewat
         AmbilLogMedis (penting untuk RS: siapa mendaftarkan pasien apa,
         kapan). Disamakan supaya trail-nya konsisten lintas semua jalur
         input, bukan cuma OCR. */
      AmbilLogMedis("DAFTAR_PASIEN","Perawat","ASISTEN_CHAT_DAFTAR",`Asisten chat — ${dataCapped.length} entri dicatat`,{jumlahEntri:dataCapped.length, sumber:"Asisten Virtual"});

      const n = dataCapped.length;
      const names = dataCapped.map(r=>r.patient||"—").join(", ");
      setMsgs(m=>[...m, {role:"assistant", text: `✓ ${n} data pasien (${names}) sudah saya catat & masuk antrian. Buka Tab Jadwal, cari badge biru "🔍 Perlu Review", lalu Edit untuk cek/lengkapi & konfirmasi.`}]);
      showToast(`✓ ${n} pendaftaran via asisten masuk antrian review`, C.p);
    } catch(e:any) {
      setMsgs(m=>[...m,{role:"assistant",text:"Terjadi kesalahan menghubungi server: " + (e?.message||"unknown error")}]);
    } finally {
      setBusy(false);
    }
  };

  return (
    <>
      {/* Bubble melayang — tombol buka/tutup */}
      <button onClick={()=>setOpen(o=>!o)} aria-label="Asisten Virtual Pendaftaran"
        style={{position:"fixed",right:18,bottom:18,width:56,height:56,borderRadius:"50%",
          background:"linear-gradient(135deg,#0284c7,#38bdf8)",border:"none",color:"#fff",
          fontSize:24,boxShadow:"0 6px 20px rgba(2,132,199,.4)",cursor:"pointer",zIndex:9500,
          display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"inherit"}}>
        {open ? "✕" : "🤖"}
      </button>

      {open && (
        <div style={{position:"fixed",right:18,bottom:84,
          width:"min(360px, calc(100vw - 32px))",height:"min(480px, calc(100vh - 140px))",
          background:C.white,borderRadius:18,boxShadow:"0 12px 40px rgba(0,0,0,.25)",zIndex:9500,
          display:"flex",flexDirection:"column",overflow:"hidden",border:"1px solid #e2e8f0"}}>

          <div style={{background:"linear-gradient(135deg,#0284c7,#1d4ed8)",color:"#fff",padding:"14px 16px",fontWeight:800,fontSize:14,display:"flex",alignItems:"center",gap:8,flexShrink:0}}>
            🤖 Asisten Pendaftaran
          </div>

          <div ref={scrollRef} style={{flex:1,overflowY:"auto",padding:"14px 12px",display:"flex",flexDirection:"column",gap:8,background:"#F8FAFC"}}>
            {msgs.map((m,i)=>(
              <div key={i} style={{alignSelf: m.role==="user"?"flex-end":"flex-start",
                maxWidth:"85%",background: m.role==="user"?"#0284c7":"#fff",
                color: m.role==="user"?"#fff":"#0f172a",
                border: m.role==="user"?"none":"1px solid #e2e8f0",
                borderRadius:14, padding:"9px 12px", fontSize:13, lineHeight:1.5,
                whiteSpace:"pre-wrap",
                boxShadow: m.role==="user"?"none":"0 1px 3px rgba(0,0,0,.06)"}}>
                {m.text}
              </div>
            ))}
            {busy && <div style={{alignSelf:"flex-start",fontSize:12,color:"#64748b",padding:"4px 8px"}}>Sedang memproses…</div>}
          </div>

          <div style={{padding:10,borderTop:"1px solid #e2e8f0",display:"flex",gap:8,background:"#fff",flexShrink:0}}>
            <textarea value={input} onChange={e=>setInput(e.target.value)}
              onKeyDown={e=>{ if(e.key==="Enter" && !e.shiftKey){ e.preventDefault(); send(); } }}
              placeholder="Contoh: Ny. Siti, 45th, appendektomi besok jam 9, dr. Andika"
              disabled={busy}
              style={{flex:1,resize:"none",height:40,padding:"9px 11px",borderRadius:10,border:"1.5px solid #e2e8f0",fontSize:13,fontFamily:"inherit",outline:"none"}}/>
            <button onClick={send} disabled={busy || !input.trim()}
              style={{width:40,height:40,borderRadius:10,border:"none",
                background: busy||!input.trim() ? "#cbd5e1":"#0284c7",
                color:"#fff",fontSize:16,cursor: busy||!input.trim() ? "not-allowed":"pointer",flexShrink:0,fontFamily:"inherit"}}>➤</button>
          </div>
        </div>
      )}
    </>
  );
}
