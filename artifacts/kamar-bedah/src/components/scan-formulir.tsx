/* src/components/scan-formulir.tsx
   Komponen scan formulir operasi pasien — 4 jalur input (kamera live,
   ambil foto, upload screenshot/galeri, paste teks), semua bermuara ke
   OCR yang sama (services/ocr.ts) lalu diteruskan ke callback onResults
   milik ViewDaftar. Dipindah dari App.tsx (langkah 6 refactor — pisah
   View/layar besar, batch 2), TIDAK ADA perubahan logika, termasuk
   seluruh FIX AUDIT & SARAN PENINGKATAN di dalamnya. */

import { useState, useEffect, useRef, useCallback } from "react";
import type { OcrExtractResult, ScanFormulirProps } from "../types";
import { C, MAX_OCR_FILE_MB, MAX_OCR_BATCH, MAX_OCR_PASTE_CHARS } from "../constants";
import { withTimeout, rotateImageFile } from "../utils";
import { runWorkerOcrScan, runWorkerOcrScanText } from "../services/ocr";
import { AmbilLogMedis } from "../services/supabase";
import { Btn, Card } from "./primitives";

/* PENAMBAHAN — Live Camera Scan: buka kamera live (getUserMedia) langsung
   di browser, tampilkan preview real-time, perawat arahkan ke formulir
   lalu tekan tombol "Ambil" saat sudah pas fokusnya — TIDAK auto-capture
   otomatis (sengaja): deteksi "gambar sudah fokus/stabil" itu kompleks dan
   rawan salah jepret di kondisi cahaya kamar bedah yang bervariasi: lebih
   aman & dapat diandalkan kalau perawat sendiri yang menentukan kapan
   gambar sudah pas, baru menekan tombol.

   Frame yang di-capture diubah jadi `File` lalu diteruskan ke
   `onCapture(file)` — pemanggil (ScanFormulir) meneruskannya ke
   `onFilePicked` yang SAMA PERSIS dipakai jalur "Ambil Foto"/"Upload
   Screenshot" — jadi TIDAK ADA duplikasi logika kompresi/preview/rotate/
   kirim-ke-OCR sama sekali, cuma sumber gambarnya saja yang berbeda. */
function LiveCameraScanModal({onCapture, onClose}: {onCapture: (file: File) => void; onClose: () => void}) {
  const videoRef = useRef<HTMLVideoElement>(null);
  const streamRef = useRef<MediaStream | null>(null);
  const [status, setStatus] = useState<"loading"|"ready"|"error">("loading");
  const [errMsg, setErrMsg] = useState("");

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        if (!navigator.mediaDevices?.getUserMedia) {
          throw new Error("Kamera live tidak didukung browser ini — gunakan tombol Ambil Foto/Upload biasa.");
        }
        const stream = await navigator.mediaDevices.getUserMedia({
          video: { facingMode: { ideal: "environment" } }, // kamera belakang untuk scan dokumen
          audio: false,
        });
        if (cancelled) { stream.getTracks().forEach(t=>t.stop()); return; }
        streamRef.current = stream;
        if (videoRef.current) {
          videoRef.current.srcObject = stream;
          await videoRef.current.play().catch(()=>{});
        }
        setStatus("ready");
      } catch (e: any) {
        if (cancelled) return;
        setErrMsg(e?.message || "Tidak bisa mengakses kamera — cek izin kamera di browser.");
        setStatus("error");
      }
    })();
    return () => {
      cancelled = true;
      streamRef.current?.getTracks().forEach(t => t.stop());
      streamRef.current = null;
    };
  }, []);

  const handleCapture = () => {
    const video = videoRef.current;
    if (!video || video.videoWidth === 0) return;
    const canvas = document.createElement("canvas");
    canvas.width = video.videoWidth;
    canvas.height = video.videoHeight;
    const ctx = canvas.getContext("2d");
    if (!ctx) return;
    ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
    canvas.toBlob((blob) => {
      if (!blob) return;
      const file = new File([blob], `live-scan-${Date.now()}.jpg`, { type: "image/jpeg" });
      onCapture(file);
    }, "image/jpeg", 0.92);
  };

  return (
    <div className="kb-modal-backdrop" style={{position:"fixed",inset:0,background:"rgba(15,23,42,.75)",zIndex:9500,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div className="kb-modal-content" style={{background:"#0B1220",borderRadius:20,padding:16,width:"100%",maxWidth:480,boxShadow:"0 20px 60px rgba(0,0,0,.5)"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
          <div style={{color:"#fff",fontWeight:800,fontSize:15}}>📷 Scan Live — Formulir Pasien</div>
          <button onClick={onClose} style={{background:"rgba(255,255,255,.12)",border:"none",borderRadius:8,color:"#fff",padding:"6px 10px",cursor:"pointer",fontSize:16}}>✕</button>
        </div>

        {status === "error" ? (
          <div style={{background:"#FEF2F2",borderRadius:12,padding:16,textAlign:"center"}}>
            <div style={{fontSize:13,color:"#B91C1C",fontWeight:600,marginBottom:10}}>⚠ {errMsg}</div>
            <Btn sm onClick={onClose}>Tutup, pakai cara biasa</Btn>
          </div>
        ) : (
          <>
            <div style={{position:"relative",borderRadius:14,overflow:"hidden",background:"#000",aspectRatio:"3/4"}}>
              <video ref={videoRef} playsInline muted style={{width:"100%",height:"100%",objectFit:"cover",display:"block"}}/>
              {/* Bingkai panduan — bantu perawat menyejajarkan formulir */}
              <div style={{position:"absolute",inset:"8% 6%",border:"2px dashed rgba(232,206,143,.85)",borderRadius:10,pointerEvents:"none"}}/>
              {status === "loading" && (
                <div style={{position:"absolute",inset:0,display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontSize:13,fontWeight:600}}>
                  ⏳ Membuka kamera…
                </div>
              )}
            </div>
            <div style={{fontSize:11,color:"rgba(255,255,255,.65)",textAlign:"center",margin:"10px 0"}}>
              Sejajarkan formulir dalam bingkai, pastikan cahaya cukup & tidak buram, lalu tekan Ambil.
            </div>
            <Btn full disabled={status!=="ready"} onClick={handleCapture} style={{background:"linear-gradient(135deg,#E8CE8F,#C9A961)",color:"#0B1220",fontWeight:800}}>
              📸 Ambil Gambar
            </Btn>
          </>
        )}
      </div>
    </div>
  );
}

export function ScanFormulir({onResults, showToast}: ScanFormulirProps) {
  const [loading, setLoading] = useState(false);
  const [showLiveCamera, setShowLiveCamera] = useState(false);
  const cameraInputRef  = useRef<HTMLInputElement>(null);
  const galleryInputRef = useRef<HTMLInputElement>(null);
  /* FIX AUDIT #38 (LOW — Double-Submit): guard sinkron via ref, pola yang
     sama seperti isSubmittingAbsenRef di modul lembur — `loading` (state
     React) saja tidak cukup karena ada celah singkat sebelum re-render di
     mana tap ganda cepat masih bisa lolos memicu 2 scan bersamaan. */
  const isProcessingRef = useRef(false);

  /* SARAN PENINGKATAN #9: tahapan proses yang lebih informatif daripada
     satu spinner generik — berguna terutama karena request ke Worker
     (upload + jalankan model vision/teks) bisa makan waktu beberapa detik
     di jaringan RS yang lambat, dan diam saja bikin user ragu apakah
     app-nya nge-hang atau memang masih memproses. */
  const [stage, setStage] = useState<""|"rotate"|"upload"|"retry">("");
  /* PENAMBAHAN #4 (indikator progres lebih informatif): hitung detik
     berjalan sejak stage "upload"/"retry" dimulai, supaya perawat punya
     kepastian sistem masih bekerja (bukan hang) — terutama di jaringan RS
     yang lambat di mana OCR bisa makan waktu >10 detik. */
  const [elapsedSec, setElapsedSec] = useState(0);
  useEffect(() => {
    if (stage !== "upload" && stage !== "retry") { setElapsedSec(0); return; }
    setElapsedSec(0);
    const t0 = Date.now();
    const iv = setInterval(() => setElapsedSec(Math.floor((Date.now() - t0) / 1000)), 1000);
    return () => clearInterval(iv);
  }, [stage]);
  const stageLabel =
    stage==="rotate" ? "Menyesuaikan orientasi gambar…" :
    stage==="upload" ? `Mengirim & membaca dengan AI… (${elapsedSec} detik)${elapsedSec>=8 ? " — jaringan lambat? biasanya tetap selesai <60 detik" : ""}` :
    stage==="retry"  ? `Koneksi lambat — mencoba sekali lagi… (${elapsedSec} detik)` : "";

  /* SARAN PENINGKATAN #7: preview + rotate SEBELUM gambar dikirim ke OCR.
     Sebelumnya foto langsung terkirim begitu dipilih — kalau miring/
     terbalik, satu-satunya cara memperbaiki adalah foto ulang dari nol.
     Sekarang ada jeda konfirmasi: user bisa memutar gambar dulu (kasus
     paling umum: foto formulir yang tidak tegak lurus) sebelum benar-benar
     mengirim ke AI, TANPA perlu mengulang proses ambil-foto/pilih-file. */
  const [previewFile, setPreviewFile] = useState<File|null>(null);
  const [previewUrl,  setPreviewUrl]  = useState<string|null>(null);
  const [rotation,    setRotation]    = useState(0);

  /* Jalur 2 (Paste Teks): kotak untuk menempel instruksi/memo teks (mis.
     copy-paste chat WhatsApp dokter) sebagai alternatif dari foto/scan.
     BUKAN pipeline terpisah — hasilnya lewat callback onResults yang
     SAMA dengan Jalur 1, sehingga otomatis melewati coerceOcrResult di
     runWorkerOcrScanText DAN buildOcrSafeApply di ViewDaftar (gerbang
     CYTO/gol.darah/alergi tetap berlaku, sama seperti hasil scan foto). */
  const [showTextPaste, setShowTextPaste] = useState(false);
  const [pasteText, setPasteText] = useState("");

  const clearPreview = useCallback(() => {
    setPreviewFile(null);
    setRotation(0);
    setPreviewUrl(prev => { if(prev) URL.revokeObjectURL(prev); return null; });
  },[]);
  // Cegah kebocoran object URL kalau komponen unmount selagi preview aktif.
  useEffect(() => () => { setPreviewUrl(prev => { if(prev) URL.revokeObjectURL(prev); return null; }); }, []);

  const onFilePicked = useCallback((file: File | undefined) => {
    if(!file) return;
    if(isProcessingRef.current) return;
    if(!file.type.startsWith("image/")){ showToast("File harus berupa gambar (foto/screenshot)",C.d); return; }
    if(file.size > MAX_OCR_FILE_MB*1024*1024){
      showToast(`Ukuran file terlalu besar (maks ${MAX_OCR_FILE_MB}MB) — coba screenshot atau kompres dulu.`, C.d);
      return;
    }
    clearPreview();
    setPreviewFile(file);
    setPreviewUrl(URL.createObjectURL(file));
  },[showToast, clearPreview]);

  type OcrResp = { ok: boolean; data?: OcrExtractResult[]; msg: string; notAFormError?: boolean };

  /* SARAN PENINGKATAN #4 (retry timeout) — diekstrak jadi helper generik
     supaya Jalur 1 (gambar) & Jalur 2 (teks) berbagi PERSIS logika yang
     sama, tidak ada risiko satu jalur diperbaiki lalu jalur lain
     tertinggal (kelas bug yang sama seperti FIX AUDIT #47 pada CYTO). */
  const attemptWithRetry = useCallback(async (attemptFn: () => Promise<OcrResp>, timeoutLabel: string): Promise<OcrResp> => {
    setStage("upload");
    try{
      return await withTimeout(attemptFn(), 60000, timeoutLabel);
    }catch(err:any){
      const isTimeout = typeof err?.message === "string" && err.message.includes("waktu tunggu habis");
      if(isTimeout){
        setStage("retry");
        try{
          return await withTimeout(attemptFn(), 60000, timeoutLabel);
        }catch(err2:any){
          return { ok:false, msg: "Gagal setelah 2x percobaan (koneksi lambat/putus): " + (err2?.message || "network error") };
        }
      }
      return { ok:false, msg: err?.message || "Gagal memproses" };
    }
  },[]);

  /* Penanganan hasil respons — SAMA untuk Jalur 1 & Jalur 2, supaya toast,
     cap MAX_OCR_BATCH, log audit, dan penanganan notAFormError konsisten
     apa pun sumber datanya (foto atau teks). */
  const handleOcrResponse = useCallback((res: OcrResp, sourceLabel: string, genericFailMsg: string) => {
    if(!res.ok || !res.data || res.data.length===0){
      /* SARAN PENINGKATAN #6: pesan yang lebih tepat sasaran kalau Worker
         (setelah diubah — lihat catatan di interface OcrExtractResult)
         mendeteksi input bukan formulir/instruksi operasi sama sekali. */
      if(res.notAFormError){
        showToast("📄 Sepertinya isi ini bukan formulir/instruksi operasi — pastikan datanya lengkap & jelas, lalu coba lagi.", C.d);
      } else {
        showToast(res.msg || genericFailMsg, C.d);
      }
      return;
    }
    const dataCapped = res.data.length > MAX_OCR_BATCH ? res.data.slice(0, MAX_OCR_BATCH) : res.data;
    /* FIX AUDIT #39: onResults() SENDIRI yang menampilkan toast akhir yang
       akurat (lihat applyOcrResult & autoSaveOcrBatch) — di sini sengaja
       TIDAK ada toast "✓ berhasil" generik yang bisa menimpa & menyesatkan. */
    onResults(dataCapped);
    AmbilLogMedis("DAFTAR_PASIEN","Perawat","OCR_SCAN_FORMULIR",`${sourceLabel} berhasil — ${dataCapped.length} entri terbaca (Cloudflare Workers AI)`,{jumlahEntri:dataCapped.length, sumber:sourceLabel});
    const totalTerbaca = res.data.length;
    if(totalTerbaca > MAX_OCR_BATCH){
      setTimeout(()=>{
        showToast(`⚠ Sebelumnya terbaca ${totalTerbaca} entri — hanya ${MAX_OCR_BATCH} pertama yang diproses. Proses ulang per bagian untuk sisanya.`, C.w);
      }, 3400);
    }
  },[onResults, showToast]);

  const processConfirmedFile = useCallback(async (file: File, deg: number) => {
    isProcessingRef.current = true;
    setLoading(true);
    try{
      let toSend = file;
      if(deg % 360 !== 0){
        setStage("rotate");
        try{ toSend = await rotateImageFile(file, deg); }
        catch{ /* Kalau rotasi gagal (mis. gambar korup), tetap lanjut pakai
                   file asli apa adanya — jangan sampai fitur rotate yang
                   sifatnya bantu-bantu malah memblokir seluruh alur scan. */ }
      }
      // Dikirim ke route /api/ocr-scan di Worker yang sama (Cloudflare
      // Workers AI, model vision) — gratis s.d. 10.000 Neuron/hari,
      // tidak ada API key eksternal, tidak ada Supabase Edge Function.
      const res = await attemptWithRetry(() => runWorkerOcrScan(toSend), "OCR scan formulir (gambar)");
      handleOcrResponse(res, "Scan formulir (foto)", "Tidak ada data yang berhasil dibaca dari gambar");
    }catch(e:any){
      showToast("Gagal memproses gambar: "+(e?.message||"unknown error"), C.d);
    }finally{
      isProcessingRef.current = false;
      setLoading(false);
      setStage("");
      clearPreview();
      if(cameraInputRef.current) cameraInputRef.current.value = "";
      if(galleryInputRef.current) galleryInputRef.current.value = "";
    }
  },[attemptWithRetry, handleOcrResponse, showToast, clearPreview]);

  const processPastedText = useCallback(async () => {
    if(isProcessingRef.current) return;
    const trimmed = pasteText.trim();
    if(!trimmed){ showToast("Kotak teks masih kosong — tempel dulu instruksi/memo-nya.", C.d); return; }
    isProcessingRef.current = true;
    setLoading(true);
    try{
      const res = await attemptWithRetry(() => runWorkerOcrScanText(trimmed), "OCR proses teks");
      handleOcrResponse(res, "Proses teks (paste)", "Tidak ada data yang berhasil diekstrak dari teks");
    }catch(e:any){
      showToast("Gagal memproses teks: "+(e?.message||"unknown error"), C.d);
    }finally{
      isProcessingRef.current = false;
      setLoading(false);
      setStage("");
      setPasteText("");
      setShowTextPaste(false);
    }
  },[pasteText, attemptWithRetry, handleOcrResponse, showToast]);

  return (
    <div style={{marginBottom:14}}>
      <input ref={cameraInputRef} type="file" accept="image/*" capture="environment" style={{display:"none"}} onChange={(ev)=>{ onFilePicked(ev.target.files?.[0]); ev.target.value=""; }}/>
      <input ref={galleryInputRef} type="file" accept="image/*" style={{display:"none"}} onChange={(ev)=>{ onFilePicked(ev.target.files?.[0]); ev.target.value=""; }}/>
      {previewFile && previewUrl ? (
        <Card style={{marginBottom:10, background:"#FAFBFD"}}>
          <div style={{fontSize:12,fontWeight:700,color:C.t,marginBottom:8}}>🔍 Periksa foto sebelum dibaca AI</div>
          <div style={{display:"flex",justifyContent:"center",alignItems:"center",background:"#111",borderRadius:10,overflow:"hidden",marginBottom:10,minHeight:180,maxHeight:320}}>
            <img src={previewUrl} alt="Preview formulir" style={{maxWidth:"100%",maxHeight:320,objectFit:"contain",transform:`rotate(${rotation}deg)`,transition:"transform .2s"}}/>
          </div>
          {!loading && (
            <div style={{display:"flex",gap:8,marginBottom:8}}>
              <Btn outline sm onClick={()=>setRotation(r=>(r+270)%360)} style={{flex:1}}>↺ Putar Kiri</Btn>
              <Btn outline sm onClick={()=>setRotation(r=>(r+90)%360)} style={{flex:1}}>↻ Putar Kanan</Btn>
            </div>
          )}
          {loading ? (
            <div style={{fontSize:12,color:C.tL,textAlign:"center",padding:"10px 0",fontWeight:600}}>⏳ {stageLabel || "Memproses…"}</div>
          ) : (
            <div style={{display:"flex",gap:8}}>
              <Btn outline color={C.g} onClick={clearPreview} style={{flex:1}}>🔄 Ambil Ulang</Btn>
              <Btn onClick={()=>processConfirmedFile(previewFile, rotation)} style={{flex:2}}>✓ Gunakan Foto Ini, Baca dengan AI</Btn>
            </div>
          )}
        </Card>
      ) : showTextPaste ? (
        <Card style={{marginBottom:10, background:"#FAFBFD"}}>
          <div style={{fontSize:12,fontWeight:700,color:C.t,marginBottom:8}}>📝 Jalur 2: Paste Instruksi/Memo (mis. chat WA dokter)</div>
          <textarea
            value={pasteText}
            onChange={(ev)=>setPasteText(ev.target.value.slice(0, MAX_OCR_PASTE_CHARS))}
            disabled={loading}
            placeholder={`Contoh: "Tolong jadwalkan operasi cito malam ini jam 22.00 untuk pasien Tn. Sitorus No RM 998811, diagnosis Apendiks Perforasi. Operator dr. Setiawan SpB."`}
            style={{width:"100%",minHeight:110,padding:10,borderRadius:8,border:`1px solid ${C.b}`,fontSize:13,resize:"vertical",fontFamily:"inherit",boxSizing:"border-box",marginBottom:6}}
          />
          <div style={{fontSize:11,color:C.tL,marginBottom:8,textAlign:"right"}}>{pasteText.length}/{MAX_OCR_PASTE_CHARS}</div>
          {loading ? (
            <div style={{fontSize:12,color:C.tL,textAlign:"center",padding:"10px 0",fontWeight:600}}>⏳ {stageLabel || "Memproses…"}</div>
          ) : (
            <div style={{display:"flex",gap:8}}>
              <Btn outline color={C.g} onClick={()=>{setShowTextPaste(false); setPasteText("");}} style={{flex:1}}>Batal</Btn>
              <Btn onClick={processPastedText} style={{flex:2}}>⚙️ Proses Teks dengan AI</Btn>
            </div>
          )}
        </Card>
      ) : (
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <Btn outline disabled={loading} onClick={()=>setShowLiveCamera(true)} style={{background:"linear-gradient(135deg,#E8CE8F22,transparent)",borderColor:"#C9A961"}}>🎥 Scan Live (Kamera)</Btn>
          <Btn outline disabled={loading} onClick={()=>cameraInputRef.current?.click()}>📷 Ambil Foto</Btn>
          <Btn outline disabled={loading} onClick={()=>galleryInputRef.current?.click()}>🖼️ Upload Screenshot/Galeri</Btn>
          <Btn outline disabled={loading} onClick={()=>setShowTextPaste(true)}>📝 Paste Teks Chat/Memo</Btn>
        </div>
      )}
      {showLiveCamera && (
        <LiveCameraScanModal
          onCapture={(file)=>{ setShowLiveCamera(false); onFilePicked(file); }}
          onClose={()=>setShowLiveCamera(false)}
        />
      )}
    </div>
  );
}
