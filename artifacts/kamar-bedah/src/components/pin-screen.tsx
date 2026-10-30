/* src/components/pin-screen.tsx
   Layar login PIN — mendukung setup PIN pertama kali, lockout percobaan
   gagal (persist ke sessionStorage), dan mode cepat "Input Lembur" tanpa
   akses penuh. Dipindah dari App.tsx (langkah 6 refactor — pisah View/
   layar besar, batch 1), TIDAK ADA perubahan logika, termasuk seluruh
   FIX AUDIT di dalamnya. */

import { useState, useEffect, useRef } from "react";
import { C, HOSPITAL, PIN_MAX_ATTEMPTS, PIN_LOCKOUT_MS, iS } from "../constants";
import { readPinLockState, writePinLockState, verifyPin } from "../utils";
import { AmbilLogMedis } from "../services/supabase";
import { LF, Btn, ensureCustomFontLink } from "./primitives";
import { AuroraBackground, MeteorsBackground, ShinyText, fireClickSpark } from "./effects";

export function PinScreen({onVerify,pinAdmin,pinPerawat,isFirstTime,lemburPegawai}: any) {
  useEffect(() => { ensureCustomFontLink(); }, []);
  const [pin,setPin]   = useState("");
  const [err,setErr]   = useState("");
  const [shake,setShk] = useState(false);
  const [checking,setChecking] = useState(false);
  const initialLock = readPinLockState();
  const [attempts,setAttempts]     = useState(initialLock.attempts);
  const [lockedUntil,setLockedUntil] = useState<number|null>(initialLock.lockedUntil);
  /* Sinkron otomatis ke sessionStorage setiap kali salah satu berubah —
     pakai useEffect (bukan wrapper setter) agar selalu menulis nilai TERBARU
     dari kedua state, tidak ada risiko closure basi menimpa nilai yang lain. */
  useEffect(()=>{ writePinLockState(attempts, lockedUntil); }, [attempts, lockedUntil]);
  // Force setup state (first time: PIN 0000)
  const [setupMode,setSetupMode] = useState(false);
  const [newAdmin,setNewAdmin]   = useState("");
  const [cfAdmin,setCfAdmin]     = useState("");
  const [newPerawat,setNewPerawat] = useState("");
  const [cfPerawat,setCfPerawat]   = useState("");
  const [setupErr,setSetupErr]   = useState("");
  const [quickLembur,setQuickLembur] = useState(false);
  /* UPDATE TAMPILAN AWAL: dua kartu pilihan (Login biasa / Input Lembur)
     ditampilkan dulu sebelum field PIN muncul. entryMode null = tampilkan
     kartu pilihan; "login" atau "lembur" = tampilkan form PIN sesuai mode
     yang dipilih. quickLembur otomatis disetel sesuai kartu yang diklik. */
  const [entryMode,setEntryMode] = useState<null|"login"|"lembur">(null);
  /* OPSI A: step alur lembur — "pilih_pegawai" dulu, baru "pin" */
  const [lemburStep,setLemburStep] = useState<"pilih_pegawai"|"pin">("pilih_pegawai");
  const [pegawaiDipilih,setPegawaiDipilih] = useState<string>("");
  const pinInputRef = useRef<HTMLInputElement>(null);

  const chooseEntry = (mode:"login"|"lembur") => {
    setEntryMode(mode);
    setQuickLembur(mode==="lembur");
    setPin(""); setErr("");
    if(mode==="login"){
      setTimeout(()=>pinInputRef.current?.focus(), 50);
    } else {
      setLemburStep("pilih_pegawai");
      /* default ke pegawai pertama kalau ada */
      setPegawaiDipilih((lemburPegawai||[])[0]?.id||"");
    }
  };
  const backToChoice = () => {
    setEntryMode(null); setPin(""); setErr("");
    setLemburStep("pilih_pegawai"); setPegawaiDipilih("");
  };
  const lanjutKePIN = () => {
    if(!pegawaiDipilih){ setErr("Pilih nama pegawai dulu"); return; }
    setErr(""); setLemburStep("pin");
    setTimeout(()=>pinInputRef.current?.focus(), 50);
  };

  /* FIX AUDIT #4: PIN sekarang bisa berbentuk hash ("salt:hash") atau
     legacy plaintext (data lama sebelum migrasi). verifyPin menangani
     keduanya secara transparan. */
  const check = async () => {
    if(checking) return;
    if(lockedUntil && Date.now() < lockedUntil){
      const sisaDetik = Math.ceil((lockedUntil - Date.now())/1000);
      setErr(`Terlalu banyak percobaan gagal. Coba lagi dalam ${sisaDetik} detik.`); setShk(true);
      setTimeout(()=>{setShk(false);},400);
      return;
    }
    /* FIX AUDIT #32 (MEDIUM — Hardcoded Bypass PIN "0000"): sebelumnya
       bypass ini HANYA dijaga oleh `isFirstTime` (dihitung dari ada/tidaknya
       PIN di Supabase). Kalau suatu saat pinFromCloud ter-null-kan lagi
       akibat bug lain (mis. gagal parse config), "0000" jadi backdoor
       de-facto ke mode setup admin. Sekarang ditambah flag terpisah di
       localStorage yang HANYA di-set sekali setelah setup benar-benar
       selesai (lihat doSetup) — begitu flag ini ada, bypass "0000" tidak
       akan aktif lagi apa pun nilai isFirstTime dari cloud. */
    const setupSudahSelesai = (()=>{ try { return localStorage.getItem("kb_setup_done_v1")==="1"; } catch { return false; } })();
    if(isFirstTime && !setupSudahSelesai && pin==="0000"){
      setSetupMode(true); setPin(""); setErr(""); return;
    }
    setChecking(true);
    try{
      if(await verifyPin(pin, pinAdmin)){
        setAttempts(0); setLockedUntil(null);
        AmbilLogMedis("KEAMANAN","Admin","LOGIN_SUKSES","Login Admin berhasil via PIN screen",{role:"admin", mode: quickLembur?"lembur_only":"penuh"});
        onVerify("admin", undefined, undefined, quickLembur, quickLembur?pegawaiDipilih:undefined); return;
      }
      if(await verifyPin(pin, pinPerawat)){
        setAttempts(0); setLockedUntil(null);
        AmbilLogMedis("KEAMANAN","Perawat","LOGIN_SUKSES","Login Perawat berhasil via PIN screen",{role:"perawat", mode: quickLembur?"lembur_only":"penuh"});
        onVerify("perawat", undefined, undefined, quickLembur, quickLembur?pegawaiDipilih:undefined); return;
      }
      const nextAttempts = attempts + 1;
      setAttempts(nextAttempts);
      AmbilLogMedis("KEAMANAN","Unknown","LOGIN_GAGAL",
        `PIN salah — percobaan ke-${nextAttempts}/${PIN_MAX_ATTEMPTS}`,
        { attempt: nextAttempts, locked: nextAttempts >= PIN_MAX_ATTEMPTS }
      );
      const willLock = nextAttempts >= PIN_MAX_ATTEMPTS;
      if(willLock){
        setLockedUntil(Date.now() + PIN_LOCKOUT_MS);
        setErr(`Terlalu banyak percobaan gagal. Coba lagi dalam ${Math.ceil(PIN_LOCKOUT_MS/1000)} detik.`);
      } else {
        setErr(`PIN salah. Coba lagi. (${nextAttempts}/${PIN_MAX_ATTEMPTS})`);
      }
      setShk(true);
      /* BUG FIX: sebelumnya setErr("") di sini SELALU dipanggil setelah 1.5 detik,
         termasuk saat baru saja masuk status terkunci (willLock=true) — akibatnya
         pesan "Coba lagi dalam 60 detik" hilang dari layar hanya 1.5 detik setelah
         muncul, padahal input masih terkunci selama sisa ~58.5 detik berikutnya.
         User jadi bingung kenapa tombol "🔒 Terkunci" tanpa keterangan apa pun.
         Sekarang: pesan lockout HANYA dihapus oleh efek countdown di bawah (yang
         memperbarui hitungan mundur tiap detik lalu membersihkannya begitu waktu
         kunci benar-benar habis) — bukan oleh timer 1.5 detik ini. */
      setTimeout(()=>{ setShk(false); setPin(""); if(!willLock) setErr(""); },1500);
    } finally {
      setChecking(false);
    }
  };
  /* Countdown live untuk pesan lockout: memperbarui sisa detik tiap 1 detik
     selama lockedUntil aktif, lalu otomatis membuka kunci & membersihkan
     pesan begitu waktunya habis — tanpa perlu user menekan tombol lagi. */
  useEffect(()=>{
    if(!lockedUntil) return;
    const iv = setInterval(()=>{
      const sisa = lockedUntil - Date.now();
      if(sisa<=0){
        setLockedUntil(null); setErr(""); setAttempts(0); clearInterval(iv);
      } else {
        setErr(`Terlalu banyak percobaan gagal. Coba lagi dalam ${Math.ceil(sisa/1000)} detik.`);
      }
    },1000);
    return ()=>clearInterval(iv);
  },[lockedUntil]);
  const doSetup = () => {
    if(newAdmin.length<4){setSetupErr("PIN Admin minimal 4 digit");return;}
    if(newAdmin!==cfAdmin){setSetupErr("PIN Admin tidak cocok");return;}
    if(newPerawat.length<4){setSetupErr("PIN Perawat minimal 4 digit");return;}
    if(newPerawat!==cfPerawat){setSetupErr("PIN Perawat tidak cocok");return;}
    if(newAdmin===newPerawat){setSetupErr("PIN Admin & Perawat tidak boleh sama");return;}
    /* FIX AUDIT #32: kunci permanen bypass "0000" begitu setup pertama selesai */
    try { localStorage.setItem("kb_setup_done_v1","1"); } catch {}
    onVerify("admin",newAdmin,newPerawat);
  };

  if(setupMode) return (
    <div className="kb-login-shell" style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"linear-gradient(135deg,#0B1220,#151F35 55%,#0B1220)",padding:16,position:"relative",overflow:"hidden"}}>
      <AuroraBackground/>
      <MeteorsBackground count={12}/>
      <style>{`
        @keyframes lgFloat{0%,100%{transform:translateY(0)}50%{transform:translateY(-6px)}}
        /* PENAMBAHAN: panel branding — HANYA tampil di layar lebar (>=1024px,
           laptop/komputer) supaya ruang kosong di kiri/kanan card login
           tidak terasa sia-sia, tanpa mengubah tampilan mobile sama sekali. */
        .kb-login-brand-panel{display:none;}
        @media (min-width:1024px){
          .kb-login-shell{gap:56px;justify-content:center!important;}
          .kb-login-brand-panel{display:flex;flex-direction:column;justify-content:center;color:#fff;max-width:420px;}
        }
      `}</style>
      <div className="kb-login-brand-panel" style={{position:"relative",zIndex:1}}>
        <div style={{fontSize:13,fontWeight:700,letterSpacing:3,textTransform:"uppercase",opacity:.75,marginBottom:12}}>RS Panti Rini</div>
        <div className="kb-font-display" style={{fontSize:38,fontWeight:800,lineHeight:1.15,marginBottom:16}}>Koordinasi Kamar Bedah, Lebih Cepat &amp; Akurat</div>
        <div style={{fontSize:15,opacity:.85,lineHeight:1.6}}>Jadwal operasi, tim medis, dan status kamar bedah — semua tersinkron real-time untuk seluruh tim perawat dan admin.</div>
      </div>
      <div style={{background:C.white,borderRadius:24,padding:"32px 28px",width:"100%",maxWidth:380,boxShadow:"0 20px 60px rgba(0,0,0,.3)",position:"relative",zIndex:1}}>
        <div style={{textAlign:"center",marginBottom:20}}>
          <img src="/logo.jpeg" alt="Logo" style={{width:72,height:72,borderRadius:"50%",objectFit:"cover",display:"block",margin:"0 auto 12px",animation:"lgFloat 3s ease-in-out infinite"}}/>
          <div style={{fontSize:16,fontWeight:800,color:C.p}}>Setup PIN Pertama Kali</div>
          <div style={{fontSize:12,color:C.tL,marginTop:4}}>Buat PIN Admin dan PIN Perawat baru</div>
        </div>
        <div style={{background:C.wBg,border:`1px solid ${C.w}`,borderRadius:10,padding:"10px 14px",marginBottom:16,fontSize:12,color:C.w,fontWeight:600}}>
          ⚠ Setelah setup, PIN <b>0000</b> tidak bisa digunakan lagi.
        </div>
        <div style={{fontSize:13,fontWeight:700,color:C.p,marginBottom:8}}>🔐 PIN Admin</div>
        <LF label="PIN Admin Baru (min. 4 digit)"><input style={iS} type="password" inputMode="numeric" maxLength={6} placeholder="PIN admin baru" value={newAdmin} onChange={e=>setNewAdmin(e.target.value)}/></LF>
        <LF label="Konfirmasi PIN Admin"><input style={iS} type="password" inputMode="numeric" maxLength={6} placeholder="Ulangi PIN admin" value={cfAdmin} onChange={e=>setCfAdmin(e.target.value)}/></LF>
        <div style={{fontSize:13,fontWeight:700,color:C.g,margin:"12px 0 8px"}}>👤 PIN Perawat</div>
        <LF label="PIN Perawat Baru (min. 4 digit)"><input style={iS} type="password" inputMode="numeric" maxLength={6} placeholder="PIN perawat baru" value={newPerawat} onChange={e=>setNewPerawat(e.target.value)}/></LF>
        <LF label="Konfirmasi PIN Perawat"><input style={iS} type="password" inputMode="numeric" maxLength={6} placeholder="Ulangi PIN perawat" value={cfPerawat} onChange={e=>setCfPerawat(e.target.value)}/></LF>
        {setupErr && <div style={{fontSize:12,color:C.d,marginBottom:10,fontWeight:600}}>⚠ {setupErr}</div>}
        <Btn full onClick={doSetup} style={{marginTop:8,padding:"13px",fontSize:15,background:C.p}}>✓ Simpan PIN & Masuk</Btn>
      </div>
    </div>
  );

  return (
    <div className="kb-login-shell" style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"linear-gradient(135deg,#0B1220,#151F35 55%,#0B1220)",padding:16,position:"relative",overflow:"hidden"}}>
      <AuroraBackground/>
      <style>{`@keyframes shk{0%,100%{transform:translateX(0)}25%,75%{transform:translateX(-7px)}50%{transform:translateX(7px)}} @keyframes lgFloat{0%,100%{transform:translateY(0) scale(1)}50%{transform:translateY(-8px) scale(1.06)}} @keyframes lgGlow{0%,100%{box-shadow:0 4px 18px rgba(22,104,95,.3),0 0 0 3px #16685F55}50%{box-shadow:0 12px 38px rgba(22,104,95,.65),0 0 0 4px #2D9A87aa}}
        .kb-login-brand-panel{display:none;}
        @media (min-width:1024px){
          .kb-login-shell{justify-content:center!important;}
          .kb-login-card{max-width:560px !important;padding:40px 44px !important;}
          .kb-login-card img{width:140px !important;height:140px !important;}
          .kb-login-card > div:first-child > div:not(.kb-font-display){font-size:15px !important;}
          .kb-login-card .kb-font-display{font-size:30px !important;}
          .kb-login-card .kb-login-judul{font-size:18px !important;}
          .kb-login-card .kb-login-pilih{padding:22px 22px !important;gap:20px !important;}
          .kb-login-card .kb-login-pilih > span:first-child{font-size:44px !important;}
          .kb-login-card .kb-login-pilih > span:last-child > div:first-child{font-size:22px !important;}
          .kb-login-card .kb-login-pilih > span:last-child > div:last-child{font-size:15px !important;}
          
        }
      `}</style>
      <div className="kb-login-card" style={{background:C.white,borderRadius:24,padding:"32px 28px",width:"100%",maxWidth:360,boxShadow:"0 20px 60px rgba(0,0,0,.3)",animation:shake?"shk .4s":"none",position:"relative",zIndex:1}}>
        <div style={{textAlign:"center",marginBottom:24}}>
          <img src="/logo.jpeg" alt="Logo Kamar Bedah" style={{width:104,height:104,borderRadius:"50%",objectFit:"cover",display:"block",margin:"0 auto 14px",animation:"lgFloat 3s ease-in-out infinite, lgGlow 3s ease-in-out infinite"}}/>
          <div style={{fontSize:11,fontWeight:700,color:C.g,letterSpacing:2,textTransform:"uppercase",marginBottom:4}}>RS Panti Rini</div>
          <div className="kb-font-display" style={{fontSize:19,fontWeight:700,color:C.navy,letterSpacing:.3,lineHeight:1.25}}><ShinyText style={{color:C.navy}}>SISTEM KOORDINASI<br/>KAMAR BEDAH</ShinyText></div>
          <div style={{fontSize:11,color:C.tL,marginTop:6}}>{HOSPITAL}</div>
        </div>

        {entryMode===null ? (
          /* UPDATE TAMPILAN AWAL: 2 kartu pilihan ditampilkan dulu sebelum PIN.
             Klik salah satu kartu akan menampilkan form PIN dengan mode yang
             sesuai (quickLembur otomatis disetel). Tidak ada perubahan pada
             logika verifikasi PIN/role itu sendiri — murni urutan tampilan. */
          <>
            <div className="kb-login-judul" style={{fontSize:13,fontWeight:600,color:C.t,textAlign:"center",marginBottom:16}}>Pilih Menu</div>
            <button
              type="button"
              className="kb-login-pilih" onClick={(e)=>{fireClickSpark(e); chooseEntry("login");}}
              style={{
                width:"100%",display:"flex",alignItems:"center",gap:14,textAlign:"left",
                background:C.pBg,border:`1.5px solid ${C.p}`,borderRadius:16,
                padding:"18px 16px",marginBottom:12,cursor:"pointer",fontFamily:"inherit"
              }}
            >
              <span style={{fontSize:28}}>🔐</span>
              <span>
                <div style={{fontSize:15,fontWeight:800,color:C.p}}>Masuk / Login</div>
                <div style={{fontSize:11,color:C.tL,marginTop:2}}>Akses penuh — Admin atau Perawat</div>
              </span>
            </button>
            <button
              type="button"
              className="kb-login-pilih" onClick={(e)=>{fireClickSpark(e); chooseEntry("lembur");}}
              style={{
                width:"100%",display:"flex",alignItems:"center",gap:14,textAlign:"left",
                background:"#FFF7E6",border:"1.5px solid #F2A93B",borderRadius:16,
                padding:"18px 16px",marginBottom:6,cursor:"pointer",fontFamily:"inherit"
              }}
            >
              <span style={{fontSize:28}}>⏰</span>
              <span>
                <div style={{fontSize:15,fontWeight:800,color:"#B8740A"}}>Input Lembur</div>
                <div style={{fontSize:11,color:C.tL,marginTop:2}}>Langsung ke Catat Lembur saja</div>
              </span>
            </button>
          </>
        ) : entryMode==="lembur" && lemburStep==="pilih_pegawai" ? (
          /* OPSI A STEP 1: pilih nama pegawai sebelum PIN ditampilkan */
          <>
            <button type="button" onClick={backToChoice} style={{display:"flex",alignItems:"center",gap:4,background:"none",border:"none",color:C.tL,fontSize:12,fontWeight:600,cursor:"pointer",fontFamily:"inherit",marginBottom:14,padding:0}}>
              ← Kembali ke Pilihan Menu
            </button>
            <div style={{fontSize:13,fontWeight:700,color:"#B8740A",textAlign:"center",marginBottom:6}}>⏰ Input Lembur</div>
            <div style={{fontSize:11,color:C.tL,textAlign:"center",marginBottom:14}}>Pilih nama kamu terlebih dahulu</div>
            {(lemburPegawai||[]).length===0 ? (
              <div style={{fontSize:12,color:C.d,background:"#FFF3CD",borderRadius:10,padding:"12px 14px",marginBottom:14,textAlign:"center"}}>
                ⚠ Belum ada data pegawai.<br/>Minta Admin untuk menambahkan data pegawai di tab Staf → Lembur terlebih dahulu.
              </div>
            ) : (
              <>
                <select
                  style={{...iS,marginBottom:12,fontSize:14,fontWeight:600,height:48}}
                  value={pegawaiDipilih}
                  onChange={e=>setPegawaiDipilih(e.target.value)}
                >
                  <option value="">— Pilih nama kamu —</option>
                  {(lemburPegawai||[]).map((p:any)=>(
                    <option key={p.id} value={p.id}>{p.name}{p.nik?` — ${p.nik}`:""}</option>
                  ))}
                </select>
                {err && <div style={{fontSize:12,color:C.d,textAlign:"center",marginBottom:10}}>⚠ {err}</div>}
                <Btn full onClick={lanjutKePIN} style={{padding:"13px",fontSize:15,marginBottom:8}}>Lanjut →</Btn>
              </>
            )}
          </>
        ) : (
          /* STEP 2 (lembur) atau mode login: form PIN */
          <>
            <button type="button" onClick={entryMode==="lembur"?()=>{setLemburStep("pilih_pegawai");setErr("");}:backToChoice} style={{display:"flex",alignItems:"center",gap:4,background:"none",border:"none",color:C.tL,fontSize:12,fontWeight:600,cursor:"pointer",fontFamily:"inherit",marginBottom:14,padding:0}}>
              ← {entryMode==="lembur"?"Kembali Pilih Nama":"Kembali ke Pilihan Menu"}
            </button>
            {entryMode==="lembur" && pegawaiDipilih && (
              <div style={{background:"#FFF7E6",border:"1px solid #F2A93B",borderRadius:10,padding:"8px 14px",marginBottom:12,fontSize:13,fontWeight:700,color:"#B8740A",textAlign:"center"}}>
                👤 {(lemburPegawai||[]).find((p:any)=>p.id===pegawaiDipilih)?.name||pegawaiDipilih}
              </div>
            )}
            <div style={{fontSize:13,fontWeight:700,color:entryMode==="lembur"?"#B8740A":C.p,textAlign:"center",marginBottom:6}}>
              {entryMode==="lembur" ? "⏰ Masukkan PIN" : "🔐 Masukkan PIN Akses"}
            </div>
            <div style={{fontSize:11,color:C.tL,textAlign:"center",marginBottom:10}}>Gunakan PIN Admin atau PIN Perawat</div>
            <input ref={pinInputRef} style={{...iS,textAlign:"center",fontSize:28,letterSpacing:12,height:56,marginBottom:12,borderRadius:12}} type="password" inputMode="numeric" maxLength={6} placeholder="••••" value={pin} disabled={checking||!!(lockedUntil&&Date.now()<lockedUntil)} onChange={e=>setPin(e.target.value)} onKeyDown={e=>e.key==="Enter"&&check()}/>
            {err && <div style={{fontSize:12,color:C.d,textAlign:"center",marginBottom:10}}>⚠ {err}</div>}
            <Btn full onClick={check} disabled={checking||!!(lockedUntil&&Date.now()<lockedUntil)} style={{marginBottom:12,padding:"13px",fontSize:15}}>{checking?"Memeriksa...":(lockedUntil&&Date.now()<lockedUntil)?"🔒 Terkunci":"Masuk"}</Btn>
            <div style={{display:"flex",justifyContent:"center"}}>
              <button onClick={()=>{}} style={{background:"none",border:"none",color:C.tL,fontSize:11,cursor:"pointer",fontFamily:"inherit"}}>Hubungi Admin jika lupa PIN</button>
            </div>
            <div style={{display:"flex",gap:6,marginTop:12,justifyContent:"center"}}>
              <span style={{background:C.pBg,color:C.p,fontSize:10,padding:"3px 10px",borderRadius:12,fontWeight:700}}>🔐 Admin</span>
              <span style={{background:C.gBg,color:C.g,fontSize:10,padding:"3px 10px",borderRadius:12,fontWeight:700}}>👤 Perawat</span>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
