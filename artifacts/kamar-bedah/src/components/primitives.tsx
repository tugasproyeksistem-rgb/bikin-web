/* src/components/primitives.tsx
   Primitif UI dasar yang dipakai di hampir seluruh tab aplikasi: elemen
   form (LF/TF/TA/SF), tombol (Btn/WaBt), Card, skeleton loading, animasi
   angka/scroll, header kecil (Row/Bdg/SH), Toggle, MiniBar chart, dan
   ErrorBoundary. Dipindah dari App.tsx (langkah 5 refactor — pisah
   components, batch 2: primitives), TIDAK ADA perubahan logika, termasuk
   seluruh FIX AUDIT di dalamnya. */

import { useState, useEffect, useRef, Component, Fragment } from "react";
import { C, iS } from "../constants";
import { TwinklingStars, RotatingText } from "./effects";

/* ─── UI PRIMITIVES ─────────────────────────────────────────────────── */

export function LF({label,req,err,children}: {label?: string; req?: boolean; err?: string; children: React.ReactNode}) {
  return (
    <div style={{marginBottom:14}}>
      {label && <div style={{fontSize:12,fontWeight:700,color:C.g,marginBottom:5,textTransform:"uppercase",letterSpacing:.4}}>
        {label}{req&&<span style={{color:C.d}}> ✱</span>}
      </div>}
      {children}
      {err && <div style={{fontSize:11,color:C.d,marginTop:4}}>⚠ {err}</div>}
    </div>
  );
}
export function TF({label,req,err,foc,onFoc,onBlr,...p}: any) {
  return (
    <LF label={label} req={req} err={err}>
      <input style={{...iS, borderColor:err?C.d:foc?C.p:C.b, boxShadow:foc?`0 0 0 3px ${C.p}22`:err?`0 0 0 3px ${C.d}22`:"none"} as React.CSSProperties}
        onFocus={onFoc} onBlur={onBlr} autoComplete="off" spellCheck={false} {...p}/>
    </LF>
  );
}
export function TA({label,minH=72,...p}: any) {
  const [f,sf] = useState(false);
  return (
    <LF label={label}>
      <textarea style={{...iS,resize:"vertical",minHeight:minH,lineHeight:1.7, borderColor:f?C.p:C.b, boxShadow:f?`0 0 0 3px ${C.p}22`:""} as React.CSSProperties}
        onFocus={()=>sf(true)} onBlur={()=>sf(false)} spellCheck={false} {...p}/>
    </LF>
  );
}
export function SF({label,req,err,options,empty="-- Pilih --",...p}: any) {
  const [f,sf] = useState(false);
  return (
    <LF label={label} req={req} err={err}>
      <div style={{position:"relative"}}>
        <select style={{...iS,appearance:"none",cursor:"pointer",paddingRight:32, borderColor:err?C.d:f?C.p:C.b, boxShadow:f?`0 0 0 3px ${C.p}22`:""} as React.CSSProperties}
          onFocus={()=>sf(true)} onBlur={()=>sf(false)} {...p}>
          <option value="">{empty}</option>
          {options.map((o: any,i: number)=><option key={i} value={typeof o==="string"?o:o.v}>{typeof o==="string"?o:o.l}</option>)}
        </select>
        <div style={{position:"absolute",right:12,top:"50%",transform:"translateY(-50%)",pointerEvents:"none",color:C.g,fontSize:11}}>▼</div>
      </div>
    </LF>
  );
}
export const Btn = ({children,onClick,color=C.p,outline,sm,full,disabled,style={}}: any) => {
  const isGrad = typeof color==="string" && color.startsWith("linear-gradient");
  const borderCol = isGrad ? "rgba(30,58,138,.4)" : (disabled?"#ccc":color);
  const bgStyle = outline ? C.white : disabled ? "#eee" : color;
  return (
    <button onClick={onClick} disabled={disabled} style={{padding:sm?"7px 12px":"10px 18px",borderRadius:10,border:`1.5px solid ${borderCol}`,background:bgStyle,color:outline?color:C.white,fontSize:sm?12:14,fontWeight:700,cursor:disabled?"not-allowed":"pointer",opacity:disabled?.45:1,width:full?"100%":"auto",fontFamily:"inherit",transition:"opacity .15s",boxShadow:isGrad&&!disabled?"0 2px 8px rgba(30,58,138,.25)":"none",...style}}>{children}</button>
  );
};
export const WaBt = ({children,onClick,sm,full,disabled,style={}}: any) => (
  <button onClick={onClick} disabled={disabled} style={{padding:sm?"7px 12px":"10px 18px",borderRadius:10,border:`1.5px solid ${disabled?"#ccc":C.wa}`,background:disabled?"#eee":C.wa,color:C.white,fontSize:sm?12:14,fontWeight:700,cursor:disabled?"not-allowed":"pointer",width:full?"100%":"auto",fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"center",gap:5,...style}}>✓ {children}</button>
);
export const Card = ({children,style={},hi,className}: any) => {
  const isCytoGlow = typeof className === "string" && className.includes("kb-cyto-glow");
  return (
    <div className={`kbCardHoverable${className?` ${className}`:""}`} style={{
      ...(isCytoGlow ? {} : {background:"#ffffff",border:`1px solid ${hi||C.b}`}),
      borderRadius:16,padding:"18px 20px",marginBottom:12,
      boxShadow:`6px 6px 16px rgba(148,163,184,.16), -4px -4px 12px rgba(255,255,255,.7), 0 1px 2px rgba(0,0,0,.03)`,
      ...style
    }}>{children}</div>
  );
};

/* PENAMBAHAN #3 (skeleton loading): pengganti visual untuk kartu jadwal
   selagi data AWAL masih dimuat dari Supabase, supaya perawat melihat
   "sedang memuat" (bentuk kartu samar berkedip) alih-alih layar
   kosong/pesan "Tidak ada jadwal" yang keliru — dan alih-alih spinner
   generik yang tidak memberi gambaran struktur konten yang akan muncul. */
const skeletonShimmerCss = `
@keyframes kbSkeletonShimmer {
  0% { background-position: -200px 0; }
  100% { background-position: calc(200px + 100%) 0; }
}
.kb-skeleton-bar {
  background: linear-gradient(90deg, #EEF1F4 25%, #E2E8F0 37%, #EEF1F4 63%);
  background-size: 400px 100%;
  animation: kbSkeletonShimmer 1.4s ease-in-out infinite;
  border-radius: 6px;
}`;
let _kbSkeletonCssInjected = false;
export function ensureSkeletonCss() {
  if (_kbSkeletonCssInjected || typeof document === "undefined") return;
  const el = document.createElement("style");
  el.id = "kb-skeleton-css";
  el.textContent = skeletonShimmerCss;
  document.head.appendChild(el);
  _kbSkeletonCssInjected = true;
}

/* PENAMBAHAN — kesan "mewah": load Google Font (Manrope + Playfair
   Display) via <link> di <head>, BUKAN via @import di dalam CSS string
   (yang akan diabaikan browser kalau bukan aturan pertama — lihat catatan
   di blok CSS global). <link> juga lebih andal: browser bisa mem-preload
   dan cache secara normal seperti resource lain. */
let _kbFontLinkInjected = false;
export function ensureCustomFontLink() {
  if (_kbFontLinkInjected || typeof document === "undefined") return;
  if (document.getElementById("kb-custom-font-link")) { _kbFontLinkInjected = true; return; }
  const preconnect1 = document.createElement("link");
  preconnect1.rel = "preconnect"; preconnect1.href = "https://fonts.googleapis.com";
  const preconnect2 = document.createElement("link");
  preconnect2.rel = "preconnect"; preconnect2.href = "https://fonts.gstatic.com"; preconnect2.crossOrigin = "anonymous";
  const link = document.createElement("link");
  link.id = "kb-custom-font-link";
  link.rel = "stylesheet";
  link.href = "https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&family=Playfair+Display:wght@600;700;800&display=swap";
  document.head.appendChild(preconnect1);
  document.head.appendChild(preconnect2);
  document.head.appendChild(link);
  _kbFontLinkInjected = true;
}
export const SkeletonCard = () => {
  ensureSkeletonCss();
  return (
    <Card>
      <div style={{display:"flex",justifyContent:"space-between",marginBottom:10}}>
        <div className="kb-skeleton-bar" style={{width:70,height:18}}/>
        <div className="kb-skeleton-bar" style={{width:50,height:14}}/>
      </div>
      <div className="kb-skeleton-bar" style={{width:"60%",height:16,marginBottom:8}}/>
      <div className="kb-skeleton-bar" style={{width:"85%",height:13,marginBottom:6}}/>
      <div className="kb-skeleton-bar" style={{width:"40%",height:13}}/>
    </Card>
  );
};
/** Render N SkeletonCard sekaligus — dipakai saat initialLoading true. */
export const SkeletonList = ({count=4}: {count?: number}) => (
  <>{Array.from({length: count}).map((_,i)=><SkeletonCard key={`skel-${i}`}/>)}</>
);

/* PENAMBAHAN #4 (number counter animation): angka menghitung naik dari 0
   ke nilai final saat pertama muncul/berubah — efek umum di dashboard
   SaaS modern (Stripe, Linear, dll), bukan sekadar tampil statis. Durasi
   tetap singkat (~600ms) supaya tidak terasa lambat/mengganggu. */
export const AnimatedNumber = ({value, duration=600}: {value:number; duration?:number}) => {
  const [display, setDisplay] = useState(0);
  const prevRef = useRef(0);
  useEffect(() => {
    const from = prevRef.current;
    const to = value;
    if (from === to) { setDisplay(to); return; }
    const t0 = performance.now();
    let raf = 0;
    const step = (now: number) => {
      const p = Math.min(1, (now - t0) / duration);
      // ease-out cubic — cepat di awal, melambat mendekati nilai akhir
      const eased = 1 - Math.pow(1 - p, 3);
      setDisplay(Math.round(from + (to - from) * eased));
      if (p < 1) raf = requestAnimationFrame(step);
      else prevRef.current = to;
    };
    raf = requestAnimationFrame(step);
    return () => cancelAnimationFrame(raf);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [value]);
  return <>{display}</>;
};

/* PENAMBAHAN #1 (parallax scroll halus): elemen dekoratif di dalam banner
   bergerak sedikit lebih lambat dari kecepatan scroll normal, memberi
   kesan kedalaman ringan. Listener di-cleanup otomatis saat unmount, dan
   pakai requestAnimationFrame throttle supaya tidak membebani scroll. */
export const HomeParallaxBanner = () => {
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    let raf = 0;
    const onScroll = () => {
      if (raf) return;
      raf = requestAnimationFrame(() => {
        raf = 0;
        const el = ref.current;
        if (!el) return;
        const scrollY = window.scrollY || document.documentElement.scrollTop || 0;
        // Gerak halus & terbatas (maks ~14px) — sekadar aksen kedalaman,
        // bukan efek besar yang bisa mengganggu keterbacaan teks.
        const offset = Math.min(14, scrollY * 0.15);
        el.style.transform = `translateY(${offset}px)`;
      });
    };
    window.addEventListener("scroll", onScroll, { passive: true });
    return () => { window.removeEventListener("scroll", onScroll); if (raf) cancelAnimationFrame(raf); };
  }, []);
  return (
    <div style={{background:"linear-gradient(135deg,#1E1B4B,#2E1065)",borderRadius:16,padding:"28px 24px",marginBottom:20,color:"#fff",boxShadow:"0 4px 20px rgba(30,27,75,.35)",overflow:"hidden",position:"relative",border:"1px solid rgba(232,206,143,.2)"}}>
      <TwinklingStars count={25}/>
      {/* Elemen dekoratif yang bergerak parallax */}
      <div ref={ref} style={{position:"absolute",top:-30,right:-30,width:140,height:140,borderRadius:"50%",background:"rgba(255,255,255,.08)",willChange:"transform"}}/>
      <div style={{position:"relative"}}>
        <div style={{fontSize:11,fontWeight:700,letterSpacing:2,color:"#E8CE8F",textTransform:"uppercase",marginBottom:6}}>Selamat Datang</div>
        <div className="kb-font-display" style={{fontSize:22,fontWeight:700,lineHeight:1.2,marginBottom:6}}>Sistem Koordinasi<br/>Kamar Bedah — <RotatingText words={["Cepat","Akurat","Terkoordinasi"]} style={{color:"#E8CE8F"}}/></div>
        <div style={{fontSize:13,opacity:.85}}>RS Panti Rini · {new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"})}</div>
      </div>
    </div>
  );
};

/* PENAMBAHAN #5 (scroll-reveal): fade-in + slide dari bawah saat elemen
   pertama masuk viewport — pakai IntersectionObserver dengan `once:true`
   (berhenti mengamati setelah sekali reveal, overhead minimal). SENGAJA
   diterapkan TERBATAS (Tab Home saja) — bukan ke list panjang seperti Tab
   Jadwal, karena banyak observer sekaligus pada list besar bisa berisiko
   ke performa scroll di HP kelas bawah. */
export const ScrollReveal = ({children, delay=0}: {children: React.ReactNode; delay?: number}) => {
  const ref = useRef<HTMLDivElement>(null);
  const [visible, setVisible] = useState(false);
  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    const obs = new IntersectionObserver((entries) => {
      if (entries[0]?.isIntersecting) { setVisible(true); obs.disconnect(); }
    }, { threshold: 0.15 });
    obs.observe(el);
    return () => obs.disconnect();
  }, []);
  return (
    <div ref={ref} style={{
      opacity: visible?1:0,
      transform: visible?"translateY(0)":"translateY(14px)",
      filter: visible?"blur(0)":"blur(4px)",
      transition: `opacity .5s ease ${delay}ms, transform .5s ease ${delay}ms, filter .5s ease ${delay}ms`,
    }}>{children}</div>
  );
};

export const Row  = ({title,right}: any) => (
  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
    <div style={{fontSize:16,fontWeight:800,color:C.t}}>{title}</div>
    {right}
  </div>
);
export const Bdg  = ({label,color,bg,icon}: any) => <span style={{background:`linear-gradient(135deg, ${bg}, ${bg}CC)`,color,fontSize:10,fontWeight:700,padding:"3px 9px",borderRadius:20,border:`1px solid ${color}33`,whiteSpace:"nowrap",display:"inline-flex",alignItems:"center",gap:4,boxShadow:`0 1px 3px ${color}22`}}>{icon && <span style={{fontSize:9}}>{icon}</span>}{label}</span>;
export const SH   = ({label,color=C.p}: any) => <div style={{fontSize:12,fontWeight:800,color,letterSpacing:.6,marginBottom:14,paddingBottom:8,borderBottom:`2px solid ${color}22`,textTransform:"uppercase"}}>{label}</div>;
export const Toggle = ({value,onChange,label,sub,color=C.p}: any) => (
  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 0",borderBottom:`1px solid ${C.b}`}}>
    <div><div style={{fontSize:14,fontWeight:700,color:C.t}}>{label}</div>{sub&&<div style={{fontSize:11,color:C.tL,marginTop:2}}>{sub}</div>}</div>
    <button onClick={()=>onChange(!value)} style={{width:50,height:28,borderRadius:14,background:value?color:"#DDE3EA",border:"none",cursor:"pointer",position:"relative",transition:"background .2s",flexShrink:0}}>
      <div style={{position:"absolute",top:4,left:value?24:4,width:20,height:20,borderRadius:10,background:"#fff",boxShadow:"0 1px 4px rgba(0,0,0,.25)",transition:"left .2s"}}/>
    </button>
  </div>
);

/* ─── MINI BAR CHART ─────────────────────────────────────────────────── */
export function MiniBar({data, colorFn, maxVal}: {data: {label: string; value: number; color?: string}[]; colorFn?: (i: number) => string; maxVal?: number}) {
  const max = maxVal ?? Math.max(...data.map(d=>d.value), 1);
  return (
    <div style={{display:"flex",flexDirection:"column",gap:6}}>
      {data.map((d,i)=>(
        <div key={i} style={{display:"flex",alignItems:"center",gap:8}}>
          <div style={{width:110,fontSize:11,color:C.tL,textAlign:"right",flexShrink:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{d.label}</div>
          <div style={{flex:1,background:"#F0F4F8",borderRadius:6,height:20,position:"relative",overflow:"hidden"}}>
            <div style={{position:"absolute",left:0,top:0,bottom:0,width:`${Math.round((d.value/max)*100)}%`,background:d.color||(colorFn?colorFn(i):C.p),borderRadius:6,transition:"width .4s"}}/>
          </div>
          <div style={{width:28,fontSize:12,fontWeight:700,color:C.t,textAlign:"left",flexShrink:0}}>{d.value}</div>
        </div>
      ))}
    </div>
  );
}

/* ─── ERROR BOUNDARY ────────────────────────────────────────────────── */
/* FIX AUDIT #14 (MEDIUM — ErrorBoundary tidak me-reset state setelah error):
   Sebelumnya "Coba Lagi" hanya men-set err:false, tanpa me-remount komponen
   child yang error. React akan mencoba merender ulang child yang sama persis
   yang sudah error — kemungkinan besar akan error lagi dan pengguna terjebak.
   Fix: tambahkan field `key` yang di-increment setiap reset. Perubahan `key`
   pada React.Fragment memaksa seluruh subtree di-unmount lalu remount bersih,
   sehingga state internal child ikut direset. */
export class ErrorBoundary extends Component<{children: React.ReactNode},{err:boolean;msg:string;key:number}> {
  state = {err:false, msg:"", key:0};
  static getDerivedStateFromError(e: any){ return {err:true, msg:e?.message||"Error"}; }

  handleReset = () => {
    this.setState(prev => ({err:false, msg:"", key: prev.key + 1}));
  };

  render(){
    if(this.state.err) return (
      <div style={{padding:24,textAlign:"center",background:C.dBg,borderRadius:14,margin:"8px 0",border:`1px solid ${C.dL}`}}>
        <div style={{fontSize:32,marginBottom:8}}>⚠️</div>
        <div style={{fontSize:14,color:C.d,fontWeight:600,marginBottom:8}}>{this.state.msg}</div>
        <Btn onClick={this.handleReset}>Coba Lagi</Btn>
      </div>
    );
    /* key change forces full remount of children — state internal child ikut direset */
    return <Fragment key={this.state.key}>{this.props.children}</Fragment>;
  }
}
