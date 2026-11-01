/* src/components/effects.tsx
   Komponen visual/animasi dekoratif (gaya "ReactBits") — background,
   hover/tilt effect, teks animasi, dan efek klik. SEMUA murni presentational
   (pure CSS/JS via inline <style> injection), TIDAK bergantung ke state,
   types, atau constants aplikasi. Dipindah dari App.tsx (langkah 5 refactor
   — pisah components, batch 1: effects), TIDAK ADA perubahan logika. */

import { useState, useEffect, useMemo, useRef } from "react";

/** Aurora — cahaya warna-warni bergerak halus di background, dipakai di
 *  layar login untuk kesan "premium/mewah". */
export const AuroraBackground = ({opacity=1}: {opacity?: number}) => {
  useEffect(() => {
    if (typeof document === "undefined" || document.getElementById("kb-aurora-css")) return;
    const el = document.createElement("style");
    el.id = "kb-aurora-css";
    el.textContent = `
      @keyframes kbAuroraDrift1{0%,100%{transform:translate(-10%,-10%) scale(1);}50%{transform:translate(8%,6%) scale(1.15);}}
      @keyframes kbAuroraDrift2{0%,100%{transform:translate(10%,8%) scale(1.1);}50%{transform:translate(-6%,-8%) scale(1);}}
      @keyframes kbAuroraDrift3{0%,100%{transform:translate(-6%,10%) scale(1);}50%{transform:translate(10%,-6%) scale(1.2);}}
    `;
    document.head.appendChild(el);
  }, []);
  return (
    <div style={{position:"absolute",inset:0,overflow:"hidden",zIndex:0,opacity,pointerEvents:"none"}}>
      <div style={{position:"absolute",top:"-15%",left:"-10%",width:"55%",height:"55%",borderRadius:"50%",
        background:"radial-gradient(circle,rgba(56,189,248,.55),transparent 70%)",
        filter:"blur(50px)",animation:"kbAuroraDrift1 14s ease-in-out infinite"}}/>
      <div style={{position:"absolute",bottom:"-20%",right:"-10%",width:"60%",height:"60%",borderRadius:"50%",
        background:"radial-gradient(circle,rgba(45,212,191,.5),transparent 70%)",
        filter:"blur(55px)",animation:"kbAuroraDrift2 17s ease-in-out infinite"}}/>
      <div style={{position:"absolute",top:"25%",right:"5%",width:"40%",height:"40%",borderRadius:"50%",
        background:"radial-gradient(circle,rgba(232,206,143,.4),transparent 70%)",
        filter:"blur(45px)",animation:"kbAuroraDrift3 20s ease-in-out infinite"}}/>
    </div>
  );
};

/* PENAMBAHAN — batch ke-3 efek ReactBits, dipilih untuk melengkapi tema
   navy+gold "mewah" dan konteks kerja klinis. Semua pure CSS/JS, tanpa
   dependency baru.

   Meteors — kilatan cahaya emas jatuh diagonal, kesan "night sky mewah"
   di background gelap (login). */
export const MeteorsBackground = ({count=14}: {count?: number}) => {
  useEffect(() => {
    if (typeof document === "undefined" || document.getElementById("kb-meteors-css")) return;
    const el = document.createElement("style");
    el.id = "kb-meteors-css";
    el.textContent = `@keyframes kbMeteorFall{0%{transform:translate(0,0) rotate(215deg);opacity:1;}70%{opacity:1;}100%{transform:translate(-420px,420px) rotate(215deg);opacity:0;}}`;
    document.head.appendChild(el);
  }, []);
  const meteors = useMemo(() => Array.from({length:count}).map((_,i)=>({
    top: Math.random()*40-10,
    left: Math.random()*100,
    delay: Math.random()*7,
    duration: 4.5 + Math.random()*3.5,
  })), [count]);
  return (
    <div style={{position:"absolute",inset:0,overflow:"hidden",zIndex:0,pointerEvents:"none"}}>
      {meteors.map((m,i) => (
        <span key={i} style={{
          position:"absolute",top:`${m.top}%`,left:`${m.left}%`,
          width:2,height:2,borderRadius:"50%",background:"#E8CE8F",
          boxShadow:"0 0 0 1px rgba(232,206,143,.15),0 0 12px 2px rgba(232,206,143,.5),-60px 60px 20px -8px rgba(232,206,143,.15)",
          animation:`kbMeteorFall ${m.duration}s linear ${m.delay}s infinite`,
        }}>
          <span style={{position:"absolute",top:"50%",right:0,height:1,width:56,
            background:"linear-gradient(90deg,rgba(232,206,143,.6),transparent)",transform:"translateY(-50%)"}}/>
        </span>
      ))}
    </div>
  );
};

/** Glare Hover — sapuan cahaya diagonal melintasi elemen saat hover,
 *  memberi kesan premium/reflektif seperti permukaan kaca/logam. */
export const GlareHover = ({children, style={}}: {children: React.ReactNode; style?: any}) => {
  useEffect(() => {
    if (typeof document === "undefined" || document.getElementById("kb-glare-css")) return;
    const el = document.createElement("style");
    el.id = "kb-glare-css";
    el.textContent = `
      .kb-glare-hover{position:relative;overflow:hidden;}
      .kb-glare-hover::before{content:"";position:absolute;top:0;left:-75%;width:50%;height:100%;
        background:linear-gradient(115deg,transparent,rgba(255,255,255,.35),transparent);
        transform:skewX(-20deg);transition:left .65s ease;pointer-events:none;z-index:2;}
      .kb-glare-hover:hover::before{left:125%;}
    `;
    document.head.appendChild(el);
  }, []);
  return <div className="kb-glare-hover" style={style}>{children}</div>;
};

/** Marquee — teks berjalan kontinu (ticker), dipakai untuk alert
 *  darurat (CYTO) supaya menarik perhatian tanpa perlu di-scroll. */
export const Marquee = ({children, speed=18}: {children: React.ReactNode; speed?: number}) => {
  useEffect(() => {
    if (typeof document === "undefined" || document.getElementById("kb-marquee-css")) return;
    const el = document.createElement("style");
    el.id = "kb-marquee-css";
    el.textContent = `@keyframes kbMarqueeScroll{from{transform:translateX(0);}to{transform:translateX(-50%);}}`;
    document.head.appendChild(el);
  }, []);
  return (
    <div style={{overflow:"hidden",whiteSpace:"nowrap",width:"100%"}}>
      <div style={{display:"inline-flex",animation:`kbMarqueeScroll ${speed}s linear infinite`}}>
        <span style={{display:"inline-block",paddingRight:48}}>{children}</span>
        <span style={{display:"inline-block",paddingRight:48}} aria-hidden="true">{children}</span>
      </div>
    </div>
  );
};

/* PENAMBAHAN — batch ke-4 efek ReactBits: Split Text, Twinkling Stars,
   Rotating Text, Squares background, Animated List item wrapper. Semua
   pure CSS/JS, tanpa dependency baru. */

/** Split Text — teks muncul kata-per-kata (fade+slide-up bertahap),
 *  dipakai untuk judul form Tab Daftar. */
export const SplitText = ({text, style={}, delayStep=40}: {text: string; style?: any; delayStep?: number}) => {
  const words = text.split(" ");
  return (
    <span style={style}>
      {words.map((w,i) => (
        <span key={i} style={{
          display:"inline-block",
          animation:`kbSplitWordIn .5s ease-out ${(i*delayStep)}ms both`,
        }}>{w}{i<words.length-1?"\u00A0":""}</span>
      ))}
    </span>
  );
};

/** Twinkling Stars — titik-titik kecil berkedip acak di background gelap,
 *  melengkapi Meteors/Aurora untuk kesan "malam mewah" yang lebih kuat. */
export const TwinklingStars = ({count=40}: {count?: number}) => {
  useEffect(() => {
    if (typeof document === "undefined" || document.getElementById("kb-twinkle-css")) return;
    const el = document.createElement("style");
    el.id = "kb-twinkle-css";
    el.textContent = `
      @keyframes kbSplitWordIn{from{opacity:0;transform:translateY(10px);}to{opacity:1;transform:translateY(0);}}
      @keyframes kbTwinkle{0%,100%{opacity:.15;}50%{opacity:1;}}
    `;
    document.head.appendChild(el);
  }, []);
  const stars = useMemo(() => Array.from({length:count}).map(() => ({
    top: Math.random()*100, left: Math.random()*100,
    size: 1 + Math.random()*1.5,
    delay: Math.random()*4, duration: 2+Math.random()*3,
  })), [count]);
  return (
    <div style={{position:"absolute",inset:0,overflow:"hidden",zIndex:0,pointerEvents:"none"}}>
      {stars.map((s,i) => (
        <span key={i} style={{position:"absolute",top:`${s.top}%`,left:`${s.left}%`,
          width:s.size,height:s.size,borderRadius:"50%",background:"#fff",
          animation:`kbTwinkle ${s.duration}s ease-in-out ${s.delay}s infinite`}}/>
      ))}
    </div>
  );
};

/** Rotating Text — kata berganti otomatis dengan transisi slide-up,
 *  dipakai di welcome banner Tab Home. */
export const RotatingText = ({words, interval=2200, style={}}: {words: string[]; interval?: number; style?: any}) => {
  const [idx, setIdx] = useState(0);
  useEffect(() => {
    const iv = setInterval(() => setIdx(i => (i+1) % words.length), interval);
    return () => clearInterval(iv);
  }, [words.length, interval]);
  return (
    <span style={{position:"relative",display:"inline-block",overflow:"hidden",verticalAlign:"bottom",...style}}>
      <span key={idx} style={{display:"inline-block",animation:"kbRotateWordIn .45s ease-out"}}>{words[idx]}</span>
      <style>{`@keyframes kbRotateWordIn{from{opacity:0;transform:translateY(100%);}to{opacity:1;transform:translateY(0);}}`}</style>
    </span>
  );
};

/** Squares Background — grid garis tipis animasi halus, kesan "dashboard
 *  data teknis" untuk halaman tabel (Tab Staf/Master Data). */
export const SquaresBackground = ({opacity=0.5}: {opacity?: number}) => {
  useEffect(() => {
    if (typeof document === "undefined" || document.getElementById("kb-squares-css")) return;
    const el = document.createElement("style");
    el.id = "kb-squares-css";
    el.textContent = `@keyframes kbSquaresDrift{from{background-position:0 0;}to{background-position:40px 40px;}}`;
    document.head.appendChild(el);
  }, []);
  return (
    <div style={{
      position:"absolute",inset:0,zIndex:0,pointerEvents:"none",opacity,
      backgroundImage:"linear-gradient(rgba(148,163,184,.15) 1px, transparent 1px), linear-gradient(90deg, rgba(148,163,184,.15) 1px, transparent 1px)",
      backgroundSize:"40px 40px",
      animation:"kbSquaresDrift 6s linear infinite",
    }}/>
  );
};

/** Animated List item — pembungkus untuk item list yang stagger fade+slide
 *  in berdasarkan index-nya, dipakai untuk hasil filter/search Tab Jadwal. */
export const AnimatedListItem = ({children, index, staggerMs=35}: {children: React.ReactNode; index: number; staggerMs?: number}) => (
  <div style={{animation:`kbTabFadeIn .35s ease-out ${Math.min(index*staggerMs,350)}ms both`}}>{children}</div>
);

/* PENAMBAHAN — perluasan efek ReactBits (batch ke-2, untuk WEB UTAMA
   setelah login, bukan hanya login screen). Semua tetap pure CSS/JS,
   tanpa dependency baru.

   4) MagneticButton — tombol "tertarik" halus mengikuti kursor saat
      didekati (mirip "Magnetic Button" ReactBits). HANYA dipakai di
      tombol non-kritis (Quick Links, dsb) — BUKAN di card data pasien,
      supaya kesan tetap serius untuk konteks kerja klinis.
   2) SpotlightCard — cahaya lembut mengikuti posisi kursor di dalam card
      saat hover (mirip "Spotlight Card" ReactBits).
   8) DecryptedText — teks "teracak" lalu terbentuk jadi kata aslinya saat
      pertama muncul (mirip "Decrypted Text" ReactBits) — dipakai TERBATAS
      di 1 label kecil saja (bukan di data pasien) supaya tidak mengganggu
      kecepatan baca informasi penting.
   9) TiltCard — kemiringan 3D halus mengikuti kursor (mirip "Tilt Card"
      ReactBits) — SENGAJA hanya diterapkan ke tombol Quick Links (bukan
      ke card jadwal/data pasien), karena efek ini bisa terasa "main-main"
      untuk konteks data medis yang serius. */
export const MagneticButton = ({children, onClick, style={}, strength=14}: any) => {
  const ref = useRef<HTMLButtonElement>(null);
  const [pos, setPos] = useState({x:0,y:0});
  const handleMove = (e: React.MouseEvent) => {
    const el = ref.current; if (!el) return;
    const rect = el.getBoundingClientRect();
    const relX = (e.clientX - rect.left - rect.width/2) / (rect.width/2);
    const relY = (e.clientY - rect.top - rect.height/2) / (rect.height/2);
    setPos({x: relX*strength, y: relY*strength});
  };
  const reset = () => setPos({x:0,y:0});
  return (
    <button ref={ref} onClick={onClick} onMouseMove={handleMove} onMouseLeave={reset}
      style={{...style, transform:`translate(${pos.x}px, ${pos.y}px)`, transition:"transform .15s ease-out"}}>
      {children}
    </button>
  );
};

export const SpotlightCard = ({children, style={}, spotlightColor="rgba(56,189,248,.25)"}: any) => {
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState({x:50,y:50});
  const [active, setActive] = useState(false);
  const handleMove = (e: React.MouseEvent) => {
    const el = ref.current; if (!el) return;
    const rect = el.getBoundingClientRect();
    setPos({x: ((e.clientX-rect.left)/rect.width)*100, y: ((e.clientY-rect.top)/rect.height)*100});
  };
  return (
    <div ref={ref} onMouseMove={handleMove} onMouseEnter={()=>setActive(true)} onMouseLeave={()=>setActive(false)}
      style={{position:"relative",overflow:"hidden",...style}}>
      <div style={{position:"absolute",inset:0,pointerEvents:"none",opacity:active?1:0,transition:"opacity .3s ease",
        background:`radial-gradient(circle at ${pos.x}% ${pos.y}%, ${spotlightColor}, transparent 60%)`}}/>
      <div style={{position:"relative",zIndex:1}}>{children}</div>
    </div>
  );
};

export const DecryptedText = ({text, style={}}: {text: string; style?: any}) => {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  const [display, setDisplay] = useState(text);
  useEffect(() => {
    let frame = 0;
    const totalFrames = 14;
    const iv = setInterval(() => {
      frame++;
      const revealCount = Math.floor((frame/totalFrames)*text.length);
      setDisplay(text.split("").map((c,i)=>{
        if (c===" ") return " ";
        return i < revealCount ? c : chars[Math.floor(Math.random()*chars.length)];
      }).join(""));
      if (frame >= totalFrames) { setDisplay(text); clearInterval(iv); }
    }, 40);
    return () => clearInterval(iv);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [text]);
  return <span style={style}>{display}</span>;
};

export const TiltCard = ({children, style={}, maxTilt=6}: any) => {
  const ref = useRef<HTMLDivElement>(null);
  const [tilt, setTilt] = useState({x:0,y:0});
  const handleMove = (e: React.MouseEvent) => {
    const el = ref.current; if (!el) return;
    const rect = el.getBoundingClientRect();
    const relX = (e.clientX - rect.left)/rect.width - 0.5;
    const relY = (e.clientY - rect.top)/rect.height - 0.5;
    setTilt({x: relY*maxTilt*-1, y: relX*maxTilt});
  };
  const reset = () => setTilt({x:0,y:0});
  return (
    <div ref={ref} onMouseMove={handleMove} onMouseLeave={reset}
      style={{...style, transform:`perspective(600px) rotateX(${tilt.x}deg) rotateY(${tilt.y}deg)`,
        transition:"transform .2s ease-out",transformStyle:"preserve-3d"}}>
      {children}
    </div>
  );
};

export const ShinyText = ({children, style={}}: {children: React.ReactNode; style?: any}) => {
  useEffect(() => {
    if (typeof document === "undefined" || document.getElementById("kb-shiny-css")) return;
    const el = document.createElement("style");
    el.id = "kb-shiny-css";
    el.textContent = `@keyframes kbShinySweep{0%{background-position:-150% 0;}100%{background-position:250% 0;}}`;
    document.head.appendChild(el);
  }, []);
  return (
    <span style={{
      backgroundImage:"linear-gradient(100deg, currentColor 40%, rgba(255,255,255,.9) 50%, currentColor 60%)",
      backgroundSize:"250% 100%",
      WebkitBackgroundClip:"text",backgroundClip:"text",color:"transparent",
      animation:"kbShinySweep 3.2s linear infinite",
      ...style,
    }}>{children}</span>
  );
};

/** Confetti burst — dipakai saat simpan data berhasil (feedback positif). */
export function fireConfettiBurst(originEl?: HTMLElement | null) {
  if (typeof document === "undefined") return;
  const rect = originEl?.getBoundingClientRect();
  const originX = rect ? rect.left + rect.width/2 : window.innerWidth/2;
  const originY = rect ? rect.top + rect.height/2 : window.innerHeight/3;
  const colors = ["#A78BFA","#22c55e","#facc15","#f472b6","#a78bfa"];
  const container = document.createElement("div");
  container.style.cssText = "position:fixed;inset:0;pointer-events:none;z-index:99998;overflow:hidden;";
  document.body.appendChild(container);
  const N = 18;
  for (let i=0; i<N; i++) {
    const p = document.createElement("div");
    const angle = (Math.PI*2*i)/N + Math.random()*0.4;
    const dist = 60 + Math.random()*70;
    const dx = Math.cos(angle)*dist, dy = Math.sin(angle)*dist - 30;
    const size = 5 + Math.random()*4;
    p.style.cssText = `position:absolute;left:${originX}px;top:${originY}px;width:${size}px;height:${size}px;
      background:${colors[i % colors.length]};border-radius:${Math.random()>0.5?"50%":"2px"};
      transform:translate(-50%,-50%);opacity:1;
      transition:transform .7s cubic-bezier(.2,.7,.3,1),opacity .7s ease;`;
    container.appendChild(p);
    requestAnimationFrame(() => {
      p.style.transform = `translate(calc(-50% + ${dx}px), calc(-50% + ${dy}px)) rotate(${Math.random()*360}deg)`;
      p.style.opacity = "0";
    });
  }
  setTimeout(() => container.remove(), 800);
}

/** Percikan kecil di titik klik (mirip "ClickSpark" ReactBits) — dipakai
 *  di onClick tombol utama untuk feedback yang lebih hidup. */
export function fireClickSpark(e: React.MouseEvent) {
  const container = document.createElement("div");
  container.style.cssText = "position:fixed;inset:0;pointer-events:none;z-index:99998;overflow:hidden;";
  document.body.appendChild(container);
  const ox = e.clientX, oy = e.clientY;
  const colors = ["#A78BFA","#2dd4bf","#facc15"];
  for (let i=0;i<8;i++){
    const p = document.createElement("div");
    const angle = (Math.PI*2*i)/8;
    const dist = 26+Math.random()*14;
    p.style.cssText = `position:absolute;left:${ox}px;top:${oy}px;width:3px;height:3px;border-radius:50%;
      background:${colors[i%colors.length]};transform:translate(-50%,-50%);opacity:1;
      transition:transform .45s ease-out,opacity .45s ease-out;`;
    container.appendChild(p);
    requestAnimationFrame(()=>{
      p.style.transform = `translate(calc(-50% + ${Math.cos(angle)*dist}px), calc(-50% + ${Math.sin(angle)*dist}px))`;
      p.style.opacity = "0";
    });
  }
  setTimeout(()=>container.remove(), 500);
}
