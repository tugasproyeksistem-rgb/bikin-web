/* src/components/date-txt.tsx
   Input tanggal custom dengan auto-slash-insertion (DD/MM/YYYY) — pengganti
   <input type="date"> native yang render format US (MM/DD/YYYY) di Windows.
   Dropdown kalender native browser tetap tersedia (disembunyikan, dibuka
   lewat tombol 📅). Dipindah dari App.tsx (langkah 6 refactor — pisah View/
   layar besar, batch 1), TIDAK ADA perubahan logika. */

import { useState, useEffect, useRef } from "react";
import { fDMY, toISODateStrict } from "../utils";

export function DateTxt({value,onChange,style,...rest}: {value:string;onChange:(v:string)=>void;style?:React.CSSProperties;[k:string]:any}) {
  const [raw,setRaw] = useState(()=>value?fDMY(value):"");
  const prevVal = useRef(value);
  const pickerRef = useRef<HTMLInputElement>(null);
  useEffect(()=>{
    if(value!==prevVal.current){
      prevVal.current=value;
      setRaw(value?fDMY(value):"");
    }
  },[value]);
  const handleChange=(e:React.ChangeEvent<HTMLInputElement>)=>{
    let v=e.target.value.replace(/[^0-9/]/g,"");
    /* auto-insert slash setelah DD dan MM */
    if(v.length===2&&raw.length===1&&!v.includes("/")) v+="/";
    if(v.length===5&&raw.length===4&&v.split("/").length===2) v+="/";
    if(v.length>10) v=v.slice(0,10);
    setRaw(v);
    if(v.length===10){const iso=toISODateStrict(v);if(iso)onChange(iso);}
    else if(v==="")onChange("");
  };
  /* Buka dropdown kalender native browser. <input type="date"> disembunyikan
     (opacity 0, menumpuk di atas tombol 📅) sehingga tidak menampilkan
     format MM/DD/YYYY yang jadi masalah — kalender ini hanya dipakai untuk
     MEMILIH tanggal lewat dropdown, bukan untuk menampilkan teksnya. */
  const openPicker = () => {
    const el = pickerRef.current;
    if(!el) return;
    if(typeof (el as any).showPicker === "function"){
      try{ (el as any).showPicker(); return; }catch{ /* fallback di bawah */ }
    }
    el.focus();
    el.click();
  };
  const handlePickerChange=(e:React.ChangeEvent<HTMLInputElement>)=>{
    const iso = e.target.value;
    if(iso){
      setRaw(fDMY(iso));
      onChange(iso);
    }
  };
  return (
    <div style={{position:"relative",display:"inline-block",width:style?.width||"100%"}}>
      <input
        {...rest}
        type="text"
        inputMode="numeric"
        value={raw}
        onChange={handleChange}
        placeholder="DD/MM/YYYY"
        maxLength={10}
        style={{...style,width:"100%",paddingRight:30,boxSizing:"border-box"}}
      />
      <input
        ref={pickerRef}
        type="date"
        value={value||""}
        onChange={handlePickerChange}
        tabIndex={-1}
        aria-hidden="true"
        style={{position:"absolute",right:2,top:"50%",transform:"translateY(-50%)",width:24,height:24,opacity:0,cursor:"pointer"}}
      />
      <button
        type="button"
        onClick={openPicker}
        tabIndex={-1}
        aria-label="Buka kalender"
        style={{position:"absolute",right:4,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",padding:4,lineHeight:0,color:"#666",pointerEvents:"none"}}
      >📅</button>
    </div>
  );
}
