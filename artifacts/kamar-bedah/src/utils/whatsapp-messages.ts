/* src/utils/whatsapp-messages.ts
   Pembangun teks pesan WhatsApp — pengingat H-1 ke dokter bedah/anestesi,
   dan alert CYTO/emergency. Dipindah dari App.tsx (langkah 6 refactor —
   diekstrak bersamaan dengan ViewKirimWA yang memakainya), TIDAK ADA
   perubahan logika. */

import { HOSPITAL } from "../constants";
import { titleCaseDokter, gWord, timeOfDayLabel, fDLower, fDMY, fTR, tmrwDate, usiaOp, fNow } from "./format";
import { sanitize } from "./security";

export function msgSurgeon(name: string, ops: any[]) {
  const s = [...ops].sort((a,b)=>a.time.localeCompare(b.time));
  if(!s.length) return "";
  const N = titleCaseDokter(name); // nama TARGET sapaan → Title Case, tanpa prefix "dr"
  const uniqAnest = [...new Set(s.map(o=>o.anesthesiologist).filter(Boolean))];
  const anestLabel = uniqAnest.length ? uniqAnest.map(a=>sanitize(a)).join(" dan ") : "-"; // nama pendukung → raw, sesuai database
  const greet = gWord();
  const shift = timeOfDayLabel(s[0]?.time);
  const tanggal = fDLower(tmrwDate());
  const jam = fTR(s[0]?.time);
  const baris = (o:any) => {
    const usia = usiaOp(o);
    const diag = sanitize(o.diagnosis||"").trim();
    return `${sanitize(o.patient)}${usia?` ${usia}`:""} rencana ${sanitize(o.procedure)}${diag?` ${diag}`:""}`;
  };
  if(s.length>1){
    const lines = s.map((o,i)=>`${i+1}. ${baris(o)}`).join("\n\n");
    const biusWord = anestLabel==="Lokal Anestesi" ? "" : "dr bius ";
    return `selamat ${greet} dokter ${N} rencana operasi besok ${shift} , ${tanggal} mulai jam ${jam} ${biusWord}dengan ${anestLabel} ada ${s.length} pasien :\n${lines}\nterima kasih banyak dokter`;
  }
  const biusWordSingle = anestLabel==="Lokal Anestesi" ? "" : "bius ";
  return `selamat ${greet} dokter ${N} rencana operasi besok ${shift} ${tanggal} mulai jam ${jam} ${biusWordSingle}dengan ${anestLabel} atas nama ${baris(s[0])}\nterima kasih banyak dokter`;
}

/* msgAnest — direfactor sesuai Spesifikasi Teknis & CONTOH OUTPUT LITERAL
   yang diberikan. Signature (name, ops) TIDAK diubah.
   Catatan keputusan implementasi:
   - Spacing mengikuti CONTOH AKHIR secara harfiah (baris-baris rapat,
     newline tunggal antar section — BUKAN double newline seperti disebut
     di teks algoritma pipeline, yang ternyata tidak konsisten dgn contoh).
   - Header "dokter [Nama]" pakai titleCaseDokter (Title Case, tanpa prefix
     "dr"), sesuai contoh "dokter Timor". Nama dokter bedah & anestesi di
     BADAN pesan tetap nilai asli dari database (mis. "dr dedi", "dr timor"),
     sesuai contoh — tidak diubah case-nya. */
export function msgAnest(name: string, ops: any[]) {
  const tanggal = fDLower(tmrwDate());
  const greet = gWord();
  const anestRaw = sanitize(name); // dipakai di badan pesan, format asli (mis. "dr timor")

  const renderShift = (shiftOps: any[]): string => {
    if(!shiftOps.length) return "";
    const byTime: Record<string, any[]> = {};
    shiftOps.forEach(o=>{ if(!byTime[o.time]) byTime[o.time]=[]; byTime[o.time].push(o); });
    return Object.keys(byTime).sort((a,b)=>a.localeCompare(b)).map(time=>{
      const opsAtTime = byTime[time];
      const bySurgeon: Record<string, any[]> = {};
      opsAtTime.forEach(o=>{ const k=(o.surgeon||"").trim().toLowerCase(); if(!bySurgeon[k]) bySurgeon[k]=[]; bySurgeon[k].push(o); });
      return Object.values(bySurgeon).map((group:any[])=>{
        const surgeonRaw = sanitize(group[0].surgeon);
        const jam = fTR(time);
        const baris = (o:any) => {
          const usia = usiaOp(o);
          const diag = sanitize(o.diagnosis||"").trim();
          return `${sanitize(o.patient)}${usia?` ${usia}`:""} rencana ${sanitize(o.procedure)}${diag?` ${diag}`:""}`;
        };
        if(group.length>1){
          const lines = group.map((o,i)=>`${i+1}. ${baris(o)}`).join("\n\n");
          return `*jam ${jam}* ${surgeonRaw} dan ${anestRaw} ada ${group.length} pasien :\n${lines}`;
        }
        return `*jam ${jam}* ${surgeonRaw} dan ${anestRaw} atas nama ${baris(group[0])}`;
      }).join("\n\n");
    }).join("\n\n");
  };

  const pagi = ops.filter((o:any)=>o.time<"14:00");
  const sore = ops.filter((o:any)=>o.time>="14:00");

  let out = `Selamat ${greet} dokter ${titleCaseDokter(name)} ,Rencana operasi besok hari ${tanggal}`;
  if(pagi.length) out += `\n*Pagi*\n${renderShift(pagi)}`;
  if(sore.length) out += `\n*Sore*\n${renderShift(sore)}`;
  if(!pagi.length && !sore.length) out += "\nBelum ada rencana operasi";
  out += "\nterima kasih banyak dokter";
  return out;
}

export function msgCyto(op: any) {
  const alg = op.allergy&&op.allergy!=="Tidak Ada" ? `⚠ ALERGI: ${sanitize(op.allergy)}\n` : "";
  const ra  = op.ruangAsal ? `Ruang Asal : ${sanitize(op.ruangAsal)}\n` : "";
  return `⚠️ LAPORAN ACARA OPERASI CYTO / EMERGENCY ⚠️\n━━━━━━━━━━━━━━━━━━━\nKamar Bedah ${HOSPITAL}\n━━━━━━━━━━━━━━━━━━━\n${fNow()}\n\nMOHON SEGERA KE KAMAR OPERASI\n\n${alg}${ra}Pasien   : ${sanitize(op.patient)}${op.age?` (${op.age} th)`:""}\nRM/Gol   : ${op.rm||"-"} / ${op.bloodType}\nDiagnosa : ${sanitize(op.diagnosis)}\nTindakan : ${sanitize(op.procedure)}\nJam      : ${op.time} WIB — ${op.room}\n\nOperator : ${op.surgeon||"-"}\nAnestesi : ${op.anesthesiologist||"-"}\nAsisten  : ${op.assistantNurse||"-"}\nInstrumen: ${op.circulatingNurse||"-"}\nOnloop   : ${op.onloopNurse||"-"}\nRR/Katim : ${op.rrKatim||"-"}\n━━━━━━━━━━━━━━━━━━━\n⚡ HARAP RESPON SEGERA ⚡\nKamar Bedah ${HOSPITAL}`;
}

/* buildLaporan: pembangun teks laporan "Kepala Jaga" (rekap operasi hari
   ini + rencana besok + siaga anestesi/perawat) untuk dikirim via WhatsApp
   ke grup/penerima laporan. Dipakai oleh ViewLaporan. */
export function buildLaporan({greeting,recipients,keterangan,todayOps,tomorrowOps,anestStandby,anestCyto,nurseStandby,pembawaHP}: any) {
  const active = [...todayOps].filter((o:any)=>o.status!=="batal");
  const batal  = todayOps.filter((o:any)=>o.status==="batal");

  /* ── Poin ③: direstrukturisasi sesuai format contoh terbaru —
     dikelompokkan PAGI (<14:00) / SORE (>=14:00) → per jam operasi →
     per dokter bedah+anestesi (sama persis pola Poin ④), tapi untuk hari
     ini ditambahkan info perawat (asisten/instrumen/onloop/rr-katim) di
     akhir tiap baris pasien dalam tanda kurung miring, karena data
     penugasan perawat hari ini sudah final (bukan rencana seperti besok). */
  const renderShiftToday = (shiftOps: any[]): string => {
    if(!shiftOps.length) return "";
    const byTime: Record<string, any[]> = {};
    shiftOps.forEach((o:any)=>{ if(!byTime[o.time]) byTime[o.time]=[]; byTime[o.time].push(o); });
    return Object.keys(byTime).sort((a,b)=>a.localeCompare(b)).map(time=>{
      const bySurgeon: Record<string, any[]> = {};
      byTime[time].forEach((o:any)=>{ const k=(o.surgeon||"").trim().toLowerCase(); if(!bySurgeon[k]) bySurgeon[k]=[]; bySurgeon[k].push(o); });
      return Object.values(bySurgeon).map((group:any[])=>{
        const surgeonDisplay = sanitize(group[0].surgeon)||"-";
        const anestDisplay = sanitize(group[0].anesthesiologist)||"-";
        const jam = fTR(time);
        const baris = (o:any) => {
          const usia = usiaOp(o);
          const diag = sanitize(o.diagnosis||"").trim();
          const cytoMark = o.opType==="cyto" ? "⚠️[CYTO]⚠️ " : "";
          const ass = sanitize(o.assistantNurse).trim();
          const inst = sanitize(o.circulatingNurse).trim();
          const onloop = sanitize(o.onloopNurse).trim();
          const rr = sanitize(o.rrKatim).trim();
          const parts: string[] = [];
          if(ass) parts.push(`ass:${ass}`);
          if(inst) parts.push(`inst:${inst}`);
          if(onloop) parts.push(`onloop:${onloop}`);
          if(rr) parts.push(`rr/katim:${rr}`);
          const roles = parts.length ? ` (${parts.join(", ")})` : "";
          return `_${cytoMark}${sanitize(o.patient)}${usia?` ${usia}`:""} rencana ${sanitize(o.procedure)}${diag?` ${diag}`:""}${roles}_`;
        };
        if(group.length>1){
          const lines = group.map((o,i)=>`${i+1}. ${baris(o)}`).join("\n");
          return `*jam ${jam}* ${surgeonDisplay} dan ${anestDisplay} ada ${group.length} pasien :\n${lines}`;
        }
        return `*jam ${jam}* ${surgeonDisplay} dan ${anestDisplay} atas nama ${baris(group[0])}`;
      }).join("\n\n");
    }).join("\n\n");
  };
  const sortedActive = [...active].sort((a:any,b:any)=>a.time.localeCompare(b.time));
  const todayPagi = sortedActive.filter((o:any)=>o.time<"14:00");
  const todaySore = sortedActive.filter((o:any)=>o.time>="14:00");
  const batalLine = batal.length?`\n\n_Batal/Tunda: ${batal.map((o:any)=>sanitize(o.patient)+(o.cancelReason?` (${sanitize(o.cancelReason)})`:""  )).join("; ")}_`:"";
  let todaySec = `Acara operasi hari ini ada ${active.length}.\n`;
  if(todayPagi.length) todaySec += `\n*PAGI*\n${renderShiftToday(todayPagi)}`;
  if(todaySore.length) todaySec += `\n\n*SORE*\n${renderShiftToday(todaySore)}`;
  if(!todayPagi.length && !todaySore.length) todaySec += "\n_Hari Ini Tidak Ada Acara Operasi_";
  todaySec += batalLine;

  /* ── Poin ④: split shift Pagi (00:00-13:59) / Sore (14:00-23:59) → group
     by jamOperasi → group by dokter bedah → IF jumlah>1 list bernumor,
     ELSE kalimat inline. formatUsia WAJIB dipakai di bagian ini. ── */
  const pagi = tomorrowOps.filter((o:any)=>o.time<"14:00");
  const sore = tomorrowOps.filter((o:any)=>o.time>="14:00");
  const renderShift = (shiftOps: any[]): string => {
    if(!shiftOps.length) return "";
    const byTime: Record<string, any[]> = {};
    shiftOps.forEach((o:any)=>{ if(!byTime[o.time]) byTime[o.time]=[]; byTime[o.time].push(o); });
    return Object.keys(byTime).sort((a,b)=>a.localeCompare(b)).map(time=>{
      const bySurgeon: Record<string, any[]> = {};
      byTime[time].forEach((o:any)=>{ const k=(o.surgeon||"").trim().toLowerCase(); if(!bySurgeon[k]) bySurgeon[k]=[]; bySurgeon[k].push(o); });
      return Object.values(bySurgeon).map((group:any[])=>{
        const surgeonDisplay = sanitize(group[0].surgeon)||"-";
        const anestDisplay = sanitize(group[0].anesthesiologist)||"-";
        const jam = fTR(time);
        const baris = (o:any) => {
          const usia = usiaOp(o);
          const diag = sanitize(o.diagnosis||"").trim();
          return `${sanitize(o.patient)}${usia?` ${usia}`:""} rencana ${sanitize(o.procedure)}${diag?` ${diag}`:""}`;
        };
        if(group.length>1){
          const lines = group.map((o,i)=>`${i+1}. ${baris(o)}`).join("\n");
          return `*jam ${jam}* ${surgeonDisplay} dan ${anestDisplay} ada ${group.length} pasien :\n${lines}`;
        }
        return `*jam ${jam}* ${surgeonDisplay} dan ${anestDisplay} atas nama ${baris(group[0])}`;
      }).join("\n\n");
    }).join("\n\n");
  };
  let tmrwSec = `Rencana operasi besok tanggal ${fDMY(tmrwDate())}:\n`;
  if(pagi.length) tmrwSec += `\n*PAGI*\n${renderShift(pagi)}`;
  if(sore.length) tmrwSec += `\n\n*SORE*\n${renderShift(sore)}`;
  if(!pagi.length&&!sore.length) tmrwSec += "\nBelum ada rencana operasi";

  /* ── Bagian berikut TIDAK diubah dari implementasi sebelumnya ── */
  const aSec  = anestStandby.filter(Boolean).length ? `\n*Siaga Anaesthesi*\n${anestStandby.filter(Boolean).map((n:string,i:number)=>`${i+1}. ${n}`).join("\n")}` : "";
  const cSec  = anestCyto.filter(Boolean).length    ? `\n\n*Siaga Anaesthesi Cyto*\n${anestCyto.filter(Boolean).map((n:string,i:number)=>`${i+1}. ${n}`).join("\n")}` : "";
  const nSec  = nurseStandby.filter(Boolean).length ? `\n\n*Perawat siaga*\n${nurseStandby.filter(Boolean).map((n:string,i:number)=>`${i+1}. ${n}`).join("\n")}` : "";
  const phSec = pembawaHP?.trim() ? `\n\nPembawa HP: ${pembawaHP}` : "";
  const ket   = keterangan?.trim() ? `\n${keterangan}\n` : "";
  return `Selamat ${greeting} ${recipients}\n${ket}\n${todaySec}\n\n\n${tmrwSec}\n${aSec}${cSec}${nSec}${phSec}\n\nDemikian laporan kamar bedah, terima kasih🙏`;
}
