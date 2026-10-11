/* src/constants/hospital.ts
   Konstanta identitas RS & data referensi form operasi (daftar ruang,
   golongan darah, label peran staf, jenis operasi, status operasi, template
   form kosong). Dipindah dari App.tsx (langkah 2 refactor — pisah
   constants), tidak ada perubahan nilai. */

import type { Operation } from "../types";

export const CONFIG = {
  HOSPITAL: "Rumah Sakit Panti Rini",
  DEFAULT_PIN: "1234",
};
export const HOSPITAL = CONFIG.HOSPITAL;
export const LOCK_MS  = 10 * 60 * 1000;

export const ROOMS = ["Kamar Operasi 1","Kamar Operasi 2","Kamar Operasi 3","Kamar Operasi 4","Kamar Operasi 5"];
export const BT    = ["A+","A-","B+","B-","AB+","AB-","O+","O-","Tidak Diketahui"];
export const ST    = { surgeon:"Dokter Bedah", anesthesiologist:"Dokter Anestesi", circulating:"PERAWAT Instrumen", anesthesia_nurse:"Perawat Anestesi", onloop:"Perawat Onloop", katim:"RR / Katim" };
export const OT    = { elektif:{label:"Elektif",c:"#1565C0",bg:"#E3F2FD"}, semi:{label:"Semi-Elektif",c:"#E65100",bg:"#FFF3E0"}, cyto:{label:"⚠ CYTO",c:"#B71C1C",bg:"#FFCDD2"} };
export const STS   = { scheduled:{l:"Terjadwal",icon:"●",c:"#1565C0",bg:"#E3F2FD"}, ongoing:{l:"Berlangsung",icon:"◐",c:"#00897B",bg:"#E0F2F1"}, done:{l:"Selesai",icon:"✓",c:"#2E7D32",bg:"#E8F5E9"}, batal:{l:"Batal/Tunda",icon:"✕",c:"#C62828",bg:"#FFEBEE"} };
export const EOP: Partial<Operation> = {patient:"",age:"",ageMonths:"",rm:"",opType:"elektif",diagnosis:"",procedure:"",ruangAsal:"",room:ROOMS[0],date:"",time:"",surgeon:"",anesthesiologist:"",assistantNurse:"",circulatingNurse:"",anesthesiaNurse:"",onloopNurse:"",rrKatim:"",allergy:"Tidak Ada",specialNeeds:"",bloodType:"O+"};
export const PAGE_SIZE = 10;
export const DEFAULT_RECIPIENT = "Suster Thresmiati CB, bu Niken, pak Jaka dan teman sejawat, mohon ijin laporan kamar bedah:";
