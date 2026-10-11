/* src/types/operation.ts
   Tipe domain data utama aplikasi: jadwal operasi, staf, roster harian,
   dan notifikasi. Dipindah dari App.tsx (langkah 1 refactor — pisah types)
   — TIDAK ADA perubahan logika, murni pindah lokasi. */

export interface Operation {
  id: string;
  patient: string;
  age?: string;
  ageMonths?: string;
  rm?: string;
  opType: "elektif" | "semi" | "cyto";
  status: "scheduled" | "ongoing" | "done" | "batal";
  date: string;
  time: string;
  room: string;
  surgeon: string;
  anesthesiologist: string;
  procedure: string;
  diagnosis: string;
  allergy: string;
  bloodType: string;
  ruangAsal?: string;
  assistantNurse?: string;
  circulatingNurse?: string;
  anesthesiaNurse?: string;
  onloopNurse?: string;
  rrKatim?: string;
  specialNeeds?: string;
  createdAt?: string;
  updated_at?: string;
  reminders?: string[];
  requests?: { id: string; text: string; time: string }[];
  cancelReason?: string;
  needsReview?: boolean;
}

export interface StaffMember {
  id: string;
  name: string;
  role: string;
  phone?: string;
  email?: string;
  specialization?: string;
  createdAt?: string;
  updated_at?: string;
}

export interface RosterEntry {
  id: string;
  staffId: string;
  date: string;
  shift: string;
  ruang?: string;
  anestJagaList?: string[];
  anestJaga?: string;
  anestCytoList?: string[];
  anestCyto?: string;
  nurses?: string[];
  pembawaHP?: string;
  createdAt?: string;
  updated_at?: string;
}

export interface Notif {
  id: string;
  type: string;
  label: string;
  patient: string;
  procedure: string;
  message: string;
  sentAt: string;
}
