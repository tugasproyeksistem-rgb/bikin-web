/* src/constants/roster.ts
   Konstanta untuk fitur roster generator: seed data awal perawat RS Panti
   Rini, kamus kode shift, nama tabel Supabase, kolom ringkasan Excel, dan
   key localStorage draft. Dipindah dari App.tsx, tidak ada perubahan nilai.
   Bergantung ke NurseMaster (types/roster.ts) untuk SEED_MASTER. */

import type { NurseMaster } from "../types";

export const MASTER_PERAWAT_LIST: Array<{nik:string; name:string; pk:string}> = [
  { nik:"201010056", name:"Jaka R",        pk:"III"  },
  { nik:"199720045", name:"Eka S",         pk:"III"  },
  { nik:"200315032", name:"Niken A",       pk:"II"   },
  { nik:"199805021", name:"Thresmiati CB", pk:"IV"   },
  { nik:"200110078", name:"Yohana P",      pk:"III"  },
  { nik:"200812034", name:"Agus W",        pk:"II"   },
  { nik:"199912067", name:"Dewi R",        pk:"III"  },
  { nik:"200508043", name:"Sari U",        pk:"II"   },
  { nik:"201205089", name:"Budi S",        pk:"I"    },
  { nik:"200001055", name:"Ratna K",       pk:"III"  },
];

/* Kamus kode shift */
export const SHIFT_CODES = { PAGI:"P", PAGI_SIAGA:"PG", SIANG:"S", SIANG_SIAGA:"SG", LIBUR:"L", LIBUR_SIAGA:"LG", CUTI:"T" } as const;

export const KB_ROSTER_MASTER_TABLE = "kb_roster_master";

/* Seed awal: data RS Panti Rini (hanya dipakai utk first-run seed Supabase
   bila tabel kb_roster_master masih kosong — BUKAN fallback offline). */
export const SEED_MASTER: NurseMaster[] = [
  { id:"m1",  nik:"201010056", name:"Jaka R",        pk:"III" },
  { id:"m2",  nik:"199720045", name:"Eka S",          pk:"III" },
  { id:"m3",  nik:"200315032", name:"Niken A",        pk:"II"  },
  { id:"m4",  nik:"199805021", name:"Thresmiati CB",  pk:"IV"  },
  { id:"m5",  nik:"200110078", name:"Yohana P",       pk:"III" },
  { id:"m6",  nik:"200812034", name:"Agus W",         pk:"II"  },
  { id:"m7",  nik:"199912067", name:"Dewi R",         pk:"III" },
  { id:"m8",  nik:"200508043", name:"Sari U",         pk:"II"  },
  { id:"m9",  nik:"201205089", name:"Budi S",         pk:"I"   },
  { id:"m10", nik:"200001055", name:"Ratna K",        pk:"III" },
];

/* Kolom ringkasan individu di sisi kanan tabel Excel (7 kolom sesuai UI) */
export const SUMMARY_COLS = ["∑P","∑PG","∑S","∑SG","∑L/LG","∑T","∑Total"] as const;

/* localStorage key untuk draft jadwal roster (offline resilience) */
export const LS_DRAFT_KEY = "rostergen_draft_v1";
