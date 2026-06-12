# Sistem Koordinasi Kamar Bedah RS Panti Rini

Aplikasi manajemen jadwal dan koordinasi operasi untuk Kamar Bedah Rumah Sakit Panti Rini. Seluruh logika ada di satu file React (App.tsx), data disimpan di localStorage + sinkronisasi Supabase/Dropbox.

## Run & Operate

- `PORT=24074 BASE_PATH=/ pnpm --filter @workspace/kamar-bedah run build` — build React app (output: `artifacts/kamar-bedah/dist/public`)
- `pnpm --filter @workspace/api-server run dev` — jalankan API+static server (port 8080, melayani `/`)
- `pnpm --dir /home/runner/workspace --filter @workspace/kamar-bedah run dev` — Vite dev server (port 24074) dengan HMR
- Setelah edit App.tsx, rebuild → restart api-server

## Stack

- pnpm workspaces, Node.js 24, TypeScript 5.9
- API/Static: Express 5 (port 8080, melayani static React build + `/api` routes)
- React 19 + Vite 7 (build → `artifacts/kamar-bedah/dist/public`)
- DB: localStorage + Supabase realtime sync + Dropbox backup
- Export: xlsx (Excel), Word (manual blob)

## Where things live

- `artifacts/kamar-bedah/src/App.tsx` — seluruh UI + logic (±3600 baris)
- `artifacts/kamar-bedah/dist/public/` — hasil build, dilayani oleh Express
- `artifacts/api-server/src/app.ts` — Express server, serve static + `/api`

## Architecture decisions

- App.tsx satu file: memudahkan portabilitas dan backup
- Express melayani built React app di `/` (bukan Vite proxy) untuk menghindari masalah routing Replit
- Data primer: localStorage perangkat; cloud backup: Supabase (realtime) + Dropbox JSON
- Perawat Onloop disimpan sebagai string ` | ` separated untuk mendukung 1-3 orang

## Product

- **Tab Jadwal**: lihat jadwal operasi hari ini
- **Tab Daftar**: input jadwal baru, termasuk **3 slot Perawat Onloop** (1 wajib, 2-3 opsional)
- **Tab Laporan/Lembur/Monitor**: manajemen laporan & monitoring suhu
- **Tab Arsip → Unduh**: download Excel per bulan dari **Lokal**, **Supabase**, atau **Dropbox**
- **Tab WA**: template pesan WhatsApp otomatis
- **Tab Staf**: manajemen daftar staf

## User preferences

_Populate as you build — explicit user instructions worth remembering across sessions._

## Gotchas

- Setelah edit App.tsx: wajib rebuild (`PORT=24074 BASE_PATH=/ pnpm --filter @workspace/kamar-bedah run build`) lalu restart api-server workflow
- Proxy Replit hanya baca `artifacts/*/.replit-artifact/artifact.toml`; kamar-bedah artifact terdaftar di extracted dir, workaround: Express serve static
- `app.get("/{*splat}", ...)` bukan `app.get("*", ...)` karena Express 5 + path-to-regexp v8

## Pointers

- See the `pnpm-workspace` skill for workspace structure, TypeScript setup, and package details
