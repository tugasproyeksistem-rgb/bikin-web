/* src/types/misc.ts
   Tipe lain-lain yang tidak masuk kategori domain besar manapun: item
   antrean offline (kb_operasi) dan event PWA install prompt browser.
   Dipindah dari App.tsx, tidak ada perubahan logika. */

export type OfflineQueueItem = { id: string; payload: any; queuedAt: number };

export interface BeforeInstallPromptEvent extends Event {
  prompt: () => Promise<void>;
  userChoice: Promise<{ outcome: "accepted" | "dismissed" }>;
}
