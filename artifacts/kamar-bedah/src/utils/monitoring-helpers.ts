/* src/utils/monitoring-helpers.ts
   Helper murni untuk fitur monitoring suhu/kelembaban kamar bedah.
   Dipindah dari App.tsx (diekstrak bersamaan dengan ViewArsip yang
   memakainya), TIDAK ADA perubahan logika. */

import type { MonitoringCfg } from "../types";

export function monIsOK(suhu:number, rh:number, cfg:MonitoringCfg):boolean {
  return suhu>=cfg.suhuMin && suhu<=cfg.suhuMax && rh>=cfg.rhMin && rh<=cfg.rhMax;
}
