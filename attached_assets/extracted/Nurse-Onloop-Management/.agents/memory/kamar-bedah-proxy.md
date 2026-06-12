---
name: Kamar Bedah proxy workaround
description: Why Express serves the React static build instead of Vite dev server routing through the Replit proxy
---

The Replit reverse proxy at port 80 only reads artifact routing from `artifacts/*/replit-artifact/artifact.toml`. The kamar-bedah artifact was imported from a user export and is registered at `attached_assets/extracted/.../kamar-bedah` — outside `artifacts/`. So the proxy never routes `/` to the Vite dev server at port 24074.

**Workaround:** Build the React app (`PORT=24074 BASE_PATH=/ pnpm --filter @workspace/kamar-bedah run build`) and have `artifacts/api-server/src/app.ts` serve the static output at `/` with `express.static` + SPA fallback.

**Why:** `createArtifact` fails if slug directory exists; `verifyAndReplaceArtifactToml` requires the artifact.toml to exist at the target path AND be registered — impossible for a new/moved artifact.

**How to apply:** After any change to `artifacts/kamar-bedah/src/App.tsx`, always rebuild the React app and restart the api-server workflow. Do NOT just restart the kamar-bedah Vite workflow — changes won't be visible through the proxy.
