/* src/components/install-prompt.tsx
   Banner ajakan install PWA — Android/desktop pakai event browser native
   (beforeinstallprompt), iOS pakai instruksi manual (Share > Add to Home
   Screen, karena Safari tidak mendukung event ini). Dipindah dari App.tsx
   (langkah 5 refactor — pisah components, batch 4), TIDAK ADA perubahan
   logika. */

import { useState, useEffect } from "react";
import type { BeforeInstallPromptEvent } from "../types";
import { KB_INSTALL_DISMISS_KEY, KB_INSTALL_DISMISS_MS } from "../constants";

function kbIsStandalone(): boolean {
  return (
    window.matchMedia("(display-mode: standalone)").matches ||
    (window.navigator as unknown as { standalone?: boolean }).standalone === true
  );
}

function kbIsIos(): boolean {
  const ua = window.navigator.userAgent;
  const isIosDevice = /iphone|ipad|ipod/i.test(ua);
  const isIpadOs = /macintosh/i.test(ua) && "ontouchend" in document;
  return isIosDevice || isIpadOs;
}

function kbWasDismissedRecently(): boolean {
  const raw = localStorage.getItem(KB_INSTALL_DISMISS_KEY);
  if (!raw) return false;
  const dismissedAt = Number(raw);
  if (Number.isNaN(dismissedAt)) return false;
  return Date.now() - dismissedAt < KB_INSTALL_DISMISS_MS;
}

export function InstallPrompt() {
  const [deferredPrompt, setDeferredPrompt] = useState<BeforeInstallPromptEvent | null>(null);
  const [showIosBanner, setShowIosBanner] = useState(false);
  const [visible, setVisible] = useState(false);

  useEffect(() => {
    if (kbIsStandalone() || kbWasDismissedRecently()) return;

    const handleBeforeInstallPrompt = (e: Event) => {
      e.preventDefault();
      setDeferredPrompt(e as BeforeInstallPromptEvent);
      setVisible(true);
    };
    window.addEventListener("beforeinstallprompt", handleBeforeInstallPrompt);

    if (kbIsIos()) {
      setShowIosBanner(true);
      setVisible(true);
    }

    return () => window.removeEventListener("beforeinstallprompt", handleBeforeInstallPrompt);
  }, []);

  const handleInstallClick = async () => {
    if (!deferredPrompt) return;
    await deferredPrompt.prompt();
    const choice = await deferredPrompt.userChoice;
    if (choice.outcome === "accepted") setVisible(false);
    setDeferredPrompt(null);
  };

  const handleDismiss = () => {
    localStorage.setItem(KB_INSTALL_DISMISS_KEY, String(Date.now()));
    setVisible(false);
  };

  if (!visible) return null;

  return (
    <div style={{
      position: "fixed", left: "12px", right: "12px", bottom: "12px", zIndex: 9999,
      background: "#16685F", color: "white", borderRadius: "12px", padding: "14px 16px",
      boxShadow: "0 4px 16px rgba(0,0,0,0.25)", display: "flex", alignItems: "center",
      gap: "12px", fontSize: "14px", lineHeight: 1.4,
    }}>
      <div style={{ flex: 1 }}>
        {showIosBanner ? (
          <span>Instal aplikasi ini: tekan tombol <strong>Share</strong> di Safari, lalu pilih <strong>"Add to Home Screen"</strong>.</span>
        ) : (
          <span>Instal Sistem Kamar Bedah ke layar utama untuk akses lebih cepat.</span>
        )}
      </div>
      {!showIosBanner && (
        <button onClick={handleInstallClick} style={{
          background: "white", color: "#16685F", border: "none", borderRadius: "8px",
          padding: "8px 14px", fontWeight: 600, fontSize: "14px", cursor: "pointer", whiteSpace: "nowrap",
        }}>Instal</button>
      )}
      <button onClick={handleDismiss} aria-label="Tutup" style={{
        background: "transparent", border: "none", color: "white", fontSize: "18px",
        cursor: "pointer", padding: "0 4px", lineHeight: 1,
      }}>×</button>
    </div>
  );
}
