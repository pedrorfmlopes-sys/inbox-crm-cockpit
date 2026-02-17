import type { SkinId } from "../settings";

export type SkinTokens = Record<string, string>;

const CLASSIC: SkinTokens = {
  "--iccc-font": "Segoe UI Variable Text, Segoe UI, system-ui, -apple-system, sans-serif",
  "--iccc-text": "#0b2d6b",
  "--iccc-text-muted": "rgba(11,45,107,0.70)",
  "--iccc-bg": "#eff6ff",

  "--iccc-card-bg": "rgba(255,255,255,0.85)",
  "--iccc-card-border": "rgba(11,45,107,0.12)",
  "--iccc-shadow": "0 1px 10px rgba(11,45,107,0.06)",

  "--iccc-bottom-bg": "rgba(255,255,255,0.68)",
  "--iccc-bottom-border": "rgba(11,45,107,0.14)",
  "--iccc-bottom-shadow": "0 8px 26px rgba(11,45,107,0.12)",
  "--iccc-bottom-radius": "18px",

  "--iccc-radius-card": "16px",
  "--iccc-radius-pill": "999px",
  "--iccc-radius-btn": "12px",

  "--iccc-pill-bg": "rgba(255,255,255,0.70)",
  "--iccc-pill-border": "rgba(11,45,107,0.16)",
  "--iccc-pill-text": "#0b2d6b",
  "--iccc-pill-active-bg": "#0b2d6b",
  "--iccc-pill-active-text": "#ffffff",
  "--iccc-pill-active-border": "rgba(11,45,107,0.20)",

  "--iccc-btn-bg": "#0b2d6b",
  "--iccc-btn-text": "#ffffff",
  "--iccc-btn-border": "rgba(11,45,107,0.20)",

  "--iccc-btn2-bg": "rgba(11,45,107,0.10)",
  "--iccc-btn2-text": "#0b2d6b",
  "--iccc-btn2-border": "rgba(11,45,107,0.20)",

  "--iccc-weight": "500",
  "--iccc-weight-strong": "600",
  "--iccc-weight-heavy": "700",
};

const MAILMAESTRO: SkinTokens = {
  "--iccc-font": "Segoe UI Variable Text, Segoe UI, system-ui, -apple-system, sans-serif",
  "--iccc-text": "#111827",
  "--iccc-text-muted": "rgba(17,24,39,0.60)",
  "--iccc-bg": "#ffffff",

  "--iccc-card-bg": "#ffffff",
  "--iccc-card-border": "rgba(17,24,39,0.10)",
  "--iccc-shadow": "0 1px 8px rgba(17,24,39,0.10)",

  "--iccc-bottom-bg": "#ffffff",
  "--iccc-bottom-border": "rgba(17,24,39,0.12)",
  "--iccc-bottom-shadow": "0 10px 30px rgba(17,24,39,0.14)",
  "--iccc-bottom-radius": "16px",

  "--iccc-radius-card": "14px",
  "--iccc-radius-pill": "10px",
  "--iccc-radius-btn": "10px",

  "--iccc-pill-bg": "#ffffff",
  "--iccc-pill-border": "rgba(17,24,39,0.12)",
  "--iccc-pill-text": "#111827",
  "--iccc-pill-active-bg": "rgba(109,40,217,0.10)",
  "--iccc-pill-active-text": "#4c1d95",
  "--iccc-pill-active-border": "rgba(109,40,217,0.25)",

  "--iccc-btn-bg": "#111827",
  "--iccc-btn-text": "#ffffff",
  "--iccc-btn-border": "rgba(17,24,39,0.18)",

  "--iccc-btn2-bg": "#ffffff",
  "--iccc-btn2-text": "#111827",
  "--iccc-btn2-border": "rgba(17,24,39,0.12)",

  "--iccc-weight": "450",
  "--iccc-weight-strong": "600",
  "--iccc-weight-heavy": "650",
};

export function getSkinTokens(id: SkinId): SkinTokens {
  return id === "mailmaestro" ? MAILMAESTRO : CLASSIC;
}

export function applySkin(id: SkinId): void {
  const tokens = getSkinTokens(id);
  const root = document.documentElement;
  try {
    root.dataset.icccSkin = id;
  } catch {
    // ignore
  }
  for (const [k, v] of Object.entries(tokens)) {
    try {
      root.style.setProperty(k, v);
    } catch {
      // ignore
    }
  }

  // Ensure base font/weight are applied even when components use inline styles.
  try {
    document.body.style.fontFamily = `var(--iccc-font)`;
    document.body.style.fontWeight = `var(--iccc-weight)`;
    document.body.style.color = `var(--iccc-text)`;
    document.body.style.background = `var(--iccc-bg)`;
  } catch {
    // ignore
  }
}
