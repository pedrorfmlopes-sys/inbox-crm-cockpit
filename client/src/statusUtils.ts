import type React from "react";

export type UnifiedStatusColorKey = "blue" | "amber" | "green" | "red";

export const UNIFIED_STATUS_COLOR_MAP: Record<UnifiedStatusColorKey, React.CSSProperties> = {
  blue: { borderColor: "rgba(59,130,246,0.34)", background: "rgba(219,234,254,0.92)", color: "#1d4ed8" },
  amber: { borderColor: "rgba(245,158,11,0.3)", background: "rgba(254,243,199,0.95)", color: "#b45309" },
  green: { borderColor: "rgba(34,197,94,0.28)", background: "rgba(220,252,231,0.95)", color: "#15803d" },
  red: { borderColor: "rgba(239,68,68,0.26)", background: "rgba(254,226,226,0.95)", color: "#b91c1c" },
};

export const UNIFIED_STATUS_LEGEND: Array<{ key: UnifiedStatusColorKey; label: string; style: React.CSSProperties }> = [
  { key: "blue", label: "Azul = Em analise", style: UNIFIED_STATUS_COLOR_MAP.blue },
  { key: "amber", label: "Amarelo = Aguarda", style: UNIFIED_STATUS_COLOR_MAP.amber },
  { key: "green", label: "Verde = Concluido", style: UNIFIED_STATUS_COLOR_MAP.green },
  { key: "red", label: "Vermelho = Bloqueado", style: UNIFIED_STATUS_COLOR_MAP.red },
];

/**
 * Retorna a configuracao de apresentacao (label e cor) para um status,
 * suportando aliases legados mas mantendo a apresentacao unificada.
 */
export function getStatusDisplayConfig(status: string | undefined): { label: string; color: UnifiedStatusColorKey; style: React.CSSProperties } {
  const normalized = String(status || "").trim().toLowerCase();
  
  // Blue Group: Novo / Em analise
  if (!normalized || ["em_analise", "aberto", "open", "novo", "new"].includes(normalized)) {
    let label = "Em analise";
    if (normalized === "aberto" || normalized === "open") label = "Aberto";
    if (normalized === "novo" || normalized === "new") label = "Novo";
    return { label, color: "blue", style: UNIFIED_STATUS_COLOR_MAP.blue };
  }

  // Amber Group: Em progresso / Aguarda
  if (["em_progresso", "progresso", "aguarda", "pending", "in_progress"].includes(normalized)) {
    let label = "Em progresso";
    if (normalized === "aguarda") label = "Aguarda";
    if (normalized === "pending") label = "Pendente";
    return { label, color: "amber", style: UNIFIED_STATUS_COLOR_MAP.amber };
  }

  // Green Group: Concluido
  if (["concluido", "concluido.", "resolvido", "done", "resolved", "completed"].includes(normalized)) {
    let label = "Concluido";
    if (normalized === "resolvido" || normalized === "resolved") label = "Resolvido";
    if (normalized === "done") label = "Finalizado";
    return { label, color: "green", style: UNIFIED_STATUS_COLOR_MAP.green };
  }

  // Red Group: Fechado / Bloqueado
  if (["bloqueado", "fechado", "closed", "blocked", "cancelado", "cancelled"].includes(normalized)) {
    let label = "Fechado";
    if (normalized === "bloqueado" || normalized === "blocked") label = "Bloqueado";
    if (normalized === "cancelado" || normalized === "cancelled") label = "Cancelado";
    return { label, color: "red", style: UNIFIED_STATUS_COLOR_MAP.red };
  }

  // Fallback para outros status customizados
  return { label: normalized, color: "blue", style: UNIFIED_STATUS_COLOR_MAP.blue };
}
