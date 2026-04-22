import type { GroupStorageMode } from "./types";

export const GROUP_STORAGE_MODE_OPTIONS: GroupStorageMode[] = [
  "supabase",
  "local_device",
  "chosen_folder",
  "hybrid",
];

export const GROUP_STORAGE_MODE_LABELS: Record<GroupStorageMode, string> = {
  supabase: "Cockpit Cloud",
  local_device: "Local neste PC (nao suportado nesta fase)",
  chosen_folder: "Pasta local / sincronizada",
  hybrid: "Hibrido (nao suportado nesta fase)",
};

export const GROUP_STORAGE_MODE_DESCRIPTIONS: Record<GroupStorageMode, string> = {
  supabase: "Persistencia final atual da app, com metadata sempre e binario no store cloud quando o payload o traz.",
  local_device: "Modo legado ainda nao fechado como opcao executavel nesta fase.",
  chosen_folder: "Persistencia binaria best-effort para pasta local ou sincronizada por caminho local/UNC, mantendo a persistencia final da app.",
  hybrid: "Modo legado ainda nao fechado como opcao executavel nesta fase.",
};
