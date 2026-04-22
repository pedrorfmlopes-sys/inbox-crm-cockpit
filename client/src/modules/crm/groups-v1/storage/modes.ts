import type { GroupStorageMode } from "./types";

export const GROUP_STORAGE_MODE_OPTIONS: GroupStorageMode[] = [
  "supabase",
  "local_device",
  "chosen_folder",
  "hybrid",
];

export const GROUP_STORAGE_MODE_LABELS: Record<GroupStorageMode, string> = {
  supabase: "Cockpit Cloud",
  local_device: "Local acessivel ao servidor",
  chosen_folder: "Pasta local / sincronizada",
  hybrid: "Hibrido",
};

export const GROUP_STORAGE_MODE_DESCRIPTIONS: Record<GroupStorageMode, string> = {
  supabase: "Persistencia final atual da app, com metadata sempre e binario no store cloud quando o payload o traz.",
  local_device: "Persistencia final central + tentativa real de mirror/binario em caminho local/UNC acessivel ao servidor.",
  chosen_folder: "Persistencia final central + mirror/binario para pasta local ou sincronizada por caminho fisico validado.",
  hybrid: "Persistencia final central com destino primario file-backed validado para worksets e binario.",
};
