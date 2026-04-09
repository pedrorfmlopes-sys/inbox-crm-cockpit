import type { GroupStorageMode } from "./types";

export const GROUP_STORAGE_MODE_OPTIONS: GroupStorageMode[] = [
  "supabase",
  "local_device",
  "chosen_folder",
  "hybrid",
];

export const GROUP_STORAGE_MODE_LABELS: Record<GroupStorageMode, string> = {
  supabase: "Tudo no Supabase",
  local_device: "Local neste PC",
  chosen_folder: "Local em pasta escolhida",
  hybrid: "Hibrido",
};

export const GROUP_STORAGE_MODE_DESCRIPTIONS: Record<GroupStorageMode, string> = {
  supabase: "Persistencia principal e promotavel no Supabase, sem sessao local passar a fonte canonica.",
  local_device: "Persistencia principal em storage local deste dispositivo, com Supabase apenas por promocao futura.",
  chosen_folder: "Persistencia principal numa pasta escolhida pelo utilizador, com Supabase apenas por promocao futura.",
  hybrid: "Persistencia principal local ou pasta escolhida, com promocao remota controlada e separada.",
};
