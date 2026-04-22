import type { GroupStorageSettings } from "./settings";
import type { ResolvedGroupStorageRuntime } from "./resolveStorageMode";
import type { GroupStorageValidationResult } from "./worksetApi";

export type GroupStorageModeCapability = {
  mode: GroupStorageSettings["mode"];
  label: string;
  supported: boolean;
  howItWrites: string;
  whereItWrites: string;
  limitations: string;
};

function normalizeText(value: string | undefined): string {
  return String(value || "").trim();
}

export function buildGroupStorageValidationPayload(
  settings: GroupStorageSettings
): Record<string, unknown> {
  return {
    mode: settings.mode,
    baseFolderPath: settings.baseFolderPath,
    localDevice: settings.localDevice,
    chosenFolder: settings.chosenFolder,
    hybrid: settings.hybrid,
    chosenFolderKind: settings.chosenFolder.kind,
    primaryTarget: settings.hybrid.primaryTarget,
  };
}

export function isGroupStorageModeActuallySupported(
  mode: GroupStorageSettings["mode"],
  validation: GroupStorageValidationResult | null | undefined
): boolean {
  if (mode === "supabase") return true;
  return validation?.supported === true;
}

export function describeGroupStorageCapabilities(args: {
  settings: GroupStorageSettings;
  runtime: ResolvedGroupStorageRuntime;
  validation?: GroupStorageValidationResult | null;
}): GroupStorageModeCapability[] {
  const validation = args.validation || null;
  const blockedReason = normalizeText(validation?.blockingReason);
  const fileBackedWhere = validation?.supported && validation.normalizedBasePath
    ? validation.normalizedBasePath
    : normalizeText(args.runtime.primaryLocation.basePath) || "por configurar";

  return [
    {
      mode: "supabase",
      label: "Cockpit Cloud",
      supported: true,
      howItWrites: "Persistencia final na app e binario no store cloud quando o payload o inclui.",
      whereItWrites: "Persistencia central `/api/links/*`.",
      limitations: "Nao cria copia local adicional por si so.",
    },
    {
      mode: "local_device",
      label: "Local acessivel ao servidor",
      supported: args.settings.mode === "local_device" ? isGroupStorageModeActuallySupported("local_device", validation) : true,
      howItWrites: "Metadata final continua central; worksets e binario tentam usar path local/UNC acessivel ao servidor.",
      whereItWrites: fileBackedWhere,
      limitations: blockedReason || "Nao representa o disco do utilizador sem bridge nativa; exige caminho acessivel ao processo do servidor.",
    },
    {
      mode: "chosen_folder",
      label: "Pasta local / sincronizada",
      supported: args.settings.mode === "chosen_folder" ? isGroupStorageModeActuallySupported("chosen_folder", validation) : true,
      howItWrites: "Metadata final continua central; ficheiros e mirrors usam a pasta configurada quando o caminho e fisico.",
      whereItWrites: fileBackedWhere,
      limitations: blockedReason || "URL web de OneDrive/SharePoint nao fecha escrita real nesta arquitetura.",
    },
    {
      mode: "hybrid",
      label: "Hibrido",
      supported: args.settings.mode === "hybrid" ? isGroupStorageModeActuallySupported("hybrid", validation) : true,
      howItWrites: "Mantem persistencia central da app e tenta espelhar worksets/binario no destino local primario.",
      whereItWrites: fileBackedWhere,
      limitations: blockedReason || "Continua a exigir path fisico acessivel ao servidor; promocao remota nao substitui IO local inexistente.",
    },
  ];
}
