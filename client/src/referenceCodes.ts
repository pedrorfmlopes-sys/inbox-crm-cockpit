import { getSettings, saveSettings, type CockpitSettingsV1, type ReferenceCodeSettings, type ReferenceEntityKey } from "./settings";
import { searchOdooDomain } from "./api";

type SupportedReferenceModel = "crm.lead" | "project.project" | "project.task" | "helpdesk.ticket";

const MODEL_TO_ENTITY: Record<SupportedReferenceModel, ReferenceEntityKey> = {
  "crm.lead": "lead",
  "project.project": "project",
  "project.task": "task",
  "helpdesk.ticket": "ticket",
};

export function getReferenceEntityKey(model: string): ReferenceEntityKey | null {
  return (MODEL_TO_ENTITY as Record<string, ReferenceEntityKey | undefined>)[model] || null;
}

export function formatReferenceCode(settings: ReferenceCodeSettings, entity: ReferenceEntityKey, sequence: number, now = new Date()): string {
  const parts = [
    String(settings.prefixes[entity] || "").trim(),
    settings.includeYear ? String(now.getFullYear()) : "",
    String(sequence).padStart(4, "0"),
  ].filter(Boolean);

  return parts.join("-");
}

export function applyReferenceCodeToTitle(title: string, code: string, position: ReferenceCodeSettings["position"]): string {
  const cleanTitle = String(title || "").trim();
  if (!code) return cleanTitle;
  if (!cleanTitle) return code;
  return position === "suffix" ? `${cleanTitle} [${code}]` : `[${code}] ${cleanTitle}`;
}

export function previewReferenceCode(settings: CockpitSettingsV1, entity: ReferenceEntityKey): string {
  const nextSequence = settings.referenceCodes.counterMode === "global"
    ? (settings.referenceCodes.counters.global || 0) + 1
    : ((settings.referenceCodes.counters.perType?.[entity] || 0) + 1);
  const code = formatReferenceCode(settings.referenceCodes, entity, nextSequence);
  return applyReferenceCodeToTitle(exampleBaseTitle(entity), code, settings.referenceCodes.position);
}

function exampleBaseTitle(entity: ReferenceEntityKey): string {
  if (entity === "lead") return "Lead do email";
  if (entity === "project") return "Projeto do email";
  if (entity === "task") return "Tarefa do email";
  return "Ticket do email";
}

function nextSequence(settings: CockpitSettingsV1, entity: ReferenceEntityKey): number {
  return settings.referenceCodes.counterMode === "global"
    ? (settings.referenceCodes.counters.global || 0) + 1
    : ((settings.referenceCodes.counters.perType?.[entity] || 0) + 1);
}

function applyReservedSequence(settings: CockpitSettingsV1, entity: ReferenceEntityKey, sequence: number): CockpitSettingsV1 {
  if (settings.referenceCodes.counterMode === "global") {
    return {
      ...settings,
      referenceCodes: {
        ...settings.referenceCodes,
        counters: {
          ...settings.referenceCodes.counters,
          global: sequence,
        },
      },
    };
  }

  return {
    ...settings,
    referenceCodes: {
      ...settings.referenceCodes,
      counters: {
        ...settings.referenceCodes.counters,
        perType: {
          ...settings.referenceCodes.counters.perType,
          [entity]: sequence,
        },
      },
    },
  };
}

async function titleExistsInOdoo(model: SupportedReferenceModel, title: string): Promise<boolean> {
  const rows = await searchOdooDomain(model, [["name", "=", title]], ["id", "name"], 1).catch(() => []);
  return Array.isArray(rows) && rows.length > 0;
}

export async function prepareReferencedRecordName(model: string, baseTitle: string): Promise<{
  title: string;
  referenceCode: string | null;
}> {
  const entity = getReferenceEntityKey(model);
  if (!entity) {
    return { title: String(baseTitle || "").trim(), referenceCode: null };
  }

  const settings = await getSettings();
  if (!settings.referenceCodes.enabled) {
    return { title: String(baseTitle || "").trim(), referenceCode: null };
  }

  let workingSettings = settings;
  let sequence = nextSequence(workingSettings, entity);

  for (let attempt = 0; attempt < 50; attempt += 1) {
    const code = formatReferenceCode(workingSettings.referenceCodes, entity, sequence);
    const candidateTitle = applyReferenceCodeToTitle(baseTitle, code, workingSettings.referenceCodes.position);
    const exists = await titleExistsInOdoo(model as SupportedReferenceModel, candidateTitle);
    if (!exists) {
      await saveSettings({
        referenceCodes: applyReservedSequence(workingSettings, entity, sequence).referenceCodes,
      });
      return { title: candidateTitle, referenceCode: code };
    }

    sequence += 1;
    workingSettings = applyReservedSequence(workingSettings, entity, sequence - 1);
  }

  throw new Error("Nao foi possivel gerar um codigo de referencia unico.");
}
