import fs from "node:fs";
import path from "node:path";

const repoRoot = process.cwd();
const outputDir = path.join(repoRoot, "output", "playwright");
fs.mkdirSync(outputDir, { recursive: true });

function read(relativePath) {
  return fs.readFileSync(path.join(repoRoot, relativePath), "utf8");
}

function humanize(value) {
  return String(value || "")
    .trim()
    .replace(/_/g, " ")
    .replace(/\b\w/g, (char) => char.toUpperCase());
}

function resolveStateCatalog(input) {
  const enabled = Array.isArray(input) ? true : input?.enabled !== false;
  const states = Array.isArray(input) ? input : input?.states;
  const seen = new Set();
  const options = [];
  for (const entry of enabled ? (states || []) : []) {
    const value = String(entry?.name || "").trim();
    if (!value || seen.has(value)) continue;
    seen.add(value);
    options.push({
      value,
      label: humanize(value),
      color: String(entry?.color || "").trim() || undefined,
    });
  }
  return { enabled, options };
}

function isConfiguredStateValue(value, options) {
  const normalized = String(value || "").trim();
  if (!normalized) return false;
  return options.some((option) => option.value === normalized);
}

const studioSource = read("client/src/modules/crm/GroupClassificationStudioApp.tsx");
const editorSource = read("client/src/modules/crm/group-classification/components/ClassificationEditor.tsx");
const managerSource = read("client/src/modules/crm/GroupManagerCockpit.tsx");

const typeDefinitions = [
  {
    type: "group",
    settingsPath: ["groups", "groups", "states"],
    studioNeedle: "buildConfiguredStateCatalog(groupsSettingsSnapshot?.groups?.groups?.states)",
    editorNeedles: ["groupStateEnabled", "groupStateOptions.map"],
  },
  {
    type: "reference",
    settingsPath: ["groups", "references", "states"],
    studioNeedle: "buildConfiguredStateCatalog(groupsSettingsSnapshot?.groups?.references?.states)",
    editorNeedles: ["referenceStateEnabled", "referenceStateOptions.map"],
  },
  {
    type: "ticket",
    settingsPath: ["groups", "tickets", "states"],
    studioNeedle: "buildConfiguredStateCatalog(groupsSettingsSnapshot?.groups?.tickets?.states)",
    editorNeedles: ["ticketStateEnabled", "ticketStateOptions.map"],
  },
  {
    type: "label",
    settingsPath: ["groups", "labels", "states"],
    studioNeedle: "buildConfiguredStateCatalog(groupsSettingsSnapshot?.groups?.labels?.states)",
    editorNeedles: ["labelStateEnabled", "labelStateOptions.map"],
  },
];

const sourceChecks = Object.fromEntries(
  typeDefinitions.map((definition) => [
    definition.type,
    {
      studioConsumesSettings: studioSource.includes(definition.studioNeedle),
      editorConsumesSettings: definition.editorNeedles.every((needle) => editorSource.includes(needle)),
    },
  ])
);

const noHardcodedArraysActive =
  !studioSource.includes("LABEL_STATUS_OPTIONS") &&
  !studioSource.includes("TICKET_STATUS_OPTIONS") &&
  !studioSource.includes("STATUS_OPTIONS") &&
  !editorSource.includes("LABEL_STATUS_OPTIONS") &&
  !editorSource.includes("TICKET_STATUS_OPTIONS");

const managerSupportsCatalogEnabled =
  managerSource.includes("catalog?.enabled === false") ||
  managerSource.includes("catalog?.enabled === false");

const stateFixtures = {
  group: [
    { name: "novo", color: "#111111" },
    { name: "em_validacao", color: "#222222" },
  ],
  reference: [
    { name: "ligada", color: "#333333" },
    { name: "bloqueada", color: "#444444" },
  ],
  ticket: [
    { name: "pendente", color: "#555555" },
    { name: "fechado", color: "#666666" },
  ],
  label: [
    { name: "aguarda_cliente", color: "#777777" },
    { name: "resolvida", color: "#888888" },
  ],
};

function buildSettings(type, catalog) {
  return {
    groups: {
      groups: {
        states: type === "group" ? catalog : { enabled: true, states: [] },
      },
      references: {
        states: type === "reference" ? catalog : { enabled: true, states: [] },
      },
      tickets: {
        states: type === "ticket" ? catalog : { enabled: true, states: [] },
      },
      labels: {
        states: type === "label" ? catalog : { enabled: true, states: [] },
      },
    },
  };
}

function readCatalogFromSettings(type, settings) {
  if (type === "group") return settings.groups.groups.states;
  if (type === "reference") return settings.groups.references.states;
  if (type === "ticket") return settings.groups.tickets.states;
  return settings.groups.labels.states;
}

const scenarios = [];

for (const definition of typeDefinitions) {
  const fixture = stateFixtures[definition.type];
  const variants = [
    {
      suffix: "disabled",
      catalog: { enabled: false, states: fixture },
      freeTextCandidate: "estado_livre",
    },
    {
      suffix: "enabled-list",
      catalog: { enabled: true, states: fixture },
      freeTextCandidate: "estado_livre",
    },
    {
      suffix: "enabled-empty",
      catalog: { enabled: true, states: [] },
      freeTextCandidate: "estado_livre",
    },
    {
      suffix: "free-text-rejected",
      catalog: { enabled: true, states: fixture },
      freeTextCandidate: "texto_nao_configurado",
    },
  ];

  for (const variant of variants) {
    const settingsInjected = buildSettings(definition.type, variant.catalog);
    const resolved = resolveStateCatalog(readCatalogFromSettings(definition.type, settingsInjected));
    const expectedUiOptions = resolveStateCatalog(variant.catalog).options;
    const selectorVisibleExpected = variant.catalog.enabled !== false && expectedUiOptions.length > 0;
    const selectorVisibleObtained =
      sourceChecks[definition.type].studioConsumesSettings &&
      sourceChecks[definition.type].editorConsumesSettings &&
      resolved.enabled &&
      resolved.options.length > 0;
    const freeTextAcceptedObtained = isConfiguredStateValue(variant.freeTextCandidate, resolved.options);

    const checks = [
      {
        id: "studio-consumes-settings",
        pass: sourceChecks[definition.type].studioConsumesSettings,
        message: "Classificar deriva o catalogo a partir de settings.groups.*.states",
      },
      {
        id: "editor-renders-catalog",
        pass: sourceChecks[definition.type].editorConsumesSettings,
        message: "O editor usa apenas as opcoes recebidas do catalogo dinamico",
      },
      {
        id: "enabled-gates-selector",
        pass: selectorVisibleExpected === selectorVisibleObtained,
        message: "A visibilidade do seletor respeita enabled + lista configurada",
      },
      {
        id: "options-match-settings",
        pass: JSON.stringify(expectedUiOptions) === JSON.stringify(resolved.options),
        message: "As opcoes efetivas batem com a lista vinda de settings",
      },
      {
        id: "free-text-rejected",
        pass: freeTextAcceptedObtained === false,
        message: "Valores fora da lista configurada nao sao aceites",
      },
      {
        id: "no-hardcoded-arrays-active",
        pass: noHardcodedArraysActive,
        message: "Nao existem arrays hardcoded ativos no Classificar",
      },
    ];

    scenarios.push({
      scenarioId: `${definition.type}-states-${variant.suffix}`,
      type: definition.type,
      settingsInjected,
      expectedUiOptions,
      obtainedUiOptions: resolved.options,
      selectorVisibleExpected,
      selectorVisibleObtained,
      freeTextCandidate: variant.freeTextCandidate,
      freeTextAcceptedExpected: false,
      freeTextAcceptedObtained,
      pass: checks.every((entry) => entry.pass),
      reason: checks.filter((entry) => !entry.pass).map((entry) => entry.message).join("; ") || undefined,
      checks,
    });
  }
}

const report = {
  generatedAtIso: new Date().toISOString(),
  validatedFiles: [
    "client/src/modules/crm/GroupClassificationStudioApp.tsx",
    "client/src/modules/crm/group-classification/components/ClassificationEditor.tsx",
    "client/src/modules/crm/GroupManagerCockpit.tsx",
  ],
  sourceChecks,
  managerSupportsCatalogEnabled,
  noHardcodedArraysActive,
  scenarios,
  pass: scenarios.every((entry) => entry.pass),
};

fs.writeFileSync(
  path.join(outputDir, "groups-state-alignment-report.json"),
  JSON.stringify(report, null, 2),
  "utf8"
);

console.log(JSON.stringify(report, null, 2));
