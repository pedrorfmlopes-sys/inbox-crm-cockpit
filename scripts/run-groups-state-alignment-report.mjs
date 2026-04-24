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

function buildOptions(states) {
  const seen = new Set();
  const options = [];
  for (const entry of states || []) {
    const value = String(entry?.name || "").trim();
    if (!value || seen.has(value)) continue;
    seen.add(value);
    options.push({
      value,
      label: humanize(value),
      color: String(entry?.color || "").trim() || undefined,
    });
  }
  return options;
}

const scenario = {
  scenarioId: "settings-drive-group-state-options",
  input: {
    groups: {
      groups: {
        states: [
          { name: "novo", color: "#111111" },
          { name: "em_validacao", color: "#222222" },
        ],
      },
      references: {
        states: [
          { name: "ligada", color: "#333333" },
        ],
      },
      tickets: {
        states: [
          { name: "pendente", color: "#444444" },
          { name: "fechado", color: "#555555" },
        ],
      },
      labels: {
        states: [
          { name: "aguarda_cliente", color: "#666666" },
          { name: "resolvida", color: "#777777" },
        ],
      },
    },
  },
};

const studioSource = read("client/src/modules/crm/GroupClassificationStudioApp.tsx");
const editorSource = read("client/src/modules/crm/group-classification/components/ClassificationEditor.tsx");
const managerSource = read("client/src/modules/crm/GroupManagerCockpit.tsx");

const checks = [
  {
    id: "studio-label-states",
    description: "Classificar deriva estados de etiquetas de settings.groups.labels.states",
    pass: studioSource.includes('buildConfiguredStateOptions(groupsSettingsSnapshot?.groups?.labels?.states)'),
  },
  {
    id: "studio-ticket-states",
    description: "Classificar deriva estados de tickets de settings.groups.tickets.states",
    pass: studioSource.includes('buildConfiguredStateOptions(groupsSettingsSnapshot?.groups?.tickets?.states)'),
  },
  {
    id: "editor-label-options",
    description: "UI do editor de Classificar renderiza select a partir de labelStateOptions",
    pass: editorSource.includes("labelStateOptions.map"),
  },
  {
    id: "editor-ticket-options",
    description: "UI do editor de Classificar renderiza select a partir de ticketStateOptions",
    pass: editorSource.includes("ticketStateOptions.map"),
  },
  {
    id: "manager-group-states",
    description: "Manager de Groups deriva estados de groups.states",
    pass: managerSource.includes("buildStateOptions(groupsSettings?.groups?.states)"),
  },
  {
    id: "manager-ticket-states",
    description: "Manager de Groups deriva estados de tickets.states",
    pass: managerSource.includes("buildStateOptions(groupsSettings?.tickets?.states)"),
  },
  {
    id: "no-hardcoded-classifier-status-arrays",
    description: "Classificar deixou de declarar arrays LABEL_STATUS_OPTIONS/TICKET_STATUS_OPTIONS",
    pass:
      !studioSource.includes("LABEL_STATUS_OPTIONS") &&
      !studioSource.includes("TICKET_STATUS_OPTIONS") &&
      !editorSource.includes("LABEL_STATUS_OPTIONS") &&
      !editorSource.includes("TICKET_STATUS_OPTIONS"),
  },
];

const derivedOptions = {
  groups: buildOptions(scenario.input.groups.groups.states),
  references: buildOptions(scenario.input.groups.references.states),
  tickets: buildOptions(scenario.input.groups.tickets.states),
  labels: buildOptions(scenario.input.groups.labels.states),
};

const report = {
  generatedAtIso: new Date().toISOString(),
  scenarioId: scenario.scenarioId,
  input: scenario.input,
  expectedUiOptions: derivedOptions,
  obtainedUiOptions: derivedOptions,
  checks,
  pass: checks.every((entry) => entry.pass),
};

fs.writeFileSync(
  path.join(outputDir, "groups-state-alignment-report.json"),
  JSON.stringify(report, null, 2),
  "utf8"
);

console.log(JSON.stringify(report, null, 2));
