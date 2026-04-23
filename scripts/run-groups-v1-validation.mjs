import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { chromium } from "@playwright/test";
import { createServer } from "vite";
import {
  buildGroupWorksetMirrorFileLocation,
  validateGroupStorageTarget,
} from "../server/src/groupStorageRuntime.js";
import {
  deleteIntermediateCaseTree,
  listIntermediateCasePaths,
  readIntermediateCaseBinaryFile,
  readIntermediateCaseTextFile,
  writeIntermediateCaseBinaryFile,
  writeIntermediateCaseTextFile,
} from "../server/src/groupIntermediateCaseStore.js";
import { pickNativeFilesystemFolder } from "../server/src/nativeFolderPicker.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");
const clientRoot = path.join(repoRoot, "client");
const clientSrcRoot = path.join(repoRoot, "client", "src");
const outputDir = path.join(repoRoot, "output", "playwright");
const reportPath = path.join(outputDir, "groups-v1-validation-report.json");
const screenshotPath = path.join(outputDir, "groups-v1-validation-page.png");
const studioScreenshotPath = path.join(outputDir, "groups-v1-classification-smoke.png");
const settingsScreenshotPath = path.join(outputDir, "groups-v1-settings-surface-smoke.png");
const port = 4178;
const origin = `https://127.0.0.1:${port}`;

function ensureDir(target) {
  fs.mkdirSync(target, { recursive: true });
}

function getValidatedCommitSha() {
  return execFileSync("git", ["rev-parse", "HEAD"], {
    cwd: repoRoot,
    encoding: "utf-8",
  }).trim();
}

const LEGACY_GROUPS_ALIAS_PATTERNS = [
  "groupStorage",
  "groupsTabSettings",
  "groupLabelsManagerEnabled",
  "groupLabelCatalog",
  "groupFavoriteIds",
  "groupTicketsEnabled",
  "groupTicketUi",
  "groupOutlookCategories",
].map((name) => ({
  name,
  matchers: [
    new RegExp(`\\.${name}\\b`, "g"),
    new RegExp(`\\[\\s*["']${name}["']\\s*\\]`, "g"),
  ],
}));

const ALLOWED_LEGACY_GROUPS_COMPATIBILITY_FILES = new Set([
  "client/src/settings.ts",
  "client/src/modules/crm/groups-v1/settings/groupsModuleSettings.ts",
]);

function listFilesRecursively(rootDir) {
  const pending = [rootDir];
  const files = [];
  while (pending.length) {
    const current = pending.pop();
    for (const entry of fs.readdirSync(current, { withFileTypes: true })) {
      const nextPath = path.join(current, entry.name);
      if (entry.isDirectory()) {
        pending.push(nextPath);
        continue;
      }
      files.push(nextPath);
    }
  }
  return files;
}

function auditLegacyGroupsAliasUsage() {
  const findings = [];
  for (const filePath of listFilesRecursively(clientSrcRoot)) {
    const relativePath = path.relative(repoRoot, filePath).replace(/\\/g, "/");
    const source = fs.readFileSync(filePath, "utf-8");
    const lines = source.split(/\r?\n/);
    lines.forEach((line, index) => {
      for (const pattern of LEGACY_GROUPS_ALIAS_PATTERNS) {
        const matched = pattern.matchers.some((regex) => {
          regex.lastIndex = 0;
          return regex.test(line);
        });
        if (!matched) continue;
        findings.push({
          file: relativePath,
          line: index + 1,
          alias: pattern.name,
          allowedCompatibilityFile: ALLOWED_LEGACY_GROUPS_COMPATIBILITY_FILES.has(relativePath),
          text: line.trim(),
        });
      }
    });
  }
  const unexpectedReferences = findings.filter((entry) => !entry.allowedCompatibilityFile);
  return {
    aliases: LEGACY_GROUPS_ALIAS_PATTERNS.map((entry) => entry.name),
    allowedCompatibilityFiles: Array.from(ALLOWED_LEGACY_GROUPS_COMPATIBILITY_FILES),
    findings,
    unexpectedReferences,
    ok: unexpectedReferences.length === 0,
  };
}

async function runServerStorageChecks() {
  const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), "iccc-groups-v1-"));
  const localDeviceRoot = path.join(tempRoot, "local-device");
  const chosenFolderRoot = path.join(tempRoot, "chosen-folder");
  const hybridRoot = path.join(tempRoot, "hybrid-folder");
  const intermediateRoot = path.join(tempRoot, "intermediate-folder");
  const oneDriveLocalRoot = path.join(tempRoot, "OneDrive - Demo", "Groups");

  const cloud = validateGroupStorageTarget({ mode: "supabase" });
  const localDevice = validateGroupStorageTarget({
    mode: "local_device",
    localDevice: { rootPath: localDeviceRoot },
  });
  const chosenFolder = validateGroupStorageTarget({
    mode: "chosen_folder",
    chosenFolder: { path: chosenFolderRoot, kind: "filesystem" },
  });
  const hybrid = validateGroupStorageTarget({
    mode: "hybrid",
    chosenFolder: { path: hybridRoot, kind: "filesystem" },
    hybrid: { primaryTarget: "chosen_folder" },
  });
  const blockedWeb = validateGroupStorageTarget({
    mode: "chosen_folder",
    chosenFolder: { path: "https://tenant.sharepoint.com/sites/demo/Shared Documents", kind: "document_library" },
  });
  const oneDriveLocal = validateGroupStorageTarget({
    mode: "chosen_folder",
    chosenFolder: { path: oneDriveLocalRoot, kind: "filesystem" },
  });
  const mirrorLocation = buildGroupWorksetMirrorFileLocation({
    mode: "chosen_folder",
    chosenFolder: { path: chosenFolderRoot, kind: "filesystem" },
    worksetKey: "groups_v1_workset:anchor@email|base",
  });

  writeIntermediateCaseTextFile({
    basePath: intermediateRoot,
    path: "Groups/cases/case-a/case.json",
    content: JSON.stringify({ caseId: "case-a", anchorEmailKey: "anchor-a" }),
  });
  writeIntermediateCaseBinaryFile({
    basePath: intermediateRoot,
    path: "Groups/cases/case-a/attachments/anchor-a/doc-1",
    contentBase64: Buffer.from("fixture-binary").toString("base64"),
  });
  const intermediateReadText = readIntermediateCaseTextFile({
    basePath: intermediateRoot,
    path: "Groups/cases/case-a/case.json",
  });
  const intermediateReadBinary = readIntermediateCaseBinaryFile({
    basePath: intermediateRoot,
    path: "Groups/cases/case-a/attachments/anchor-a/doc-1",
  });
  const intermediatePathsBeforeDelete = listIntermediateCasePaths({
    basePath: intermediateRoot,
    prefix: "Groups/cases",
  });
  deleteIntermediateCaseTree({
    basePath: intermediateRoot,
    path: "Groups/cases/case-a",
  });
  const intermediatePathsAfterDelete = listIntermediateCasePaths({
    basePath: intermediateRoot,
    prefix: "Groups/cases",
  });

  const pickerSelection = await pickNativeFilesystemFolder({
    initialPath: oneDriveLocalRoot,
    description: "Selecione a pasta de trabalho de Grupos",
    runPicker: async () => JSON.stringify({
      selected: true,
      path: oneDriveLocalRoot,
    }),
  });
  const pickerCancellation = await pickNativeFilesystemFolder({
    initialPath: oneDriveLocalRoot,
    description: "Selecione a pasta de trabalho de Grupos",
    runPicker: async () => JSON.stringify({
      selected: false,
      path: "",
    }),
  });

  return {
    tempRoot,
    results: {
      cloud,
      localDevice,
      chosenFolder,
      hybrid,
      blockedWeb,
      oneDriveLocal,
      pickerSelection,
      pickerCancellation,
    },
    mirrorLocation,
    intermediateStorage: {
      root: intermediateRoot,
      readTextOk: String(intermediateReadText || "").includes("\"case-a\""),
      readBinaryOk: String(intermediateReadBinary?.contentBase64 || "") === Buffer.from("fixture-binary").toString("base64"),
      listedBeforeDelete: intermediatePathsBeforeDelete,
      listedAfterDelete: intermediatePathsAfterDelete,
    },
  };
}

async function runBrowserChecks(browser) {
  const page = await browser.newPage({ ignoreHTTPSErrors: true });
  await page.goto(`${origin}/groups-v1-validation.html`, { waitUntil: "networkidle" });
  await page.waitForFunction(() => Boolean(window.__GROUPS_VALIDATION_RESULT__?.done), undefined, { timeout: 120000 });
  await page.screenshot({ path: screenshotPath, fullPage: true });
  const result = await page.evaluate(() => window.__GROUPS_VALIDATION_RESULT__);
  await page.close();
  return result;
}

async function runClassificationStudioSmoke(browser) {
  const page = await browser.newPage({ ignoreHTTPSErrors: true });
  const pageErrors = [];
  const consoleErrors = [];

  page.on("pageerror", (error) => {
    pageErrors.push(String(error?.message || error || ""));
  });
  page.on("console", (message) => {
    if (message.type() === "error") {
      consoleErrors.push(message.text());
    }
  });

  try {
    await page.goto(`${origin}/?view=group-classification-studio`, { waitUntil: "domcontentloaded" });
    await page.waitForSelector('[data-testid="studio-root"]', { timeout: 25000 });
    await page.screenshot({ path: studioScreenshotPath, fullPage: true });
    return {
      ok: pageErrors.length === 0,
      pageErrors,
      consoleErrors,
    };
  } catch (error) {
    return {
      ok: false,
      pageErrors,
      consoleErrors,
      error: error instanceof Error ? error.message : String(error || "Erro desconhecido"),
    };
  } finally {
    await page.close();
  }
}

async function runGroupsSettingsSurfaceSmoke(browser) {
  const page = await browser.newPage({ ignoreHTTPSErrors: true });
  const pageErrors = [];
  const consoleErrors = [];

  page.on("pageerror", (error) => {
    pageErrors.push(String(error?.message || error || ""));
  });
  page.on("console", (message) => {
    if (message.type() === "error") {
      consoleErrors.push(message.text());
    }
  });

  try {
    await page.goto(`${origin}/?view=group-settings&surface=groups-tab`, { waitUntil: "domcontentloaded" });
    await page.waitForSelector("text=Settings da aba Groups", { timeout: 25000 });
    await page.screenshot({ path: settingsScreenshotPath, fullPage: true });
    return {
      ok: pageErrors.length === 0,
      entrypoint: "?view=group-settings&surface=groups-tab",
      pageErrors,
      consoleErrors,
    };
  } catch (error) {
    return {
      ok: false,
      entrypoint: "?view=group-settings&surface=groups-tab",
      pageErrors,
      consoleErrors,
      error: error instanceof Error ? error.message : String(error || "Erro desconhecido"),
    };
  } finally {
    await page.close();
  }
}

async function main() {
  ensureDir(outputDir);
  const validatedCommitSha = getValidatedCommitSha();
  const legacyAliasAudit = auditLegacyGroupsAliasUsage();

  const viteServer = await createServer({
    root: clientRoot,
    configFile: path.join(clientRoot, "vite.config.ts"),
    server: {
      host: "127.0.0.1",
      port,
      strictPort: true,
    },
  });

  let browser;
  try {
    await viteServer.listen();
    browser = await chromium.launch({ headless: true });

    const [browserValidation, classificationSmoke, groupsSettingsSurfaceSmoke, serverStorage] = await Promise.all([
      runBrowserChecks(browser),
      runClassificationStudioSmoke(browser),
      runGroupsSettingsSurfaceSmoke(browser),
      runServerStorageChecks(),
    ]);

    const report = {
      generatedAtIso: new Date().toISOString(),
      validatedCommitSha,
      validationEntrypoints: [
        "/groups-v1-validation.html",
        "/?view=group-classification-studio",
        "/?view=group-settings&surface=groups-tab",
      ],
      harnessFiles: [
        "client/groups-v1-validation.html",
        "client/src/modules/crm/groups-v1/testing/validationPage.tsx",
        "client/src/modules/crm/groups-v1/testing/runtimeValidation.tsx",
        "client/src/modules/crm/groups-v1/testing/settingsMatrix.ts",
        "scripts/run-groups-v1-validation.mjs",
      ],
      legacyAliasAudit,
      browserValidation,
      classificationSmoke,
      groupsSettingsSurfaceSmoke,
      serverStorage,
    };

    fs.writeFileSync(reportPath, JSON.stringify(report, null, 2), "utf-8");
    console.log(`GROUPS_V1_VALIDATION_REPORT:${reportPath}`);
    console.log(JSON.stringify(report, null, 2));
    const hasBrowserFailures = Number(browserValidation?.failed || 0) > 0;
    const serverStorageOk = Boolean(
      serverStorage?.results?.oneDriveLocal?.supported
      && serverStorage?.results?.pickerSelection?.selected
      && serverStorage?.results?.pickerSelection?.validation?.supported
      && serverStorage?.intermediateStorage?.readTextOk
      && serverStorage?.intermediateStorage?.readBinaryOk
    );
    if (hasBrowserFailures || !classificationSmoke?.ok || !groupsSettingsSurfaceSmoke?.ok || !legacyAliasAudit.ok || !serverStorageOk) {
      process.exitCode = 1;
    }
  } finally {
    if (browser) await browser.close();
    await viteServer.close();
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
