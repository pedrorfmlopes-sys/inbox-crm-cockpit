import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { chromium } from "@playwright/test";
import { createServer } from "vite";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");
const clientRoot = path.join(repoRoot, "client");
const outputDir = path.join(repoRoot, "output", "playwright");
const reportPath = path.join(outputDir, "graph-drive-self-test-report.json");
const screenshotPath = path.join(outputDir, "graph-drive-self-test-page.png");
const port = 4180;
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

async function runSelfTest(browser) {
  const page = await browser.newPage({ ignoreHTTPSErrors: true });
  const pageErrors = [];
  const consoleErrors = [];
  const popupSummaries = [];

  page.on("pageerror", (error) => {
    pageErrors.push(String(error?.message || error || ""));
  });
  page.on("console", (message) => {
    if (message.type() === "error") {
      consoleErrors.push(message.text());
    }
  });
  page.on("popup", (popup) => {
    const summary = {
      openedAtIso: new Date().toISOString(),
      initialUrl: popup.url(),
      finalUrl: "",
      closed: false,
    };
    popupSummaries.push(summary);
    popup.on("close", () => {
      summary.closed = true;
      summary.finalUrl = popup.url();
    });
  });

  try {
    await page.goto(`${origin}/?view=graph-drive-self-test`, { waitUntil: "domcontentloaded" });
    await page.waitForSelector('[data-testid="graph-drive-self-test-run"]', { timeout: 20000 });
    await page.click('[data-testid="graph-drive-self-test-run"]');
    await page.waitForFunction(
      () => Boolean(window.__GRAPH_DRIVE_SELF_TEST_RESULT__?.done),
      undefined,
      { timeout: 90000 },
    );
    await page.screenshot({ path: screenshotPath, fullPage: true });
    const result = await page.evaluate(() => window.__GRAPH_DRIVE_SELF_TEST_RESULT__);
    return {
      ok: true,
      entrypoint: "/?view=graph-drive-self-test",
      pageErrors,
      consoleErrors,
      popupSummaries,
      result,
    };
  } catch (error) {
    return {
      ok: false,
      entrypoint: "/?view=graph-drive-self-test",
      pageErrors,
      consoleErrors,
      popupSummaries,
      error: error instanceof Error ? error.message : String(error || "Erro desconhecido"),
      result: await page.evaluate(() => window.__GRAPH_DRIVE_SELF_TEST_RESULT__ || null).catch(() => null),
    };
  } finally {
    await page.close();
  }
}

async function main() {
  ensureDir(outputDir);
  const validatedCommitSha = getValidatedCommitSha();
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
    const selfTest = await runSelfTest(browser);
    const report = {
      generatedAtIso: new Date().toISOString(),
      validatedCommitSha,
      entrypoints: ["/?view=graph-drive-self-test"],
      harnessFiles: [
        "client/src/modules/auth/GraphDriveSelfTestApp.tsx",
        "client/src/office.ts",
        "client/src/main.tsx",
        "scripts/run-graph-drive-self-test.mjs",
      ],
      selfTest,
      screenshotPath,
    };

    fs.writeFileSync(reportPath, JSON.stringify(report, null, 2), "utf-8");
    console.log(`GRAPH_DRIVE_SELF_TEST_REPORT:${reportPath}`);
    console.log(JSON.stringify(report, null, 2));

    if (!selfTest.ok) {
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
