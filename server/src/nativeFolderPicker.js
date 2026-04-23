import path from "node:path";
import { execFile } from "node:child_process";
import { validateGroupStorageTarget } from "./groupStorageRuntime.js";

function normalizeText(value) {
  return String(value || "").trim();
}

function isExistingDirectory(targetPath) {
  try {
    return Boolean(targetPath) && path.isAbsolute(targetPath);
  } catch {
    return false;
  }
}

function buildPowerShellFolderPickerScript(input = {}) {
  const description = JSON.stringify(
    normalizeText(input.description) || "Selecione a pasta de trabalho de Grupos"
  );
  const initialPath = JSON.stringify(normalizeText(input.initialPath));

  return [
    "Add-Type -AssemblyName System.Windows.Forms",
    `$description = ${description}`,
    `$initialPath = ${initialPath}`,
    "$dialog = New-Object System.Windows.Forms.FolderBrowserDialog",
    "$dialog.Description = $description",
    "$dialog.ShowNewFolderButton = $true",
    "if ($initialPath -and [System.IO.Directory]::Exists($initialPath)) {",
    "  $dialog.SelectedPath = $initialPath",
    "}",
    "$result = @{ selected = $false; path = '' }",
    "if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {",
    "  $result.selected = $true",
    "  $result.path = $dialog.SelectedPath",
    "}",
    "$result | ConvertTo-Json -Compress",
  ].join("\n");
}

function runPowerShellFolderPicker(script) {
  return new Promise((resolve, reject) => {
    execFile(
      "powershell.exe",
      [
        "-NoProfile",
        "-STA",
        "-EncodedCommand",
        Buffer.from(String(script || ""), "utf16le").toString("base64"),
      ],
      {
        windowsHide: false,
        timeout: 15 * 60 * 1000,
        maxBuffer: 1024 * 1024,
      },
      (error, stdout, stderr) => {
        if (error) {
          reject(new Error(normalizeText(stderr) || normalizeText(error.message) || "Falha ao abrir o seletor de pasta."));
          return;
        }
        resolve(String(stdout || "").trim());
      }
    );
  });
}

export async function pickNativeFilesystemFolder(input = {}) {
  const runPicker = typeof input.runPicker === "function"
    ? input.runPicker
    : runPowerShellFolderPicker;

  if (process.platform !== "win32") {
    return {
      supported: false,
      selected: false,
      cancelled: false,
      path: "",
      normalizedPath: "",
      validation: null,
      picker: null,
      reason: "O picker nativo desta fase usa o seletor de pasta do Windows via backend local.",
    };
  }

  const output = await runPicker(
    buildPowerShellFolderPickerScript({
      description: input.description,
      initialPath: input.initialPath,
    })
  );

  let parsed = null;
  try {
    parsed = JSON.parse(String(output || "{}"));
  } catch {
    throw new Error("O seletor de pasta devolveu uma resposta invalida.");
  }

  const selected = parsed?.selected === true;
  const pickedPath = normalizeText(parsed?.path);
  if (!selected || !pickedPath) {
    return {
      supported: true,
      selected: false,
      cancelled: true,
      path: "",
      normalizedPath: "",
      validation: null,
      picker: "windows_folder_browser",
      reason: "Selecao cancelada pelo utilizador.",
    };
  }

  const normalizedPath = isExistingDirectory(pickedPath)
    ? path.resolve(pickedPath)
    : pickedPath;
  const validation = validateGroupStorageTarget({
    mode: "chosen_folder",
    chosenFolder: {
      path: normalizedPath,
      kind: "filesystem",
    },
  });

  return {
    supported: true,
    selected: true,
    cancelled: false,
    path: pickedPath,
    normalizedPath,
    validation,
    picker: "windows_folder_browser",
    reason: validation?.supported
      ? "Pasta local escolhida via seletor nativo do Windows."
      : validation?.blockingReason || "A pasta escolhida nao passou na validacao real do servidor.",
  };
}
