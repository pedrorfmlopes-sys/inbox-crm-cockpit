import fs from "node:fs";
import os from "node:os";
import path from "node:path";

function normalizeString(value) {
  return String(value || "").trim();
}

function normalizeStorageMode(value) {
  const normalized = normalizeString(value).toLowerCase();
  if (normalized === "local_device" || normalized === "chosen_folder" || normalized === "hybrid") {
    return normalized;
  }
  return "supabase";
}

function normalizeChosenFolderKind(value, basePath) {
  const normalized = normalizeString(value).toLowerCase();
  if (normalized === "document_library" || normalized === "filesystem") {
    return normalized;
  }
  return looksLikeWebUrl(basePath) ? "document_library" : "filesystem";
}

export function looksLikeWebUrl(value) {
  return /^https?:\/\//i.test(normalizeString(value));
}

export function sanitizePathSegment(value) {
  return normalizeString(value)
    .replace(/[\\/:*?"<>|]+/g, "_")
    .replace(/\s+/g, "_")
    .replace(/^_+|_+$/g, "");
}

export function resolveGroupStorageInput(input = {}) {
  const mode = normalizeStorageMode(input.mode);
  const localRootPath = normalizeString(input.localDevice?.rootPath || input.localRootPath);
  const chosenFolderPath = normalizeString(input.chosenFolder?.path || input.chosenFolderPath || input.baseFolderPath);
  const chosenFolderKind = normalizeChosenFolderKind(
    input.chosenFolder?.kind || input.chosenFolderKind,
    chosenFolderPath
  );
  const primaryTarget = normalizeString(input.hybrid?.primaryTarget || input.primaryTarget) === "local_device"
    ? "local_device"
    : "chosen_folder";

  if (mode === "supabase") {
    return {
      mode,
      provider: "cloud",
      basePath: "",
      chosenFolderKind,
      primaryTarget,
      fileBacked: false,
    };
  }

  if (mode === "local_device") {
    return {
      mode,
      provider: "local",
      basePath: localRootPath,
      chosenFolderKind,
      primaryTarget,
      fileBacked: true,
    };
  }

  if (mode === "chosen_folder") {
    return {
      mode,
      provider: chosenFolderKind === "document_library" ? "onedrive" : "local",
      basePath: chosenFolderPath,
      chosenFolderKind,
      primaryTarget,
      fileBacked: true,
    };
  }

  const hybridBasePath = primaryTarget === "local_device" ? localRootPath : chosenFolderPath;
  return {
    mode,
    provider: primaryTarget === "local_device"
      ? "local"
      : chosenFolderKind === "document_library"
        ? "onedrive"
        : "local",
    basePath: hybridBasePath,
    chosenFolderKind,
    primaryTarget,
    fileBacked: true,
  };
}

export function validateGroupStorageTarget(input = {}) {
  const resolved = resolveGroupStorageInput(input);
  const notes = [];

  if (!resolved.fileBacked) {
    return {
      mode: resolved.mode,
      provider: resolved.provider,
      fileBacked: false,
      supported: true,
      basePath: "",
      normalizedBasePath: "",
      isWebUrl: false,
      requiresServerAccessiblePath: false,
      canStoreManifest: true,
      canStoreBinary: true,
      pickerAvailable: false,
      pickerBlockedReason:
        "O host atual nao expoe um picker de pasta reutilizavel que entregue um caminho seguro ao backend.",
      architecturalBlocker: null,
      requiredChange: null,
      notes: ["Modo cloud: a persistencia final continua centralizada na app."],
    };
  }

  const basePath = normalizeString(resolved.basePath);
  if (!basePath) {
    return {
      mode: resolved.mode,
      provider: resolved.provider,
      fileBacked: true,
      supported: false,
      basePath: "",
      normalizedBasePath: "",
      isWebUrl: false,
      requiresServerAccessiblePath: true,
      canStoreManifest: false,
      canStoreBinary: false,
      pickerAvailable: false,
      pickerBlockedReason:
        "Sem bridge nativa, o add-in nao consegue entregar ao backend um caminho local do utilizador atraves de um picker real.",
      blockingReason: "Define primeiro um caminho local, pasta sincronizada ou UNC acessivel ao processo do servidor.",
      architecturalBlocker: null,
      requiredChange: null,
      notes,
    };
  }

  if (looksLikeWebUrl(basePath)) {
    return {
      mode: resolved.mode,
      provider: resolved.provider,
      fileBacked: true,
      supported: false,
      basePath,
      normalizedBasePath: "",
      isWebUrl: true,
      requiresServerAccessiblePath: true,
      canStoreManifest: false,
      canStoreBinary: false,
      pickerAvailable: false,
      pickerBlockedReason:
        "Mesmo com picker browser, o host atual nao fornece um caminho fisico que o backend consiga usar para escrita.",
      blockingReason:
        "OneDrive/SharePoint por URL web nao e suportado nesta arquitetura porque a escrita final atual usa filesystem no servidor, nao Graph/SharePoint API.",
      architecturalBlocker: "web_document_library_requires_graph_backend",
      requiredChange:
        "Adicionar autenticacao Graph/SharePoint com scopes Files/Sites, resolucao de site/drive/item e um fluxo real de upload/download por API em vez de filesystem.",
      notes: [
        "Os manifests do add-in expostos no repo so declaram ReadWriteMailbox.",
        "O runtime Graph atual do cliente so pede Mail.Read, User.Read e People.Read.",
        "Para fechar URL web de verdade seria necessaria uma integracao dedicada com Graph/SharePoint e autenticacao associada.",
      ],
    };
  }

  const normalizedBasePath = path.resolve(basePath);
  const probeDir = path.join(normalizedBasePath, ".icc_probe");
  const probeFile = path.join(
    probeDir,
    `${Date.now()}-${Math.random().toString(16).slice(2)}.tmp`
  );

  try {
    fs.mkdirSync(probeDir, { recursive: true });
    fs.writeFileSync(probeFile, "icc-storage-probe", "utf-8");
    const readBack = fs.readFileSync(probeFile, "utf-8");
    if (readBack !== "icc-storage-probe") {
      throw new Error("Leitura de validacao devolveu conteudo inesperado.");
    }
    fs.unlinkSync(probeFile);
    try {
      fs.rmSync(probeDir, { recursive: true, force: true });
    } catch {
      // best effort
    }
    notes.push("Caminho validado com escrita e leitura real no host atual do servidor.");
    if (normalizedBasePath.startsWith(os.tmpdir())) {
      notes.push("O caminho validado esta numa pasta temporaria; nao e recomendado para uso persistente.");
    }
    return {
      mode: resolved.mode,
      provider: resolved.provider,
      fileBacked: true,
      supported: true,
      basePath,
      normalizedBasePath,
      isWebUrl: false,
      requiresServerAccessiblePath: true,
      canStoreManifest: true,
      canStoreBinary: true,
      pickerAvailable: false,
      pickerBlockedReason:
        "O add-in nao tem picker nativo que entregue ao backend um caminho local do utilizador; nesta arquitetura usa-se path manual validado no servidor.",
      architecturalBlocker: null,
      requiredChange: null,
      notes,
    };
  } catch (error) {
    try {
      if (fs.existsSync(probeFile)) fs.unlinkSync(probeFile);
      if (fs.existsSync(probeDir)) fs.rmSync(probeDir, { recursive: true, force: true });
    } catch {
      // ignore cleanup error
    }
    return {
      mode: resolved.mode,
      provider: resolved.provider,
      fileBacked: true,
      supported: false,
      basePath,
      normalizedBasePath,
      isWebUrl: false,
      requiresServerAccessiblePath: true,
      canStoreManifest: false,
      canStoreBinary: false,
      pickerAvailable: false,
      pickerBlockedReason:
        "Sem bridge nativa, o add-in nao consegue escolher e entregar automaticamente um caminho local do utilizador ao backend.",
      blockingReason: `O servidor nao conseguiu escrever no destino configurado: ${normalizeString(error?.message) || "erro desconhecido"}`,
      architecturalBlocker: null,
      requiredChange: null,
      notes,
    };
  }
}

export function buildGroupWorksetMirrorFileLocation(input = {}) {
  const resolved = resolveGroupStorageInput(input);
  const basePath = normalizeString(resolved.basePath);
  if (!resolved.fileBacked || !basePath || looksLikeWebUrl(basePath)) {
    return null;
  }

  const normalizedBasePath = path.resolve(basePath);
  const worksetKey = sanitizePathSegment(input.worksetKey || "workset");
  if (!worksetKey) return null;

  const relativePath = path.posix.join("InboxCockpit", "Groups", "worksets", `${worksetKey}.json`);
  const filePath = path.resolve(normalizedBasePath, relativePath);
  const relativeCheck = path.relative(normalizedBasePath, filePath);
  if (!relativeCheck || relativeCheck.startsWith("..") || path.isAbsolute(relativeCheck)) {
    throw new Error("O caminho do manifesto saiu da pasta base permitida.");
  }

  return {
    basePath: normalizedBasePath,
    relativePath: relativePath.replace(/\\/g, "/"),
    filePath,
  };
}
