import fs from "node:fs";
import path from "node:path";
import { validateGroupStorageTarget } from "./groupStorageRuntime.js";

const INTERMEDIATE_CASE_ROOT_SEGMENTS = ["InboxCockpit", "Groups", "intermediate-cases"];

function normalizeText(value) {
  return String(value || "").trim();
}

function toPosixPath(value) {
  return normalizeText(value).replace(/\\/g, "/").replace(/^\/+|\/+$/g, "");
}

function ensureFilesystemLocation(basePath) {
  const validation = validateGroupStorageTarget({
    mode: "chosen_folder",
    chosenFolder: {
      path: basePath,
      kind: "filesystem",
    },
  });
  if (!validation.supported) {
    throw new Error(validation.blockingReason || "A pasta intermédia configurada não está acessível.");
  }
  return validation.normalizedBasePath;
}

function getIntermediateRoot(basePath) {
  const normalizedBasePath = ensureFilesystemLocation(basePath);
  return path.resolve(normalizedBasePath, ...INTERMEDIATE_CASE_ROOT_SEGMENTS);
}

function resolveIntermediateStoragePath(basePath, relativePath = "") {
  const rootPath = getIntermediateRoot(basePath);
  const normalizedRelativePath = toPosixPath(relativePath);
  const filePath = path.resolve(rootPath, normalizedRelativePath);
  const relativeCheck = path.relative(rootPath, filePath);
  if (relativeCheck.startsWith("..") || path.isAbsolute(relativeCheck)) {
    throw new Error("O caminho do storage intermédio saiu da pasta base permitida.");
  }
  return {
    rootPath,
    normalizedRelativePath,
    filePath,
  };
}

function ensureParentDir(filePath) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
}

function listFilesRecursive(currentDir, rootDir, rows) {
  if (!fs.existsSync(currentDir)) return;
  for (const entry of fs.readdirSync(currentDir, { withFileTypes: true })) {
    const nextPath = path.join(currentDir, entry.name);
    if (entry.isDirectory()) {
      listFilesRecursive(nextPath, rootDir, rows);
      continue;
    }
    rows.push(path.relative(rootDir, nextPath).replace(/\\/g, "/"));
  }
}

export function readIntermediateCaseTextFile(input = {}) {
  const { filePath } = resolveIntermediateStoragePath(input.basePath, input.path);
  if (!fs.existsSync(filePath)) return null;
  return fs.readFileSync(filePath, "utf-8");
}

export function writeIntermediateCaseTextFile(input = {}) {
  const { filePath } = resolveIntermediateStoragePath(input.basePath, input.path);
  ensureParentDir(filePath);
  fs.writeFileSync(filePath, String(input.content || ""), "utf-8");
}

export function readIntermediateCaseBinaryFile(input = {}) {
  const { filePath } = resolveIntermediateStoragePath(input.basePath, input.path);
  if (!fs.existsSync(filePath)) return null;
  const buffer = fs.readFileSync(filePath);
  return {
    contentBase64: buffer.toString("base64"),
    contentType: normalizeText(input.contentType) || "application/octet-stream",
  };
}

export function writeIntermediateCaseBinaryFile(input = {}) {
  const { filePath } = resolveIntermediateStoragePath(input.basePath, input.path);
  const contentBase64 = normalizeText(input.contentBase64).replace(/^data:[^,]+,/, "");
  if (!contentBase64) {
    throw new Error("Conteúdo binário intermédio vazio.");
  }
  ensureParentDir(filePath);
  fs.writeFileSync(filePath, Buffer.from(contentBase64, "base64"));
}

export function deleteIntermediateCaseTree(input = {}) {
  const { filePath } = resolveIntermediateStoragePath(input.basePath, input.path);
  if (!fs.existsSync(filePath)) return false;
  const stats = fs.statSync(filePath);
  if (stats.isDirectory()) {
    fs.rmSync(filePath, { recursive: true, force: true });
  } else {
    fs.unlinkSync(filePath);
  }
  return true;
}

export function listIntermediateCasePaths(input = {}) {
  const { rootPath, normalizedRelativePath, filePath } = resolveIntermediateStoragePath(
    input.basePath,
    input.prefix || ""
  );
  const rows = [];
  const scanRoot = normalizedRelativePath ? filePath : rootPath;
  if (!fs.existsSync(scanRoot)) return [];
  const stats = fs.statSync(scanRoot);
  if (stats.isFile()) {
    return [path.relative(rootPath, scanRoot).replace(/\\/g, "/")];
  }
  listFilesRecursive(scanRoot, rootPath, rows);
  const normalizedPrefix = toPosixPath(input.prefix || "");
  return normalizedPrefix
    ? rows.filter((entry) => entry === normalizedPrefix || entry.startsWith(`${normalizedPrefix}/`))
    : rows;
}

export function describeIntermediateCaseStorage(input = {}) {
  const rootPath = getIntermediateRoot(input.basePath);
  return {
    basePath: ensureFilesystemLocation(input.basePath),
    rootPath,
  };
}
