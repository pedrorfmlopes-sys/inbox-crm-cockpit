import type { IntermediateCaseBinaryStorageAdapter } from "./intermediateCaseRepository";

const DB_NAME = "InboxCockpit_GroupsIntermediateCase";
const DB_VERSION = 1;
const TEXT_STORE = "intermediate_case_text";
const BINARY_STORE = "intermediate_case_binary";

type StoredTextRecord = {
  key: string;
  namespace: string;
  path: string;
  content: string;
  updatedAt: string;
};

type StoredBinaryRecord = {
  key: string;
  namespace: string;
  path: string;
  blob: Blob;
  updatedAt: string;
};

function makeStorageKey(namespace: string, path: string): string {
  return `${namespace}::${path}`;
}

function openIntermediateCaseDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(TEXT_STORE)) {
        const store = db.createObjectStore(TEXT_STORE, { keyPath: "key" });
        store.createIndex("namespace", "namespace", { unique: false });
        store.createIndex("path", "path", { unique: false });
      }
      if (!db.objectStoreNames.contains(BINARY_STORE)) {
        const store = db.createObjectStore(BINARY_STORE, { keyPath: "key" });
        store.createIndex("namespace", "namespace", { unique: false });
        store.createIndex("path", "path", { unique: false });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function runStoreRequest<T>(
  db: IDBDatabase,
  storeName: string,
  mode: IDBTransactionMode,
  run: (store: IDBObjectStore) => IDBRequest<T>
): Promise<T> {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, mode);
    const store = tx.objectStore(storeName);
    const request = run(store);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
    tx.onerror = () => reject(tx.error);
  });
}

function loadNamespaceRecords<T extends { namespace: string; path: string }>(
  db: IDBDatabase,
  storeName: string,
  namespace: string
): Promise<T[]> {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, "readonly");
    const store = tx.objectStore(storeName);
    const index = store.index("namespace");
    const request = index.getAll(namespace);
    request.onsuccess = () => resolve((request.result || []) as T[]);
    request.onerror = () => reject(request.error);
    tx.onerror = () => reject(tx.error);
  });
}

function deleteNamespacePrefix(
  db: IDBDatabase,
  storeName: string,
  namespace: string,
  prefix: string
): Promise<void> {
  return new Promise((resolve, reject) => {
    const normalizedPrefix = `${prefix.replace(/\/+$/, "")}/`;
    const tx = db.transaction(storeName, "readwrite");
    const store = tx.objectStore(storeName);
    const index = store.index("namespace");
    const request = index.getAll(namespace);
    request.onsuccess = () => {
      const rows = (request.result || []) as Array<{ key: string; path: string }>;
      for (const row of rows) {
        if (row.path === prefix || row.path.startsWith(normalizedPrefix)) {
          store.delete(row.key);
        }
      }
    };
    request.onerror = () => reject(request.error);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

export function createIndexedDbIntermediateCaseStorageAdapter(input: {
  namespace: string;
}): IntermediateCaseBinaryStorageAdapter {
  const namespace = String(input.namespace || "").trim();

  return {
    async readText(path) {
      const db = await openIntermediateCaseDb();
      const key = makeStorageKey(namespace, path);
      const row = await runStoreRequest<StoredTextRecord | undefined>(db, TEXT_STORE, "readonly", (store) =>
        store.get(key) as IDBRequest<StoredTextRecord | undefined>
      );
      return row?.content || null;
    },
    async writeText(path, content) {
      const db = await openIntermediateCaseDb();
      const key = makeStorageKey(namespace, path);
      await runStoreRequest(db, TEXT_STORE, "readwrite", (store) =>
        store.put({
          key,
          namespace,
          path,
          content,
          updatedAt: new Date().toISOString(),
        } satisfies StoredTextRecord)
      );
    },
    async readBinary(path) {
      const db = await openIntermediateCaseDb();
      const key = makeStorageKey(namespace, path);
      const row = await runStoreRequest<StoredBinaryRecord | undefined>(db, BINARY_STORE, "readonly", (store) =>
        store.get(key) as IDBRequest<StoredBinaryRecord | undefined>
      );
      return row?.blob || null;
    },
    async writeBinary(path, content) {
      const db = await openIntermediateCaseDb();
      const key = makeStorageKey(namespace, path);
      const blob = content instanceof Blob
        ? content
        : content instanceof ArrayBuffer
          ? new Blob([content])
          : new Blob([content.buffer.slice(content.byteOffset, content.byteOffset + content.byteLength)]);
      await runStoreRequest(db, BINARY_STORE, "readwrite", (store) =>
        store.put({
          key,
          namespace,
          path,
          blob,
          updatedAt: new Date().toISOString(),
        } satisfies StoredBinaryRecord)
      );
    },
    async deleteTree(path) {
      const db = await openIntermediateCaseDb();
      await Promise.all([
        deleteNamespacePrefix(db, TEXT_STORE, namespace, path),
        deleteNamespacePrefix(db, BINARY_STORE, namespace, path),
      ]);
    },
    async listPaths(prefix) {
      const db = await openIntermediateCaseDb();
      const rows = await loadNamespaceRecords<StoredTextRecord>(db, TEXT_STORE, namespace);
      return rows.map((row) => row.path).filter((path) => path.startsWith(prefix));
    },
  };
}
