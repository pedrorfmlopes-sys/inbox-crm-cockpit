/**
 * client/src/modules/crm/excelProvider.ts
 * IndexedDB Engine for Project Protection (Bypass IT Logic)
 */

const DB_NAME = "CockpitV2_Moat";
const STORE_NAME = "project_protection";
const DB_VERSION = 1;

export interface ProtectedProject {
    id?: number;
    projectName: string;
    distributor: string;
    constructionCo?: string;
    location?: string;
    refArticles?: string[];
    updatedAt: string;
}

export async function initDb(): Promise<IDBDatabase> {
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(DB_NAME, DB_VERSION);

        request.onupgradeneeded = (event: any) => {
            const db = event.target.result;
            if (!db.objectStoreNames.contains(STORE_NAME)) {
                const store = db.createObjectStore(STORE_NAME, { keyPath: "id", autoIncrement: true });
                store.createIndex("projectName", "projectName", { unique: false });
                store.createIndex("location", "location", { unique: false });
            }
        };

        request.onsuccess = (event: any) => resolve(event.target.result);
        request.onerror = (event: any) => reject(event.target.error);
    });
}

export async function saveProjects(projects: ProtectedProject[]): Promise<void> {
    const db = await initDb();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, "readwrite");
        const store = tx.objectStore(STORE_NAME);

        // Clear existing for now (Master Spec implies one source of truth table)
        store.clear();

        projects.forEach(p => store.add({ ...p, updatedAt: new Date().toISOString() }));

        tx.oncomplete = () => resolve();
        tx.onerror = (event: any) => reject(event.target.error);
    });
}

export async function getAllProtectedProjects(): Promise<ProtectedProject[]> {
    const db = await initDb();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, "readonly");
        const store = tx.objectStore(STORE_NAME);
        const request = store.getAll();

        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
}

/**
 * AI Header Mapping helper
 * Uses Gemini to map user-uploaded columns to our internal schema
 */
export async function mapHeadersAi(headers: string[]): Promise<Record<string, string>> {
    // In a real implementation, this would call /api/ai/generate with a mapping prompt
    // For now, we simulate the mapping to stay within technical scope
    const mapping: Record<string, string> = {};
    headers.forEach(h => {
        const lh = h.toLowerCase();
        if (lh.includes("projeto") || lh.includes("obra") || lh.includes("designação")) mapping[h] = "projectName";
        if (lh.includes("distribuidor") || lh.includes("cliente")) mapping[h] = "distributor";
        if (lh.includes("construtora") || lh.includes("empreiteiro")) mapping[h] = "constructionCo";
        if (lh.includes("local") || lh.includes("cidade") || lh.includes("morada")) mapping[h] = "location";
        if (lh.includes("ref") || lh.includes("artigo") || lh.includes("material")) mapping[h] = "refArticles";
    });
    return mapping;
}
