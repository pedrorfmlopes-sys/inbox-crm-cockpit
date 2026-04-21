import { getIntermediateCaseJsonPath, INTERMEDIATE_CASES_ROOT } from "./intermediateCasePaths";
import { buildIntermediateCaseSummary } from "./intermediateCaseSummary";
import { parseIntermediateCase, serializeIntermediateCase } from "./intermediateCaseSerialization";
import type { IntermediateCase, IntermediateCaseSummary } from "./intermediateCaseTypes";

export interface IntermediateCaseStorageAdapter {
  readText(path: string): Promise<string | null>;
  writeText(path: string, content: string): Promise<void>;
  deletePath(path: string): Promise<void>;
  listPaths(prefix: string): Promise<string[]>;
}

export interface IntermediateCaseRepository {
  readCase(caseId: string): Promise<IntermediateCase | null>;
  writeCase(caseValue: IntermediateCase): Promise<IntermediateCase>;
  deleteCase(caseId: string): Promise<boolean>;
  listCases(): Promise<IntermediateCaseSummary[]>;
  findCaseByEmailKey(emailKey: string): Promise<IntermediateCase | null>;
}

export function createIntermediateCaseRepository(adapter: IntermediateCaseStorageAdapter): IntermediateCaseRepository {
  return {
    async readCase(caseId) {
      const raw = await adapter.readText(getIntermediateCaseJsonPath(caseId));
      return raw ? parseIntermediateCase(raw) : null;
    },
    async writeCase(caseValue) {
      const path = getIntermediateCaseJsonPath(caseValue.caseId);
      await adapter.writeText(path, serializeIntermediateCase(caseValue));
      return caseValue;
    },
    async deleteCase(caseId) {
      await adapter.deletePath(getIntermediateCaseJsonPath(caseId));
      return true;
    },
    async listCases() {
      const paths = await adapter.listPaths(INTERMEDIATE_CASES_ROOT);
      const cases = await Promise.all(
        paths
          .filter((path) => path.endsWith("/case.json"))
          .map(async (path) => {
            const raw = await adapter.readText(path);
            const parsed = raw ? parseIntermediateCase(raw) : null;
            return parsed ? buildIntermediateCaseSummary(parsed) : null;
          })
      );
      return cases.filter(Boolean) as IntermediateCaseSummary[];
    },
    async findCaseByEmailKey(emailKey) {
      const paths = await adapter.listPaths(INTERMEDIATE_CASES_ROOT);
      for (const path of paths.filter((entry) => entry.endsWith("/case.json"))) {
        const raw = await adapter.readText(path);
        const parsed = raw ? parseIntermediateCase(raw) : null;
        if (parsed?.emails.some((email) => email.emailKey === emailKey)) {
          return parsed;
        }
      }
      return null;
    },
  };
}

export function createInMemoryIntermediateCaseStorageAdapter(
  seed?: Record<string, string>
): IntermediateCaseStorageAdapter {
  const store = new Map<string, string>(Object.entries(seed || {}));
  return {
    async readText(path) {
      return store.get(path) || null;
    },
    async writeText(path, content) {
      store.set(path, content);
    },
    async deletePath(path) {
      store.delete(path);
    },
    async listPaths(prefix) {
      return Array.from(store.keys()).filter((path) => path.startsWith(prefix));
    },
  };
}
