import { supportsIntermediateCaseBinaryStorage, type IntermediateCaseStorageAdapter } from "./intermediateCaseRepository";
import type { IntermediateCase } from "./intermediateCaseTypes";

export type IntermediateAttachmentBinarySource = {
  path: string;
  contentBase64: string;
  contentType?: string;
};

function decodeBase64ToBytes(base64: string): Uint8Array | null {
  const raw = String(base64 || "").trim().replace(/^data:[^,]+,/, "");
  if (!raw) return null;
  try {
    const binary = atob(raw);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) {
      bytes[index] = binary.charCodeAt(index);
    }
    return bytes;
  } catch {
    return null;
  }
}

export async function persistIntermediateCaseAttachmentBinaries(args: {
  adapter: IntermediateCaseStorageAdapter;
  caseValue: IntermediateCase;
  binarySources: IntermediateAttachmentBinarySource[];
}): Promise<number> {
  if (!supportsIntermediateCaseBinaryStorage(args.adapter)) return 0;
  const binarySourceMap = new Map(
    args.binarySources
      .filter((entry) => entry.path && entry.contentBase64)
      .map((entry) => [entry.path, entry])
  );
  let writes = 0;
  for (const email of args.caseValue.emails) {
    for (const attachment of email.attachments) {
      const path = String(attachment.localRef?.value || "").trim();
      if (!path) continue;
      const source = binarySourceMap.get(path);
      if (!source) continue;
      const bytes = decodeBase64ToBytes(source.contentBase64);
      if (!bytes) continue;
      await args.adapter.writeBinary(path, new Blob([bytes], { type: source.contentType || attachment.contentType || "application/octet-stream" }));
      writes += 1;
    }
  }
  return writes;
}
