import type { IntermediateCase } from "./intermediateCaseTypes";

export const INTERMEDIATE_CASES_ROOT = "Groups/cases";
export const INTERMEDIATE_CASE_FILE_NAME = "case.json";
export const INTERMEDIATE_CASE_ATTACHMENTS_DIR = "attachments";

function sanitizeSegment(value: string): string {
  return String(value || "").trim().replace(/[\\/:*?"<>|]+/g, "_").replace(/\s+/g, "_");
}

export function getIntermediateCaseFolder(caseId: string): string {
  return `${INTERMEDIATE_CASES_ROOT}/${sanitizeSegment(caseId)}`;
}

export function getIntermediateCaseJsonPath(caseId: string): string {
  return `${getIntermediateCaseFolder(caseId)}/${INTERMEDIATE_CASE_FILE_NAME}`;
}

export function getIntermediateCaseAttachmentsFolder(caseId: string): string {
  return `${getIntermediateCaseFolder(caseId)}/${INTERMEDIATE_CASE_ATTACHMENTS_DIR}`;
}

export function getIntermediateCaseEmailAttachmentsFolder(caseId: string, emailKey: string): string {
  return `${getIntermediateCaseAttachmentsFolder(caseId)}/${sanitizeSegment(emailKey)}`;
}

export function getIntermediateCaseAttachmentPath(
  caseId: string,
  emailKey: string,
  attachmentKey: string,
  attachmentName?: string
): string {
  const baseName = sanitizeSegment(attachmentName || attachmentKey);
  return `${getIntermediateCaseEmailAttachmentsFolder(caseId, emailKey)}/${sanitizeSegment(attachmentKey)}-${baseName}`;
}

export function buildIntermediateCaseStorageLayout(input: Pick<IntermediateCase, "caseId" | "emails">) {
  return {
    caseFolder: getIntermediateCaseFolder(input.caseId),
    caseJsonPath: getIntermediateCaseJsonPath(input.caseId),
    attachmentsFolder: getIntermediateCaseAttachmentsFolder(input.caseId),
    emailAttachmentFolders: input.emails.map((email) => ({
      emailKey: email.emailKey,
      folder: getIntermediateCaseEmailAttachmentsFolder(input.caseId, email.emailKey),
    })),
  };
}
