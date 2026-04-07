import { type RelatedEmailEntry } from "@/api";

/**
 * Escapes HTML special characters in a string.
 */
export function escapeHtml(value: string): string {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * Sanitizes HTML string for preview, removing potentially dangerous elements and attributes.
 */
export function sanitizeEmailPreviewHtml(html: string): string {
  const raw = String(html || "")
    .replace(/<!--[\s\S]*?-->/g, " ")
    .replace(/<\?xml[\s\S]*?\?>/gi, " ")
    .replace(/<\/?(xml|o:[^>\s]+|v:[^>\s]+)[^>]*>/gi, " ")
    .trim();
  if (!raw) return "";

  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(raw, "text/html");
    doc.querySelectorAll("script, noscript, iframe, object, embed, form, link[rel='stylesheet'], meta[http-equiv], base, svg").forEach((node) => node.remove());
    doc.querySelectorAll<HTMLElement>("*").forEach((element) => {
      Array.from(element.attributes).forEach((attribute) => {
        const name = String(attribute.name || "").toLowerCase();
        const value = String(attribute.value || "").trim();
        if (!name) return;
        if (name.startsWith("on")) {
          element.removeAttribute(attribute.name);
          return;
        }
        if (name === "style" && /url\s*\(/i.test(value)) {
          element.removeAttribute(attribute.name);
          return;
        }
        if (!["src", "href", "poster", "background", "data"].includes(name)) return;
        if (/^(cid|javascript|vbscript|file|ms-appx|about):/i.test(value)) {
          if (element.tagName === "IMG") {
            const fallbackLabel = element.getAttribute("alt") || element.getAttribute("title") || "Imagem inline indisponivel neste preview.";
            element.setAttribute("alt", fallbackLabel);
          }
          element.removeAttribute(attribute.name);
        }
      });
    });
    return String(doc.body?.innerHTML || "")
      .replace(/<!--[\s\S]*?-->/g, " ")
      .replace(/<script[\s\S]*?<\/script>/gi, "")
      .trim();
  } catch {
    return raw
      .replace(/<!--[\s\S]*?-->/g, " ")
      .replace(/<script[\s\S]*?<\/script>/gi, "")
      .replace(/<noscript[\s\S]*?<\/noscript>/gi, "")
      .replace(/<svg[\s\S]*?<\/svg>/gi, "")
      .replace(/\s(on\w+)=(".*?"|'.*?'|[^\s>]+)/gi, "")
      .replace(/\s(style)=(".*?url\s*\(.*?\).*?"|'.*?url\s*\(.*?\).*?'|[^\s>]+)/gi, "")
      .replace(/\s(src|href|poster|background|data)=("cid:[^"]*"|'cid:[^']*'|cid:[^\s>]+)/gi, "");
  }
}

/**
 * Builds the final HTML for email body preview, prioritizing HTML and fallbacking to plain text with escaping.
 */
export function buildEmailPreviewHtml(email: RelatedEmailEntry | null): string {
  const html = String(email?.bodyHtml || "").trim();
  if (html) {
    const sanitizedHtml = sanitizeEmailPreviewHtml(html);
    if (sanitizedHtml) {
      return `<div style="padding:18px;color:#172b4d;font:14px/1.5 'Segoe UI',sans-serif;word-break:break-word">${sanitizedHtml}</div>`;
    }
  }
  const text = String(email?.bodyText || "").trim();
  if (!text) return "";
  return `<pre style="margin:0;padding:18px;color:#172b4d;background:#fff;font:14px/1.55 'Segoe UI',sans-serif;white-space:pre-wrap;word-break:break-word">${escapeHtml(text)}</pre>`;
}

/**
 * Decodes base64 content to string using TextDecoder.
 */
export function decodeBase64Text(content: string): string {
  try {
    const binary = globalThis.atob(String(content || "").trim());
    const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
    return new TextDecoder("utf-8").decode(bytes);
  } catch {
    return "";
  }
}

/**
 * Strips 'data:xxx/yyy;base64,' prefix from a string.
 */
export function stripDataUrlPrefix(value: string): string {
  const raw = String(value || "").trim();
  const separatorIndex = raw.indexOf(",");
  if (raw.startsWith("data:") && separatorIndex >= 0) return raw.slice(separatorIndex + 1);
  return raw;
}

/**
 * Checks if the current environment allows using the Office Online Web Viewer (requires public HTTPS).
 */
export function canUseOfficeWebViewer(): boolean {
  try {
    const url = new URL(window.location.origin);
    const hostname = String(url.hostname || "").trim().toLowerCase();
    return Boolean(
      /^https?:$/i.test(url.protocol)
      && hostname
      && hostname !== "localhost"
      && hostname !== "127.0.0.1"
    );
  } catch {
    return false;
  }
}

/**
 * Builds the Office Web Viewer embed URL for a given public source URL.
 */
export function buildOfficePreviewUrl(sourceUrl: string): string {
  const normalizedSourceUrl = String(sourceUrl || "").trim();
  if (!normalizedSourceUrl || !canUseOfficeWebViewer()) return "";
  return `https://view.officeapps.live.com/op/embed.aspx?src=${encodeURIComponent(normalizedSourceUrl)}`;
}

/**
 * Converts a data URL to Uint8Array after stripping the prefix.
 */
export function dataUrlToUint8Array(dataUrl: string): Uint8Array {
  const base64 = stripDataUrlPrefix(dataUrl);
  const binary = globalThis.atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}
