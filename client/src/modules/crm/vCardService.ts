/**
 * vCardService.ts
 * Generates .vcf files and Visual HTML Cards for Odoo contacts.
 */

export interface ContactMetadata {
    name: string;
    email?: string;
    phone?: string;
    mobile?: string;
    company?: string;
    jobTitle?: string;
    website?: string;
}

/**
 * Generates a VCF string for a contact.
 */
export function generateVCard(m: ContactMetadata): string {
    const vcard = [
        "BEGIN:VCARD",
        "VERSION:3.0",
        `FN:${m.name}`,
        m.company ? `ORG:${m.company}` : "",
        m.jobTitle ? `TITLE:${m.jobTitle}` : "",
        m.email ? `EMAIL;TYPE=INTERNET,WORK:${m.email}` : "",
        m.phone ? `TEL;TYPE=WORK,VOICE:${m.phone}` : "",
        m.mobile ? `TEL;TYPE=CELL,VOICE:${m.mobile}` : "",
        m.website ? `URL:${m.website}` : "",
        "END:VCARD"
    ].filter(Boolean).join("\n");

    return vcard;
}

/**
 * Downloads a contact as a .vcf file.
 */
export function downloadVCard(m: ContactMetadata) {
    const content = generateVCard(m);
    const blob = new Blob([content], { type: "text/vcard" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `${m.name.replace(/\s+/g, "_")}.vcf`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

/**
 * Generates an HTML string for the "Visual Card".
 * Designed for 13px high-density layout.
 */
export function generateVisualCardHtml(m: ContactMetadata): string {
    return `
<div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 13px; line-height: 1.2; color: #1e293b; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px; max-width: 300px; background: #ffffff; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
    <div style="font-weight: 800; color: #0f172a; margin-bottom: 4px;">${m.name}</div>
    ${m.jobTitle ? `<div style="color: #64748b; font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.02em; margin-bottom: 8px;">${m.jobTitle}</div>` : ""}
    <div style="display: grid; gap: 4px;">
        ${m.company ? `<div style="display: flex; gap: 6px;"><b>Empresa:</b> ${m.company}</div>` : ""}
        ${m.email ? `<div style="display: flex; gap: 6px;"><b>Email:</b> ${m.email}</div>` : ""}
        ${m.phone ? `<div style="display: flex; gap: 6px;"><b>Tel:</b> ${m.phone}</div>` : ""}
        ${m.mobile ? `<div style="display: flex; gap: 6px;"><b>Móvel:</b> ${m.mobile}</div>` : ""}
    </div>
</div>
`.trim();
}

/**
 * Copies the visual card (as rich text/HTML) to the clipboard.
 * Bypass for IT attachment restrictions.
 */
export async function copyVisualCardToClipboard(m: ContactMetadata) {
    const html = generateVisualCardHtml(m);
    const text = `${m.name}\n${m.jobTitle || ""}\n${m.company || ""}\n${m.email || ""}\n${m.phone || ""}\n${m.mobile || ""}`.trim();

    try {
        const type = "text/html";
        const blob = new Blob([html], { type });
        const data = [new ClipboardItem({
            [type]: blob,
            "text/plain": new Blob([text], { type: "text/plain" })
        })];
        await navigator.clipboard.write(data);
        return true;
    } catch (err) {
        console.error("Failed to copy visual card:", err);
        // Fallback to plain text
        await navigator.clipboard.writeText(text);
        return false;
    }
}
