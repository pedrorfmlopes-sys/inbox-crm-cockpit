import { PDFParse } from 'pdf-parse';

export async function extractTextFromPdfBuffer(buffer) {
    let parser = null;
    try {
        parser = new PDFParse({ data: buffer });
        const result = await parser.getText();
        console.log(`[pdf] extracted ${result.text?.length || 0} characters`);
        return result.text || "";
    } catch (e) {
        console.error("[pdf] Extraction failed:", e.message);
        return "";
    } finally {
        if (parser) {
            try { await parser.destroy(); } catch (err) { /* ignore */ }
        }
    }
}
