
import * as pdfjsLib from 'pdfjs-dist';

// Point to the worker file in public folder (copied from node_modules)
// This avoids CDN issues in Outlook add-in environment
pdfjsLib.GlobalWorkerOptions.workerSrc = `/pdf.worker.min.mjs`;

export async function extractTextFromPdf(file: File): Promise<string> {
    try {
        const arrayBuffer = await file.arrayBuffer();
        const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
        const pdf = await loadingTask.promise;

        let fullText = "";

        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();

            // Join items with a space, but try to respect layout slightly
            const pageText = textContent.items
                .map((item: any) => item.str)
                .join(" ");

            fullText += `--- PÁGINA ${i} ---\n${pageText}\n\n`;
        }

        return fullText;
    } catch (error) {
        console.error("PDF Extraction Error:", error);
        return `[Erro ao ler PDF: ${error instanceof Error ? error.message : String(error)}]`;
    }
}
