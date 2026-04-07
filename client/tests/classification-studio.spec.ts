import { test, expect, type Page, type ConsoleMessage } from '@playwright/test';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const REPORT_DIR = path.resolve(__dirname, 'reports');
const SCREENSHOT_DIR = path.resolve(REPORT_DIR, 'screenshots');

if (!fs.existsSync(SCREENSHOT_DIR)) {
  fs.mkdirSync(SCREENSHOT_DIR, { recursive: true });
}

test.describe('Classification Studio Local Validation', () => {
  let consoleErrors: { type: string, text: string }[] = [];
  let pageErrors: string[] = [];

  test.beforeEach(async ({ page }) => {
    consoleErrors = [];
    pageErrors = [];
    page.on('console', (msg: ConsoleMessage) => {
      if (msg.type() === 'error') {
        consoleErrors.push({ type: 'console.error', text: msg.text() });
      }
    });
    page.on('pageerror', (err: Error) => {
      pageErrors.push(err.message);
    });
  });

  async function takeScreenshot(page: Page, name: string) {
    await page.screenshot({ path: path.join(SCREENSHOT_DIR, `${name}.png`) });
  }

  test('Teste A — abertura do Studio', async ({ page }) => {
    // ABRIR STUDIO
    await page.goto('/?view=group-classification-studio');
    
    // Confirmar que o root aparece
    const studioRoot = page.locator('[data-testid="studio-root"]');
    await expect(studioRoot).toBeVisible({ timeout: 15000 });
    
    // Confirmar ausência de white screen (se o root está visível, não há white screen total)
    expect(pageErrors.length).toBe(0);
    
    await takeScreenshot(page, 'studio-opened');
  });

  test('Teste B — modo Normal', async ({ page }) => {
    await page.goto('/?view=group-classification-studio');
    await page.waitForSelector('[data-testid="studio-root"]');

    // Mudar para Normal se não for o padrão
    const normalBtn = page.locator('[data-testid="mode-normal-button"]');
    await normalBtn.click();

    // Validar presença visível de cards principais
    await expect(page.locator('[data-testid="emails-card"]')).toBeVisible();
    await expect(page.locator('[data-testid="quick-documents-card"]')).toBeVisible();
    await expect(page.locator('[data-testid="preview-pane"]')).toBeVisible();
    
    // Por defeito está em modo SUMÁRIO
    await expect(page.locator('[data-testid="classification-summary"]')).toBeVisible();

    // Abrir o editor clicando num tile (ex: principal)
    await page.locator('[data-testid="summary-tile-principal"]').click();
    await expect(page.locator('[data-testid="classification-editor"]')).toBeVisible();

    await takeScreenshot(page, 'mode-normal');
  });

  test('Teste C — modo Avançado', async ({ page }) => {
    await page.goto('/?view=group-classification-studio');
    await page.waitForSelector('[data-testid="studio-root"]');

    // Alternar para modo Avançado
    const advancedBtn = page.locator('[data-testid="mode-advanced-button"]');
    await advancedBtn.click();

    // Abrir o editor
    await page.locator('[data-testid="summary-tile-principal"]').click();
    await expect(page.locator('[data-testid="classification-editor"]')).toBeVisible();

    // Confirmar ausência de crash (page errors)
    expect(pageErrors.length).toBe(0);

    await takeScreenshot(page, 'mode-advanced');
  });

  test('Teste D — cards visuais e alturas', async ({ page }) => {
    await page.goto('/?view=group-classification-studio');
    await page.waitForSelector('[data-testid="studio-root"]');

    const results: any = { emailsHeight: null, docsHeight: null };

    // Verificar Emails
    const emailList = page.locator('[data-testid="emails-list"]');
    if (await emailList.isVisible()) {
        const firstEmail = emailList.locator('> div').first();
        if (await firstEmail.isVisible()) {
            const box = await firstEmail.boundingBox();
            results.emailsHeight = box?.height;
        }
    }

    // Verificar Documentos (se existir lista similar)
    const docList = page.locator('[data-testid="quick-documents-card"] .topCardScroll'); // Fallback se não tiver test-id específico
    if (await docList.isVisible()) {
        const firstDoc = docList.locator('> div').first();
        if (await firstDoc.isVisible()) {
            const box = await firstDoc.boundingBox();
            results.docsHeight = box?.height;
        }
    }

    console.log('MEASURED_HEIGHTS_JSON:' + JSON.stringify(results));
    await takeScreenshot(page, 'cards-visualization');
  });

  test('Teste E — modal “Aplicar a...”', async ({ page }) => {
    await page.goto('/?view=group-classification-studio');
    await page.waitForSelector('[data-testid="studio-root"]');

    // Mudar para o editor primeiro pois o botão save pode estar disabled sem alterações
    await page.locator('[data-testid="summary-tile-principal"]').click();
    await expect(page.locator('[data-testid="classification-editor"]')).toBeVisible();

    // Escrever algo para permitir guardar (enable button)
    const searchInput = page.locator('[data-testid="principal-search-input"]');
    await searchInput.fill('Teste Manual');

    // Clicar no botão principal de guardar
    const saveBtn = page.locator('[data-testid="main-save-button"]');
    await saveBtn.click();
    
    // Confirmar que o modal aparece
    const applyDialog = page.locator('[data-testid="apply-dialog"]');
    await expect(applyDialog).toBeVisible();
    
    // Confirmar botões principais do modal
    await expect(page.locator('[data-testid="apply-dialog-confirm"]')).toBeVisible();
    await expect(page.locator('[data-testid="apply-dialog-cancel"]').first()).toBeVisible();

    await takeScreenshot(page, 'apply-dialog-opened');

    // Fechar
    await page.locator('[data-testid="apply-dialog-cancel"]').first().click();
    await expect(applyDialog).not.toBeVisible();
  });

  test('Teste F — preview', async ({ page }) => {
    await page.goto('/?view=group-classification-studio');
    await page.waitForSelector('[data-testid="studio-root"]');

    const previewPane = page.locator('[data-testid="preview-pane"]');
    await expect(previewPane).toBeVisible();

    await takeScreenshot(page, 'preview-pane-visible');
  });

  test.afterAll(async () => {
      // Registrar erros para o relatório final se sobrar algum contexto
  });
});
