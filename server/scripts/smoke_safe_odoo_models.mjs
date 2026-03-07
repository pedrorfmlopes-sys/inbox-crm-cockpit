import { odooClientFromEnv } from "../src/odoo.js";

const MODELS = [
  { model: "res.partner", domain: [], wanted: ["id", "name", "email", "mobile", "__bad_field__"] },
  { model: "crm.lead", domain: [], wanted: ["id", "name", "email_from", "__bad_field__"] },
  { model: "project.project", domain: [], wanted: ["id", "name", "partner_id", "__bad_field__"] },
  { model: "project.task", domain: [], wanted: ["id", "name", "project_id", "stage_id", "__bad_field__"] },
];

function hasEnv() {
  return Boolean(process.env.ODOO_URL && process.env.ODOO_DB && (process.env.ODOO_USERNAME || process.env.ODOO_USER) && (process.env.ODOO_API_KEY || process.env.ODOO_PASS || process.env.ODOO_PASSWORD));
}

async function main() {
  if (!hasEnv()) {
    console.log("skip: ODOO env vars not configured; smoke script not executed against live Odoo");
    return;
  }

  const odoo = await odooClientFromEnv();

  for (const row of MODELS) {
    const fields = await odoo.schema.getModelFields(row.model);
    console.log(`schema:${row.model}:count=${fields.size}`);

    const records = await odoo.safeSearchRead(row.model, row.domain, row.wanted, 1);
    console.log(`safeSearchRead:${row.model}:ok=${Array.isArray(records)}`);
  }

  console.log("ok: safe odoo smoke finished");
}

main().catch((e) => {
  console.error("fail:", e?.message || e);
  process.exit(1);
});
