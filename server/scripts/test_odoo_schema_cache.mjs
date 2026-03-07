import { createOdooSchemaCache } from "../src/odoo_schema_cache.js";

function assert(cond, msg) {
  if (!cond) throw new Error(msg);
}

async function testFiltersInvalidFieldOnRetry() {
  let fetchCount = 0;
  const cache = createOdooSchemaCache({
    ttlMs: 60_000,
    fetchFields: async (_model) => {
      fetchCount += 1;
      // first schema is stale and includes mobile; second refresh removes it
      if (fetchCount === 1) return { id: {}, name: {}, email: {}, mobile: {} };
      return { id: {}, name: {}, email: {} };
    },
  });

  let calls = 0;
  const result = await cache.safeSearchRead(
    "res.partner",
    [["email", "=", "a@b.com"]],
    ["id", "name", "email", "mobile"],
    1,
    {},
    async (_model, _domain, fields) => {
      calls += 1;
      if (fields.includes("mobile")) {
        throw new Error("ValueError: Invalid field 'mobile' on 'res.partner'");
      }
      return [{ id: 1, name: "A", email: "a@b.com" }];
    }
  );

  assert(calls === 2, "safeSearchRead should retry once after invalid field");
  assert(fetchCount >= 2, "schema should be refreshed after invalid field error");
  assert(Array.isArray(result) && result.length === 1, "safeSearchRead should return records after retry");
}

async function testSanitizeWrite() {
  const cache = createOdooSchemaCache({
    fetchFields: async () => ({ name: {}, email: {}, phone: {} }),
  });

  const sanitized = await cache.sanitizeWriteData("res.partner", {
    name: "X",
    email: "x@y.com",
    mobile: "+351999999999",
    unknown_field: "bad",
  });

  assert("name" in sanitized, "name should remain");
  assert("email" in sanitized, "email should remain");
  assert(!("mobile" in sanitized), "mobile should be dropped when absent from schema");
  assert(!("unknown_field" in sanitized), "unknown field should be dropped");
}

async function main() {
  await testFiltersInvalidFieldOnRetry();
  await testSanitizeWrite();
  console.log("ok: odoo schema cache tests passed");
}

main().catch((e) => {
  console.error("fail:", e?.message || e);
  process.exit(1);
});
