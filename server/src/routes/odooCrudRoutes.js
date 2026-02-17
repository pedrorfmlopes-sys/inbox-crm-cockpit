import express from "express";

export function createOdooCrudRouter({ odooJsonRpc }) {
  const router = express.Router();

  router.post("/api/odoo/read", async (req, res) => {
    try {
      const { model, ids, fields } = req.body || {};
      if (!model || !Array.isArray(ids) || !Array.isArray(fields)) return res.status(400).send("Invalid payload");
      const result = await odooJsonRpc("/web/dataset/call_kw", {
        model,
        method: "read",
        args: [ids, fields],
        kwargs: {},
      });
      res.json(result);
    } catch (e) {
      res.status(500).send(String(e?.message || e));
    }
  });

  router.post("/api/odoo/write", async (req, res) => {
    try {
      const { model, id, values } = req.body || {};
      if (!model || !id || typeof values !== "object") return res.status(400).send("Invalid payload");
      const result = await odooJsonRpc("/web/dataset/call_kw", {
        model,
        method: "write",
        args: [[id], values],
        kwargs: {},
      });
      res.json(Boolean(result));
    } catch (e) {
      res.status(500).send(String(e?.message || e));
    }
  });

  router.post("/api/odoo/search_domain", async (req, res) => {
    try {
      const { model, domain, fields, limit } = req.body || {};
      if (!model || !Array.isArray(domain) || !Array.isArray(fields)) return res.status(400).send("Invalid payload");
      const result = await odooJsonRpc("/web/dataset/call_kw", {
        model,
        method: "search_read",
        args: [domain],
        kwargs: { fields, limit: limit ?? 20 },
      });
      res.json(result);
    } catch (e) {
      res.status(500).send(String(e?.message || e));
    }
  });

  return router;
}
