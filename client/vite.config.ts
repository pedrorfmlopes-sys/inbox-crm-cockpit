import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import fs from "node:fs";
import path from "node:path";
import os from "node:os";

const certDir = path.join(os.homedir(), ".office-addin-dev-certs");

function attachProxyDebug(proxy, label) {
  proxy.on("error", (err, _req, _res) => {
    // This is the #1 cause of "HTTP 500:" with empty body in the UI (proxy ECONNREFUSED).
    console.error(`[proxy:${label}] ERROR`, err?.code || "", err?.message || err);
  });
  proxy.on("proxyReq", (proxyReq, req, _res) => {
    console.log(`[proxy:${label}] -> ${req.method} ${req.url}  (target: ${proxyReq.protocol}//${proxyReq.host})`);
  });
  proxy.on("proxyRes", (proxyRes, req, _res) => {
    console.log(`[proxy:${label}] <- ${proxyRes.statusCode} ${req.method} ${req.url}`);
  });
}

export default defineConfig({
  plugins: [react()],
  server: {
    host: true,
    port: 5174,
    strictPort: true,
    https:
      fs.existsSync(path.join(certDir, "localhost.key")) && fs.existsSync(path.join(certDir, "localhost.crt"))
        ? {
          key: fs.readFileSync(path.join(certDir, "localhost.key")),
          cert: fs.readFileSync(path.join(certDir, "localhost.crt")),
        }
        : undefined,
    proxy: {
      "/api": {
        // Use 127.0.0.1 to avoid IPv6/localhost resolution weirdness on some setups
        target: "http://127.0.0.1:7071",
        changeOrigin: true,
        secure: false,
        configure: (proxy) => attachProxyDebug(proxy, "api"),
      },
      "/health": {
        target: "http://127.0.0.1:7071",
        changeOrigin: true,
        secure: false,
        configure: (proxy) => attachProxyDebug(proxy, "health"),
      },
    },
  },
  build: { outDir: "dist" },
});
