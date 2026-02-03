import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import fs from "fs";
import path from "path";

// Check if dev certs exist
const certPath = path.join(
  process.env.HOME || "",
  ".office-addin-dev-certs",
  "localhost.crt"
);
const keyPath = path.join(
  process.env.HOME || "",
  ".office-addin-dev-certs",
  "localhost.key"
);

const httpsConfig =
  fs.existsSync(certPath) && fs.existsSync(keyPath)
    ? {
        key: fs.readFileSync(keyPath),
        cert: fs.readFileSync(certPath),
      }
    : undefined;

export default defineConfig({
  plugins: [react()],
  base: "/",
  build: {
    outDir: "dist",
    sourcemap: true,
    rollupOptions: {
      input: {
        taskpane: "taskpane.html",
      },
    },
  },
  server: {
    port: 3000,
    https: httpsConfig,
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
    proxy: {
      "/api": {
        target: "https://engine.futuresearch.ai",
        changeOrigin: true,
        rewrite: (path) => path.replace(/^\/api/, ""),
        secure: true,
      },
    },
  },
});
