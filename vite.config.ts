import { defineConfig, Plugin } from "vite";
import fs from "node:fs";
import path from "node:path";
import { execSync } from "node:child_process";

/** Dev-only endpoint: POST /__save?name=<file> writes the body to samples/out (browser-mode testing). */
function devSave(): Plugin {
  return {
    name: "dev-save",
    configureServer(server) {
      server.middlewares.use("/__save", (req, res) => {
        if (req.method !== "POST") { res.statusCode = 405; res.end(); return; }
        const url = new URL(req.url || "", "http://localhost");
        const name = path.basename(decodeURIComponent(url.searchParams.get("name") || "out.bin"));
        const dir = path.resolve(process.cwd(), "samples", "out");
        fs.mkdirSync(dir, { recursive: true });
        const chunks: Buffer[] = [];
        req.on("data", (c: Buffer) => chunks.push(c));
        req.on("end", () => {
          fs.writeFileSync(path.join(dir, name), Buffer.concat(chunks));
          res.setHeader("content-type", "application/json");
          res.end(JSON.stringify({ ok: true, path: path.join(dir, name), size: Buffer.concat(chunks).length }));
        });
      });
    },
  };
}

const pkg = JSON.parse(fs.readFileSync(path.resolve(__dirname, "package.json"), "utf-8")) as { version: string };
let commit = "";
try { commit = execSync("git rev-parse --short HEAD", { stdio: ["ignore", "pipe", "ignore"] }).toString().trim(); } catch { /* not a git checkout */ }
const buildInfo = `${commit || "local"}, ${new Date().toISOString().slice(0, 10)}`;

// Tauri expects a fixed port; fail if that port is not available.
export default defineConfig({
  clearScreen: false,
  plugins: [devSave()],
  define: { __APP_VERSION__: JSON.stringify(pkg.version), __BUILD_INFO__: JSON.stringify(buildInfo) },
  server: {
    port: 1420,
    strictPort: true,
    watch: { ignored: ["**/src-tauri/**", "**/samples/**"] },
  },
  build: {
    target: ["es2022", "chrome110", "safari16"],
    minify: "esbuild",
    sourcemap: false,
    outDir: "dist",
    emptyOutDir: true,
    chunkSizeWarningLimit: 1500,
  },
  esbuild: { legalComments: "none" },
});
