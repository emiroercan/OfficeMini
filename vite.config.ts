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

/** `--mode ext`: copy the extension's static half (manifest, service worker, popup, icons) into dist-ext. */
function extensionFiles(): Plugin {
  return {
    name: "extension-files",
    apply: "build",
    closeBundle() {
      const out = path.resolve(__dirname, "dist-ext");
      const src = path.resolve(__dirname, "extension");
      fs.cpSync(src, out, { recursive: true });
      // The manifest's version is the extension's identity in Chrome; keep it on package.json.
      const mf = path.join(out, "manifest.json");
      const m = JSON.parse(fs.readFileSync(mf, "utf-8"));
      m.version = pkg.version;
      fs.writeFileSync(mf, JSON.stringify(m, null, 2) + "\n");
      // Icons are the app's own, so there is one set to maintain.
      const icons = path.join(out, "icons");
      fs.mkdirSync(icons, { recursive: true });
      for (const [from, to] of [["32x32.png", "32.png"], ["128x128.png", "128.png"], ["icon.png", "512.png"]]) {
        fs.copyFileSync(path.resolve(__dirname, "src-tauri/icons", from), path.join(icons, to));
      }
      fs.copyFileSync(path.join(icons, "32.png"), path.join(icons, "16.png"));
      fs.copyFileSync(path.join(icons, "128.png"), path.join(icons, "48.png"));
    },
  };
}

// Three targets from one source tree: the default build is the Tauri app, `--mode web` is the
// browser build that talks to the File System Access API instead of the Rust backend, and
// `--mode ext` is that same browser build packaged as a Chrome extension. The flag is a
// compile-time constant, so each build folds the other one's branches away entirely.
export default defineConfig(({ mode }) => {
  const ext = mode === "ext";
  const web = mode === "web" || ext;
  return {
  clearScreen: false,
  // An extension page is loaded from chrome-extension://<id>/index.html, so its assets are
  // addressed relative to the page rather than from a server root.
  base: ext ? "./" : "/",
  plugins: [devSave(), ...(ext ? [extensionFiles()] : [])],
  define: { __APP_VERSION__: JSON.stringify(pkg.version), __BUILD_INFO__: JSON.stringify(buildInfo), __WEB_BUILD__: JSON.stringify(web) },
  server: {
    port: 1420,
    strictPort: true,
    watch: { ignored: ["**/src-tauri/**", "**/samples/**"] },
  },
  build: {
    target: ["es2022", "chrome110", "safari16"],
    minify: "esbuild",
    sourcemap: false,
    // The web builds never overwrite the app's dist: the Tauri bundle only ever ships ../dist.
    outDir: ext ? "dist-ext" : web ? "dist-web" : "dist",
    emptyOutDir: true,
    chunkSizeWarningLimit: 1500,
  },
  esbuild: { legalComments: "none" },
  };
});
