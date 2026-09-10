// Zip dist-ext into an archive the Chrome Web Store accepts: run `npm run pack:ext`.
// The store wants the manifest at the root of the zip, not inside a folder.

import { zipSync } from "fflate";
import fs from "node:fs";
import path from "node:path";

const root = path.resolve(import.meta.dirname, "..");
const dir = path.join(root, "dist-ext");
if (!fs.existsSync(path.join(dir, "manifest.json"))) {
  console.error("dist-ext/manifest.json is missing - run `npm run build:ext` first.");
  process.exit(1);
}

const files = {};
(function walk(abs, rel) {
  for (const entry of fs.readdirSync(abs, { withFileTypes: true })) {
    const next = path.join(abs, entry.name);
    const key = rel ? rel + "/" + entry.name : entry.name;
    if (entry.isDirectory()) walk(next, key);
    else files[key] = new Uint8Array(fs.readFileSync(next));
  }
})(dir, "");

const version = JSON.parse(fs.readFileSync(path.join(dir, "manifest.json"), "utf-8")).version;
const out = path.join(root, `officemini-extension-${version}.zip`);
fs.writeFileSync(out, zipSync(files, { level: 9 }));

const kb = (n) => (n / 1024).toFixed(0) + " kB";
console.log(`${path.relative(root, out)}  ${kb(fs.statSync(out).size)}  (${Object.keys(files).length} files)`);
