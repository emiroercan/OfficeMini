// End-to-end test of the Chrome extension: the real extension in a real Chromium, driven over the
// DevTools protocol. Every case is a genuine click or a genuine navigation - nothing in the
// extension is stubbed, which is the point: stubbed checks passed while export buttons, blob
// downloads and local files all failed in a real browser.
//
//   npm run test:ext                   build, then run in every browser found (in parallel)
//   node scripts/e2e-ext.mjs edge      one browser: edge | chrome   (OM_EDGE / OM_CHROME = exe)
//
// Edge loads the extension with --load-extension. Branded Chrome has ignored that flag since 137,
// so there it is loaded with Extensions.loadUnpacked over --remote-debugging-pipe - and Chrome
// disables an extension loaded that way the moment file access is switched off
// (disableReasons.unsupportedDeveloperExtension). The file-access phases therefore run in Edge
// only. An ordinary "Load unpacked" install is not affected.
import http from "node:http";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { zipSync, strToU8 } from "fflate";
import { launch, targets, connect, attach, sleep } from "./cdp.mjs";

const ROOT = path.resolve(import.meta.dirname, "..");
const EXT = path.join(ROOT, "dist-ext");

const CANDIDATES = {
  edge: [process.env.OM_EDGE, "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe", "C:/Program Files/Microsoft/Edge/Application/msedge.exe",
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge", "/usr/bin/microsoft-edge"],
  chrome: [process.env.OM_CHROME, "C:/Program Files/Google/Chrome/Application/chrome.exe", path.join(process.env.LOCALAPPDATA || "/nonexistent", "Google/Chrome/Application/chrome.exe"),
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome", "/usr/bin/google-chrome"],
};
const findExe = (tag) => (CANDIDATES[tag] || []).find((p) => p && fs.existsSync(p)) || null;

// ---- fixtures: generated, so the test never depends on anyone's documents ----------

const XLSX_A1 = "OfficeMini e2e", CSV_A1 = "Müşteri";
const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

function minimalXlsx() {
  const xml = (s) => strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s);
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  return zipSync({
    "[Content_Types].xml": xml('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      + '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'),
    "_rels/.rels": xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": xml(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": xml('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
      + `<row r="1"><c r="A1" t="inlineStr"><is><t>${XLSX_A1}</t></is></c><c r="B1"><v>42</v></c></row></sheetData></worksheet>`),
  });
}
const XLSX = Buffer.from(minimalXlsx());
const CSV = Buffer.from(`${CSV_A1},Şehir\nAyşe,İzmir\nMehmet,Ankara\n`, "utf-8");

const PAGE = `<!doctype html><meta charset=utf-8><title>download cases</title>
<style>body{font:16px system-ui;line-height:2.2}</style>
<a id="plain" href="/files/Plain.xlsx">plain .xlsx link</a><br>
<a id="cd" href="/export?id=7">export endpoint: Content-Disposition, octet-stream, no extension in the URL</a><br>
<a id="csv" href="/files/list.csv">.csv link</a><br>
<button id="blob">blob download, revoked at once</button><br>
<button id="blobslow">blob download, revoked later</button><br>
<a id="pdf" href="/files/manual.pdf">a .pdf, which is not ours</a>
<script>
async function blobDl(name, now) {
  const b = await (await fetch('/files/Plain.xlsx')).blob();
  const u = URL.createObjectURL(b);
  const a = document.createElement('a'); a.href = u; a.download = name; document.body.append(a); a.click();
  if (now) URL.revokeObjectURL(u); else setTimeout(() => URL.revokeObjectURL(u), 40000);
}
document.getElementById('blob').onclick = () => blobDl('Blob Now.xlsx', true);
document.getElementById('blobslow').onclick = () => blobDl('Blob Later.xlsx', false);
</script>`;

function serve(port) {
  const server = http.createServer((req, res) => {
    const u = new URL(req.url, "http://x");
    if (u.pathname === "/page.html") { res.writeHead(200, { "content-type": "text/html; charset=utf-8" }); return res.end(PAGE); }
    if (u.pathname === "/files/Plain.xlsx") { res.writeHead(200, { "content-type": XLSX_MIME }); return res.end(XLSX); }
    if (u.pathname === "/export") { res.writeHead(200, { "content-type": "application/octet-stream", "content-disposition": 'attachment; filename="Export.xlsx"' }); return res.end(XLSX); }
    if (u.pathname === "/files/list.csv") { res.writeHead(200, { "content-type": "text/csv; charset=utf-8" }); return res.end(CSV); }
    if (u.pathname === "/files/manual.pdf") { res.writeHead(200, { "content-type": "application/octet-stream", "content-disposition": 'attachment; filename="manual.pdf"' }); return res.end("%PDF-1.4\n%fake\n"); }
    res.writeHead(404); res.end();
  });
  return new Promise((r) => server.listen(port, "127.0.0.1", () => r(server)));
}

// ---- one browser ---------------------------------------------------------------

async function suite(tag, exe, { port, web }) {
  const WORK = path.join(os.tmpdir(), "officemini-e2e-" + tag);
  const DL = path.join(WORK, "downloads"), LOCAL = path.join(WORK, "local");
  fs.rmSync(WORK, { recursive: true, force: true });
  fs.mkdirSync(DL, { recursive: true });
  fs.mkdirSync(LOCAL, { recursive: true });
  const writeLocal = () => { fs.writeFileSync(path.join(LOCAL, "Local Book.xlsx"), XLSX); fs.writeFileSync(path.join(LOCAL, "local list.csv"), CSV); };
  writeLocal();

  const rows = [], notes = [];
  const server = await serve(web);
  const pipe = tag === "chrome";
  const b = await launch({
    exe, port, profile: path.join(WORK, "profile"), pipe,
    args: pipe ? [] : [`--load-extension=${EXT}`, `--disable-extensions-except=${EXT}`],
    prepare: (profile) => {
      // A realistic download setup: a fixed folder and no "ask where to save" prompt.
      fs.mkdirSync(path.join(profile, "Default"), { recursive: true });
      fs.writeFileSync(path.join(profile, "Default", "Preferences"), JSON.stringify({
        download: { default_directory: DL, prompt_for_download: false, directory_upgrade: true },
        savefile: { default_directory: DL },
      }));
    },
  });

  let cdp = null;
  try {
    if (pipe) await b.pipeSend("Extensions.loadUnpacked", { path: EXT });
    cdp = connect(b.version.webSocketDebuggerUrl);
    await cdp.ready;

    // -- the service worker, found again after a reload --
    let sw = null, extId = null;
    const swLog = [];
    cdp.on((m) => {
      if (!sw || m.sessionId !== sw.sessionId) return;
      if (m.method === "Runtime.exceptionThrown") swLog.push("EXCEPTION " + (m.params.exceptionDetails.exception?.description || m.params.exceptionDetails.text));
      if (m.method === "Runtime.consoleAPICalled" && m.params.type === "error") swLog.push("console.error " + m.params.args.map((a) => a.value ?? a.description).join(" "));
    });
    /** Find (and listen to) the worker; `wake` opens the popup, which messages it, if it is asleep. */
    async function findSW({ wake = true } = {}) {
      sw = null;
      for (let round = 0; round < 2 && !sw; round++) {
        for (let i = 0; i < 24 && !sw; i++) {
          for (const t of await targets(port)) {
            if (t.type !== "service_worker" || !t.url.endsWith("/background.js")) continue;
            try {
              const s = await attach(cdp, t.id);
              if ((await s.evaluate("chrome.runtime.getManifest().name")) === "OfficeMini") { sw = s; extId = new URL(t.url).host; break; }
            } catch { /* a worker that is going away */ }
          }
          if (!sw) await sleep(250);
        }
        if (!sw && wake && extId && round === 0) {
          const { targetId } = await cdp.send("Target.createTarget", { url: `chrome-extension://${extId}/popup.html` });
          await sleep(1500);
          await cdp.send("Target.closeTarget", { targetId }).catch(() => {});
        }
      }
      if (sw) await sw.send("Runtime.enable");
      return !!sw;
    }
    if (!(await findSW({ wake: false }))) throw new Error("the OfficeMini service worker never started - is dist-ext built?");

    const settings = (patch) => sw.evaluate(`(async () => { const k = 'settings'; const cur = (await chrome.storage.sync.get(k))[k] || {}; await chrome.storage.sync.set({ [k]: { ...cur, ...${JSON.stringify(patch)} } }); })()`);

    /** Flip "Allow access to file URLs" the way the chrome://extensions toggle does. */
    async function setFileAccess(on) {
      const { targetId } = await cdp.send("Target.createTarget", { url: "chrome://extensions/" });
      const page = await attach(cdp, targetId);
      await sleep(1500);
      await page.evaluate(`new Promise((res) => chrome.developerPrivate.updateExtensionConfiguration({ extensionId: ${JSON.stringify(extId)}, fileAccess: ${on} }, res))`);
      await cdp.send("Target.closeTarget", { targetId }).catch(() => {});
      await sleep(2500);
      const left = (await targets(port)).filter((t) => t.url.includes(extId)).map((t) => t.type + ":" + t.url.split(extId)[1].split("?")[0]);
      notes.push(`after fileAccess=${on}: ${left.length ? left.join(", ") : "nothing of ours running"}`);
    }

    // -- helpers --
    const pages = async () => (await targets(port)).filter((t) => t.type === "page");
    const ours = async (file) => (await pages()).filter((t) => t.url.startsWith(`chrome-extension://${extId}/${file}`));
    const dlFiles = () => fs.readdirSync(DL).filter((f) => !/\.(crdownload|tmp)$/.test(f)).sort();

    async function newTab(url) {
      const { targetId } = await cdp.send("Target.createTarget", { url });
      const s = await attach(cdp, targetId);
      await s.send("Page.enable").catch(() => {});
      await sleep(1500);
      return { targetId, ...s };
    }
    async function click(tab, id) {
      const pt = await tab.evaluate(`(() => { const r = document.getElementById(${JSON.stringify(id)}).getBoundingClientRect(); return { x: r.x + r.width / 2, y: r.y + r.height / 2 }; })()`);
      await tab.send("Page.bringToFront");
      for (const type of ["mousePressed", "mouseReleased"]) await tab.send("Input.dispatchMouseEvent", { type, x: pt.x, y: pt.y, button: "left", clickCount: 1 });
    }
    async function inspectEditor(t) {
      const s = await attach(cdp, t.id);
      for (let i = 0; i < 40; i++) {
        const v = await s.evaluate(`(() => { const o = window.om; if (!o || !o.app || !o.app.path) return null;
          const a1 = o.sheet ? (o.sheet().cells.get(0) || {}).v : undefined;
          return { name: o.app.path.split('/').pop(), a1 }; })()`).catch(() => null);
        if (v) return v;
        await sleep(250);
      }
      return { name: null, error: "editor did not load: " + t.url.slice(0, 120) };
    }
    async function reset() {
      for (const t of await pages()) if (t.url !== "about:blank") await cdp.send("Target.closeTarget", { targetId: t.id }).catch(() => {});
      for (const f of fs.readdirSync(DL)) fs.rmSync(path.join(DL, f), { force: true });
      writeLocal();
    }
    async function observe(wait = 6000) {
      await sleep(wait);
      const editors = [];
      for (const t of await ours("index.html")) editors.push(await inspectEditor(t));
      return { editors, explainer: (await ours("file-access.html")).length > 0, downloads: dlFiles() };
    }
    function verdict(phase, name, got, want, errs) {
      const fails = [];
      const names = got.editors.map((e) => e.name);
      if (want.editor === null && names.length) fails.push("unexpected editor " + names.join(","));
      if (want.editor && !(names.length === 1 && names[0] === want.editor)) fails.push(`editor ${JSON.stringify(names)} != ["${want.editor}"]`);
      if (want.a1 !== undefined && got.editors[0] && got.editors[0].a1 !== want.a1) fails.push(`A1 ${JSON.stringify(got.editors[0].a1)} != "${want.a1}"`);
      if (!!want.explainer !== got.explainer) fails.push("explainer " + got.explainer);
      if (JSON.stringify(got.downloads) !== JSON.stringify(want.downloads)) fails.push(`Downloads ${JSON.stringify(got.downloads)} != ${JSON.stringify(want.downloads)}`);
      for (const f of ["Local Book.xlsx", "local list.csv"]) if (!fs.existsSync(path.join(LOCAL, f))) fails.push("LOCAL ORIGINAL DELETED: " + f);
      for (const e of got.editors) if (e.error) fails.push(e.error);
      fails.push(...errs);
      rows.push({ phase, case: name, result: fails.length ? "FAIL" : "pass", detail: fails.join("; ") });
    }
    async function run(phase, name, act, want) {
      await reset();
      const before = swLog.length;
      await act();
      verdict(phase, name, await observe(), want, swLog.slice(before));
    }

    const pageUrl = `http://127.0.0.1:${web}/page.html`;
    const webCase = (id) => async () => { const tab = await newTab(pageUrl); await click(tab, id); };
    const localCase = (file) => async () => { await newTab(pathToFileURL(path.join(LOCAL, file)).href); };

    // -- phase 1: file access on, defaults (nothing kept in Downloads) --
    const on0 = await sw.evaluate("chrome.extension.isAllowedFileSchemeAccess()");
    const P1 = on0 ? "access on" : "access on (WAS OFF)";
    await run(P1, "plain link", webCase("plain"), { editor: "Plain.xlsx", a1: XLSX_A1, downloads: [] });
    await run(P1, "export endpoint", webCase("cd"), { editor: "Export.xlsx", a1: XLSX_A1, downloads: [] });
    await run(P1, "csv link", webCase("csv"), { editor: "list.csv", a1: CSV_A1, downloads: [] });
    await run(P1, "blob, revoked at once", webCase("blob"), { editor: "Blob Now.xlsx", a1: XLSX_A1, downloads: [] });
    await run(P1, "blob, revoked later", webCase("blobslow"), { editor: "Blob Later.xlsx", a1: XLSX_A1, downloads: [] });
    await run(P1, "pdf is left alone", webCase("pdf"), { editor: null, downloads: ["manual.pdf"] });
    await run(P1, "local .xlsx into a tab", localCase("Local Book.xlsx"), { editor: "Local Book.xlsx", a1: XLSX_A1, downloads: [] });
    await run(P1, "local .csv into a tab", localCase("local list.csv"), { editor: "local list.csv", a1: CSV_A1, downloads: [] });

    // -- phase 2: keep a copy --
    await settings({ keepCopy: true });
    await run("keep a copy", "export endpoint", webCase("cd"), { editor: "Export.xlsx", a1: XLSX_A1, downloads: ["Export.xlsx"] });
    await run("keep a copy", "blob, revoked at once", webCase("blob"), { editor: "Blob Now.xlsx", a1: XLSX_A1, downloads: ["Blob Now.xlsx"] });
    await settings({ keepCopy: false });

    // -- phase 3: master switch off --
    await settings({ enabled: false });
    await run("switched off", "export endpoint", webCase("cd"), { editor: null, downloads: ["Export.xlsx"] });
    await run("switched off", "local .xlsx into a tab", localCase("Local Book.xlsx"), { editor: null, downloads: ["Local Book.xlsx"] });
    await settings({ enabled: true });

    if (pipe) {
      rows.push({ phase: "access off / on", case: "(not run in Chrome)", result: "skip", detail: "Chrome disables a pipe-loaded extension when file access is switched off; Edge covers these" });
    } else {
      // -- phase 4: file access off --
      await setFileAccess(false);
      await findSW();
      const P4 = (sw && (await sw.evaluate("chrome.extension.isAllowedFileSchemeAccess()")) === false) ? "access off" : "access off (NOT OFF)";
      await run(P4, "plain link (fetched again)", webCase("plain"), { editor: "Plain.xlsx", a1: XLSX_A1, downloads: [] });
      await run(P4, "export endpoint (fetched again)", webCase("cd"), { editor: "Export.xlsx", a1: XLSX_A1, downloads: [] });
      await run(P4, "blob -> explainer, file kept", webCase("blob"), { editor: null, explainer: true, downloads: ["Blob Now.xlsx"] });
      await run(P4, "local .xlsx -> explainer", localCase("Local Book.xlsx"), { editor: null, explainer: true, downloads: [] });

      // -- phase 5: turn it on with the explainer open. Look BEFORE waking anything: the file
      // has to open on its own, because that is what a person gets.
      const before = swLog.length;
      await setFileAccess(true);
      verdict("access turned on", "pending opens with no help", await observe(5000), { editor: "Local Book.xlsx", a1: XLSX_A1, downloads: [] }, swLog.slice(before));
    }
  } catch (e) {
    rows.push({ phase: "harness", case: "-", result: "FAIL", detail: e.message });
  } finally {
    try { if (cdp) await cdp.send("Browser.close"); } catch { /* closing */ }
    await sleep(500);
    try { b.proc.kill(); } catch { /* gone */ }
    server.close();
  }
  return { browser: b.version.Browser, rows, notes };
}

// ---- main -----------------------------------------------------------------------

if (!fs.existsSync(path.join(EXT, "manifest.json"))) {
  console.error("dist-ext/manifest.json is missing - run `npm run build:ext` first.");
  process.exit(2);
}
const wanted = process.argv[2] ? [process.argv[2]] : Object.keys(CANDIDATES);
const todo = [];
for (const tag of wanted) {
  const exe = findExe(tag);
  if (exe) todo.push([tag, exe]); else console.log(`${tag}: not found${CANDIDATES[tag] ? "" : " (use edge or chrome)"}, skipped`);
}
if (!todo.length) { console.error("No browser to test with."); process.exit(2); }

const results = await Promise.all(todo.map(([tag, exe], i) => suite(tag, exe, { port: 9361 + i, web: 9411 + i })));
let failed = 0, passed = 0;
for (const r of results) {
  console.log("\n" + r.browser);
  console.table(r.rows);
  for (const n of r.notes) console.log("  note:", n);
  failed += r.rows.filter((x) => x.result === "FAIL").length;
  passed += r.rows.filter((x) => x.result === "pass").length;
}
console.log(failed ? `\n${failed} FAILED, ${passed} passed` : `\nall ${passed} passed`);
process.exit(failed ? 1 : 0);
