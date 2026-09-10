// Entry point: pick the editor for the requested file. Words (docx/md) keeps
// the existing shell in main.ts; Sheets (xlsx/csv) loads lazily so the word
// processor's start-up cost is unchanged.
import { cliArgs, extname } from "./files";

const SHEET_EXT = new Set(["xlsx", "xlsm", "xltx", "csv", "tsv"]);

async function start() {
  // The browser extension hands a document in through the query string; turn it into the
  // ordinary "?file=" the rest of the app reads, before either editor loads.
  if (__WEB_BUILD__) await (await import("./files-web")).adoptEntryUrl();
  const params = new URLSearchParams(location.search);
  let file = params.get("file");
  const kind = params.get("new");
  if (!file && !kind) {
    try { const args = await cliArgs(); if (args.length) file = args[0]; } catch { /* browser mode */ }
  }
  // "new:xlsx" / "new:docx" pseudo paths let a window be opened directly on a blank editor.
  const newKind = file && file.startsWith("new:") ? file.slice(4) : kind;
  const wantSheets = newKind === "xlsx" || newKind === "csv" || (!newKind && file ? SHEET_EXT.has(extname(file)) : false);
  if (wantSheets) {
    document.title = "OfficeMini";
    await import("./sheets/app");
  } else {
    await import("./main");
  }
}

start();
