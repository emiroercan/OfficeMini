// Which downloads the extension takes over. Shared by the service worker and the popup.

const KEY = "settings";

/** File types are offered as groups, because that is how people think about them. */
export const GROUPS = {
  word: { label: "Word documents", exts: ["docx", "docm", "dotx"] },
  excel: { label: "Excel workbooks", exts: ["xlsx", "xlsm", "xltx"] },
  csv: { label: "CSV and TSV", exts: ["csv", "tsv"] },
  text: { label: "Markdown and text", exts: ["md", "markdown", "txt"] },
};

// Markdown and text are off by default: the browser shows those links rather than downloading
// them, so the setting would mostly catch files someone deliberately asked to keep.
//
// keepCopy is off by default because the point of the extension is "open it instead of
// downloading it". The copy is only removed once the editor holds its own, and turning this on
// leaves every download exactly where Chrome put it.
export const DEFAULTS = { enabled: true, keepCopy: false, groups: { word: true, excel: true, csv: true, text: false } };

export async function loadSettings() {
  const got = await chrome.storage.sync.get(KEY);
  const s = got[KEY] || {};
  return { enabled: s.enabled !== false, keepCopy: s.keepCopy === true, groups: { ...DEFAULTS.groups, ...(s.groups || {}) } };
}

export async function saveSettings(s) {
  await chrome.storage.sync.set({ [KEY]: s });
}

/** The flat set of extensions the current settings ask for. */
export function activeExtensions(s) {
  const out = new Set();
  for (const [name, group] of Object.entries(GROUPS)) if (s.groups[name]) for (const e of group.exts) out.add(e);
  return out;
}
