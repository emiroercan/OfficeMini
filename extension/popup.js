// The toolbar popup: three ways into the editor, and the switches for taking over downloads.

import { GROUPS, loadSettings, saveSettings } from "./settings.js";

const editor = (query) => chrome.runtime.getURL("index.html") + (query ? "?" + query : "");

function openTab(url) {
  chrome.tabs.create({ url });
  window.close();
}

document.getElementById("version").textContent = "v" + chrome.runtime.getManifest().version;
document.getElementById("new-doc").addEventListener("click", () => openTab(editor("new=docx")));
document.getElementById("new-sheet").addEventListener("click", () => openTab(editor("new=xlsx")));
// The picker needs a user gesture in the page that shows it, so this opens the editor's welcome
// screen and the Open button there raises the dialog.
document.getElementById("open").addEventListener("click", () => openTab(editor("")));
document.getElementById("access-btn").addEventListener("click", () => openTab("chrome://extensions/?id=" + chrome.runtime.id));

const enabled = document.getElementById("enabled");
const opts = document.getElementById("opts");
const groupBox = document.getElementById("groups");
const keep = document.getElementById("keep");

let state = await loadSettings();

async function update(patch) {
  state = { ...state, ...patch };
  await saveSettings(state);
  render();
}

function render() {
  enabled.checked = state.enabled;
  keep.checked = state.keepCopy;
  opts.classList.toggle("off", !state.enabled);
  groupBox.innerHTML = "";
  for (const [name, group] of Object.entries(GROUPS)) {
    const label = document.createElement("label");
    const box = document.createElement("input");
    box.type = "checkbox";
    box.checked = !!state.groups[name];
    box.addEventListener("change", () => update({ groups: { ...state.groups, [name]: box.checked } }));
    const text = document.createElement("span");
    text.textContent = group.label;
    const exts = document.createElement("span");
    exts.className = "exts";
    exts.textContent = group.exts.map((e) => "." + e).join(" ");
    label.append(box, text, exts);
    groupBox.append(label);
  }
}

enabled.addEventListener("change", () => update({ enabled: enabled.checked }));
keep.addEventListener("change", () => update({ keepCopy: keep.checked }));
document.getElementById("access").hidden = await chrome.extension.isAllowedFileSchemeAccess();
// Opening the popup wakes the worker, which finishes a file that was waiting on that switch.
chrome.runtime.sendMessage({ type: "resume" }).catch(() => {});

render();
