// The toolbar popup: three ways into the editor, and the switch for taking over downloads.

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

const enabled = document.getElementById("enabled");
const groupBox = document.getElementById("groups");

let state = await loadSettings();

function render() {
  enabled.checked = state.enabled;
  groupBox.classList.toggle("off", !state.enabled);
  groupBox.innerHTML = "";
  for (const [name, group] of Object.entries(GROUPS)) {
    const label = document.createElement("label");
    const box = document.createElement("input");
    box.type = "checkbox";
    box.checked = !!state.groups[name];
    box.addEventListener("change", async () => {
      state = { ...state, groups: { ...state.groups, [name]: box.checked } };
      await saveSettings(state);
    });
    const text = document.createElement("span");
    text.textContent = group.label;
    const exts = document.createElement("span");
    exts.className = "exts";
    exts.textContent = group.exts.map((e) => "." + e).join(" ");
    label.append(box, text, exts);
    groupBox.append(label);
  }
}

enabled.addEventListener("change", async () => {
  state = { ...state, enabled: enabled.checked };
  await saveSettings(state);
  render();
});

render();
