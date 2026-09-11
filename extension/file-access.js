// Shown when a download or a local file could not be read because file access is off.
// Watches for the switch and, once it is on, asks the service worker to finish opening the file.

const q = new URLSearchParams(location.search);
document.getElementById("name").textContent = "“" + (q.get("name") || "your file") + "”";
document.getElementById("go").addEventListener("click", () => {
  chrome.tabs.create({ url: "chrome://extensions/?id=" + chrome.runtime.id });
});

async function watch() {
  if (!(await chrome.extension.isAllowedFileSchemeAccess())) { setTimeout(watch, 1000); return; }
  document.getElementById("status").textContent = "Opening…";
  try { await chrome.runtime.sendMessage({ type: "resume" }); } catch { /* the worker restarted; it resumes on its own */ }
  window.close();
}
watch();
