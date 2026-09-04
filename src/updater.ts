// Self-update from GitHub Releases via the Tauri updater plugin. Builds are
// signed in CI; the app verifies the signature before installing.
import { isTauri, showMessage } from "./files";
import { showDialog, closeDialog } from "./ui/dialogs";
import { el } from "./ui/widgets";

let checking = false;
let lastCheckedAt = 0;

export async function currentVersion(): Promise<string> {
  if (!isTauri) return "dev";
  const { getVersion } = await import("@tauri-apps/api/app");
  return getVersion();
}

/**
 * Check GitHub for a newer release. `interactive` shows "up to date" / error
 * messages; the silent startup check only speaks up when an update exists.
 */
export async function checkForUpdates(interactive: boolean): Promise<void> {
  if (!isTauri) { if (interactive) await showMessage("Updates are only available in the installed app.", "OfficeMini"); return; }
  if (checking) return;
  if (!interactive && Date.now() - lastCheckedAt < 6 * 3600 * 1000) return;
  checking = true;
  lastCheckedAt = Date.now();
  try {
    const { check } = await import("@tauri-apps/plugin-updater");
    const update = await check({ timeout: 15000 });
    if (!update) { if (interactive) await showMessage(`OfficeMini ${await currentVersion()} is up to date.`, "OfficeMini"); return; }
    offerUpdate(update);
  } catch (e) {
    console.warn("update check failed", e);
    if (interactive) await showMessage("Could not check for updates.\n\n" + String((e as Error).message || e), "OfficeMini", "warning");
  } finally {
    checking = false;
  }
}

function offerUpdate(update: { version: string; currentVersion: string; body?: string | null; date?: string | null; downloadAndInstall(cb?: (ev: any) => void): Promise<void> }) {
  const notes = (update.body || "").trim();
  const progress = el("div", { style: { marginTop: "10px", display: "none" } },
    el("div", { style: { height: "6px", background: "var(--ui-border)", borderRadius: "3px", overflow: "hidden" } }, el("div", { class: "bar", style: { height: "100%", width: "0%", background: "var(--ui-accent)", transition: "width .15s" } })),
    el("div", { class: "label", style: { color: "var(--ui-muted)", marginTop: "4px" } }, "Downloading…"));
  const body = el("div", null,
    el("p", null, `Version ${update.version} is available (you have ${update.currentVersion}).`),
    notes ? el("pre", { style: { whiteSpace: "pre-wrap", font: "12px var(--ui-font)", maxHeight: "200px", overflow: "auto", background: "var(--ui-input)", padding: "8px", borderRadius: "4px" } }, notes.slice(0, 4000)) : null,
    el("p", { style: { color: "var(--ui-muted)" } }, "The app will restart after installing. Unsaved changes are kept until you confirm."),
    progress,
  );
  let installing = false;
  showDialog("Update available", body, [
    { label: "Later" },
    { label: "Update now", primary: true, action: () => {
      if (installing) return false;
      installing = true;
      progress.style.display = "block";
      const bar = progress.querySelector<HTMLElement>(".bar")!, label = progress.querySelector<HTMLElement>(".label")!;
      let total = 0, got = 0;
      (async () => {
        try {
          await update.downloadAndInstall((ev: any) => {
            if (ev.event === "Started") { total = ev.data?.contentLength || 0; }
            else if (ev.event === "Progress") { got += ev.data?.chunkLength || 0; if (total) bar.style.width = Math.min(100, Math.round((got / total) * 100)) + "%"; label.textContent = total ? `Downloading… ${Math.round(got / 1024)} / ${Math.round(total / 1024)} KB` : `Downloading… ${Math.round(got / 1024)} KB`; }
            else if (ev.event === "Finished") { bar.style.width = "100%"; label.textContent = "Installing…"; }
          });
          label.textContent = "Restarting…";
          const { relaunch } = await import("@tauri-apps/plugin-process");
          await relaunch();
        } catch (e) {
          closeDialog();
          await showMessage("The update could not be installed.\n\n" + String((e as Error).message || e), "OfficeMini", "error");
        }
      })();
      return false; // keep the dialog open while downloading
    } },
  ]);
}
