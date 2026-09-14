// Minimal Chrome DevTools Protocol helpers for driving a browser on a throwaway profile.
// Used by scripts/e2e-ext.mjs; nothing here is part of any build.
import { spawn } from "node:child_process";
import fs from "node:fs";

export const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/**
 * Launch a browser on a throwaway profile. `pipe: true` also opens --remote-debugging-pipe, which
 * branded Chrome needs for Extensions.loadUnpacked (it ignores --load-extension since 137); the
 * returned `pipeSend` talks over it while everything else uses the ordinary WebSocket port.
 */
export async function launch({ exe, port, profile, args = [], url = "about:blank", prepare, pipe = false, headless = !process.env.OM_HEADED }) {
  fs.rmSync(profile, { recursive: true, force: true });
  fs.mkdirSync(profile, { recursive: true });
  if (prepare) prepare(profile);
  const proc = spawn(exe, [
    // Headless by default: a test browser must never pop up on the desktop of someone who is
    // using the machine. OM_HEADED=1 shows the window, for watching a run on purpose.
    ...(headless ? ["--headless=new"] : []),
    ...(pipe ? ["--remote-debugging-pipe", "--enable-unsafe-extension-debugging"] : []),
    `--remote-debugging-port=${port}`, `--user-data-dir=${profile}`,
    "--no-first-run", "--no-default-browser-check", "--disable-sync", "--disable-background-networking",
    "--window-size=1280,900", ...args, url,
  ], { stdio: pipe ? ["ignore", "ignore", "ignore", "pipe", "pipe"] : "ignore" });
  let pipeSend = null;
  if (pipe) {
    const out = proc.stdio[3], inp = proc.stdio[4];
    let buf = "", n = 0;
    const pend = new Map();
    inp.on("data", (d) => {
      buf += d.toString();
      let i;
      while ((i = buf.indexOf("\0")) >= 0) {
        const m = JSON.parse(buf.slice(0, i)); buf = buf.slice(i + 1);
        if (m.id && pend.has(m.id)) { const { res, rej } = pend.get(m.id); pend.delete(m.id); m.error ? rej(new Error(m.error.message)) : res(m.result); }
      }
    });
    pipeSend = (method, params = {}) => new Promise((res, rej) => { const k = ++n; pend.set(k, { res, rej }); out.write(JSON.stringify({ id: k, method, params }) + "\0"); });
  }
  for (let i = 0; i < 120; i++) {
    try { const r = await fetch(`http://127.0.0.1:${port}/json/version`); if (r.ok) return { proc, version: await r.json(), pipeSend }; } catch { /* not up yet */ }
    await sleep(150);
  }
  proc.kill();
  throw new Error("browser did not start: " + exe);
}

export async function targets(port) { return (await fetch(`http://127.0.0.1:${port}/json/list`)).json(); }

export function connect(wsUrl) {
  const ws = new WebSocket(wsUrl);
  let id = 0;
  const pending = new Map();
  const listeners = new Set();
  ws.onmessage = (m) => {
    const msg = JSON.parse(m.data);
    if (msg.id && pending.has(msg.id)) {
      const { res, rej } = pending.get(msg.id); pending.delete(msg.id);
      msg.error ? rej(new Error(msg.error.message + (msg.error.data ? " " + msg.error.data : ""))) : res(msg.result);
    } else for (const l of listeners) l(msg);
  };
  const ready = new Promise((res, rej) => { ws.onopen = res; ws.onerror = rej; });
  const send = (method, params = {}, sessionId) => new Promise((res, rej) => {
    const i = ++id; pending.set(i, { res, rej });
    ws.send(JSON.stringify({ id: i, method, params, ...(sessionId ? { sessionId } : {}) }));
  });
  return { ready, send, on: (l) => { listeners.add(l); return () => listeners.delete(l); }, close: () => ws.close() };
}

/** Attach to a target and return a session-bound `send` plus an evaluate helper. */
export async function attach(cdp, targetId) {
  const { sessionId } = await cdp.send("Target.attachToTarget", { targetId, flatten: true });
  const send = (m, p = {}) => cdp.send(m, p, sessionId);
  const evaluate = async (expression) => {
    const r = await send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
    if (r.exceptionDetails) throw new Error("page threw: " + (r.exceptionDetails.exception?.description || r.exceptionDetails.text));
    return r.result.value;
  };
  return { sessionId, send, evaluate };
}
