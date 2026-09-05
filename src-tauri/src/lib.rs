use std::sync::atomic::{AtomicUsize, Ordering};
use tauri::ipc::{InvokeBody, Request, Response};
use tauri::{AppHandle, Manager, WebviewUrl, WebviewWindowBuilder};

static WINDOW_COUNTER: AtomicUsize = AtomicUsize::new(1);

/// Read a file and return its raw bytes (zero JSON overhead).
#[tauri::command]
fn read_file(path: String) -> Result<Response, String> {
    let bytes = std::fs::read(&path).map_err(|e| format!("{}: {}", path, e))?;
    Ok(Response::new(bytes))
}

/// Write raw bytes to a file. The path travels in the `x-path` header (URL-encoded),
/// the body is the file content. Writes to a temp file first, then renames.
#[tauri::command]
fn write_file(request: Request<'_>) -> Result<(), String> {
    let path = request
        .headers()
        .get("x-path")
        .and_then(|v| v.to_str().ok())
        .map(|s| url_decode(s))
        .ok_or("missing path")?;
    let data: &[u8] = match request.body() {
        InvokeBody::Raw(b) => b.as_slice(),
        InvokeBody::Json(_) => return Err("expected raw body".into()),
    };
    let tmp = format!("{}.omtmp", path);
    std::fs::write(&tmp, data).map_err(|e| format!("{}: {}", tmp, e))?;
    if let Err(e) = std::fs::rename(&tmp, &path) {
        // Rename over an open/locked file can fail on Windows; fall back to a direct write.
        let _ = std::fs::remove_file(&tmp);
        std::fs::write(&path, data).map_err(|e2| format!("{}: {} / {}", path, e, e2))?;
    }
    Ok(())
}

#[tauri::command]
fn read_text_file(path: String) -> Result<String, String> {
    std::fs::read_to_string(&path).map_err(|e| format!("{}: {}", path, e))
}

#[tauri::command]
fn file_exists(path: String) -> bool {
    std::path::Path::new(&path).exists()
}

/// Command-line arguments (file paths to open) for the first window.
#[tauri::command]
fn cli_args() -> Vec<String> {
    std::env::args().skip(1).filter(|a| !a.starts_with('-')).collect()
}

/// Open another document in a new window of this process (fast: no new process).
#[tauri::command]
fn open_window(app: AppHandle, path: Option<String>) -> Result<String, String> {
    let n = WINDOW_COUNTER.fetch_add(1, Ordering::SeqCst);
    let label = format!("doc{}", n);
    let mut url = String::from("index.html");
    if let Some(p) = path {
        url.push_str("?file=");
        url.push_str(&url_encode(&p));
    }
    let mut builder = WebviewWindowBuilder::new(&app, &label, WebviewUrl::App(url.into()))
        .title("OfficeMini")
        .inner_size(1100.0, 820.0)
        .min_inner_size(520.0, 360.0)
        .visible(false)
        .zoom_hotkeys_enabled(false);
    // Cascade relative to the focused window.
    if let Some(w) = app.webview_windows().values().find(|w| w.is_focused().unwrap_or(false)) {
        if let Ok(pos) = w.outer_position() {
            builder = builder.position((pos.x + 32) as f64, (pos.y + 32) as f64);
        }
    }
    let win = builder.build().map_err(|e| e.to_string())?;
    harden_webview(&win);
    Ok(label)
}

/// Persist small JSON settings (recent files, window prefs) in the app config dir.
#[tauri::command]
fn load_settings(app: AppHandle) -> Result<String, String> {
    let dir = app.path().app_config_dir().map_err(|e| e.to_string())?;
    let p = dir.join("settings.json");
    Ok(std::fs::read_to_string(p).unwrap_or_else(|_| "{}".into()))
}

#[tauri::command]
fn save_settings(app: AppHandle, json: String) -> Result<(), String> {
    let dir = app.path().app_config_dir().map_err(|e| e.to_string())?;
    std::fs::create_dir_all(&dir).map_err(|e| e.to_string())?;
    std::fs::write(dir.join("settings.json"), json).map_err(|e| e.to_string())
}

#[derive(serde::Serialize)]
struct FileInfo {
    name: String,
    path: String,
    size: u64,
    mtime: u64,
}

fn mtime_secs(md: &std::fs::Metadata) -> u64 {
    md.modified()
        .ok()
        .and_then(|t| t.duration_since(std::time::UNIX_EPOCH).ok())
        .map(|d| d.as_secs())
        .unwrap_or(0)
}

/// Per-user data directory for recovery copies (created on demand).
#[tauri::command]
fn recovery_dir(app: AppHandle) -> Result<String, String> {
    let dir = app.path().app_data_dir().map_err(|e| e.to_string())?.join("recovery");
    std::fs::create_dir_all(&dir).map_err(|e| e.to_string())?;
    Ok(dir.to_string_lossy().into_owned())
}

#[tauri::command]
fn list_files(dir: String) -> Result<Vec<FileInfo>, String> {
    let mut out = Vec::new();
    let rd = match std::fs::read_dir(&dir) {
        Ok(rd) => rd,
        Err(_) => return Ok(out),
    };
    for entry in rd.flatten() {
        if let Ok(md) = entry.metadata() {
            if md.is_file() {
                out.push(FileInfo {
                    name: entry.file_name().to_string_lossy().into_owned(),
                    path: entry.path().to_string_lossy().into_owned(),
                    size: md.len(),
                    mtime: mtime_secs(&md),
                });
            }
        }
    }
    Ok(out)
}

#[tauri::command]
fn delete_file(path: String) -> Result<(), String> {
    match std::fs::remove_file(&path) {
        Ok(()) => Ok(()),
        Err(e) if e.kind() == std::io::ErrorKind::NotFound => Ok(()),
        Err(e) => Err(format!("{}: {}", path, e)),
    }
}

/// Modification time in seconds since the epoch, or null when the file does not exist.
#[tauri::command]
fn file_mtime(path: String) -> Option<u64> {
    std::fs::metadata(&path).ok().map(|md| mtime_secs(&md))
}

/// Open a path with the OS default handler (used for hyperlinks / "show in folder").
#[tauri::command]
fn open_external(url: String) -> Result<(), String> {
    #[cfg(target_os = "windows")]
    {
        std::process::Command::new("cmd")
            .args(["/C", "start", "", &url])
            .spawn()
            .map_err(|e| e.to_string())?;
    }
    #[cfg(target_os = "linux")]
    {
        std::process::Command::new("xdg-open").arg(&url).spawn().map_err(|e| e.to_string())?;
    }
    #[cfg(target_os = "macos")]
    {
        std::process::Command::new("open").arg(&url).spawn().map_err(|e| e.to_string())?;
    }
    Ok(())
}

/// WebView2 ships browser accelerator keys (Ctrl+P print, Ctrl+F find bar, F5 reload,
/// Ctrl+S "save page"...) that would fight the editor's own shortcuts. Turn them off,
/// along with the stock context menu (the app draws its own). Editing keys
/// (Ctrl+C/V/X/Z/Y/A) are unaffected by this setting.
#[cfg(windows)]
fn harden_webview(win: &tauri::WebviewWindow) {
    use webview2_com::Microsoft::Web::WebView2::Win32::ICoreWebView2Settings3;
    use windows::core::Interface;
    let _ = win.with_webview(|webview| unsafe {
        let core = match webview.controller().CoreWebView2() {
            Ok(c) => c,
            Err(_) => return,
        };
        let settings = match core.Settings() {
            Ok(s) => s,
            Err(_) => return,
        };
        let _ = settings.SetAreDefaultContextMenusEnabled(false);
        let _ = settings.SetIsZoomControlEnabled(false);
        if let Ok(s3) = settings.cast::<ICoreWebView2Settings3>() {
            let _ = s3.SetAreBrowserAcceleratorKeysEnabled(false);
        }
    });
}

#[cfg(target_os = "linux")]
fn harden_webview(win: &tauri::WebviewWindow) {
    use gtk::prelude::*;
    // The webview paints every pixel of the window, so GTK's own CSS background fill is wasted
    // work: a full-window software memset on every redraw, and at the 2x buffer scale GTK3 uses
    // on fractional-scale Wayland outputs it dominated the main thread (sampled with gdb).
    // An app-paintable window skips that fill.
    if let Ok(gtk_win) = win.gtk_window() {
        gtk_win.set_app_paintable(true);
    }
}

#[cfg(not(any(windows, target_os = "linux")))]
fn harden_webview(_win: &tauri::WebviewWindow) {}

fn url_encode(s: &str) -> String {
    let mut out = String::new();
    for b in s.bytes() {
        match b {
            b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'-' | b'_' | b'.' | b'~' => out.push(b as char),
            _ => out.push_str(&format!("%{:02X}", b)),
        }
    }
    out
}

fn url_decode(s: &str) -> String {
    let bytes = s.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%' && i + 2 < bytes.len() {
            if let Ok(v) = u8::from_str_radix(&s[i + 1..i + 3], 16) {
                out.push(v);
                i += 3;
                continue;
            }
        }
        if bytes[i] == b'+' {
            out.push(b' ');
        } else {
            out.push(bytes[i]);
        }
        i += 1;
    }
    String::from_utf8_lossy(&out).into_owned()
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_updater::Builder::new().build())
        .plugin(tauri_plugin_process::init())
        .plugin(tauri_plugin_window_state::Builder::default().build())
        .invoke_handler(tauri::generate_handler![
            read_file,
            write_file,
            read_text_file,
            file_exists,
            cli_args,
            open_window,
            load_settings,
            save_settings,
            open_external,
            recovery_dir,
            list_files,
            delete_file,
            file_mtime
        ])
        .setup(|app| {
            // The first window is created from tauri.conf.json (hidden); the frontend
            // shows it once the document is rendered to avoid a white flash.
            if let Some(win) = app.get_webview_window("main") {
                harden_webview(&win);
            }
            Ok(())
        })
        .run(tauri::generate_context!())
        .expect("error while running OfficeMini");
}
