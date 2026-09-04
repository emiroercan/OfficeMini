// Prevents an additional console window on Windows in release builds.
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

fn main() {
    // On Wayland the GTK3/WebKitGTK stack spends seconds of main-thread CPU compositing the
    // window after launch (flashing, laggy startup on hybrid-GPU laptops, measured ~8.7 s vs
    // ~0.7 s under X11 on Fedora 43). Run through XWayland unless the user chose a backend.
    #[cfg(target_os = "linux")]
    if std::env::var_os("GDK_BACKEND").is_none()
        && std::env::var_os("WAYLAND_DISPLAY").is_some()
        && std::env::var_os("DISPLAY").is_some()
    {
        std::env::set_var("GDK_BACKEND", "x11");
    }
    officemini_lib::run()
}
