// Prevents an additional console window on Windows in release builds.
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

fn main() {
    // Linux: stay on the session's native backend. Wayland gives per-pixel touchpad scrolling
    // and pinch gestures, which XWayland cannot deliver; set GDK_BACKEND=x11 to override.
    officemini_lib::run()
}
