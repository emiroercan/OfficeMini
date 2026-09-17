// ============================================================================
// macOS ONLY - native menu bar builder.
// ----------------------------------------------------------------------------
// The frontend (src/native-menu.ts) owns the menu definitions and sends them here
// as a JSON model; this file turns that model into the real macOS NSMenu and sets
// it. Keeping it in its own module makes the Mac-specific work easy to find/review.
//
// Model shape (one entry per top-level menu):
//   [{ "title": "File", "items": [ Node, ... ] }, ...]
// Node:
//   { "sep": true }                                   -> separator
//   { "role": "cut|copy|paste|selectAll", "label" }   -> native predefined item (WebKit clipboard)
//   { "label", "submenu": [Node, ...], "enabled" }    -> submenu
//   { "id", "label", "accel"?, "enabled", "checked"? }-> clickable item (checked -> CheckMenuItem)
//
// This module wraps the model with the native App menu (with a routed Quit) and a
// native Window menu. Clicking a clickable item fires on_menu_event in lib.rs, which
// emits "menu-action" with the item id; the frontend runs the matching action.
// ============================================================================

use serde_json::Value;
use tauri::menu::{CheckMenuItem, IsMenuItem, Menu, MenuItem, PredefinedMenuItem, Submenu};
use tauri::{AppHandle, Manager, Runtime};

/// Build one menu node (leaf, separator, predefined role, or nested submenu).
fn build_item<R: Runtime>(app: &AppHandle<R>, node: &Value) -> tauri::Result<Box<dyn IsMenuItem<R>>> {
    if node.get("sep").and_then(Value::as_bool).unwrap_or(false) {
        return Ok(Box::new(PredefinedMenuItem::separator(app)?));
    }
    let label = node.get("label").and_then(Value::as_str).unwrap_or("");
    let enabled = node.get("enabled").and_then(Value::as_bool).unwrap_or(true);

    // Clipboard items must be native predefined items or WebKit will not honour them.
    if let Some(role) = node.get("role").and_then(Value::as_str) {
        return Ok(match role {
            "cut" => Box::new(PredefinedMenuItem::cut(app, Some(label))?),
            "copy" => Box::new(PredefinedMenuItem::copy(app, Some(label))?),
            "paste" => Box::new(PredefinedMenuItem::paste(app, Some(label))?),
            "selectAll" => Box::new(PredefinedMenuItem::select_all(app, Some(label))?),
            _ => Box::new(PredefinedMenuItem::separator(app)?),
        });
    }

    // Submenu (recurse).
    if let Some(children) = node.get("submenu").and_then(Value::as_array) {
        let submenu = Submenu::new(app, label, enabled)?;
        for child in children {
            let item = build_item(app, child)?;
            submenu.append(item.as_ref())?;
        }
        return Ok(Box::new(submenu));
    }

    let id = node.get("id").and_then(Value::as_str).unwrap_or("");
    let accel = node.get("accel").and_then(Value::as_str);
    // A `checked` field makes it a toggle (Bold ✓, Wrap ✓, ...).
    if let Some(checked) = node.get("checked").and_then(Value::as_bool) {
        return Ok(Box::new(CheckMenuItem::with_id(app, id, label, enabled, checked, accel)?));
    }
    Ok(Box::new(MenuItem::with_id(app, id, label, enabled, accel)?))
}

/// Rebuild and set the whole native menu bar: App menu + the model's menus + Window menu.
pub fn set_app_menu<R: Runtime>(app: &AppHandle<R>, model: &Value) -> tauri::Result<()> {
    let menu = Menu::new(app)?;

    // App menu (OfficeMini). Quit is a routed item, not the predefined one, so the frontend's
    // unsaved-changes flow runs before the app quits.
    let app_menu = Submenu::new(app, "OfficeMini", true)?;
    app_menu.append(&PredefinedMenuItem::about(app, Some("About OfficeMini"), None)?)?;
    app_menu.append(&PredefinedMenuItem::separator(app)?)?;
    app_menu.append(&PredefinedMenuItem::hide(app, None)?)?;
    app_menu.append(&PredefinedMenuItem::hide_others(app, None)?)?;
    app_menu.append(&PredefinedMenuItem::show_all(app, None)?)?;
    app_menu.append(&PredefinedMenuItem::separator(app)?)?;
    app_menu.append(&MenuItem::with_id(app, "quit", "Quit OfficeMini", true, Some("Cmd+Q"))?)?;
    menu.append(&app_menu)?;

    // The document menus supplied by the frontend (File, Edit, View, ...).
    if let Some(menus) = model.as_array() {
        for m in menus {
            let title = m.get("title").and_then(Value::as_str).unwrap_or("");
            let submenu = Submenu::new(app, title, true)?;
            if let Some(items) = m.get("items").and_then(Value::as_array) {
                for node in items {
                    let item = build_item(app, node)?;
                    submenu.append(item.as_ref())?;
                }
            }
            menu.append(&submenu)?;
        }
    }

    // Window menu (native).
    let window = Submenu::new(app, "Window", true)?;
    window.append(&PredefinedMenuItem::minimize(app, None)?)?;
    window.append(&PredefinedMenuItem::maximize(app, None)?)?;
    window.append(&PredefinedMenuItem::fullscreen(app, None)?)?;
    window.append(&PredefinedMenuItem::separator(app)?)?;
    window.append(&PredefinedMenuItem::close_window(app, None)?)?;
    menu.append(&window)?;

    app.set_menu(menu)?;
    Ok(())
}
