//! Desktop-Hülle von Auswertung Light.
//!
//! Die Fachlogik liegt im TypeScript-Frontend. Rust stellt nur die
//! systemnahen Dienste bereit: SQLite-Datenbank, Datei-Dialoge,
//! Dateizugriff inkl. Beobachtung der Zeitmessungs-Datei und Drucken.

/// Öffnet den Druckdialog des Betriebssystems für das aktuelle Fenster.
/// `window.print()` wird nicht von allen WebViews unterstützt.
#[tauri::command]
fn drucken(window: tauri::WebviewWindow) -> Result<(), String> {
    window.print().map_err(|e| e.to_string())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_sql::Builder::default().build())
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .plugin(tauri_plugin_opener::init())
        .invoke_handler(tauri::generate_handler![drucken])
        .run(tauri::generate_context!())
        .expect("Fehler beim Starten von Auswertung Light");
}
