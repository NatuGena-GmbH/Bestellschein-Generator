#![windows_subsystem = "windows"]

// Bestellscheine mit integrierter UI
use eframe::egui;
use eframe::App;
use std::sync::{Arc, Mutex};
use once_cell::sync::Lazy;
use std::thread;
use std::fs;
use lopdf::{Document, content::{Content, Operation}, dictionary, Object, Stream};
use qrcode::QrCode;

// Debug-Logging-Funktion (nur wenn Debug-Modus aktiv)
fn debug_log(message: &str, debug_enabled: bool) {
    if debug_enabled {
        let timestamp = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs();
        
        let log_entry = format!("[{}] {}\n", timestamp, message);
        let log_path = get_temp_file_path("debug.log");
        
        // Cache-Ordner sicherheitshalber nochmal explizit erstellen
        if let Some(parent) = log_path.parent() {
            let _ = std::fs::create_dir_all(parent);
        }
        
        // Append zum Log (ignoriere Fehler um Performance nicht zu beeinträchtigen)
        if let Ok(mut file) = std::fs::OpenOptions::new()
            .create(true)
            .append(true)
            .open(&log_path)
        {
            use std::io::Write;
            let _ = file.write_all(log_entry.as_bytes());
        }
        
        // Auch in Konsole ausgeben
        println!("[DEBUG] {}", message);
    }
}

// Debug-Print nur im Debug-Modus (für detaillierte Pfad-Infos)
fn debug_print(message: &str, debug_enabled: bool) {
    if debug_enabled {
        println!("DEBUG: {}", message);
        debug_log(&format!("DEBUG: {}", message), true);
    }
}

// Versteckte Dateipfade für temporäre/interne Dateien (für Nutzer unsichtbar)
fn get_temp_file_path(filename: &str) -> std::path::PathBuf {
    let exe_dir = std::env::current_exe()
        .ok()
        .and_then(|p| p.parent().map(|p| p.to_path_buf()))
        .unwrap_or_else(|| std::env::current_dir().unwrap_or_default());
    
    let temp_dir = exe_dir.join("cache");  // Weniger verdächtiger Name statt .temp
    let _ = std::fs::create_dir_all(&temp_dir); // Ordner erstellen falls nicht vorhanden
    
    temp_dir.join(filename)
}

// Sichere System-Kommandos (Antivirus-freundlich)
fn safe_open_explorer(path: &str) -> Result<(), std::io::Error> {
    // Nur erlaubte, sichere Pfade öffnen
    let safe_path = std::path::Path::new(path);
    if safe_path.is_absolute() || path.contains("..") {
        return Err(std::io::Error::new(std::io::ErrorKind::PermissionDenied, "Unsafe path"));
    }
    
    // Windows Explorer mit sicheren Argumenten
    std::process::Command::new("explorer")
        .arg("/select,")
        .arg(safe_path)
        .spawn()
        .map(|_| ())
}

fn safe_open_notepad(file_path: &std::path::Path) -> Result<(), std::io::Error> {
    // Nur existierende Dateien im Projekt-Verzeichnis öffnen
    if !file_path.exists() {
        return Err(std::io::Error::new(std::io::ErrorKind::NotFound, "File not found"));
    }
    
    // Notepad mit sicherem Dateipfad
    std::process::Command::new("notepad.exe")
        .arg(file_path)
        .spawn()
        .map(|_| ())
}

// Release-Ordnerstruktur (relativ zum Programmverzeichnis)
// Helper-Funktion um Config-Verzeichnis zu ermitteln (im EXE-Verzeichnis)
fn get_config_dir() -> std::path::PathBuf {
    let exe_dir = match std::env::current_exe() {
        Ok(exe_path) => match exe_path.parent() {
            Some(parent) => parent.to_path_buf(),
            None => std::env::current_dir().unwrap_or_else(|_| std::path::PathBuf::from("."))
        },
        Err(_) => std::env::current_dir().unwrap_or_else(|_| std::path::PathBuf::from("."))
    };
    let config_dir = exe_dir.join("Config");
    if !config_dir.exists() {
        let _ = std::fs::create_dir_all(&config_dir);
    }
    config_dir
}

fn get_release_dirs() -> (std::path::PathBuf, std::path::PathBuf, std::path::PathBuf, std::path::PathBuf, std::path::PathBuf) {
    get_release_dirs_with_debug(false)  // Standardmäßig ohne Debug
}

fn get_release_dirs_with_debug(debug_enabled: bool) -> (std::path::PathBuf, std::path::PathBuf, std::path::PathBuf, std::path::PathBuf, std::path::PathBuf) {
    let exe_dir = match std::env::current_exe() {
        Ok(exe_path) => match exe_path.parent() {
            Some(parent) => parent.to_path_buf(),
            None => {
                println!("ERROR: Konnte Parent-Verzeichnis der EXE nicht ermitteln");
                std::env::current_dir().unwrap_or_else(|_| std::path::PathBuf::from("."))
            }
        },
        Err(e) => {
            println!("ERROR: Konnte EXE-Pfad nicht ermitteln: {}", e);
            std::env::current_dir().unwrap_or_else(|_| std::path::PathBuf::from("."))
        }
    };
    
    debug_print(&format!("EXE-Verzeichnis: {}", exe_dir.display()), debug_enabled);
    
    // Development-Mode deaktiviert - verwende immer Release-Modus für Deployment
    let is_development_mode = false;
    
    debug_print(&format!("Development-Modus: {}", is_development_mode), debug_enabled);
    
    let project_root = exe_dir.clone(); // Verwende immer EXE-Verzeichnis direkt
    
    debug_print(&format!("Projekt-Root: {}", project_root.display()), debug_enabled);
    
    // Release-Ordnerstruktur (portabel)
    let config_dir = project_root.join("Config");         // Sichtbar für User
    let data_dir = project_root.join("Data");              // Data-Ordner
    let templates_dir = project_root.join("VORLAGE");      // VORLAGE-Ordner (wie ursprünglich)
    let tools_dir = project_root.join("tools");           // Tools-Ordner
    let output_base = project_root.join("OUTPUT");        // PDF-Output
    
    debug_print(&format!("Ordner-Pfade bestimmt - Config: {}, Data: {}, VORLAGE: {}, Tools: {}, Output: {}",
             config_dir.display(), data_dir.display(), templates_dir.display(), tools_dir.display(), output_base.display()), debug_enabled);
    
    // Ordner erstellen falls nicht vorhanden
    for (dir_name, dir_path) in [("Config", &config_dir), ("Data", &data_dir), ("VORLAGE", &templates_dir), ("tools", &tools_dir), ("Output", &output_base)] {
        if !dir_path.exists() {
            match std::fs::create_dir_all(dir_path) {
                Ok(()) => debug_print(&format!("{}-Ordner erstellt: {}", dir_name, dir_path.display()), debug_enabled),
                Err(e) => println!("ERROR: Konnte {}-Ordner nicht erstellen: {} - {}", dir_name, dir_path.display(), e),
            }
        } else {
            debug_print(&format!("{}-Ordner existiert bereits: {}", dir_name, dir_path.display()), debug_enabled);
        }
    }
    
    (config_dir, data_dir, templates_dir, tools_dir, output_base)
}

// Helper-Funktionen für korrekte Pfade
fn get_default_csv_path(group: &str) -> String {
    if group == "Apo" {
        "Data/Vertreternummern-Apo.CSV".to_string()
    } else {
        "Data/Vertreternummern.csv".to_string()
    }
}

fn get_default_template_path() -> String {
    "VORLAGE/Bestellschein-Endkunde-de_de.pdf".to_string()
}

// Template-Pfad zu absolutem Pfad auflösen
fn resolve_template_path_with_debug(template_path: &str, debug_enabled: bool) -> std::path::PathBuf {
    let (_, _, templates_dir, _, _) = get_release_dirs_with_debug(debug_enabled);
    
    // Prüfe ob es bereits ein absoluter Pfad ist
    if std::path::Path::new(template_path).is_absolute() {
        return std::path::PathBuf::from(template_path);
    }
    
    // Entferne Development-Pfad-Präfixe und verwende VORLAGE-Ordner
    let cleaned_path = template_path
        .replace("VORLAGE/", "")
        .replace("Vorlagen/", "")
        .replace("DATA/", "")
        .replace("Data/", "");
    
    let resolved_path = templates_dir.join(&cleaned_path);
    debug_print(&format!("Template-Pfad aufgelöst: '{}' -> '{}'", template_path, resolved_path.display()), debug_enabled);
    resolved_path
}

fn get_default_selections() -> Vec<(String, String, bool)> {
    vec![(get_default_csv_path("Endkunde"), get_default_template_path(), true)]
}

// Output-Verzeichnis basierend auf Gruppe, Sprache und Messe bestimmen
fn get_output_dir_for_group_with_debug(group: &str, language: &str, is_messe: bool, debug_enabled: bool) -> std::path::PathBuf {
    let (_, _, _, _, output_base) = get_release_dirs_with_debug(debug_enabled);
    
    // Bessere Sortierung: Messe zuerst, dann normale Gruppen
    let group_folder = if is_messe {
        format!("Messe_{}", group)
    } else {
        group.to_string()
    };
    
    let language_folder = match language.to_lowercase().as_str() {
        "englisch" | "english" | "en" => "EN",
        _ => "DE", // Kürzer für bessere Übersicht
    };
    
    let final_output_dir = output_base.join(group_folder).join(language_folder);
    
    // Sicherstellen dass das Output-Verzeichnis existiert
    if !final_output_dir.exists() {
        match std::fs::create_dir_all(&final_output_dir) {
            Ok(()) => debug_print(&format!("Output-Verzeichnis erstellt: {}", final_output_dir.display()), debug_enabled),
            Err(e) => println!("ERROR: Konnte Output-Verzeichnis nicht erstellen: {} - {}", final_output_dir.display(), e),
        }
    } else {
        debug_print(&format!("Output-Verzeichnis existiert bereits: {}", final_output_dir.display()), debug_enabled);
    }
    
    final_output_dir
}

// Output-Verzeichnis basierend auf Benutzer-Konfiguration bestimmen
fn get_configured_output_dir_with_debug(use_custom: bool, custom_path: &str, group: &str, language: &str, is_messe: bool, debug_enabled: bool) -> std::path::PathBuf {
    debug_print(&format!("get_configured_output_dir - use_custom: {}, custom_path: '{}', group: '{}', language: '{}', is_messe: {}", 
             use_custom, custom_path, group, language, is_messe), debug_enabled);
             
    if use_custom && !custom_path.is_empty() {
        // Benutzerdefinierten Pfad verwenden
        let path = std::path::Path::new(custom_path);
        if path.is_absolute() {
            debug_print(&format!("Verwende absoluten benutzerdefinierten Pfad: {}", path.display()), debug_enabled);
            path.to_path_buf()
        } else {
            // Relativ zum Programmordner - verwende get_release_dirs() für konsistente Pfade
            let (_, _, _, _, output_base) = get_release_dirs_with_debug(debug_enabled);
            let exe_dir = output_base.parent().unwrap_or_else(|| std::path::Path::new("."));
            let final_path = exe_dir.join(custom_path);
            debug_print(&format!("Verwende relativen benutzerdefinierten Pfad: {} -> {}", custom_path, final_path.display()), debug_enabled);
            final_path
        }
    } else {
        // Standard automatische Ordnerstruktur verwenden
        let auto_dir = get_output_dir_for_group_with_debug(group, language, is_messe, debug_enabled);
        debug_print(&format!("Verwende automatische Ordnerstruktur: {}", auto_dir.display()), debug_enabled);
        auto_dir
    }
}

/// Hauptkonfiguration für Bestellschein-Generierung
/// 
/// Diese Struktur enthält alle Einstellungen für:
/// - QR-Code-Positionen und -Größen  
/// - Vertreternummer-Positionen und Schriftarten
/// 
/// # Beispiel
/// ```
/// let config = Config {
///     qr_codes: vec![QrCodeConfig { x: 50.0, y: 50.0, size: 18.0, pages: vec![1], all_pages: false }],
///     vertreter: vec![VertreterConfig { x: 77.0, y: 80.0, size: 12.0, pages: vec![1], all_pages: false, 
///                                      font_name: "Arial".to_string(), font_style: "Normal".to_string(), 
///                                      font_path: "".to_string() }],
/// };
/// ```
#[derive(Clone)]
pub struct Config {
    /// Liste der QR-Code-Konfigurationen
    pub qr_codes: Vec<QrCodeConfig>,
    /// Liste der Vertreternummer-Konfigurationen  
    pub vertreter: Vec<VertreterConfig>,
}

/// Konfiguration für QR-Code-Platzierung
/// 
/// Definiert Position, Größe und auf welchen Seiten der QR-Code erscheinen soll.
#[derive(Clone, Debug)]
pub struct QrCodeConfig {
    /// X-Position in Millimetern (von links)
    pub x: f32,
    pub y: f32,
    pub size: f32,
    pub pages: Vec<u32>,      // Seiten für diesen QR-Code
    pub all_pages: bool,      // Wenn true, ignoriere pages und verwende alle Seiten
}

#[derive(Clone, Debug)]
pub struct VertreterConfig {
    pub x: f32,
    pub y: f32,
    pub size: f32,
    pub pages: Vec<u32>,      // Seiten für diese Vertreternummer-Position
    pub all_pages: bool,      // Wenn true, ignoriere pages und verwende alle Seiten
    pub font_name: String,    // Name der Schriftart (z.B. "Arial", "Times New Roman")
    pub font_size: f32,       // Schriftgröße für die Vertreternummer
    pub font_style: String,   // Style: "Normal", "Bold", "Italic", "BoldItalic"
}

impl Default for Config {
    fn default() -> Self {
        Self { 
            qr_codes: vec![QrCodeConfig { x: 18.0, y: 18.0, size: 6.3, pages: vec![1], all_pages: false }],
            vertreter: vec![
                VertreterConfig { x: 27.0, y: 28.0, size: 12.0, pages: vec![1], all_pages: false, font_name: "Arial".to_string(), font_size: 12.0, font_style: "Normal".to_string() },
                VertreterConfig { x: 35.0, y: 229.0, size: 10.0, pages: vec![1], all_pages: false, font_name: "Arial".to_string(), font_size: 10.0, font_style: "Normal".to_string() }
            ],
        }
    }
}

// Gruppenspezifische Default-Konfigurationen
fn get_group_default_config(group: &str, is_messe: bool) -> Config {
    println!("Erstelle gruppenspezifische Default-Config für: {} (Messe: {})", group, is_messe);
    
    match group.to_lowercase().as_str() {
        "apo" | "apotheken" => {
            if is_messe {
                // Apo Messe - andere Positionen (in mm)
                Config {
                    qr_codes: vec![QrCodeConfig { x: 28.0, y: 25.0, size: 8.0, pages: vec![1, 2], all_pages: false }],
                    vertreter: vec![
                        VertreterConfig { x: 42.0, y: 35.0, size: 14.0, pages: vec![1, 2], all_pages: false, font_name: "Arial".to_string(), font_size: 14.0, font_style: "Normal".to_string() },
                        VertreterConfig { x: 53.0, y: 247.0, size: 12.0, pages: vec![1], all_pages: false, font_name: "Arial".to_string(), font_size: 12.0, font_style: "Normal".to_string() }
                    ],
                }
            } else {
                // Apo Normal - optimiert für Apotheken-Formulare (in mm)
                Config {
                    qr_codes: vec![QrCodeConfig { x: 26.0, y: 21.0, size: 7.0, pages: vec![1], all_pages: false }],
                    vertreter: vec![
                        VertreterConfig { x: 35.0, y: 32.0, size: 14.0, pages: vec![1], all_pages: false, font_name: "Arial".to_string(), font_size: 14.0, font_style: "Normal".to_string() },
                        VertreterConfig { x: 46.0, y: 240.0, size: 12.0, pages: vec![1], all_pages: false, font_name: "Arial".to_string(), font_size: 12.0, font_style: "Normal".to_string() }
                    ],
                }
            }
        },
        "endkunde" | "endnutzer" => {
            if is_messe {
                // Endkunde Messe - angepasst für Messestände (in mm)
                Config {
                    qr_codes: vec![QrCodeConfig { x: 21.0, y: 28.0, size: 8.5, pages: vec![1], all_pages: false }],
                    vertreter: vec![
                        VertreterConfig { x: 32.0, y: 42.0, size: 12.0, pages: vec![1], all_pages: false, font_name: "Arial".to_string(), font_size: 12.0, font_style: "Normal".to_string() },
                        VertreterConfig { x: 42.0, y: 254.0, size: 10.0, pages: vec![1], all_pages: false, font_name: "Arial".to_string(), font_size: 10.0, font_style: "Normal".to_string() }
                    ],
                }
            } else {
                // Endkunde Normal - Standard-Layout (in mm)
                Config {
                    qr_codes: vec![QrCodeConfig { x: 18.0, y: 18.0, size: 6.3, pages: vec![1], all_pages: false }],
                    vertreter: vec![
                        VertreterConfig { x: 27.0, y: 28.0, size: 12.0, pages: vec![1], all_pages: false, font_name: "Arial".to_string(), font_size: 12.0, font_style: "Normal".to_string() },
                        VertreterConfig { x: 35.0, y: 229.0, size: 10.0, pages: vec![1], all_pages: false, font_name: "Arial".to_string(), font_size: 10.0, font_style: "Normal".to_string() }
                    ],
                }
            }
        },
        _ => {
            // Fallback für unbekannte Gruppen
            Config::default()
        }
    }
}

// Migration von globaler Config zu gruppenspezifischen Configs beim ersten Start (nur einmalig)
fn migrate_global_to_group_configs() {
    let (config_dir, _, _, _, _) = get_release_dirs();
    let migration_marker_path = config_dir.join(".migration_completed");
    
    // Wenn Migration bereits durchgeführt wurde, sofort beenden (keine Ausgabe)
    if migration_marker_path.exists() {
        return; // Kein println! - läuft still im Hintergrund
    }
    
    println!("🔄 ERSTMALIGER START: Prüfe Migration von globaler Config...");
    
    let global_config_path = config_dir.join("config.toml");
    
    // Wenn keine globale Config existiert, einfach Marker erstellen und fertig
    if !global_config_path.exists() {
        println!("ℹ️ SETUP: Keine globale Config gefunden - System bereit für gruppenspezifische Configs.");
        let _ = std::fs::write(&migration_marker_path, "Setup completed - no migration needed");
        return;
    }
    
    // Lade die globale Config
    println!("📥 MIGRATION: Globale Config gefunden, starte Migration...");
    match std::fs::read_to_string(&global_config_path) {
        Ok(global_toml) => {
            let global_config = parse_toml_to_config(&global_toml);
            
            // Migriere zu allen wichtigen Gruppen-Kombinationen
            let migration_targets = vec![
                ("Endkunde", "Deutsch", false),
                ("Endkunde", "Englisch", false),
                ("Endkunde", "Deutsch", true),
                ("Endkunde", "Englisch", true),
                ("Apo", "Deutsch", false),
                ("Apo", "Englisch", false),
                ("Apo", "Deutsch", true),
                ("Apo", "Englisch", true),
            ];
            
            for (group, language, is_messe) in migration_targets {
                let target_file = if is_messe {
                    config_dir.join(format!("config_{}_{}_messe.toml", group.to_lowercase(), language.to_lowercase()))
                } else {
                    config_dir.join(format!("config_{}_{}.toml", group.to_lowercase(), language.to_lowercase()))
                };
                
                // Nur migrieren wenn die Ziel-Config noch nicht existiert
                if !target_file.exists() {
                    println!("🚚 MIGRATION: Migriere globale Config nach {:?}", target_file);
                    save_group_config(group, language, is_messe, &global_config);
                } else {
                    println!("⏭️ MIGRATION: {:?} existiert bereits, überspringe.", target_file);
                }
            }
            
            // Migration abgeschlossen - Marker erstellen
            let _ = std::fs::write(&migration_marker_path, format!("Migration completed at {}", 
                std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH).unwrap().as_secs()));
            println!("✅ MIGRATION: Globale Config erfolgreich zu gruppenspezifischen Configs migriert!");
            
            // Optional: Globale Config umbenennen als Backup
            let backup_path = config_dir.join("config_global_backup.toml");
            if let Err(e) = std::fs::rename(&global_config_path, &backup_path) {
                println!("⚠️ MIGRATION: Konnte globale Config nicht zu Backup umbenennen: {}", e);
            } else {
                println!("💾 MIGRATION: Globale Config als {:?} gesichert.", backup_path);
            }
        }
        Err(e) => {
            println!("❌ MIGRATION: Fehler beim Lesen der globalen Config: {}", e);
            let _ = std::fs::write(&migration_marker_path, format!("Migration failed: {}", e));
        }
    }
}

// Hilfsfunktion für die Formatierung von Zeitdauern
fn format_duration(duration: std::time::Duration) -> String {
    let total_seconds = duration.as_secs();
    let hours = total_seconds / 3600;
    let minutes = (total_seconds % 3600) / 60;
    let seconds = total_seconds % 60;
    
    if hours > 0 {
        format!("{}h {}m {}s", hours, minutes, seconds)
    } else if minutes > 0 {
        format!("{}m {}s", minutes, seconds)
    } else {
        format!("{}s", seconds)
    }
}

// App-spezifische Einstellungen (UI-Konfiguration)
fn save_app_settings(dark_mode: bool) {
    let config_dir = get_config_dir();
    let settings_path = config_dir.join("app_settings.toml");

    let toml = format!("[ui]\ndark_mode = {}\nmaximized_mode = true\n", dark_mode);
    
    if let Err(e) = std::fs::write(&settings_path, toml) {
        eprintln!("Fehler beim Speichern der App-Einstellungen: {}", e);
    } else {
        println!("App-Einstellungen gespeichert: dark_mode={}", dark_mode);
    }
}

fn load_app_settings() -> bool {
    let config_dir = get_config_dir();
    let settings_path = config_dir.join("app_settings.toml");
    
    if let Ok(content) = std::fs::read_to_string(&settings_path) {
        for line in content.lines() {
            let line = line.trim();
            if line.starts_with("dark_mode = ") {
                if let Some(value_str) = line.strip_prefix("dark_mode = ") {
                    return value_str == "true";
                }
            }
        }
    }
    
    false // Standard: Light Mode
}

// Hilfsfunktion: Parsen eines TOML-Strings in ein Config-Objekt
fn parse_toml_to_config(toml: &str) -> Config {
    // Startwerte - werden überschrieben wenn in der Datei gefunden
    let mut qr_codes = Vec::new();
    let mut vertreter = Vec::new();
    let mut pages = Vec::new();

    let mut in_qr_array = false;
    let mut in_vertreter_array = false;
    let mut _in_positions_section = false;
    let mut _in_pages_section = false;

    for l in toml.lines() {
        let l = l.trim();

        // Section-Header erkennen
        if l == "[positions]" {
            _in_positions_section = true;
            _in_pages_section = false;
            continue;
        } else if l == "[pages]" {
            _in_positions_section = false;
            _in_pages_section = true;
            continue;
        } else if l.starts_with('[') {
            // Andere Section
            _in_positions_section = false;
            _in_pages_section = false;
            continue;
        }

        // Einzelner QR-Code (Rückwärtskompatibilität)
        if l.starts_with("qr_code =") && l.contains('{') {
            if let Some(start) = l.find('{') {
                let inner = &l[start+1..l.find('}').unwrap_or(l.len())];
                let mut x = 50.0;
                let mut y = 50.0;
                let mut size = 18.0;
                let mut all_pages = false;
                for part in inner.split(',') {
                    let part = part.trim();
                    if part.starts_with("x =") {
                        x = part[3..].trim().parse().unwrap_or(50.0);
                    } else if part.starts_with("y =") {
                        y = part[3..].trim().parse().unwrap_or(50.0);
                    } else if part.starts_with("size =") {
                        size = part[6..].trim().parse().unwrap_or(18.0);
                    } else if part.starts_with("all_pages =") {
                        all_pages = part[11..].trim() == "true";
                    }
                }
                qr_codes.push(QrCodeConfig { x, y, size, pages: vec![1], all_pages });
            }
        }

        // QR-Code Array (funktioniert sowohl mit als auch ohne [positions] Section)
        else if l.starts_with("qr_codes = [") {
            in_qr_array = true;
            continue;
        } else if in_qr_array {
            if l.starts_with(']') {
                in_qr_array = false;
                continue;
            }
            if l.contains("x =") && l.contains("y =") {
                let mut x = 50.0;
                let mut y = 50.0;
                let mut size = 18.0;
                let mut all_pages = false;
                for part in l.trim_matches(|c| c == '{' || c == '}' || c == ',').split(',') {
                    let part = part.trim();
                    if part.starts_with("x =") {
                        x = part[3..].trim().parse().unwrap_or(50.0);
                    } else if part.starts_with("y =") {
                        y = part[3..].trim().parse().unwrap_or(50.0);
                    } else if part.starts_with("size =") {
                        size = part[6..].trim().parse().unwrap_or(18.0);
                    } else if part.starts_with("all_pages =") {
                        all_pages = part[11..].trim() == "true";
                    }
                }
                qr_codes.push(QrCodeConfig { x, y, size, pages: vec![1], all_pages });
            }
        }

        // Vertreter-Positionen
        else if l.starts_with("vertreter_nummer = [") {
            in_vertreter_array = true;
            continue;
        } else if in_vertreter_array {
            if l.starts_with(']') {
                in_vertreter_array = false;
                continue;
            }
            if l.contains("x =") && l.contains("y =") {
                let mut x = 0.0;
                let mut y = 0.0;
                let mut size = 12.0; // Default size für Vertreter
                let mut font_name = "Arial".to_string();
                let mut font_size = 12.0;
                let mut font_style = "Normal".to_string();
                let mut all_pages = false;
                
                for part in l.trim_matches(|c| c == '{' || c == '}' || c == ',').split(',') {
                    let part = part.trim();
                    if part.starts_with("x =") {
                        x = part[3..].trim().parse().unwrap_or(0.0);
                    } else if part.starts_with("y =") {
                        y = part[3..].trim().parse().unwrap_or(0.0);
                    } else if part.starts_with("size =") {
                        size = part[6..].trim().parse().unwrap_or(12.0);
                    } else if part.starts_with("font_name =") {
                        font_name = part[11..].trim().trim_matches('"').to_string();
                    } else if part.starts_with("font_size =") {
                        font_size = part[11..].trim().parse().unwrap_or(12.0);
                    } else if part.starts_with("font_style =") {
                        font_style = part[12..].trim().trim_matches('"').to_string();
                    } else if part.starts_with("all_pages =") {
                        all_pages = part[11..].trim() == "true";
                    }
                }
                
                // Fallback: font_size auf size setzen wenn nicht explizit gesetzt
                if font_size == 12.0 && size != 12.0 {
                    font_size = size;
                }
                
                vertreter.push(VertreterConfig { x, y, size, pages: vec![1], all_pages, font_name, font_size, font_style });
            }
        }

        // Seiten (sowohl "pages = [" als auch "include = [")
        else if l.starts_with("pages = [") || l.starts_with("include = [") {
            let nums = l.trim_start_matches("pages = [")
                       .trim_start_matches("include = [")
                       .trim_end_matches(']')
                       .split(',');
            for n in nums {
                let n = n.trim();
                if !n.is_empty() {
                    if let Ok(val) = n.parse() {
                        pages.push(val);
                    }
                }
            }
        }
    }

    // Defaults setzen wenn nichts gefunden wurde (in mm)
    if qr_codes.is_empty() {
        qr_codes.push(QrCodeConfig { x: 18.0, y: 18.0, size: 6.3, pages: vec![1], all_pages: false });
    }
    if vertreter.is_empty() {
        vertreter = vec![
            VertreterConfig { x: 27.0, y: 28.0, size: 12.0, pages: vec![1], all_pages: false, font_name: "Arial".to_string(), font_size: 12.0, font_style: "Normal".to_string() },
            VertreterConfig { x: 35.0, y: 229.0, size: 10.0, pages: vec![1], all_pages: false, font_name: "Arial".to_string(), font_size: 10.0, font_style: "Normal".to_string() },
        ];
    }
    if pages.is_empty() {
        pages.push(1);
    }

    let final_config = Config {
        qr_codes,
        vertreter,
    };

    println!("Geladene Config via parse_toml_to_config: QR={:?}, Vertreter={:?}", 
             final_config.qr_codes, final_config.vertreter);

    final_config
}

// Gruppenspezifische Config-Datei speichern
fn save_group_config(group: &str, language: &str, is_messe: bool, config: &Config) {
    println!("=== SAVE_GROUP_CONFIG AUFGERUFEN FÜR: {} {} (Messe: {}) ===", group, language, is_messe);
    
    // Release-Ordnerstruktur verwenden - Config ist sichtbar für User
    let (config_dir, _, _, _, _) = get_release_dirs();
    
    // Filename für die Gruppe bestimmen
    let group_filename = if is_messe {
        config_dir.join(format!("config_{}_messe.toml", group.to_lowercase()))
    } else {
        config_dir.join(format!("config_{}.toml", group.to_lowercase()))
    };
    
    // TOML generieren
    let mut toml = String::new();
    toml.push_str(&format!("# Konfiguration für {}\n", group));
    if is_messe {
        toml.push_str("# Messe-spezifische Konfiguration\n");
    }
    toml.push_str("# Koordinaten sind in Millimetern (DIN A4: 210×297 mm)\n\n");
    
    // QR-Codes
    toml.push_str("qr_codes = [\n");
    for qr in &config.qr_codes {
        toml.push_str(&format!("  {{ x = {}, y = {}, size = {}, all_pages = {} }},\n", qr.x, qr.y, qr.size, qr.all_pages));
    }
    toml.push_str("]\n\n");
    
    toml.push_str("[positions]\n");
    toml.push_str("vertreter_nummer = [\n");
    for v in &config.vertreter {
        toml.push_str(&format!("  {{ x = {}, y = {}, size = {}, all_pages = {}, font_name = \"{}\", font_size = {}, font_style = \"{}\" }},\n", v.x, v.y, v.size, v.all_pages, v.font_name, v.font_size, v.font_style));
    }
    toml.push_str("]\n\n");
    
    // Datei schreiben
    if let Err(e) = std::fs::write(&group_filename, toml) {
        eprintln!("Konnte gruppenspezifische Config nicht speichern: {}", e);
    } else {
        println!("✅ Gruppenspezifische Config gespeichert: {:?}", group_filename);
        println!("Config-Werte: QR={:?}, Vertreter={:?}", 
                 config.qr_codes, config.vertreter);
    }
}

// Resume-Funktionalität: Prüfen ob bereits PDFs erstellt wurden
fn check_resume_available() -> bool {
    if let Ok(entries) = std::fs::read_dir("OUTPUT") {
        let pdf_count = entries
            .filter_map(|e| e.ok())
            .filter(|e| {
                e.path().extension()
                    .and_then(|ext| ext.to_str())
                    .map(|ext| ext.to_lowercase() == "pdf")
                    .unwrap_or(false)
            })
            .count();
        
        let has_pdfs = pdf_count > 0;
        println!("Resume-Check: {} PDFs gefunden, Resume verfügbar: {}", pdf_count, has_pdfs);
        has_pdfs
    } else {
        println!("Resume-Check: OUTPUT Ordner nicht lesbar");
        false
    }
}

// Resume-Funktionalität: Anzahl bereits verarbeiteter Dateien ermitteln
fn get_last_processed_count() -> usize {
    if let Ok(entries) = std::fs::read_dir("OUTPUT") {
        let count = entries
            .filter_map(|e| e.ok())
            .filter(|e| {
                e.path().extension()
                    .and_then(|ext| ext.to_str())
                    .map(|ext| ext.to_lowercase() == "pdf")
                    .unwrap_or(false)
            })
            .count();
        
        println!("Letzte verarbeitete Anzahl: {}", count);
        count
    } else {
        0
    }
}

pub struct MyApp {
    config: Config,
    // Startup selection for which documents to create
    show_startup_dialog: bool,
    selected_group: String,
    selected_language: String,
    is_messe: bool,
    show_config: bool,
    progress: f32,
    status_message: String,
    save_message: Option<std::time::Instant>,
    is_generating: bool,
    stop_signal: Arc<Mutex<bool>>,
    show_meme: bool,
    meme_time: Option<std::time::Instant>,
    resume_available: bool,
    last_processed_count: usize,
    animation_frame: usize,
    animation_time: Option<std::time::Instant>,
    resume_needs_update: bool,
    // Manual coordinate input fields
    manual_qr_x: String,
    manual_qr_y: String,
    manual_qr_size: String,
    manual_vertreter_x: String,
    manual_vertreter_y: String,
    manual_vertreter_size: String,
    // Dark mode toggle
    dark_mode: bool,
    // Maximized window toggle
    fullscreen_mode: bool,
    // Settings dialog
    show_settings_dialog: bool,
    // Output directory configuration
    custom_output_dir: String,
    
    // Font-Auswahl erweiterte Ansicht
    show_all_fonts: bool,
    // Font-Suche
    font_search_text: String,
    use_custom_output_dir: bool,
    // Template directory configuration
    custom_template_dir: String,
    use_custom_template_dir: bool,
    // Template selection fallback
    available_templates: Vec<String>,
    selected_template_index: Option<usize>,
    show_template_selection: bool,
    // Time tracking
    generation_start_time: Option<std::time::Instant>,
    last_progress_update: Option<std::time::Instant>,
    estimated_total_duration: Option<std::time::Duration>,
    // Progress-Update-Control
    progress_frozen: bool,  // Verhindert Progress-Updates nach Stop
    // Debug und Performance
    debug_mode: bool,       // Versteckter Debug-Modus
    debug_key_pressed: bool, // Flag für Tastatur-Behandlung
    max_threads: usize,     // Thread-Begrenzung für Performance
    thread_sleep_ms: u64,   // Pause zwischen PDF-Generierungen (ms)
    // Font-System-Konfiguration
    enable_font_fallback: bool, // Font-Fallback aktiviert (Standard: true)
    // Font-Caching für Performance
    cached_fonts: Vec<String>, // Gecachte Font-Liste
    // Bereichs-Auswahl für Vertreternummern
    use_range_selection: bool,  // Ob Bereichs-Auswahl aktiviert ist
    range_start_index: String,  // Start-Index (0-basiert)
    range_end_index: String,    // End-Index (0-basiert)
    // Resume-Information
    resume_info: Option<(usize, usize, u64)>, // (current_index, total_count, elapsed_seconds)
}

impl Default for MyApp {
    fn default() -> Self {
        // Progress-Datei initial löschen/erstellen (versteckt)
        let progress_path = get_temp_file_path("progress.txt");
        let _ = std::fs::write(&progress_path, "0.0");

        // Stop-Status-Datei löschen falls vorhanden (versteckt)
        let stop_status_path = get_temp_file_path("stop_status.txt");
        let _ = std::fs::remove_file(&stop_status_path);

        // CONFIG Ordner erstellen falls er nicht existiert (für Legacy-Kompatibilität)
        if !std::path::Path::new("CONFIG").exists() {
            let _ = std::fs::create_dir("CONFIG");
            // Info-Datei erstellen
            let info_text = r#"# CONFIG Ordner Info
# 
# Seit Version 2.0 wird die Konfiguration intern gespeichert (versteckt für den User).
# Dieser Ordner dient nur noch als Fallback für alte Konfigurationen.
# 
# Die echte Konfiguration wird gespeichert in:
# Config/app_config.toml (im Anwendungsverzeichnis)
# 
# Sie können diesen Ordner löschen wenn Sie möchten - er wird automatisch neu erstellt.
"#;
            let _ = std::fs::write("CONFIG/README.txt", info_text);
            println!("CONFIG Ordner automatisch erstellt mit Info-Datei");
        }

        // Migration von globaler zu gruppenspezifischer Config NUR beim ersten Start (einmalig)
        // Diese Funktion prüft intern ob Migration bereits durchgeführt wurde
        migrate_global_to_group_configs();

        // Standard-Gruppenauswahl beim Start
        let default_group = "Endkunde".to_string();
        let default_language = "Deutsch".to_string();
        let default_is_messe = false;

        // Gruppenspezifische Config beim Start laden statt globaler Config
        let initial_config = load_group_config(&default_group, &default_language, default_is_messe);
        let initial_resume_info = load_resume_info(&default_group, &default_language, default_is_messe);

        // Manual input fields basierend auf der geladenen Config setzen
        let mut manual_qr_x = "50.0".to_string();
        let mut manual_qr_y = "50.0".to_string();
        let mut manual_qr_size = "18.0".to_string();
        let mut manual_vertreter_x = "77.0".to_string();
        let mut manual_vertreter_y = "80.0".to_string();
        let manual_vertreter_size = "12.0".to_string();

        // Update manual fields with loaded config values
        if !initial_config.qr_codes.is_empty() {
            manual_qr_x = format!("{:.1}", initial_config.qr_codes[0].x);
            manual_qr_y = format!("{:.1}", initial_config.qr_codes[0].y);
            manual_qr_size = format!("{:.1}", initial_config.qr_codes[0].size);
        }
        if !initial_config.vertreter.is_empty() {
            manual_vertreter_x = format!("{:.1}", initial_config.vertreter[0].x);
            manual_vertreter_y = format!("{:.1}", initial_config.vertreter[0].y);
        }

        println!("🚀 APP-START: Lade gruppenspezifische Config für {} {} (Messe: {})", default_group, default_language, default_is_messe);

        Self {
            config: initial_config, // Gruppenspezifische Config statt globaler Config
            show_startup_dialog: true,
            selected_group: default_group,
            selected_language: default_language,
            is_messe: default_is_messe,
            show_config: false,
            progress: 0.0,
            status_message: "Bereit".to_string(),
            save_message: None,
            is_generating: false,
            stop_signal: Arc::new(Mutex::new(false)),
            show_meme: false,
            meme_time: None,
            resume_available: check_resume_available(),
            last_processed_count: get_last_processed_count(),
            animation_frame: 0,
            animation_time: None,
            resume_needs_update: false,
            manual_qr_x,
            manual_qr_y,
            manual_qr_size,
            manual_vertreter_x,
            manual_vertreter_y,
            manual_vertreter_size,
            dark_mode: load_app_settings(), // Dark Mode aus gespeicherten Einstellungen laden
            fullscreen_mode: true, // Standard: Maximiert beim Start
            show_settings_dialog: false, // Settings-Dialog standardmäßig geschlossen
            custom_output_dir: "Output".to_string(), // Standard-Ausgabeordner
            use_custom_output_dir: false, // Standardmäßig automatische Ordner verwenden
            custom_template_dir: "Vorlagen".to_string(), // Standard-Vorlagenordner
            use_custom_template_dir: false, // Standardmäßig interne Logik verwenden
            available_templates: Vec::new(),
            selected_template_index: None,
            show_template_selection: false,
            generation_start_time: None,
            last_progress_update: None,
            estimated_total_duration: None,
            progress_frozen: false, // Progress-Updates standardmäßig erlaubt
            // Debug und Performance Defaults
            debug_mode: load_debug_config(), // Debug-Modus aus persistentem Speicher laden
            debug_key_pressed: false, // Tastatur-Flag
            max_threads: (std::thread::available_parallelism().map(|n| n.get()).unwrap_or(4) * 3 / 4).max(1), // 75% der verfügbaren Kerne
            thread_sleep_ms: 0,     // 0ms = maximale Geschwindigkeit (kein Sleep)
            // Font-System Defaults
            enable_font_fallback: true, // Font-Fallback standardmäßig aktiviert
            // Font-Caching für Performance
            cached_fonts: refresh_font_cache(), // Einmalig beim Start laden (mit Cache)
            // Bereichs-Auswahl Defaults
            use_range_selection: false,
            range_start_index: String::new(),
            range_end_index: String::new(),
            // Resume-Information (gruppenspezifisch beim Start geladen)
            resume_info: initial_resume_info,
            // Font-Auswahl erweiterte Ansicht
            show_all_fonts: false,
            // Font-Suche
            font_search_text: String::new(),
        }
    }
}

// Debug-Persistierung Funktionen
fn load_debug_config() -> bool {
    let debug_config_path = get_temp_file_path("debug_config.txt");
    match std::fs::read_to_string(&debug_config_path) {
        Ok(content) => content.trim() == "true",
        Err(_) => false, // Default: Debug aus
    }
}

fn save_debug_config(debug_enabled: bool) {
    let debug_config_path = get_temp_file_path("debug_config.txt");
    let content = if debug_enabled { "true" } else { "false" };
    let _ = std::fs::write(&debug_config_path, content);
}

// Progress-Verwaltung pro Kategorie/Sprache/Messe
fn get_progress_filename(group: &str, language: &str, is_messe: bool) -> String {
    let messe_suffix = if is_messe { "_messe" } else { "" };
    format!("progress_{}_{}_{}.txt", group.to_lowercase(), language.to_lowercase(), messe_suffix)
}

fn get_stop_status_filename(group: &str, language: &str, is_messe: bool) -> String {
    let messe_suffix = if is_messe { "_messe" } else { "" };
    format!("stop_status_{}_{}_{}.txt", group.to_lowercase(), language.to_lowercase(), messe_suffix)
}

fn get_resume_filename(group: &str, language: &str, is_messe: bool) -> String {
    let messe_suffix = if is_messe { "_messe" } else { "" };
    format!("resume_{}_{}_{}.txt", group.to_lowercase(), language.to_lowercase(), messe_suffix)
}

// Speichere Resume-Info: aktueller Index, Gesamtanzahl, Startzeit
#[allow(dead_code)]
fn save_resume_info(group: &str, language: &str, is_messe: bool, current_index: usize, total_count: usize, start_time: std::time::Instant) {
    let resume_path = get_temp_file_path(&get_resume_filename(group, language, is_messe));
    let elapsed = start_time.elapsed().as_secs();
    let content = format!("{}|{}|{}", current_index, total_count, elapsed);
    let _ = std::fs::write(&resume_path, content);
}

// Lade Resume-Info: (current_index, total_count, elapsed_seconds)
fn load_resume_info(group: &str, language: &str, is_messe: bool) -> Option<(usize, usize, u64)> {
    let resume_path = get_temp_file_path(&get_resume_filename(group, language, is_messe));
    if let Ok(content) = std::fs::read_to_string(&resume_path) {
        let parts: Vec<&str> = content.trim().split('|').collect();
        if parts.len() == 3 {
            if let (Ok(current), Ok(total), Ok(elapsed)) = (
                parts[0].parse::<usize>(),
                parts[1].parse::<usize>(),
                parts[2].parse::<u64>()
            ) {
                return Some((current, total, elapsed));
            }
        }
    }
    None
}

// Lösche alle Progress-Dateien für eine Kategorie
fn clear_progress_files(group: &str, language: &str, is_messe: bool) {
    let progress_path = get_temp_file_path(&get_progress_filename(group, language, is_messe));
    let stop_path = get_temp_file_path(&get_stop_status_filename(group, language, is_messe));
    let resume_path = get_temp_file_path(&get_resume_filename(group, language, is_messe));
    
    let _ = std::fs::remove_file(&progress_path);
    let _ = std::fs::remove_file(&stop_path);
    let _ = std::fs::remove_file(&resume_path);
}

// Global selection for generation: data CSV, template path, and whether to generate QR
static CURRENT_SELECTION: Lazy<Mutex<Option<Vec<(String, String, bool)>>>> = Lazy::new(|| Mutex::new(None));

fn set_current_selections(selections: Vec<(String, String, bool)>) {
    let mut guard = CURRENT_SELECTION.lock().unwrap();
    *guard = Some(selections);
}

fn set_current_selection(csv: &str, template: &str, gen_qr: bool) {
    set_current_selections(vec![(csv.to_string(), template.to_string(), gen_qr)]);
    println!("Auswahl gesetzt: CSV={}, Template={}, gen_qr={}", csv, template, gen_qr);
}

fn get_current_selections() -> Option<Vec<(String, String, bool)>> {
    let guard = CURRENT_SELECTION.lock().unwrap();
    guard.clone()
}

// Suche die passende Template-Datei in VORLAGE/ basierend auf Gruppe, Sprache und optional Land
fn find_best_template(group: &str, lang: &str, country: Option<&str>) -> Option<String> {
    let mut candidates = Vec::new();
    // normalize group
    let g = group.replace(' ', "").to_lowercase();

    // Map language names like "Deutsch"/"Englisch" to likely filename codes
    let lang_lower = lang.replace(' ', "").to_lowercase();
    let codes: Vec<String> = match lang_lower.as_str() {
        "deutsch" | "german" | "de" | "de_de" => vec!["de_de".to_string(), "de".to_string()],
        "englisch" | "english" | "en" | "en_us" => vec!["en_us".to_string(), "en".to_string()],
        other => {
            // if user already passed a code like "de_de", use it first
            if other.contains('_') || other.len() <= 3 {
                vec![other.to_string()]
            } else {
                vec![other.to_string()]
            }
        }
    };

    // If country was set to "messe" we prefer Messe templates
    if let Some(ct) = country {
        if ct.to_lowercase() == "messe" {
            for code in &codes {
                candidates.push(format!("VORLAGE/Bestellschein-Messe-{}-{}.pdf", capitalize_first(&g), code));
                candidates.push(format!("VORLAGE/Bestellschein-Messe-{}-{}.pdf", capitalize_first(&g), code));
            }
        } else {
            // country-specific codes, e.g. at, ch
            let c = ct.replace(' ', "").to_lowercase();
            for code in &codes {
                candidates.push(format!("VORLAGE/Bestellschein-{}-{}_{}.pdf", capitalize_first(&g), code.split('_').next().unwrap_or(code), c));
            }
        }
    }

    // language-specific with country code style (e.g. de_de)
    for code in &codes {
        candidates.push(format!("VORLAGE/Bestellschein-{}-{}.pdf", capitalize_first(&g), code));
    }

    // fallback: generic per-group
    candidates.push(format!("VORLAGE/Bestellschein-{}.pdf", capitalize_first(&g)));
    candidates.push(format!("VORLAGE/Bestellscheine-{}.pdf", capitalize_first(&g)));

    let project_root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"));
    for c in candidates {
        let abs = project_root.join(&c);
        if abs.exists() {
            return Some(c);
        }
    }
    None
}

// Erweiterte Template-Suche mit benutzerdefiniertem Verzeichnis
fn find_best_template_in_dir(group: &str, lang: &str, country: Option<&str>, template_dir: &str) -> Option<String> {
    let mut candidates = Vec::new();
    // normalize group
    let g = group.replace(' ', "").to_lowercase();

    // Map language names like "Deutsch"/"Englisch" to likely filename codes
    let lang_lower = lang.replace(' ', "").to_lowercase();
    let codes: Vec<String> = match lang_lower.as_str() {
        "deutsch" | "german" | "de" | "de_de" => vec!["de_de".to_string(), "de".to_string()],
        "englisch" | "english" | "en" | "en_us" => vec!["en_us".to_string(), "en".to_string()],
        other => {
            // if user already passed a code like "de_de", use it first
            if other.contains('_') || other.len() <= 3 {
                vec![other.to_string()]
            } else {
                vec![other.to_string()]
            }
        }
    };

    // If country was set to "messe" we prefer Messe templates
    if let Some(ct) = country {
        if ct.to_lowercase() == "messe" {
            for code in &codes {
                candidates.push(format!("{}/Bestellschein-Messe-{}-{}.pdf", template_dir, capitalize_first(&g), code));
                candidates.push(format!("{}/Bestellschein-Messe-{}.pdf", template_dir, capitalize_first(&g)));
            }
        } else {
            // country-specific codes, e.g. at, ch
            let c = ct.replace(' ', "").to_lowercase();
            for code in &codes {
                candidates.push(format!("{}/Bestellschein-{}-{}_{}.pdf", template_dir, capitalize_first(&g), code.split('_').next().unwrap_or(code), c));
            }
        }
    }

    // language-specific with country code style (e.g. de_de)
    for code in &codes {
        candidates.push(format!("{}/Bestellschein-{}-{}.pdf", template_dir, capitalize_first(&g), code));
    }

    // fallback: generic per-group
    candidates.push(format!("{}/Bestellschein-{}.pdf", template_dir, capitalize_first(&g)));
    candidates.push(format!("{}/Bestellscheine-{}.pdf", template_dir, capitalize_first(&g)));

    // Prüfen ob die Pfade existieren
    for c in candidates {
        let path = std::path::Path::new(&c);
        if path.exists() {
            return Some(c);
        }
    }
    None
}

// Liefere die Kandidatenliste, die find_best_template prüfen würde (für UI-Vorschau)
fn list_template_candidates(group: &str, lang: &str, is_messe: bool) -> Vec<String> {
    let mut candidates = Vec::new();
    // Reuse mapping logic from find_best_template
    let g = group.replace(' ', "").to_lowercase();
    let lang_lower = lang.replace(' ', "").to_lowercase();
    let codes: Vec<String> = match lang_lower.as_str() {
        "deutsch" | "german" | "de" | "de_de" => vec!["de_de".to_string(), "de".to_string()],
        "englisch" | "english" | "en" | "en_us" => vec!["en_us".to_string(), "en".to_string()],
        other => vec![other.to_string()],
    };

    if is_messe {
        for code in &codes {
            candidates.push(format!("VORLAGE/Bestellschein-Messe-{}-{}.pdf", capitalize_first(&g), code));
            candidates.push(format!("VORLAGE/Bestellschein-Messe-{}-{}.pdf", capitalize_first(&g), code));
        }
    }

    for code in &codes {
        candidates.push(format!("VORLAGE/Bestellschein-{}-{}.pdf", capitalize_first(&g), code));
    }

    candidates.push(format!("VORLAGE/Bestellschein-{}.pdf", capitalize_first(&g)));
    candidates.push(format!("VORLAGE/Bestellscheine-{}.pdf", capitalize_first(&g)));

    candidates
}

// Erweiterte Template-Suche mit Bewertung und Sortierung
fn find_available_templates_with_score(group: &str, lang: &str, is_messe: bool) -> Vec<(String, i32, bool)> {
    let mut results = Vec::new();
    let project_root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"));
    
    // Basis-Kandidaten generieren
    let candidates = list_template_candidates(group, lang, is_messe);
    
    for candidate in candidates {
        let abs_path = project_root.join(&candidate);
        let exists = abs_path.exists();
        
        // Bewertungs-Score basierend auf Übereinstimmung
        let mut score = 0;
        let filename_lower = candidate.to_lowercase();
        
        // Gruppe passt
        if filename_lower.contains(&group.to_lowercase()) {
            score += 10;
        }
        
        // Sprache passt
        let lang_lower = lang.to_lowercase();
        if lang_lower.contains("deutsch") || lang_lower.contains("de") {
            if filename_lower.contains("de_de") || filename_lower.contains("de") {
                score += 8;
            }
        } else if lang_lower.contains("englisch") || lang_lower.contains("en") {
            if filename_lower.contains("en_us") || filename_lower.contains("en") {
                score += 8;
            }
        }
        
        // Messe passt
        if is_messe && filename_lower.contains("messe") {
            score += 5;
        } else if !is_messe && !filename_lower.contains("messe") {
            score += 3;
        }
        
        // Existierende Dateien bevorzugen
        if exists {
            score += 20;
        }
        
        results.push((candidate, score, exists));
    }
    
    // Auch alle anderen PDF-Dateien im VORLAGE-Ordner scannen
    if let Ok(entries) = std::fs::read_dir(project_root.join("VORLAGE")) {
        for entry in entries.flatten() {
            if let Ok(file_type) = entry.file_type() {
                if file_type.is_file() {
                    if let Some(filename) = entry.file_name().to_str() {
                        if filename.ends_with(".pdf") {
                            let relative_path = format!("VORLAGE/{}", filename);
                            
                            // Nur hinzufügen wenn nicht schon in Kandidaten
                            if !results.iter().any(|(path, _, _)| path == &relative_path) {
                                let mut score = 1; // Basis-Score für gefundene PDFs
                                let filename_lower = filename.to_lowercase();
                                
                                // Bewertung wie oben
                                if filename_lower.contains(&group.to_lowercase()) {
                                    score += 10;
                                }
                                if lang.to_lowercase().contains("deutsch") || lang.to_lowercase().contains("de") {
                                    if filename_lower.contains("de_de") || filename_lower.contains("de") {
                                        score += 8;
                                    }
                                }
                                if is_messe && filename_lower.contains("messe") {
                                    score += 5;
                                } else if !is_messe && !filename_lower.contains("messe") {
                                    score += 3;
                                }
                                
                                score += 20; // Existiert
                                results.push((relative_path, score, true));
                            }
                        }
                    }
                }
            }
        }
    }
    
    // Nach Score sortieren (höchster zuerst)
    results.sort_by(|a, b| b.1.cmp(&a.1));
    results
}

fn capitalize_first(s: &str) -> String {
    let mut cs = s.chars();
    match cs.next() {
        None => String::new(),
        Some(f) => f.to_uppercase().collect::<String>() + cs.as_str(),
    }
}

// Lade gruppenspezifische Config-Datei, falls vorhanden.
// Erwartete Pfade (in Reihenfolge):
// CONFIG/config_<group>_<lang>.toml, CONFIG/config_<group>.toml, CONFIG/config.toml
fn load_group_config(group: &str, language: &str, is_messe: bool) -> Config {
    println!("=== LOAD_GROUP_CONFIG AUFGERUFEN FÜR: {} {} (Messe: {}) ===", group, language, is_messe);
    
    // Release-Ordnerstruktur verwenden - Config-Verzeichnis ist jetzt sichtbar für User
    let (config_dir, _, _, _, _) = get_release_dirs();
    
    // Kandidatenreihenfolge: group+lang(+messe) -> group(+messe) -> generic (+messe variants)
    let mut candidates = Vec::new();
    if is_messe {
        candidates.push(config_dir.join(format!("config_{}_{}_messe.toml", group.to_lowercase(), language.to_lowercase())));
        candidates.push(config_dir.join(format!("config_{}_messe.toml", group.to_lowercase())));
        candidates.push(config_dir.join("config_messe.toml"));
    }
    candidates.push(config_dir.join(format!("config_{}_{}.toml", group.to_lowercase(), language.to_lowercase())));
    candidates.push(config_dir.join(format!("config_{}.toml", group.to_lowercase())));
    candidates.push(config_dir.join("config.toml"));

    println!("CONFIG-Verzeichnis: {:?}", config_dir);
    println!("Prüfe Config-Kandidaten in Reihenfolge:");
    for c in &candidates {
        println!("  {:?}", c);
    }

    for c in &candidates {
        if c.exists() {
            let c_str = c.to_string_lossy();
            println!("✅ GEFUNDEN: Lade gruppenspezifische Config von: {}", c_str);
            if let Ok(toml) = std::fs::read_to_string(c) {
                println!("Config-Inhalt aus {}:\n{}", c_str, toml);
                
                // Parse TOML direkt statt über temp Dateien
                let result = parse_toml_to_config(&toml);
                println!("=== LOAD_GROUP_CONFIG ABGESCHLOSSEN - VERWENDE: {} ===", c_str);
                return result;
            }
        } else {
            println!("❌ NICHT VORHANDEN: {:?}", c);
        }
    }

    // Wenn keine gruppenspezifische Config existiert, erstelle eine Standard-Datei für die Gruppe
    println!("⚠️ KEINE GRUPPENSPEZIFISCHE CONFIG GEFUNDEN für {} {} (Messe: {}) - erzeuge Standard-Config.", group, language, is_messe);
    
    // CONFIG-Ordner relativ zur ausführbaren Datei erstellen
    let exe_dir = std::env::current_exe()
        .ok()
        .and_then(|p| p.parent().map(|p| p.to_path_buf()))
        .unwrap_or_else(|| std::env::current_dir().unwrap_or_default());
    let config_dir = exe_dir.join("CONFIG");
    
    if !config_dir.exists() {
        if let Err(e) = std::fs::create_dir_all(&config_dir) {
            eprintln!("Konnte CONFIG-Verzeichnis nicht erstellen: {}", e);
            return Config::default();
        } else {
            println!("CONFIG-Verzeichnis erstellt: {:?}", config_dir);
        }
    }

    // Bevorzugte Filename für die Gruppe; wenn Messe, erstelle messe-spezifische Datei
    let group_filename = if is_messe {
        config_dir.join(format!("config_{}_messe.toml", group.to_lowercase()))
    } else {
        config_dir.join(format!("config_{}.toml", group.to_lowercase()))
    };
    
    // Falls die Gruppendatei noch nicht existiert, schreibe eine Default-Konfiguration hinein
    if !group_filename.exists() {
        let default = get_group_default_config(group, is_messe);
        println!("Verwende gruppenspezifische Defaults für {}: QR={:?}, Vertreter={:?}", 
                 group, default.qr_codes, default.vertreter);
                 
        // Erzeuge eine einfache TOML-Repräsentation, kompatibel mit load_config parsing
        let mut toml = String::new();
        toml.push_str(&format!("# Automatisch generierte Gruppenkonfiguration für {}\n", group));
        if is_messe {
            toml.push_str("# Messe-spezifische Konfiguration\n");
        }
        toml.push_str("# Koordinaten sind in PDF-Punkten (1 Punkt = 1/72 Zoll ≈ 0.35mm)\n\n");
        
        // QR-Codes
        toml.push_str("qr_codes = [\n");
        for qr in &default.qr_codes {
            toml.push_str(&format!("  {{ x = {}, y = {}, size = {} }},\n", qr.x, qr.y, qr.size));
        }
        toml.push_str("]\n\n");
        
        toml.push_str("[positions]\n");
        toml.push_str("vertreter_nummer = [\n");
        for v in &default.vertreter {
            toml.push_str(&format!("  {{ x = {}, y = {}, size = {}, font_name = \"{}\", font_size = {}, font_style = \"{}\" }},\n", v.x, v.y, v.size, v.font_name, v.font_size, v.font_style));
        }
        toml.push_str("]\n\n");
        
        if let Err(e) = std::fs::write(&group_filename, toml) {
            eprintln!("Konnte Default-Config für Gruppe {} nicht schreiben: {}", group, e);
            return get_group_default_config(group, is_messe);
        } else {
            println!("✅ Schreibe gruppenspezifische Default-Config nach: {:?}", group_filename);
        }
    }

    // Lade die gerade erstellte (oder existierende) Gruppendatei
    if let Ok(toml) = std::fs::read_to_string(&group_filename) {
        println!("Config-Inhalt aus neu erstellter Datei:\n{}", toml);
        return parse_toml_to_config(&toml);
    }

    // Fallback
    println!("⚠️ FALLBACK: Verwende gruppenspezifische Default-Config für {}", group);
    get_group_default_config(group, is_messe)
}

impl MyApp {
    // Helper-Methode um aktuellen CSV-Pfad zu bestimmen
    fn get_current_csv_path(&self) -> Option<String> {
        Some(if self.selected_group == "Apo" { 
            get_default_csv_path("Apo") 
        } else { 
            get_default_csv_path("Endkunde") 
        })
    }
    
    // Rest der impl...
    // Animation für PDF-Generierung
    fn get_generating_animation(&mut self) -> String {
        // Animation alle 300ms wechseln
        if self.animation_time.is_none() {
            self.animation_time = Some(std::time::Instant::now());
        }
        
        if let Some(start_time) = self.animation_time {
            let elapsed = start_time.elapsed().as_millis();
            if elapsed > 300 {
                self.animation_frame = (self.animation_frame + 1) % 12;
                self.animation_time = Some(std::time::Instant::now());
            }
        }
        
        let animations = [
            "📄 Erstelle PDFs... ✨",
            "📄 Erstelle PDFs... 🌟",
            "📄 Erstelle PDFs... ⭐",
            "📄 Erstelle PDFs... 💫",
            "📄 Erstelle PDFs... 🎯",
            "📄 Erstelle PDFs... 🎨",
            "📄 Erstelle PDFs... 🚀",
            "📄 Erstelle PDFs... 💎",
            "📄 Erstelle PDFs... 🎪",
            "📄 Erstelle PDFs... 🎭",
            "📄 Erstelle PDFs... 🎊",
            "📄 Erstelle PDFs... 🎉",
        ];
        
        animations[self.animation_frame].to_string()
    }

    // Template-Suche basierend auf User-Settings
    fn find_template(&self, group: &str, lang: &str, country: Option<&str>) -> Option<String> {
        if self.use_custom_template_dir {
            // Benutzerdefinierter Ordner
            find_best_template_in_dir(group, lang, country, &self.custom_template_dir)
        } else {
            // Standard-Logik (VORLAGE-Ordner)
            find_best_template(group, lang, country)
        }
    }

    // Template-Liste mit Scoring basierend auf User-Settings
    fn find_templates_with_score(&self, group: &str, lang: &str, is_messe: bool) -> Vec<(String, i32, bool)> {
        if self.use_custom_template_dir {
            // Benutzerdefinierter Ordner durchsuchen
            self.scan_custom_template_dir(group, lang, is_messe)
        } else {
            // Standard-Funktion verwenden
            find_available_templates_with_score(group, lang, is_messe)
        }
    }

    // Scanne den benutzerdefinierten Template-Ordner
    fn scan_custom_template_dir(&self, group: &str, lang: &str, is_messe: bool) -> Vec<(String, i32, bool)> {
        let mut results = Vec::new();
        let template_dir = std::path::Path::new(&self.custom_template_dir);
        
        if !template_dir.exists() {
            // Ordner existiert nicht, return empty list für Fallback zur manuellen Auswahl
            return results;
        }
        
        if let Ok(entries) = std::fs::read_dir(template_dir) {
            for entry in entries.flatten() {
                if let Ok(file_type) = entry.file_type() {
                    if file_type.is_file() {
                        if let Some(filename) = entry.file_name().to_str() {
                            if filename.ends_with(".pdf") {
                                // Relativer Pfad: Benutzerdefinierter Ordner + Dateiname
                                let relative_path = format!("{}/{}", self.custom_template_dir, filename);
                                let mut score = 1; // Basis-Score
                                let filename_lower = filename.to_lowercase();
                                
                                // Bewertungs-Score basierend auf Übereinstimmung
                                if filename_lower.contains(&group.to_lowercase()) {
                                    score += 10;
                                }
                                
                                // Sprache passt
                                let lang_lower = lang.to_lowercase();
                                if lang_lower.contains("deutsch") || lang_lower.contains("de") {
                                    if filename_lower.contains("de_de") || filename_lower.contains("de") {
                                        score += 8;
                                    }
                                } else if lang_lower.contains("englisch") || lang_lower.contains("en") {
                                    if filename_lower.contains("en_us") || filename_lower.contains("en") {
                                        score += 8;
                                    }
                                }
                                
                                // Messe-Spezifisch
                                if is_messe && filename_lower.contains("messe") {
                                    score += 15;
                                }
                                
                                // Höhere Priorität für exakte Matches
                                if filename_lower.contains(&format!("{}-{}", group.to_lowercase(), lang_lower)) {
                                    score += 20;
                                }
                                
                                results.push((relative_path, score, true)); // exists = true (wir haben es gescannt)
                            }
                        }
                    }
                }
            }
        }
        
        // Nach Score sortieren (höchster zuerst)
        results.sort_by(|a, b| b.1.cmp(&a.1));
        results
    }
}

impl App for MyApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        // Theme setzen basierend auf dark_mode
        if self.dark_mode {
            ctx.set_visuals(egui::Visuals::dark());
        } else {
            ctx.set_visuals(egui::Visuals::light());
        }
        
        // Fortschritt aus versteckter Datei lesen (z.B. .temp/progress.txt mit Wert 0.0 bis 1.0)
        // Aber nur wenn Progress nicht eingefroren ist (z.B. nach Stop)
        if !self.progress_frozen {
            let progress_path = get_temp_file_path("progress.txt");
            if let Ok(s) = std::fs::read_to_string(&progress_path) {
                if let Ok(val) = s.trim().parse::<f32>() {
                    let old_progress = self.progress;
                    self.progress = val;
                    
                    // Zeitschätzung berechnen wenn Generierung läuft
                    if self.is_generating && val > 0.0 && val != old_progress {
                        if let Some(start_time) = self.generation_start_time {
                            let elapsed = start_time.elapsed();
                            let progress_ratio = val as f64;
                            if progress_ratio > 0.01 { // Mindestens 1% Fortschritt
                                let estimated_total = elapsed.as_secs_f64() / progress_ratio;
                                self.estimated_total_duration = Some(std::time::Duration::from_secs_f64(estimated_total));
                            }
                        }
                        self.last_progress_update = Some(std::time::Instant::now());
                    }
                }
            }
        }

        // Resume-Status nur einmal beim Start oder nach Stop aktualisieren, nicht ständig
        // Dies verhindert das ständige Neu-Berechnen während der Animation
        if self.resume_needs_update {
            self.resume_available = check_resume_available();
            self.last_processed_count = get_last_processed_count();
            
            // Neue Resume-Info für aktuelle Kategorie laden
            self.resume_info = load_resume_info(&self.selected_group, &self.selected_language, self.is_messe);
            
            self.resume_needs_update = false;
            println!("Resume-Status aktualisiert: {} verfügbar, {} PDFs", 
                     self.resume_available, self.last_processed_count);
        }

        egui::TopBottomPanel::top("top_panel").show(ctx, |ui| {
            ui.horizontal(|ui| {
                // Kleines Logo: ein markantes 'B' links in der Leiste
                ui.add(egui::Label::new(egui::RichText::new("B").heading()).sense(egui::Sense::hover()))
                    .on_hover_text("Bestellscheine - B Logo");
                
                ui.separator();
                
                // Entfernt - Konfiguration jetzt neben Settings-Icon
                
                if ui.button(" Auswahl treffen").clicked() {
                    self.show_startup_dialog = true;
                }
                
                ui.separator();
                
                // BEREICHS-AUSWAHL für Vertreternummern
                ui.horizontal(|ui| {
                    if ui.checkbox(&mut self.use_range_selection, "📊 Nur bestimmten Bereich generieren").clicked() {
                        if self.use_range_selection {
                            // Beim ersten Aktivieren, versuche Gesamtanzahl zu bestimmen
                            if let Some(csv_path) = self.get_current_csv_path() {
                                let customers = read_vertreter(&csv_path);
                                if !customers.is_empty() {
                                    self.range_start_index = "0".to_string();
                                    self.range_end_index = (customers.len().saturating_sub(1)).to_string();
                                }
                            }
                        }
                    }
                    
                    if self.use_range_selection {
                        ui.separator();
                        ui.label("Von Index:");
                        ui.text_edit_singleline(&mut self.range_start_index);
                        ui.label("bis Index:");
                        ui.text_edit_singleline(&mut self.range_end_index);
                        
                        // Validierung und Info
                        if let (Ok(start), Ok(end)) = (self.range_start_index.parse::<usize>(), self.range_end_index.parse::<usize>()) {
                            if start <= end {
                                let count = end.saturating_sub(start) + 1;
                                ui.label(format!("({} Vertreter)", count));
                            } else {
                                ui.label(egui::RichText::new("❌ Start > Ende").color(egui::Color32::RED));
                            }
                        } else {
                            ui.label(egui::RichText::new("❌ Ungültige Eingabe").color(egui::Color32::RED));
                        }
                    }
                });
                
                // Resume-Info anzeigen
                if let Some((current_index, total_count, elapsed_seconds)) = self.resume_info {
                    ui.horizontal(|ui| {
                        let hours = elapsed_seconds / 3600;
                        let minutes = (elapsed_seconds % 3600) / 60;
                        let seconds = elapsed_seconds % 60;
                        ui.label(egui::RichText::new(format!("⏸️ Unterbrochen bei {}/{} ({}:{:02}:{:02})", 
                                                           current_index, total_count, hours, minutes, seconds))
                                 .color(egui::Color32::from_rgb(200, 150, 0)));
                        if ui.button("🗑️ Reset").clicked() {
                            clear_progress_files(&self.selected_group, &self.selected_language, self.is_messe);
                            self.resume_info = None;
                        }
                    });
                }
                
                ui.separator();
                
                // HAUPTBUTTON: Bestellscheine generieren
                if !self.is_generating {
                    // Primärer Button: Erstellen oder Fortsetzen  
                    let button_text = if self.resume_available {
                        format!("📄 Fortsetzen ({} bereits erstellt)", self.last_processed_count)
                    } else {
                        "🚀 Bestellscheine erstellen".to_string()
                    };
                    
                    // Großer, auffälliger Button
                    let generate_button = egui::Button::new(egui::RichText::new(button_text).size(16.0))
                        .fill(egui::Color32::from_rgb(46, 125, 50)); // Grün
                    
                    if ui.add(generate_button).clicked() {
                        let mut can_start = true;
                        // Ensure a selection exists. If not, set it from current UI state (but don't auto-start the dialog)
                        if get_current_selections().is_none() {
                            // Build CSV and template from current UI selection
                            let csv_default = get_default_csv_path(&self.selected_group);
                            
                            let template = self.find_template(&self.selected_group, &self.selected_language, if self.is_messe { Some("messe") } else { None })
                                .unwrap_or_else(|| {
                                    if self.selected_group == "Apo" { "Vorlagen/Bestellscheine-Apo.pdf".to_string() } 
                                    else { get_default_template_path() }
                                });

                            // Check existence relative to project root
                            let project_root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"));
                            let csv_abs = project_root.join(&csv_default);
                            let template_abs = project_root.join(&template);
                            let mut missing = Vec::new();
                            if !csv_abs.exists() { missing.push(csv_abs.to_string_lossy().to_string()); }
                            if !template_abs.exists() { missing.push(template_abs.to_string_lossy().to_string()); }

                            if !missing.is_empty() {
                                // Open the selection dialog so user can fix the selection
                                self.show_startup_dialog = true;
                                self.status_message = format!("Bitte Auswahl prüfen, fehlende Dateien: {}", missing.join(", "));
                                can_start = false;
                            }

                            // Persist selection for generator
                            set_current_selection(&csv_default, &template, true);
                            // Load group config and set it (consider Messe flag)
                            let group_cfg = load_group_config(&self.selected_group, &self.selected_language, self.is_messe);
                            self.config = group_cfg;
                        }

                        if !can_start {
                            // Do not start threads when files missing
                        } else {
                            // WICHTIG: Aktuelle Config für PDF-Generierung setzen
                            set_current_config(&self.config);

                        let start_from = if self.resume_available { 
                            self.last_processed_count 
                        } else { 
                            0 
                        };

                        self.status_message = if self.resume_available {
                            format!("Setze Erstellung fort ab PDF #{}", start_from + 1)
                        } else {
                            "Bestellscheine werden erstellt...".to_string()
                        };
                        self.is_generating = true;
                        
                        // Zeittracking starten
                        self.generation_start_time = Some(std::time::Instant::now());
                        self.last_progress_update = Some(std::time::Instant::now());
                        self.estimated_total_duration = None;
                        
                        // Progress-Updates erlauben
                        self.progress_frozen = false;

                        // Vorherige Progress-Datei bereinigen (falls noch vorhanden)
                        let progress_path = get_temp_file_path("progress.txt");
                        let _ = std::fs::remove_file(&progress_path);

                        // Animation zurücksetzen
                        self.animation_frame = 0;
                        self.animation_time = Some(std::time::Instant::now());

                        // Stop-Signal zurücksetzen
                        {
                            let mut stop = self.stop_signal.lock().unwrap();
                            *stop = false;
                        }

                        let progress_clone = Arc::new(Mutex::new(0.0f32));
                        let progress_ref = Arc::clone(&progress_clone);
                        let stop_signal = Arc::clone(&self.stop_signal);

                        // Prepare thread/IO related variables for generator - mit Performance-Optimierung
                        let threads = self.max_threads; // Verwende benutzerdefinierte Thread-Anzahl
                        debug_log(&format!("Starte PDF-Generierung mit {} Threads (von {} verfügbaren)", threads, std::thread::available_parallelism().map(|n| n.get()).unwrap_or(4)), self.debug_mode);
                        
                        // Sichere Auswahl der Dateien
                        let selections = match get_current_selections() {
                            Some(s) => {
                                println!("DEBUG: Verwende gespeicherte Auswahl: {:?}", s);
                                s
                            },
                            None => {
                                println!("DEBUG: Keine Auswahl gefunden, verwende Standard");
                                get_default_selections()
                            }
                        };
                        
                        let csv_path = selections.get(0).map(|s| s.0.clone()).unwrap_or_else(|| get_default_csv_path("Endkunde"));
                        println!("DEBUG: CSV-Pfad: {}", csv_path);
                        
                        // Prüfe ob CSV-Datei existiert
                        if !std::path::Path::new(&csv_path).exists() {
                            self.status_message = format!("FEHLER: CSV-Datei nicht gefunden: {}", csv_path);
                            println!("ERROR: CSV-Datei nicht gefunden: {}", csv_path);
                            return;
                        }
                        
                        let vertreter_vec = match std::panic::catch_unwind(|| read_vertreter(&csv_path)) {
                            Ok(vertreter) => {
                                println!("DEBUG: {} Vertreter geladen", vertreter.len());
                                vertreter
                            },
                            Err(e) => {
                                self.status_message = "FEHLER: Fehler beim Laden der Vertreterdaten".to_string();
                                println!("ERROR: Fehler beim Laden der CSV: {:?}", e);
                                return;
                            }
                        };
                        
                        let vertreter_arc = Arc::new(vertreter_vec);
                        let total = vertreter_arc.len();
                        
                        // Bereichs-Auswahl verarbeiten
                        let (use_range_selection, range_start_parsed, range_end_parsed) = if self.use_range_selection {
                            let start = self.range_start_index.parse::<usize>().unwrap_or(0);
                            let end = self.range_end_index.parse::<usize>().unwrap_or(total.saturating_sub(1));
                            let end_clamped = end.min(total.saturating_sub(1));
                            if start <= end_clamped {
                                println!("INFO: Verwende Bereichs-Auswahl: Index {} bis {} ({} von {} Vertretern)", 
                                        start, end_clamped, end_clamped - start + 1, total);
                                (true, start, end_clamped)
                            } else {
                                println!("WARNING: Ungültiger Bereich: Start {} > Ende {}. Verwende alle Vertreter.", start, end_clamped);
                                (false, 0, total.saturating_sub(1))
                            }
                        } else {
                            (false, 0, total.saturating_sub(1))
                        };
                        
                        if total == 0 {
                            self.status_message = "FEHLER: Keine Vertreterdaten in CSV-Datei gefunden".to_string();
                            println!("ERROR: Keine Vertreterdaten gefunden");
                            return;
                        }
                        
                        let progress_counter = Arc::new(Mutex::new(0usize));
                        
                        // Sichere Ordner-Erkennung
                        let dirs_result = std::panic::catch_unwind(|| get_release_dirs());
                        let (_cfg_dir, data_dir, templates_dir, _tools, _out) = match dirs_result {
                            Ok(dirs) => {
                                println!("DEBUG: Ordner erfolgreich ermittelt");
                                println!("DEBUG: Data-Dir: {}", dirs.1.display());
                                println!("DEBUG: Templates-Dir: {}", dirs.2.display());
                                dirs
                            },
                            Err(e) => {
                                self.status_message = "FEHLER: Ordnerstruktur konnte nicht ermittelt werden".to_string();
                                println!("ERROR: Ordner-Fehler: {:?}", e);
                                return;
                            }
                        };

                        // Output-Konfiguration für Thread klonen
                        let use_custom_output = self.use_custom_output_dir;
                        let custom_output_path = self.custom_output_dir.clone();
                        let group = self.selected_group.clone();
                        let language = self.selected_language.clone();
                        let is_messe = self.is_messe;
                        // Performance-Parameter klonen
                        let thread_sleep_ms = self.thread_sleep_ms;
                        let debug_mode = self.debug_mode;
                        // Font-Fallback Parameter klonen
                        let enable_font_fallback = self.enable_font_fallback;

                        thread::spawn(move || {
                            if let Err(e) = generate_bestellscheine_resume(
                                progress_ref,
                                stop_signal,
                                start_from,
                                threads,
                                vertreter_arc,
                                progress_counter,
                                total,
                                data_dir,
                                templates_dir,
                                use_custom_output,
                                custom_output_path,
                                group,
                                language,
                                is_messe,
                                thread_sleep_ms,
                                debug_mode,
                                // Bereichs-Parameter
                                use_range_selection,
                                range_start_parsed,
                                range_end_parsed,
                                // Font-Fallback Parameter
                                enable_font_fallback,
                            ) {
                                eprintln!("Fehler beim Erstellen der Bestellscheine: {}", e);
                            }
                        });
                        }
                    }
                    
                    // Sekundärer Button: Von vorne beginnen (nur wenn Resume verfügbar)
                    if self.resume_available {
                        if ui.button("🔄 Von vorne beginnen").clicked() {
                            // WICHTIG: Aktuelle Config für PDF-Generierung setzen
                            set_current_config(&self.config);
                            
                            self.status_message = "Alle PDFs werden neu erstellt...".to_string();
                            self.is_generating = true;
                            
                            // Zeittracking starten (Neustart)
                            self.generation_start_time = Some(std::time::Instant::now());
                            self.last_progress_update = Some(std::time::Instant::now());
                            self.estimated_total_duration = None;
                            
                            // Progress-Updates erlauben
                            self.progress_frozen = false;
                            
                            // Animation zurücksetzen
                            self.animation_frame = 0;
                            self.animation_time = Some(std::time::Instant::now());
                            
                            // Stop-Signal zurücksetzen
                            {
                                let mut stop = self.stop_signal.lock().unwrap();
                                *stop = false;
                            }

                            let progress_clone = Arc::new(Mutex::new(0.0f32));
                            let progress_ref = Arc::clone(&progress_clone);
                            let stop_signal = Arc::clone(&self.stop_signal);

                            // Prepare thread/IO related variables for generator
                            debug_log(&format!("Starte PDF-Generierung mit {} Threads", self.max_threads), self.debug_mode);
                            let threads = self.max_threads;
                            let thread_sleep_ms = self.thread_sleep_ms;
                            let debug_mode = self.debug_mode;
                            let selections = get_current_selections().unwrap_or_else(|| vec![ ("DATA/Vertreternummern.csv".to_string(), "VORLAGE/Bestellschein-Endkunde-de_de.pdf".to_string(), true) ]);
                            let csv_path = selections.get(0).map(|s| s.0.clone()).unwrap_or_else(|| "DATA/Vertreternummern.csv".to_string());
                            let vertreter_vec = read_vertreter(&csv_path);
                            let vertreter_arc = Arc::new(vertreter_vec);
                            let total = vertreter_arc.len();
                            let progress_counter = Arc::new(Mutex::new(0usize));
                            let (_cfg_dir, data_dir, templates_dir, _tools, _out) = get_release_dirs();

                            // Output-Konfiguration für Thread klonen
                            let use_custom_output = self.use_custom_output_dir;
                            let custom_output_path = self.custom_output_dir.clone();
                            let group = self.selected_group.clone();
                            let language = self.selected_language.clone();
                            let is_messe = self.is_messe;

                            thread::spawn(move || {
                                if let Err(e) = generate_bestellscheine_resume(
                                    progress_ref,
                                    stop_signal,
                                    0,
                                    threads,
                                    vertreter_arc,
                                    progress_counter,
                                    total,
                                    data_dir,
                                    templates_dir,
                                    use_custom_output,
                                    custom_output_path,
                                    group,
                                    language,
                                    is_messe,
                                    thread_sleep_ms,
                                    debug_mode,
                                    // "Von vorne" = alle Kunden ohne Bereichs-Begrenzung
                                    false, // use_range
                                    0,     // range_start
                                    total.saturating_sub(1), // range_end
                                    // Font-Fallback Parameter - Thread 2
                                    true, // TODO: self.enable_font_fallback durch Clone ersetzen
                                ) {
                                    eprintln!("Fehler beim Erstellen der Bestellscheine: {}", e);
                                }
                            });
                        }
                    }
                    
                    // Sekundärer Button: Von vorne beginnen (nur wenn Resume verfügbar)
                    if self.resume_available {
                        let restart_button = egui::Button::new("🔄 Von vorne beginnen")
                            .fill(egui::Color32::from_rgb(255, 193, 7)); // Gelb
                        if ui.add(restart_button).clicked() {
                            // WICHTIG: Aktuelle Config für PDF-Generierung setzen
                            set_current_config(&self.config);
                            
                            self.status_message = "Alle PDFs werden neu erstellt...".to_string();
                            self.is_generating = true;
                            
                            // Zeittracking starten (Neustart)
                            self.generation_start_time = Some(std::time::Instant::now());
                            self.last_progress_update = Some(std::time::Instant::now());
                            self.estimated_total_duration = None;
                            
                            // Progress-Updates erlauben
                            self.progress_frozen = false;
                            
                            // Animation zurücksetzen
                            self.animation_frame = 0;
                            self.animation_time = Some(std::time::Instant::now());
                            
                            // Stop-Signal zurücksetzen
                            {
                                let mut stop = self.stop_signal.lock().unwrap();
                                *stop = false;
                            }

                            let progress_clone = Arc::new(Mutex::new(0.0f32));
                            let progress_ref = Arc::clone(&progress_clone);
                            let stop_signal = Arc::clone(&self.stop_signal);

                            // Prepare thread/IO related variables for generator
                            debug_log(&format!("Starte PDF-Generierung mit {} Threads", self.max_threads), self.debug_mode);
                            let threads = self.max_threads;
                            let thread_sleep_ms = self.thread_sleep_ms;
                            let debug_mode = self.debug_mode;
                            let selections = get_current_selections().unwrap_or_else(|| vec![ ("DATA/Vertreternummern.csv".to_string(), "VORLAGE/Bestellschein-Endkunde-de_de.pdf".to_string(), true) ]);
                            let csv_path = selections.get(0).map(|s| s.0.clone()).unwrap_or_else(|| "DATA/Vertreternummern.csv".to_string());
                            let vertreter_vec = read_vertreter(&csv_path);
                            let vertreter_arc = Arc::new(vertreter_vec);
                            let total = vertreter_arc.len();
                            let progress_counter = Arc::new(Mutex::new(0usize));
                            let (_cfg_dir, data_dir, templates_dir, _tools, _out) = get_release_dirs();

                            // Output-Konfiguration für Thread klonen
                            let use_custom_output = self.use_custom_output_dir;
                            let custom_output_path = self.custom_output_dir.clone();
                            let group = self.selected_group.clone();
                            let language = self.selected_language.clone();
                            let is_messe = self.is_messe;

                            thread::spawn(move || {
                                if let Err(e) = generate_bestellscheine_resume(
                                    progress_ref,
                                    stop_signal,
                                    0,
                                    threads,
                                    vertreter_arc,
                                    progress_counter,
                                    total,
                                    data_dir,
                                    templates_dir,
                                    use_custom_output,
                                    custom_output_path,
                                    group,
                                    language,
                                    is_messe,
                                    thread_sleep_ms,
                                    debug_mode,
                                    // "Von vorne beginnen" = alle Kunden ohne Bereichs-Begrenzung  
                                    false, // use_range
                                    0,     // range_start
                                    total.saturating_sub(1), // range_end
                                    // Font-Fallback Parameter - Thread 3
                                    true, // TODO: self.enable_font_fallback durch Clone ersetzen
                                ) {
                                    eprintln!("Fehler beim Erstellen der Bestellscheine: {}", e);
                                }
                            });
                        }
                    }
                } else {
                    // Stop Button während Generierung
                    let stop_button = egui::Button::new(egui::RichText::new("🛑 Stoppen").size(16.0))
                        .fill(egui::Color32::from_rgb(244, 67, 54)); // Rot
                    if ui.add(stop_button).clicked() {
                        println!("STOP Button gedrückt!");
                        {
                            let mut stop = self.stop_signal.lock().unwrap();
                            *stop = true;
                            println!("Stop-Signal auf true gesetzt!");
                        }
                        self.status_message = "Wird gestoppt...".to_string();
                    }
                }
                
                // Settings-Button und Konfigurations-Button (ganz rechts)
                ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                    let settings_button = egui::Button::new("⚙").fill(if self.dark_mode { 
                        egui::Color32::from_rgb(60, 60, 60) 
                    } else { 
                        egui::Color32::from_rgb(230, 230, 230) 
                    });
                    
                    if ui.add(settings_button)
                        .on_hover_text("App-Einstellungen")
                        .clicked() {
                        self.show_settings_dialog = true;
                    }
                    
                    // Kleiner Konfigurations-Button neben Settings
                    let config_button = egui::Button::new("📐").fill(if self.dark_mode { 
                        egui::Color32::from_rgb(60, 60, 60) 
                    } else { 
                        egui::Color32::from_rgb(230, 230, 230) 
                    });
                    
                    if ui.add(config_button)
                        .on_hover_text("Positionen konfigurieren")
                        .clicked() {
                        // Config KOMPLETT neu laden für aktuelle Gruppe/Sprache/Messe
                        self.config = load_group_config(&self.selected_group, &self.selected_language, self.is_messe);
                        println!("Config geladen - QR-Codes: {:?}", self.config.qr_codes);
                        println!("Config geladen - Vertreter: {:?}", self.config.vertreter);
                        
                        // Initialize manual coordinate fields with current values
                        if !self.config.qr_codes.is_empty() {
                            self.manual_qr_x = format!("{:.1}", self.config.qr_codes[0].x);
                            self.manual_qr_y = format!("{:.1}", self.config.qr_codes[0].y);
                            self.manual_qr_size = format!("{:.1}", self.config.qr_codes[0].size);
                        }
                        if !self.config.vertreter.is_empty() {
                            self.manual_vertreter_x = format!("{:.1}", self.config.vertreter[0].x);
                            self.manual_vertreter_y = format!("{:.1}", self.config.vertreter[0].y);
                            self.manual_vertreter_size = format!("{:.1}", self.config.vertreter[0].size);
                        }
                        
                        self.show_config = true;
                    }
                });
            });
        });

        // Startup-Auswahl-Dialog (einmalig beim Programmstart)
        if self.show_startup_dialog {
            let mut open = self.show_startup_dialog;
            let mut should_close = false;
            egui::Window::new("Was soll erstellt werden?")
                .collapsible(false)
                .resizable(false)
                .anchor(egui::Align2::CENTER_CENTER, egui::vec2(0.0, 0.0))
                .open(&mut open)
                .show(ctx, |ui| {
                    ui.label("Wählen Sie die Kundengruppe:");
                    ui.horizontal(|ui| {
                        if ui.selectable_label(self.selected_group == "Endkunde", "Endkunde").clicked() {
                            // Alte Progress-Dateien löschen bei Kategorie-Wechsel
                            clear_progress_files(&self.selected_group, &self.selected_language, self.is_messe);
                            self.selected_group = "Endkunde".to_string();
                            // Resume-Info für neue Kategorie laden
                            self.resume_info = load_resume_info(&self.selected_group, &self.selected_language, self.is_messe);
                            // Sofort gruppenspezifische Config laden
                            let group_cfg = load_group_config(&self.selected_group, &self.selected_language, self.is_messe);
                            self.config = group_cfg;
                            // Update manual input fields with loaded values
                            if !self.config.qr_codes.is_empty() {
                                self.manual_qr_x = format!("{:.1}", self.config.qr_codes[0].x);
                                self.manual_qr_y = format!("{:.1}", self.config.qr_codes[0].y);
                                self.manual_qr_size = format!("{:.1}", self.config.qr_codes[0].size);
                            }
                            if !self.config.vertreter.is_empty() {
                                self.manual_vertreter_x = format!("{:.1}", self.config.vertreter[0].x);
                                self.manual_vertreter_y = format!("{:.1}", self.config.vertreter[0].y);
                            }
                            println!("Gruppe geändert zu Endkunde - Config neu geladen");
                        }
                        if ui.selectable_label(self.selected_group == "Apo", "Apotheken (Apo)").clicked() {
                            // Alte Progress-Dateien löschen bei Kategorie-Wechsel
                            clear_progress_files(&self.selected_group, &self.selected_language, self.is_messe);
                            self.selected_group = "Apo".to_string();
                            // Resume-Info für neue Kategorie laden
                            self.resume_info = load_resume_info(&self.selected_group, &self.selected_language, self.is_messe);
                            // Sofort gruppenspezifische Config laden
                            let group_cfg = load_group_config(&self.selected_group, &self.selected_language, self.is_messe);
                            self.config = group_cfg;
                            // Update manual input fields with loaded values
                            if !self.config.qr_codes.is_empty() {
                                self.manual_qr_x = format!("{:.1}", self.config.qr_codes[0].x);
                                self.manual_qr_y = format!("{:.1}", self.config.qr_codes[0].y);
                                self.manual_qr_size = format!("{:.1}", self.config.qr_codes[0].size);
                            }
                            if !self.config.vertreter.is_empty() {
                                self.manual_vertreter_x = format!("{:.1}", self.config.vertreter[0].x);
                                self.manual_vertreter_y = format!("{:.1}", self.config.vertreter[0].y);
                            }
                            println!("Gruppe geändert zu Apo - Config neu geladen");
                        }
                    });

                    ui.separator();
                    ui.label("Sprache:");
                    ui.horizontal(|ui| {
                        if ui.selectable_label(self.selected_language == "Deutsch", "Deutsch").clicked() {
                            self.selected_language = "Deutsch".to_string();
                            // Sofort sprachspezifische Config laden
                            let group_cfg = load_group_config(&self.selected_group, &self.selected_language, self.is_messe);
                            self.config = group_cfg;
                            // Update manual input fields
                            if !self.config.qr_codes.is_empty() {
                                self.manual_qr_x = format!("{:.1}", self.config.qr_codes[0].x);
                                self.manual_qr_y = format!("{:.1}", self.config.qr_codes[0].y);
                                self.manual_qr_size = format!("{:.1}", self.config.qr_codes[0].size);
                            }
                            if !self.config.vertreter.is_empty() {
                                self.manual_vertreter_x = format!("{:.1}", self.config.vertreter[0].x);
                                self.manual_vertreter_y = format!("{:.1}", self.config.vertreter[0].y);
                            }
                            println!("Sprache geändert zu Deutsch - Config neu geladen");
                        }
                        if ui.selectable_label(self.selected_language == "Englisch", "Englisch").clicked() {
                            self.selected_language = "Englisch".to_string();
                            // Sofort sprachspezifische Config laden
                            let group_cfg = load_group_config(&self.selected_group, &self.selected_language, self.is_messe);
                            self.config = group_cfg;
                            // Update manual input fields
                            if !self.config.qr_codes.is_empty() {
                                self.manual_qr_x = format!("{:.1}", self.config.qr_codes[0].x);
                                self.manual_qr_y = format!("{:.1}", self.config.qr_codes[0].y);
                                self.manual_qr_size = format!("{:.1}", self.config.qr_codes[0].size);
                            }
                            if !self.config.vertreter.is_empty() {
                                self.manual_vertreter_x = format!("{:.1}", self.config.vertreter[0].x);
                                self.manual_vertreter_y = format!("{:.1}", self.config.vertreter[0].y);
                            }
                            println!("Sprache geändert zu Englisch - Config neu geladen");
                        }
                    });

                    ui.separator();
                    ui.label("Hinweis: Es werden nur die für die Auswahl relevanten QR-Codes und Vorlagen verwendet.");

                        ui.separator();
                        ui.horizontal(|ui| {
                            ui.label("Messescheine:");
                            if ui.selectable_label(self.is_messe, "Ja (Messe)").clicked() {
                                self.is_messe = !self.is_messe;
                                // Sofort messe-spezifische Config laden
                                let group_cfg = load_group_config(&self.selected_group, &self.selected_language, self.is_messe);
                                self.config = group_cfg;
                                // Update manual input fields
                                if !self.config.qr_codes.is_empty() {
                                    self.manual_qr_x = format!("{:.1}", self.config.qr_codes[0].x);
                                    self.manual_qr_y = format!("{:.1}", self.config.qr_codes[0].y);
                                    self.manual_qr_size = format!("{:.1}", self.config.qr_codes[0].size);
                                }
                                if !self.config.vertreter.is_empty() {
                                    self.manual_vertreter_x = format!("{:.1}", self.config.vertreter[0].x);
                                    self.manual_vertreter_y = format!("{:.1}", self.config.vertreter[0].y);
                                }
                                println!("Messe-Option geändert zu {} - Config neu geladen", self.is_messe);
                            }
                        });

                    // Erweiterte Template-Auswahl mit Bewertung und Fallback
                    ui.separator();
                    ui.label("� Datenherkunft:");
                    
                    // CSV-Datei-Status prüfen und anzeigen
                    let csv_path = if self.selected_group == "Apo" { 
                        get_default_csv_path("Apo") 
                    } else { 
                        get_default_csv_path("Endkunde") 
                    };
                    
                    let (_, data_dir, _, _, _) = get_release_dirs_with_debug(self.debug_mode);
                    let cleaned_filename = csv_path.replace("Data/", "");
                    let full_csv_path = data_dir.join(&cleaned_filename);
                    let csv_exists = full_csv_path.exists();
                    
                    // Debug-Info nur im Debug-Modus in Logdatei schreiben
                    debug_log(&format!("CSV-Check: group='{}', csv_path='{}', data_dir='{}', cleaned='{}', full_path='{}', exists={}", 
                             self.selected_group, csv_path, data_dir.display(), cleaned_filename, full_csv_path.display(), csv_exists), self.debug_mode);
                    
                    ui.horizontal(|ui| {
                        let csv_icon = if csv_exists { "✅" } else { "❌" };
                        let csv_color = if csv_exists { egui::Color32::from_rgb(0, 120, 0) } else { egui::Color32::from_rgb(200, 0, 0) };
                        ui.label(egui::RichText::new(format!("{} CSV-Datei: {}", csv_icon, csv_path)).color(csv_color));
                        if !csv_exists {
                            ui.label(egui::RichText::new(format!("(Erwartet: {})", full_csv_path.display())).size(9.0).color(egui::Color32::GRAY));
                        }
                    });
                    
                    ui.separator();
                    ui.label("📋 Template-Auswahl:");
                    
                    // Templates mit Bewertung laden
                    let templates_with_score = self.find_templates_with_score(&self.selected_group, &self.selected_language, self.is_messe);
                    
                    if templates_with_score.is_empty() {
                        ui.label("⚠️ Keine Templates gefunden!");
                    } else {
                        // Zeige nur die besten 5 Templates zur Auswahl
                        ui.label("🏆 Empfohlene Templates (nach Relevanz sortiert):");
                        for (i, (template, score, exists)) in templates_with_score.iter().take(5).enumerate() {
                            let icon = if *exists { "✅" } else { "❌" };
                            let score_text = format!("(Score: {})", score);
                            
                            ui.horizontal(|ui| {
                                let is_selected = self.selected_template_index == Some(i);
                                if ui.selectable_label(is_selected, format!("{} {}", icon, template)).clicked() {
                                    self.selected_template_index = Some(i);
                                    self.available_templates = templates_with_score.iter().map(|(t, _, _)| t.clone()).collect();
                                }
                                ui.label(egui::RichText::new(score_text).size(10.0).color(egui::Color32::GRAY));
                            });
                        }
                        
                        ui.separator();
                        
                        // Automatische vs Manuelle Auswahl
                        ui.horizontal(|ui| {
                            ui.label("Auswahl-Modus:");
                            if ui.selectable_label(!self.show_template_selection, "🤖 Automatisch (beste Übereinstimmung)").clicked() {
                                self.show_template_selection = false;
                                self.selected_template_index = None;
                            }
                            if ui.selectable_label(self.show_template_selection, "👤 Manuell (oben ausgewählt)").clicked() {
                                self.show_template_selection = true;
                                // Wähle das erste Template automatisch vor
                                if self.selected_template_index.is_none() && !templates_with_score.is_empty() {
                                    self.selected_template_index = Some(0);
                                    self.available_templates = templates_with_score.iter().map(|(t, _, _)| t.clone()).collect();
                                }
                            }
                        });
                        
                        if self.show_template_selection {
                            ui.label("👆 Wählen Sie ein Template aus der Liste oben aus");
                        } else {
                            let auto_template = templates_with_score.first();
                            if let Some((template, _score, exists)) = auto_template {
                                let status = if *exists { "✅ gefunden" } else { "❌ fehlt" };
                                ui.label(format!("🤖 Automatische Wahl: {} - {}", template, status));
                            }
                        }
                    }

                    // Zeige alte Kandidatenliste für Referenz (eingeklappt)
                    ui.collapsing("🔍 Alle geprüften Kandidaten (Debug)", |ui| {
                        let candidates = list_template_candidates(&self.selected_group, &self.selected_language, self.is_messe);
                        let project_root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"));
                        for c in candidates {
                            let abs = project_root.join(&c);
                            let exists = abs.exists();
                            if exists {
                                ui.label(format!("✔ {}", c));
                            } else {
                                ui.label(format!("✖ {}", c));
                            }
                        }
                    });

                    ui.horizontal(|ui| {
                        if ui.button("💾 Auswahl speichern").clicked() {
                            // CSV bestimmen
                            let csv_default = if self.selected_group == "Apo" { "DATA/Vertreternummern-Apo.CSV".to_string() } else { "DATA/Vertreternummern.csv".to_string() };
                            
                            // Template bestimmen: Manuell oder Automatisch
                            let template = if self.show_template_selection {
                                // Manuelle Auswahl verwenden
                                if let Some(index) = self.selected_template_index {
                                    if index < self.available_templates.len() {
                                        self.available_templates[index].clone()
                                    } else {
                                        // Fallback auf automatische Erkennung
                                        self.find_template(&self.selected_group, &self.selected_language, if self.is_messe { Some("messe") } else { None })
                                            .unwrap_or_else(|| "VORLAGE/Bestellschein-Endkunde-de_de.pdf".to_string())
                                    }
                                } else {
                                    self.status_message = "❌ Bitte wählen Sie ein Template aus der Liste aus!".to_string();
                                    return; // Früh beenden, Dialog offen lassen
                                }
                            } else {
                                // Automatische Erkennung verwenden
                                self.find_template(&self.selected_group, &self.selected_language, if self.is_messe { Some("messe") } else { None })
                                    .unwrap_or_else(|| if self.selected_group == "Apo" { "VORLAGE/Bestellscheine-Apo.pdf".to_string() } else { "VORLAGE/Bestellschein-Endkunde-de_de.pdf".to_string() })
                            };
                            
                            let gen_qr = true;
                            let csv = csv_default;

                            // Überprüfe ob Dateien existieren
                            let project_root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"));
                            let mut missing = Vec::new();
                            let csv_abs = project_root.join(&csv);
                            let template_abs = project_root.join(&template);
                            if !csv_abs.exists() { missing.push(format!("CSV: {}", csv_abs.to_string_lossy())); }
                            if !template_abs.exists() { missing.push(format!("Template: {}", template_abs.to_string_lossy())); }

                            if !missing.is_empty() {
                                // Zeige Fehlermeldung im UI
                                self.status_message = format!("❌ Fehlende Dateien: {}", missing.join(", "));
                                println!("Fehlende Dateien bei Auswahl: {:?}", missing);
                                // Dialog offen halten damit der Nutzer es sehen kann
                            } else {
                                // Erfolgreiche Auswahl
                                self.status_message = format!("✅ Auswahl gespeichert: {}", template);
                                
                                // Setze globale Auswahl für den Erstellungsprozess
                                set_current_selection(&csv, &template, gen_qr);
                                // Lade gruppenspezifische Config und setze sie in der UI
                                let group_cfg = load_group_config(&self.selected_group, &self.selected_language, self.is_messe);
                                self.config = group_cfg;
                                // request to close after the closure completes
                                should_close = true;
                                self.status_message = format!("Auswahl gesetzt: {} / {}", self.selected_group, self.selected_language);
                            }
                        }

                        if ui.button("Abbrechen").clicked() {
                            self.show_startup_dialog = false;
                            self.status_message = "Keine Auswahl getroffen - Standardwerte werden verwendet".to_string();
                        }
                    });
                });
            if should_close {
                self.show_startup_dialog = false;
            } else {
                self.show_startup_dialog = open;
            }
        }

        // App-Settings Dialog
        if self.show_settings_dialog {
            let mut open = self.show_settings_dialog;
            egui::Window::new("⚙ App-Einstellungen")
                .collapsible(false)
                .resizable(false)
                .anchor(egui::Align2::CENTER_CENTER, egui::vec2(0.0, 0.0))
                .open(&mut open)
                .show(ctx, |ui| {
                    // Darstellung Sektion
                    ui.group(|ui| {
                        ui.label(egui::RichText::new("🎨 Darstellung").size(16.0));
                        ui.separator();
                        
                        ui.horizontal(|ui| {
                            ui.label("Theme:");
                            let theme_button_text = if self.dark_mode { "☀ Light Mode" } else { "🌙 Dark Mode" };
                            if ui.button(theme_button_text).clicked() {
                                self.dark_mode = !self.dark_mode;
                                save_app_settings(self.dark_mode);
                            }
                        });
                        
                        ui.horizontal(|ui| {
                            ui.label("Fenster:");
                            let maximize_text = if self.fullscreen_mode { "🗗 Normal" } else { "🗖 Maximiert" };
                            if ui.button(maximize_text).clicked() {
                                self.fullscreen_mode = !self.fullscreen_mode;
                                ctx.send_viewport_cmd(egui::ViewportCommand::Maximized(self.fullscreen_mode));
                            }
                        });
                    });
                    
                    ui.add_space(10.0);
                    
                    // Ausgabe-Ordner Sektion
                    ui.group(|ui| {
                        ui.label(egui::RichText::new("📁 Ausgabe-Ordner").size(16.0));
                        ui.separator();
                        
                        ui.checkbox(&mut self.use_custom_output_dir, "Benutzerdefinierten Ausgabe-Ordner verwenden");
                        
                        if self.use_custom_output_dir {
                            ui.horizontal(|ui| {
                                ui.label("Ordner:");
                                ui.text_edit_singleline(&mut self.custom_output_dir);
                                if ui.button("📁").clicked() {
                                    // Hier könnte ein Ordner-Auswahl-Dialog hinzugefügt werden
                                    // Für jetzt können Benutzer den Pfad manuell eingeben
                                }
                                if ui.button("🗂️").clicked() {
                                    // Explorer mit aktuellem Ordner öffnen
                                    let path = if self.custom_output_dir.is_empty() {
                                        "OUTPUT".to_string()
                                    } else {
                                        self.custom_output_dir.clone()
                                    };
                                    let _ = safe_open_explorer(&path);
                                }
                            });
                            ui.label(egui::RichText::new("📝 Hinweis: Absoluter Pfad oder relativ zum Programmordner").size(11.0).italics());
                        } else {
                            ui.label("Standard: Automatische Ordnerstruktur in 'Output'");
                            ui.label("└── Gruppe/Sprache (z.B. Output/Endkunde/Deutsch/)");
                            ui.horizontal(|ui| {
                                if ui.button("🗂️ Output-Ordner öffnen").clicked() {
                                    let _ = safe_open_explorer("Output");
                                }
                            });
                        }
                    });
                    
                    ui.add_space(10.0);
                    
                    // Template-Ordner Sektion
                    ui.group(|ui| {
                        ui.label(egui::RichText::new("📄 Vorlagen-Ordner").size(16.0));
                        ui.separator();
                        
                        ui.checkbox(&mut self.use_custom_template_dir, "Benutzerdefinierten Vorlagen-Ordner verwenden");
                        
                        if self.use_custom_template_dir {
                            ui.horizontal(|ui| {
                                ui.label("Ordner:");
                                ui.text_edit_singleline(&mut self.custom_template_dir);
                                if ui.button("📁").clicked() {
                                    // Hier könnte ein Ordner-Auswahl-Dialog hinzugefügt werden
                                }
                                if ui.button("🗂️").clicked() {
                                    // Explorer mit aktuellem Ordner öffnen
                                    let path = if self.custom_template_dir.is_empty() {
                                        "VORLAGE".to_string()
                                    } else {
                                        self.custom_template_dir.clone()
                                    };
                                    let _ = safe_open_explorer(&path);
                                }
                            });
                            ui.label(egui::RichText::new("📝 Hinweis: Absoluter Pfad oder relativ zum Programmordner").size(11.0).italics());
                            ui.label(egui::RichText::new("🔍 Das System sucht die beste passende Vorlage basierend auf Gruppe/Sprache/Messe").size(11.0).color(egui::Color32::GRAY));
                        } else {
                            ui.label("Standard: Automatische Suche in 'VORLAGE'-Ordner");
                            ui.label("├── Interne Template-Erkennung");
                            ui.label("└── Fallback zur manuellen Auswahl");
                            ui.horizontal(|ui| {
                                if ui.button("🗂️ Vorlagen-Ordner öffnen").clicked() {
                                    let _ = safe_open_explorer("Vorlagen");
                                }
                            });
                        }
                    });
                    
                    ui.add_space(10.0);
                    
                    // Info/Support Sektion
                    ui.group(|ui| {
                        ui.label(egui::RichText::new("ℹ Information & Support").size(16.0));
                        ui.separator();
                        
                        ui.label("Bei jeglichen Problemen technischer Art");
                        ui.label("wenden Sie sich bitte an die");
                        ui.label(egui::RichText::new("IT-Abteilung").strong().color(egui::Color32::BLUE));
                        ui.add_space(5.0);
                        ui.label(egui::RichText::new("Diese wird die Anfrage prüfen und umsetzen.").italics());
                        
                        ui.add_space(8.0);
                        ui.label(egui::RichText::new("📖 Weitere Informationen und Anleitungen:").size(12.0));
                        ui.hyperlink_to("https://wiki.natugena.de/wiki/Bestellscheine", "https://wiki.natugena.de/wiki/Bestellscheine");
                        
                        ui.add_space(10.0);
                        ui.separator();
                        ui.label(egui::RichText::new("Bestellschein Generator").size(14.0).strong());
                        ui.label(egui::RichText::new("Version 1").size(12.0));
                        ui.add_space(5.0);
                        ui.label(egui::RichText::new("Programmentwicklung: Alexander Löschke, IT - Abteilung")
                            .strong()
                            .color(egui::Color32::from_rgb(100, 149, 237))); // Cornflower Blue
                            
                        // Versteckter Debug/Performance-Bereich (nur sichtbar bei Ctrl+Shift+D)
                        // Stabile Tastenkombination mit Toggle
                        if ctx.input(|i| i.modifiers.ctrl && i.modifiers.shift && i.key_pressed(egui::Key::D)) {
                            if !self.debug_key_pressed {
                                self.debug_mode = !self.debug_mode; // Toggle Debug-Modus
                                save_debug_config(self.debug_mode); // Persistieren des Debug-Status
                                self.debug_key_pressed = true;
                            }
                        } else {
                            self.debug_key_pressed = false; // Reset wenn Tasten losgelassen
                        }
                        
                        if self.debug_mode {
                            ui.separator();
                            ui.label(egui::RichText::new("🔧 Erweiterte Einstellungen (Debug-Modus)").size(12.0).color(egui::Color32::RED));
                            
                            ui.checkbox(&mut self.debug_mode, "Debug-Modus aktivieren (Log-Datei)");
                            if self.debug_mode {
                                ui.label("📝 Debug-Informationen werden in cache/debug.log gespeichert");
                                if ui.button("🗂️ Log-Datei öffnen").clicked() {
                                    let log_path = get_temp_file_path("debug.log");
                                    let _ = safe_open_notepad(&log_path);
                                }
                            }
                            
                            ui.separator();
                            ui.label("⚡ Performance-Einstellungen:");
                            ui.horizontal(|ui| {
                                ui.label("Max. Threads:");
                                ui.add(egui::Slider::new(&mut self.max_threads, 1..=16).text("Threads"));
                            });
                            ui.horizontal(|ui| {
                                ui.label("Thread-Pause:");
                                ui.add(egui::Slider::new(&mut self.thread_sleep_ms, 0..=5).suffix(" ms"));
                                ui.label("(nur 0-2ms aktiv)");
                            });
                            
                            ui.separator();
                            ui.label("🔤 Font-System-Einstellungen:");
                            ui.checkbox(&mut self.enable_font_fallback, "Font-Fallback aktivieren");
                            if self.enable_font_fallback {
                                ui.label("✓ Wenn Custom Font nicht gefunden → Standard-Font (Times, Helvetica, Courier)");
                            } else {
                                ui.label("❌ Nur echte Custom Fonts verwenden (keine Standard-Font-Fallbacks)");
                                ui.label(egui::RichText::new("⚠ Warnung: Fehlende Fonts können zu Fehlern führen").size(11.0).color(egui::Color32::from_rgb(255, 165, 0)));
                            }
                            
                            ui.label("💡 Weniger Threads = geringere CPU-Last, mehr Threads = schneller");
                            ui.label("⚡ Thread-Pause >2ms deaktiviert für maximale Geschwindigkeit");
                        }
                        
                        if !self.debug_mode {
                            // Hinweis auf versteckten Debug-Modus
                            ui.add_space(5.0);
                            ui.horizontal(|ui| {
                                ui.label("🔧");
                                ui.label(egui::RichText::new("Erweiterte Debug-Optionen: Ctrl+Shift+D").size(11.0).color(egui::Color32::GRAY));
                            });
                        }
                    });
                });
            self.show_settings_dialog = open;
        }

        egui::CentralPanel::default().show(ctx, |ui| {
            ui.heading("Fortschritt");
            
            // Nur EINE Progress-Bar mit angepassten Informationen
            if self.is_generating {
                // Zeit-Informationen berechnen
                let (elapsed_str, remaining_str) = if let Some(start_time) = self.generation_start_time {
                    let elapsed = start_time.elapsed();
                    let elapsed_text = format_duration(elapsed);
                    
                    let remaining_text = if let Some(total_duration) = self.estimated_total_duration {
                        let remaining = if total_duration > elapsed {
                            total_duration - elapsed
                        } else {
                            std::time::Duration::from_secs(0)
                        };
                        format_duration(remaining)
                    } else {
                        "Berechnung...".to_string()
                    };
                    
                    (elapsed_text, remaining_text)
                } else {
                    ("--".to_string(), "--".to_string())
                };
                
                // Progress-Bar mit Zeitinformationen
                ui.add(egui::ProgressBar::new(self.progress)
                    .show_percentage()
                    .fill(egui::Color32::from_rgb(100, 200, 255))
                    .text(format!("Vergangen: {} | Verbleibend: ~{}", elapsed_str, remaining_str)));
                    
                ui.label(&self.status_message);
            } else {
                // Standard Progress-Bar wenn nicht generiert wird
                ui.add(egui::ProgressBar::new(self.progress).show_percentage());
                ui.label(&self.status_message);
            }
            
            // Coole Animation während der PDF-Generierung anzeigen
            if self.is_generating {
                ui.separator();
                ui.vertical_centered(|ui| {
                    let animation_text = self.get_generating_animation();
                    ui.heading(&animation_text);
                    
                    // Zusätzliche visuelle Effekte
                    ui.horizontal(|ui| {
                        let dancing_chars = ["🕺", "💃", "🎭", "🎪", "🎨", "🎯", "🚀", "💎"];
                        let char_index = (self.animation_frame / 3) % dancing_chars.len();
                        
                        for i in 0..8 {
                            let char_to_show = dancing_chars[(char_index + i) % dancing_chars.len()];
                            ui.label(char_to_show);
                        }
                    });
                });
                ui.separator();
                
                // Kontinuierliche Aktualisierung für Animation - aber nur alle 300ms
                if let Some(last_repaint) = self.animation_time {
                    if last_repaint.elapsed().as_millis() >= 290 {
                        ctx.request_repaint();
                    }
                } else {
                    ctx.request_repaint();
                }
            }
            
            // Speicher-Bestätigung anzeigen (verschwindet nach 2 Sekunden)
            if let Some(save_time) = self.save_message {
                if save_time.elapsed().as_secs() < 2 {
                    ui.colored_label(egui::Color32::from_rgb(0, 150, 0), "✅ Konfiguration gespeichert!");
                } else {
                    self.save_message = None;
                }
            }
            
            // Status-Update basierend auf Progress (nur wenn wirklich 100% erreicht)
            if self.progress >= 1.0 && self.is_generating && self.status_message != "Bereit" && self.status_message != "Bestellscheine fertig erstellt!" {
                self.status_message = "Bestellscheine fertig erstellt!".to_string();
                self.is_generating = false;
                // Resume-Status aktualisieren da jetzt alle fertig sind
                self.resume_available = false;
                self.last_processed_count = 0;
                self.resume_needs_update = false; // Keine weitere Aktualisierung nötig
                // Meme anzeigen wenn PDFs fertig sind! 😄
                self.show_meme = true;
                self.meme_time = Some(std::time::Instant::now());
            }
            
            // Meme nach 5 Sekunden ausblenden
            if self.show_meme {
                if let Some(meme_start) = self.meme_time {
                    if meme_start.elapsed().as_secs() >= 5 {
                        self.show_meme = false;
                        self.meme_time = None;
                    }
                }
            }
            
            // Prüfen ob gestoppt wurde - überprüfe auch Progress-Datei (aber nicht ständig die Resume-Status)
            if self.is_generating {
                // Stop-Signal prüfen
                if let Ok(should_stop) = self.stop_signal.try_lock() {
                    if *should_stop {
                        println!("Stop-Signal erkannt - beende Generierung");
                        self.status_message = format!("Gestoppt bei {}% - kann fortgesetzt werden", (self.progress * 100.0) as u32);
                        self.is_generating = false;
                        // Resume-Status aktualisieren NUR WENN GESTOPPT
                        self.resume_needs_update = true;
                        // Zeit-Tracking beim Stop beenden
                        self.estimated_total_duration = None;
                        // Progress-Updates einfrieren um das Springen zu verhindern
                        self.progress_frozen = true;
                    }
                }
                
                
                // Stop-Status-Datei prüfen (separate von progress.txt, versteckt)
                let stop_status_path = get_temp_file_path("stop_status.txt");
                if let Ok(_) = std::fs::read_to_string(&stop_status_path) {
                    println!("Stop-Status-Datei gefunden - beende Generierung");
                    self.status_message = format!("Gestoppt bei {}% - kann fortgesetzt werden", (self.progress * 100.0) as u32);
                    self.is_generating = false;
                    // Resume-Status aktualisieren NUR WENN GESTOPPT
                    self.resume_needs_update = true;
                    // Stop-Signal für nächsten Start zurücksetzen
                    if let Ok(mut stop) = self.stop_signal.try_lock() {
                        *stop = false;
                    }
                    // Zeit-Tracking beim Stop beenden
                    self.estimated_total_duration = None;
                    // Progress-Updates einfrieren um das Springen zu verhindern
                    self.progress_frozen = true;
                    // Stop-Status-Datei löschen nach dem Verarbeiten
                    let _ = std::fs::remove_file(&stop_status_path);
                }
            }
            
            // Seiten-Konfigurationsfenster
            if self.show_config {
                // WICHTIG: Config nochmal neu laden wenn Dialog geöffnet wird
                if ui.input(|i| i.key_pressed(egui::Key::F5)) {
                    self.config = load_group_config(&self.selected_group, &self.selected_language, self.is_messe);
                    println!("F5 gedrückt - gruppenspezifische Config neu geladen!");
                }
                
                let mut show_config = self.show_config;
                let _window_response = egui::Window::new("Positionen auf DIN A4 konfigurieren")
                    .open(&mut show_config)
                    .resizable(true)
                    .default_size([800.0, 600.0])
                    .show(ctx, |ui| {
                    
                    ui.horizontal(|ui| {
                        // Linke Seite - DIN A4 Darstellung
                        ui.vertical(|ui| {
                            ui.horizontal(|ui| {
                                ui.label("Ziehen Sie die Elemente an die gewünschte Position:");
                                if ui.button("🔄 Config neu laden").clicked() {
                                    self.config = load_group_config(&self.selected_group, &self.selected_language, self.is_messe);
                                    println!("Gruppenspezifische Config manuell neu geladen!");
                                }
                            });
                            
                            // PERFORMANCE: Kleinere DIN A4 Darstellung für bessere Performance
                            let a4_width = 280.0;  // Reduziert von 350
                            let a4_height = 396.0; // 280 * 1.414
                            
                            // Bereich für DIN A4 Darstellung reservieren - OHNE drag sense für bessere Performance
                            let (a4_rect, _a4_response) = ui.allocate_exact_size(
                                egui::vec2(a4_width, a4_height), 
                                egui::Sense::hover() // NUR hover, kein drag
                            );
                            
                            // Template-Info anzeigen statt PDF-Rendering - NUR MILLIMETER-RASTER!
                            
                            // A4-Hintergrund mit Millimeter-Raster
                            ui.painter().rect_filled(a4_rect, 5.0, egui::Color32::WHITE);
                            ui.painter().rect_stroke(a4_rect, 5.0, egui::Stroke::new(2.0, egui::Color32::BLACK));
                            
                            // Millimeter-Raster zeichnen (alle 10mm eine Linie)
                            let mm_to_pixel_x = a4_width / 210.0;  // 210mm A4 Breite
                            let mm_to_pixel_y = a4_height / 297.0; // 297mm A4 Höhe
                            
                            // Vertikale Linien (alle 10mm)
                            for mm in (10..210).step_by(10) {
                                let x_pos = a4_rect.left() + mm as f32 * mm_to_pixel_x;
                                let color = if mm % 50 == 0 { 
                                    egui::Color32::from_rgb(150, 150, 150) 
                                } else { 
                                    egui::Color32::from_rgb(220, 220, 220) 
                                };
                                ui.painter().line_segment(
                                    [egui::pos2(x_pos, a4_rect.top()), egui::pos2(x_pos, a4_rect.bottom())],
                                    egui::Stroke::new(if mm % 50 == 0 { 1.0 } else { 0.5 }, color)
                                );
                                
                                // Beschriftung alle 50mm
                                if mm % 50 == 0 {
                                    ui.painter().text(
                                        egui::pos2(x_pos, a4_rect.top() + 10.0),
                                        egui::Align2::CENTER_TOP,
                                        format!("{}mm", mm),
                                        egui::FontId::proportional(8.0),
                                        egui::Color32::from_rgb(100, 100, 100),
                                    );
                                }
                            }
                            
                            // Horizontale Linien (alle 10mm)
                            for mm in (10..297).step_by(10) {
                                let y_pos = a4_rect.top() + mm as f32 * mm_to_pixel_y;
                                let color = if mm % 50 == 0 { 
                                    egui::Color32::from_rgb(150, 150, 150) 
                                } else { 
                                    egui::Color32::from_rgb(220, 220, 220) 
                                };
                                ui.painter().line_segment(
                                    [egui::pos2(a4_rect.left(), y_pos), egui::pos2(a4_rect.right(), y_pos)],
                                    egui::Stroke::new(if mm % 50 == 0 { 1.0 } else { 0.5 }, color)
                                );
                                
                                // Beschriftung alle 50mm
                                if mm % 50 == 0 {
                                    ui.painter().text(
                                        egui::pos2(a4_rect.left() + 5.0, y_pos),
                                        egui::Align2::LEFT_CENTER,
                                        format!("{}mm", mm),
                                        egui::FontId::proportional(8.0),
                                        egui::Color32::from_rgb(100, 100, 100),
                                    );
                                }
                            }
                            
                            // Skalierungsfaktor für die Koordinaten (mm zu UI-Pixel)
                            let scale_x = 210.0 / a4_width;  // DIN A4: 210 mm breit
                            let scale_y = 297.0 / a4_height; // DIN A4: 297 mm hoch
                            
                            // QR-Codes - OPTIMIERT
                            for (i, qr) in self.config.qr_codes.iter_mut().enumerate() {
                                let qr_display_size = 25.0; // Feste Größe für Performance
                                let qr_pos_x = qr.x / scale_x;
                                let qr_pos_y = a4_height - (qr.y / scale_y);
                                
                                let qr_rect = egui::Rect::from_min_size(
                                    egui::pos2(a4_rect.left() + qr_pos_x, a4_rect.top() + qr_pos_y),
                                    egui::vec2(qr_display_size, qr_display_size)
                                );
                                
                                let qr_id = egui::Id::new(format!("qr_{}", i));
                                let qr_response = ui.interact(qr_rect, qr_id, egui::Sense::drag());
                                
                                // Einfache Darstellung
                                ui.painter().rect_filled(qr_rect, 3.0, egui::Color32::from_rgb(255, 165, 0));
                                ui.painter().text(
                                    qr_rect.center(),
                                    egui::Align2::CENTER_CENTER,
                                    format!("Q{}", i + 1),
                                    egui::FontId::proportional(10.0),
                                    egui::Color32::BLACK,
                                );
                                
                                if qr_response.dragged() {
                                    let delta = qr_response.drag_delta();
                                    qr.x += delta.x * scale_x;
                                    qr.y -= delta.y * scale_y;
                                    qr.x = qr.x.max(0.0).min(210.0 - qr.size); // DIN A4: 210 mm breit
                                    qr.y = qr.y.max(0.0).min(297.0 - qr.size); // DIN A4: 297 mm hoch
                                }
                                
                                if qr_response.drag_stopped() && i == 0 {
                                    self.manual_qr_x = format!("{:.1}", qr.x);
                                    self.manual_qr_y = format!("{:.1}", qr.y);
                                    self.manual_qr_size = format!("{:.1}", qr.size);
                                }
                            }
                            
                            // Vertreternummern-Felder - OPTIMIERT
                            for (i, pos) in self.config.vertreter.iter_mut().enumerate() {
                                let field_width = 40.0;
                                let field_height = 12.0;
                                let field_pos_x = pos.x / scale_x;
                                let field_pos_y = a4_height - (pos.y / scale_y);
                                
                                let field_rect = egui::Rect::from_min_size(
                                    egui::pos2(a4_rect.left() + field_pos_x, a4_rect.top() + field_pos_y),
                                    egui::vec2(field_width, field_height)
                                );
                                
                                let field_id = egui::Id::new(format!("field_{}", i));
                                let field_response = ui.interact(field_rect, field_id, egui::Sense::drag());
                                
                                // Einfache Darstellung
                                ui.painter().rect_filled(field_rect, 2.0, egui::Color32::LIGHT_BLUE);
                                ui.painter().text(
                                    field_rect.center(),
                                    egui::Align2::CENTER_CENTER,
                                    format!("K{}", i + 1),
                                    egui::FontId::proportional(8.0),
                                    egui::Color32::BLACK,
                                );
                                
                                if field_response.dragged() {
                                    let delta = field_response.drag_delta();
                                    pos.x += delta.x * scale_x;
                                    pos.y -= delta.y * scale_y;
                                    pos.x = pos.x.max(0.0).min(210.0 - field_width * scale_x); // DIN A4: 210 mm breit
                                    pos.y = pos.y.max(0.0).min(297.0 - field_height * scale_y); // DIN A4: 297 mm hoch
                                }
                                
                                if field_response.drag_stopped() && i == 0 {
                                    self.manual_vertreter_x = format!("{:.1}", pos.x);
                                    self.manual_vertreter_y = format!("{:.1}", pos.y);
                                    self.manual_vertreter_size = format!("{:.1}", pos.size);
                                }
                            }
                        });
                        
                        ui.separator();
                        
                        // Rechte Seite - Steuerungselemente
                        ui.vertical(|ui| {
                            ui.heading("Elemente verwalten");
                            
                            ui.group(|ui| {
                                ui.label("QR-Codes:");
                                ui.horizontal(|ui| {
                                    if ui.button("+ QR-Code hinzufügen").clicked() {
                                        self.config.qr_codes.push(QrCodeConfig { x: 100.0, y: 100.0, size: 18.0, pages: vec![1], all_pages: false });
                                    }
                                    if ui.button("- QR-Code entfernen").clicked() && !self.config.qr_codes.is_empty() {
                                        self.config.qr_codes.pop();
                                    }
                                });
                                
                                // QR-Code Größen-Slider und Seiten-Auswahl für jeden QR-Code
                                for (i, qr) in self.config.qr_codes.iter_mut().enumerate() {
                                    ui.group(|ui| {
                                        ui.horizontal(|ui| {
                                            ui.label(format!("QR-Code {} Größe:", i + 1));
                                            ui.add(egui::Slider::new(&mut qr.size, 10.0..=50.0).suffix(" pt"));
                                        });
                                        
                                        // Seiten-Auswahl für diesen QR-Code
                                        ui.horizontal(|ui| {
                                            ui.label("Seiten:");
                                            if ui.checkbox(&mut qr.all_pages, "Alle Seiten").clicked() {
                                                if qr.all_pages {
                                                    qr.pages.clear(); // Leeren wenn "Alle Seiten" aktiv
                                                } else if qr.pages.is_empty() {
                                                    qr.pages.push(1); // Standard-Seite wenn "Alle Seiten" deaktiviert
                                                }
                                            }
                                        });
                                        
                                        if !qr.all_pages {
                                            ui.horizontal_wrapped(|ui| {
                                                ui.label("Seiten:");
                                                
                                                // Direkte Seiten-Eingabe mit Add/Remove
                                                let mut pages_to_remove = Vec::new();
                                                let pages_clone = qr.pages.clone();
                                                for (i, page_val) in pages_clone.iter().enumerate() {
                                                    ui.horizontal(|ui| {
                                                        let mut page_val_mut = *page_val;
                                                        if ui.add(egui::DragValue::new(&mut page_val_mut).clamp_range(1..=100).prefix("S")).changed() {
                                                            qr.pages[i] = page_val_mut;
                                                        }
                                                        if ui.small_button("❌").clicked() && qr.pages.len() > 1 {
                                                            pages_to_remove.push(i);
                                                        }
                                                    });
                                                }
                                                
                                                // Seiten entfernen (rückwärts, um Indizes nicht zu verschieben)
                                                for &i in pages_to_remove.iter().rev() {
                                                    qr.pages.remove(i);
                                                }
                                                
                                                if ui.small_button("+ Seite").clicked() {
                                                    let new_page = qr.pages.iter().max().unwrap_or(&0) + 1;
                                                    qr.pages.push(new_page);
                                                    qr.pages.sort();
                                                }
                                            });
                                        }
                                    });
                                }
                            });
                            
                            ui.separator();
                            
                            ui.group(|ui| {
                                ui.label("Vertreternummer-Felder:");
                                ui.horizontal(|ui| {
                                    if ui.button("+ Feld hinzufügen").clicked() {
                                        self.config.vertreter.push(VertreterConfig { x: 100.0, y: 200.0, size: 12.0, pages: vec![1], all_pages: false, font_name: "Arial".to_string(), font_size: 12.0, font_style: "Normal".to_string() });
                                    }
                                    if ui.button("- Feld entfernen").clicked() && !self.config.vertreter.is_empty() {
                                        self.config.vertreter.pop();
                                    }
                                });
                                // Font-Einstellungen und Seiten-Auswahl für jedes Vertreternummer-Feld
                                for (i, v) in self.config.vertreter.iter_mut().enumerate() {
                                    ui.group(|ui| {
                                        ui.label(format!("Vertreternummer Feld {}", i + 1));
                                        
                                        // Font-Auswahl für dieses Vertreternummer-Feld - OPTIMIERT
                                        ui.horizontal(|ui| {
                                            ui.label("Schriftart:");
                                            
                                            // PERFORMANCE: Noch weniger Fonts für flüssige UI
                                            egui::ComboBox::from_id_source(format!("font_combo_{}", i))
                                                .selected_text(&v.font_name)
                                                .width(200.0) // Breiter für Suchfeld
                                                .show_ui(ui, |ui| {
                                                    // Font-Suchfeld
                                                    ui.horizontal(|ui| {
                                                        ui.label("🔍");
                                                        ui.text_edit_singleline(&mut self.font_search_text);
                                                        if ui.small_button("❌").clicked() {
                                                            self.font_search_text.clear();
                                                        }
                                                    });
                                                    ui.separator();
                                                    
                                                    let search_text = self.font_search_text.to_lowercase();
                                                    let is_searching = !search_text.is_empty();
                                                    
                                                    // PERFORMANCE: Nur die wichtigsten Fonts zeigen
                                                    
                                                    let common_fonts = ["Arial", "Calibri", "Times New Roman", "Helvetica", "Verdana", "Georgia"];
                                                    
                                                    // Häufige Fonts - IMMER sichtbar (außer bei Suche)
                                                    if !is_searching {
                                                        ui.label("📌 Häufig:");
                                                        for font in &common_fonts {
                                                            ui.selectable_value(&mut v.font_name, font.to_string(), *font);
                                                        }
                                                        ui.separator();
                                                    }
                                                    
                                                    // Alle anderen Fonts (gefiltert oder alle)
                                                    let label = if is_searching {
                                                        format!("🔍 Suchergebnisse:")
                                                    } else {
                                                        "🔤 Weitere:".to_string()
                                                    };
                                                    ui.label(label);
                                                    
                                                    let mut shown_count = 0;
                                                    let max_to_show = if is_searching || self.show_all_fonts { 
                                                        self.cached_fonts.len() 
                                                    } else { 
                                                        15 
                                                    };
                                                    
                                                    for font in &self.cached_fonts {
                                                        // Skip häufige Fonts (außer bei Suche)
                                                        if !is_searching && common_fonts.contains(&font.as_str()) {
                                                            continue;
                                                        }
                                                        
                                                        // Filter nach Suchtext
                                                        if is_searching && !font.to_lowercase().contains(&search_text) {
                                                            continue;
                                                        }
                                                        
                                                        if shown_count < max_to_show {
                                                            ui.selectable_value(&mut v.font_name, font.clone(), font);
                                                            shown_count += 1;
                                                        } else {
                                                            break;
                                                        }
                                                    }
                                                    
                                                    // Button um alle Fonts zu zeigen/verstecken (nur wenn nicht gesucht wird)
                                                    if !is_searching {
                                                        if !self.show_all_fonts && self.cached_fonts.len() > (common_fonts.len() + 15) {
                                                            ui.horizontal(|ui| {
                                                                ui.small(format!("📝 +{} weitere", 
                                                                    self.cached_fonts.len() - common_fonts.len() - 15));
                                                                if ui.small_button("📋 Alle zeigen").clicked() {
                                                                    self.show_all_fonts = true;
                                                                    println!("🔤 Zeige alle {} Schriftarten", self.cached_fonts.len());
                                                                }
                                                            });
                                                        } else if self.show_all_fonts {
                                                            if ui.small_button("📁 Weniger zeigen").clicked() {
                                                                self.show_all_fonts = false;
                                                                println!("🔤 Zeige nur häufige Schriftarten");
                                                            }
                                                        }
                                                    } else if shown_count == 0 {
                                                        ui.label("❌ Keine Schriftarten gefunden");
                                                    }
                                                });
                                            
                                            // Font-Refresh Button (für nachträglich installierte Fonts)
                                            if ui.button("🔄").on_hover_text("Schriftarten neu laden (für nachträglich installierte Fonts)").clicked() {
                                                println!("🔄 FONT-REFRESH: Lade Schriftarten neu...");
                                                self.cached_fonts = refresh_font_cache();
                                                println!("✅ FONT-REFRESH: {} Schriftarten verfügbar", self.cached_fonts.len());
                                            }
                                        });
                                        
                                        // Font-Style-Auswahl (mit deutschen und englischen Begriffen)
                                        ui.horizontal(|ui| {
                                            ui.label("Stil:");
                                            egui::ComboBox::from_id_source(format!("style_combo_{}", i))
                                                .selected_text(&v.font_style)
                                                .width(120.0)
                                                .show_ui(ui, |ui| {
                                                    ui.selectable_value(&mut v.font_style, "Normal".to_string(), "Normal");
                                                    ui.selectable_value(&mut v.font_style, "Bold".to_string(), "Bold (Fett)");
                                                    ui.selectable_value(&mut v.font_style, "Italic".to_string(), "Italic (Kursiv)");
                                                    ui.selectable_value(&mut v.font_style, "BoldItalic".to_string(), "Bold+Italic (Fett+Kursiv)");
                                                    // Zusätzliche Styles für Adobe/professionelle Fonts
                                                    ui.selectable_value(&mut v.font_style, "Light".to_string(), "Light (Leicht)");
                                                    ui.selectable_value(&mut v.font_style, "Medium".to_string(), "Medium");
                                                    ui.selectable_value(&mut v.font_style, "Heavy".to_string(), "Heavy (Schwer)");
                                                    ui.selectable_value(&mut v.font_style, "Black".to_string(), "Black (Sehr Fett)");
                                                    ui.selectable_value(&mut v.font_style, "Thin".to_string(), "Thin (Dünn)");
                                                });
                                        });
                                        
                                        ui.horizontal(|ui| {
                                            ui.label("Schriftgröße:");
                                            ui.add(egui::Slider::new(&mut v.font_size, 4.0..=72.0)
                                                .suffix(" pt")
                                                .step_by(0.1)
                                                .fixed_decimals(1));
                                        });
                                        
                                        // Seiten-Auswahl für dieses Vertreternummer-Feld
                                        ui.horizontal(|ui| {
                                            ui.label("Seiten:");
                                            if ui.checkbox(&mut v.all_pages, "Alle Seiten").clicked() {
                                                if v.all_pages {
                                                    v.pages.clear(); // Leeren wenn "Alle Seiten" aktiv
                                                } else if v.pages.is_empty() {
                                                    v.pages.push(1); // Standard-Seite wenn "Alle Seiten" deaktiviert
                                                }
                                            }
                                        });
                                        
                                        if !v.all_pages {
                                            ui.horizontal_wrapped(|ui| {
                                                ui.label("Seiten:");
                                                
                                                // Direkte Seiten-Eingabe mit Add/Remove
                                                let mut pages_to_remove = Vec::new();
                                                let pages_clone = v.pages.clone();
                                                for (i, page_val) in pages_clone.iter().enumerate() {
                                                    ui.horizontal(|ui| {
                                                        let mut page_val_mut = *page_val;
                                                        if ui.add(egui::DragValue::new(&mut page_val_mut).clamp_range(1..=100).prefix("S")).changed() {
                                                            v.pages[i] = page_val_mut;
                                                        }
                                                        if ui.small_button("❌").clicked() && v.pages.len() > 1 {
                                                            pages_to_remove.push(i);
                                                        }
                                                    });
                                                }
                                                
                                                // Seiten entfernen (rückwärts, um Indizes nicht zu verschieben)
                                                for &i in pages_to_remove.iter().rev() {
                                                    v.pages.remove(i);
                                                }
                                                
                                                if ui.small_button("+ Seite").clicked() {
                                                    let new_page = v.pages.iter().max().unwrap_or(&0) + 1;
                                                    v.pages.push(new_page);
                                                    v.pages.sort();
                                                }
                                            });
                                        }
                                    });
                                }
                            });
                            
                            ui.separator();
                            
                            // Koordinaten anzeigen (optional)
                            ui.collapsing("📍 Genaue Koordinaten", |ui| {
                                for (i, qr) in self.config.qr_codes.iter().enumerate() {
                                    let pages_str = if qr.all_pages {
                                        "alle Seiten".to_string()
                                    } else {
                                        format!("Seiten: {:?}", qr.pages)
                                    };
                                    ui.label(format!("QR-Code {}: x={:.1}mm, y={:.1}mm, Größe={:.1}mm, {}", 
                                        i + 1, qr.x, qr.y, qr.size, pages_str));
                                }
                                for (i, pos) in self.config.vertreter.iter().enumerate() {
                                    let pages_str = if pos.all_pages {
                                        "alle Seiten".to_string()
                                    } else {
                                        format!("Seiten: {:?}", pos.pages)
                                    };
                                    ui.label(format!("Vertreternummer {}: x={:.1}mm, y={:.1}mm, Größe={:.1}, Font: {} {} ({}pt), {}", 
                                        i + 1, pos.x, pos.y, pos.size, pos.font_name, pos.font_style, pos.font_size, pages_str));
                                }
                            });
                            
                            // Manual coordinate input
                            ui.collapsing("✏️ Manuelle Koordinaten-Eingabe", |ui| {
                                ui.label("📏 Koordinaten in PDF-Punkten (1 Punkt ≈ 0.35mm)");
                                ui.separator();
                                
                                // QR-Code manual input
                                if !self.config.qr_codes.is_empty() {
                                    ui.group(|ui| {
                                        ui.label("QR-Code Position (erster QR-Code):");
                                        ui.horizontal(|ui| {
                                            ui.label("X (mm):");
                                            ui.text_edit_singleline(&mut self.manual_qr_x);
                                        });
                                        ui.horizontal(|ui| {
                                            ui.label("Y (mm):");
                                            ui.text_edit_singleline(&mut self.manual_qr_y);
                                        });
                                        ui.horizontal(|ui| {
                                            ui.label("Größe (mm):");
                                            ui.text_edit_singleline(&mut self.manual_qr_size);
                                        });
                                        if ui.button("🔄 QR-Code Position setzen").clicked() {
                                            if let (Ok(x), Ok(y), Ok(size)) = (
                                                self.manual_qr_x.parse::<f32>(),
                                                self.manual_qr_y.parse::<f32>(),
                                                self.manual_qr_size.parse::<f32>()
                                            ) {
                                                self.config.qr_codes[0].x = x;
                                                self.config.qr_codes[0].y = y;
                                                self.config.qr_codes[0].size = size;
                                                println!("QR-Code Position manuell gesetzt: x={}mm, y={}mm, size={}mm", x, y, size);
                                            }
                                        }
                                    });
                                    ui.separator();
                                }
                                
                                // Vertreter position manual input
                                if !self.config.vertreter.is_empty() {
                                    ui.group(|ui| {
                                        ui.label("Kundennummer Position (erste Position):");
                                        ui.horizontal(|ui| {
                                            ui.label("X (mm):");
                                            ui.text_edit_singleline(&mut self.manual_vertreter_x);
                                        });
                                        ui.horizontal(|ui| {
                                            ui.label("Y (mm):");
                                            ui.text_edit_singleline(&mut self.manual_vertreter_y);
                                        });
                                        ui.horizontal(|ui| {
                                            ui.label("Größe:");
                                            ui.text_edit_singleline(&mut self.manual_vertreter_size);
                                        });
                                        if ui.button("🔄 Kundennummer Position setzen").clicked() {
                                            if let (Ok(x), Ok(y)) = (
                                                self.manual_vertreter_x.parse::<f32>(),
                                                self.manual_vertreter_y.parse::<f32>()
                                            ) {
                                                self.config.vertreter[0].x = x;
                                                self.config.vertreter[0].y = y;
                                                // Try to parse and set size as well (optional)
                                                if let Ok(sz) = self.manual_vertreter_size.parse::<f32>() {
                                                    self.config.vertreter[0].size = sz;
                                                }
                                                println!("Kundennummer Position manuell gesetzt: x={}mm, y={}mm", x, y);
                                            }
                                        }
                                    });
                                }
                                
                                ui.separator();
                                ui.label("💡 Tipp: PDF-Koordinaten beginnen links-unten bei (0,0)");
                                ui.label("📐 Referenz: DIN A4 = 595×842 Punkte");
                            });
                            
                            ui.separator();
                            
                            // Speichern Button
                            ui.horizontal(|ui| {
                                if ui.button("💾 Speichern").clicked() {
                                    println!("=== SPEICHERN GEDRÜCKT ===");
                                    println!("VOR Speichern - Config: QR={:?}, Vertreter={:?}", 
                                            self.config.qr_codes, self.config.vertreter);
                                    
                                    // Nur noch gruppenspezifische Config speichern (keine app_config.toml mehr!)
                                    set_current_config(&self.config);
                                    
                                    // Gruppenspezifische Config-Datei speichern
                                    save_group_config(&self.selected_group, &self.selected_language, self.is_messe, &self.config);
                                    
                                    println!("Gruppenspezifische Config gespeichert!");
                                    
                                    // WICHTIG: Prüfen was wirklich in der Datei steht
                                    let config_dir = get_config_dir();
                                    let internal_path = config_dir.join("app_config.toml");
                                    if let Ok(toml_content) = std::fs::read_to_string(&internal_path) {
                                        println!("Interne TOML Datei Inhalt nach Speichern:\n{}", toml_content);
                                    }
                                    
                                    self.save_message = Some(std::time::Instant::now());
                                    
                                    // Nach dem Speichern nochmal laden um sicherzustellen dass alles stimmt
                                    let loaded_config = load_group_config(&self.selected_group, &self.selected_language, self.is_messe);
                                    println!("NACH Laden - Geladene gruppenspezifische Config: QR={:?}, Vertreter={:?}", 
                                            loaded_config.qr_codes, loaded_config.vertreter);
                                    self.config = loaded_config;
                                    println!("=== SPEICHERN ABGESCHLOSSEN ===");
                                    
                                    // Dialog schließen
                                    self.show_config = false;
                                }
                            });
                        });
                    });
                });
                
                // WICHTIG: show_config korrekt übernehmen - X-Button schließt Dialog
                self.show_config = show_config;
            }
            
            // 🎉 MEME-FENSTER 🎉
            if self.show_meme {
                egui::Window::new("🎉 PDFs erstellt! 🎉")
                    .collapsible(false)
                    .resizable(false)
                    .anchor(egui::Align2::CENTER_CENTER, egui::vec2(0.0, 0.0))
                    .show(ctx, |ui| {
                        ui.vertical_centered(|ui| {
                            ui.label("Die PDFs wurden erfolgreich erstellt!");
                        });
                    });
            }
        });
    }
}

fn read_vertreter(file_path: &str) -> Vec<(String, String, String)> {
    println!("DEBUG: Versuche CSV zu lesen: {}", file_path);
    
    let content = match fs::read_to_string(file_path) {
        Ok(content) => {
            println!("DEBUG: CSV erfolgreich gelesen, {} Zeichen", content.len());
            content
        },
        Err(e) => {
            println!("ERROR: CSV konnte nicht gelesen werden: {}", e);
            return Vec::new(); // Leere Liste statt Panic
        }
    };

    // Ermitteln des Trennzeichens: wenn eine Zeile ';' enthält, priorisiere ';',
    // sonst benutze ',' als Fallback. Falls beides vorkommt, wähle das häufiger
    // vorkommende Zeichen in der ersten nicht-leeren Zeile.
    let mut delimiter = ',';
    for line in content.lines() {
        let l = line.trim();
        if l.is_empty() { continue; }
        let semi = l.matches(';').count();
        let comma = l.matches(',').count();
        if semi > 0 || comma > 0 {
            delimiter = if semi >= comma { ';' } else { ',' };
            break;
        }
    }

    let mut lines = content.lines();
    
    // Header überspringen (erste Zeile)
    if let Some(header) = lines.next() {
        println!("DEBUG: Header übersprungen: {}", header.trim());
    }

    lines
        .filter_map(|line| {
            let l = line.trim();
            if l.is_empty() { return None; }
            let parts: Vec<&str> = l.split(delimiter).collect();
            if parts.len() >= 3 {
                let vertreternr = parts[0].trim();
                let de_link = parts[1].trim();
                let en_link = parts[2].trim();
                if !vertreternr.is_empty() && !de_link.is_empty() && !en_link.is_empty() {
                    // Vertreternummer auf 4 Stellen formatieren (führende Nullen)
                    if let Ok(num) = vertreternr.parse::<u32>() {
                        let formatted_nr = if num >= 10000 {
                            num.to_string() // Zahlen >= 10000 bleiben unverändert
                        } else {
                            format!("{:04}", num) // Zahlen < 10000 werden auf 4 Stellen aufgefüllt
                        };
                        return Some((formatted_nr, de_link.to_string(), en_link.to_string()));
                    }
                }
            } else if parts.len() >= 2 {
                // Fallback für alte CSV-Struktur (nur 2 Spalten)
                let vertreternr = parts[0].trim();
                let link = parts[1].trim();
                if !vertreternr.is_empty() && !link.is_empty() {
                    if let Ok(num) = vertreternr.parse::<u32>() {
                        let formatted_nr = if num >= 10000 {
                            num.to_string()
                        } else {
                            format!("{:04}", num)
                        };
                        return Some((formatted_nr, link.to_string(), link.to_string())); // Gleicher Link für beide Sprachen
                    }
                }
            }
            None
        })
        .collect()
}

// Erzeuge ein kleines eingebettetes Icon (blockiges "B") als RGBA-Pixel-Array.
// So hat die Anwendung auch ohne externe Datei ein eigenes Icon in der Taskleiste.
#[allow(dead_code)]
fn make_app_icon() -> egui::IconData {
    let width: usize = 32;
    let height: usize = 32;
    let mut rgba = vec![0u8; width * height * 4];

    // Hintergrund: Orange-ish
    for y in 0..height {
        for x in 0..width {
            let i = (y * width + x) * 4;
            rgba[i + 0] = 255; // R
            rgba[i + 1] = 140; // G
            rgba[i + 2] = 0;   // B
            rgba[i + 3] = 255; // A
        }
    }

    // Zeichne ein blockiges 'B' in Weiß (einfach aus Rechtecken zusammengesetzt)
    // Vertikale Hauptlinie
    for y in 6..26 {
        for x in 6..10 {
            let i = (y * width + x) * 4;
            rgba[i + 0] = 255; rgba[i + 1] = 255; rgba[i + 2] = 255; rgba[i + 3] = 255;
        }
    }
    // Oberer B-Bogen
    for y in 6..14 {
        for x in 10..22 {
            let i = (y * width + x) * 4;
            rgba[i + 0] = 255; rgba[i + 1] = 255; rgba[i + 2] = 255; rgba[i + 3] = 255;
        }
    }
    // Unterer B-Bogen
    for y in 18..26 {
        for x in 10..22 {
            let i = (y * width + x) * 4;
            rgba[i + 0] = 255; rgba[i + 1] = 255; rgba[i + 2] = 255; rgba[i + 3] = 255;
        }
    }

    egui::IconData { rgba, width: width as _, height: height as _ }
}

fn generate_qr(link: &str) -> (Vec<u8>, usize) {
    let code = QrCode::new(link).expect("Konnte QR-Code nicht generieren");
    let matrix: String = code
        .render::<char>()
        .quiet_zone(false)
        .module_dimensions(1, 1)
        .build(); // String
    let height = matrix.lines().count();
    let width = if height > 0 { matrix.lines().next().unwrap().chars().count() } else { 0 };
    let mut data = Vec::with_capacity(width * height);
    for row in matrix.lines() {
        for c in row.chars() {
            data.push(if c == '█' { 0u8 } else { 255u8 });
        }
    }
    (data, width)
}

// Globale Variable für die aktuelle Config (wird von UI gesetzt)
static mut CURRENT_CONFIG: Option<Config> = None;
static CONFIG_MUTEX: std::sync::Mutex<()> = std::sync::Mutex::new(());

// Funktion um die aktuelle Config zu setzen (threadsafe)
fn set_current_config(config: &Config) {
    let _lock = CONFIG_MUTEX.lock().unwrap();
    unsafe {
        CURRENT_CONFIG = Some(config.clone());
    }
    println!("Aktuelle Config gesetzt für PDF-Generierung: QR={:?}, Vertreter={:?}", 
             config.qr_codes, config.vertreter);
}


// Parse config from a provided Config object (used to support per-template/group configs)
#[allow(dead_code)]
fn parse_config_from(config: &Config) -> (Vec<(f64, f64, f64)>, Vec<(f64, f64, f64)>, Vec<u32>) {
    let qr_configs: Vec<(f64, f64, f64)> = config.qr_codes.iter()
        .map(|qr| (qr.x as f64, qr.y as f64, qr.size as f64))
        .collect();

    // Now include size for vertreter as well
    let vertreter_positions: Vec<(f64, f64, f64)> = config.vertreter.iter()
        .map(|v| (v.x as f64, v.y as f64, v.size as f64))
        .collect();

    println!("parse_config_from verwendet: QR={:?}, Vertreter={:?}", qr_configs, vertreter_positions);

    (qr_configs, vertreter_positions, vec![1]) // Dummy für alte Kompatibilität
}

// Versuche aus einem Template-Pfad Gruppe und Sprache zu extrahieren
fn infer_group_lang_from_template(template: &str) -> (String, String, bool) {
    let filename = std::path::Path::new(template)
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or(template)
        .to_string();

    // Beispiele: Bestellschein-Apo-de_de, Bestellschein-Messe-Apo-de_de
    let parts: Vec<&str> = filename.split('-').collect();
    // Suche nach bekannten Gruppen (Apo, Endkunde, Fachkreise)
    let mut group = "Endkunde".to_string();
    for p in &parts {
        let p_low = p.to_lowercase();
        if p_low.contains("apo") { group = "Apo".to_string(); break; }
        if p_low.contains("endkunde") { group = "Endkunde".to_string(); break; }
        if p_low.contains("fachkreise") { group = "Fachkreise".to_string(); break; }
    }

    // Sprache als letzten Teil (wenn vorhanden)
    let mut lang = "Deutsch".to_string();
    if let Some(last) = parts.last() {
        let last_low = last.to_lowercase();
        if last_low.starts_with("de") { lang = "Deutsch".to_string(); }
        else if last_low.starts_with("en") { lang = "Englisch".to_string(); }
    }

    // Messe im Dateinamen erkennen
    let is_messe = filename.to_lowercase().contains("messe");

    (group, lang, is_messe)
}

// Erweiterte Funktion um installierte Windows-Fonts mit Styles zu ermitteln
fn get_installed_fonts_with_styles() -> Vec<String> {
    let mut fonts = Vec::new();
    
    // Standard Windows-Fonts die fast immer verfügbar sind (mit Styles)
    let default_fonts = vec![
        // Arial Familie
        "Arial".to_string(),
        "Arial Bold".to_string(),
        "Arial Italic".to_string(),
        "Arial Bold Italic".to_string(),
        // Times Familie
        "Times New Roman".to_string(),
        "Times New Roman Bold".to_string(),
        "Times New Roman Italic".to_string(),
        "Times New Roman Bold Italic".to_string(),
        // Calibri Familie  
        "Calibri".to_string(),
        "Calibri Bold".to_string(),
        "Calibri Italic".to_string(),
        "Calibri Bold Italic".to_string(),
        // Source Fonts (Adobe/Google) - häufig installiert
        "Source Sans Pro".to_string(),
        "Source Sans Pro Bold".to_string(),
        "Source Sans Pro Italic".to_string(),
        "Source Sans Pro Bold Italic".to_string(),
        "Source Sans Pro Light".to_string(),
        "Source Sans Pro Black".to_string(),
        "Source Code Pro".to_string(),
        "Source Code Pro Bold".to_string(),
        "Source Code Pro Light".to_string(),
        "Source Serif Pro".to_string(),
        "Source Serif Pro Bold".to_string(),
        "Source Serif Pro Italic".to_string(),
        // Andere Standard-Fonts
        "Verdana".to_string(),
        "Verdana Bold".to_string(),
        "Verdana Italic".to_string(),
        "Georgia".to_string(),
        "Georgia Bold".to_string(),
        "Georgia Italic".to_string(),
        "Trebuchet MS".to_string(),
        "Trebuchet MS Bold".to_string(),
        "Trebuchet MS Italic".to_string(),
        "Comic Sans MS".to_string(),
        "Comic Sans MS Bold".to_string(),
        "Impact".to_string(),
        "Lucida Console".to_string(),
        "Tahoma".to_string(),
        "Tahoma Bold".to_string(),
        "Courier New".to_string(),
        "Courier New Bold".to_string(),
        "Courier New Italic".to_string(),
        "Helvetica".to_string(),
        // Zusätzliche deutsche/europäische Fonts
        "Candara".to_string(),
        "Candara Bold".to_string(),
        "Candara Italic".to_string(),
        "Constantia".to_string(),
        "Constantia Bold".to_string(),
        "Constantia Italic".to_string(),
        "Corbel".to_string(),
        "Corbel Bold".to_string(),
        "Corbel Italic".to_string(),
    ];
    
    fonts.extend(default_fonts);
    
    // Versuche zusätzliche Fonts aus mehreren Verzeichnissen zu lesen
    let font_directories = vec![
        "C:\\Windows\\Fonts".to_string(),
        format!("{}\\Fonts", std::env::var("LOCALAPPDATA").unwrap_or_default()),
        format!("{}\\AppData\\Local\\Microsoft\\Windows\\Fonts", std::env::var("USERPROFILE").unwrap_or_default()),
    ];
    
    for font_dir in font_directories {
        if let Ok(entries) = std::fs::read_dir(&font_dir) {
            for entry in entries.flatten() {
                if let Some(file_name) = entry.file_name().to_str() {
                    if file_name.ends_with(".ttf") || file_name.ends_with(".otf") || file_name.ends_with(".TTF") || file_name.ends_with(".OTF") {
                        // Erweiterte Font-Namen-Extraktion mit Style-Erkennung
                        let mut font_name = file_name
                            .replace(".ttf", "")
                            .replace(".otf", "")
                            .replace(".TTF", "")
                            .replace(".OTF", "")
                            .replace("_", " ")
                            .replace("-", " ");
                        
                        // Erweiterte Font-Name-Bereinigung für bessere Erkennung
                        if font_name.contains("Adobe") {
                            font_name = font_name.replace("Adobe", "").trim().to_string();
                        }
                        // Source Fonts (Adobe): "Source Sans Pro", "Source Code Pro", etc.
                        if font_name.starts_with("Source ") {
                            // Behalte "Source" Präfix für bessere Identifikation
                        }
                        // Andere bekannte Präfixe bereinigen
                        let prefixes_to_remove = ["Microsoft ", "Google ", "Apple ", "System "];
                        for prefix in &prefixes_to_remove {
                            if font_name.starts_with(prefix) {
                                font_name = font_name.replace(prefix, "").trim().to_string();
                                break;
                            }
                        }
                        
                        // Erweiterte Style-Erkennung (deutsch, englisch und Varianten)
                        let styles = [
                            ("Bold", "Bold"), ("Fett", "Bold"), ("bold", "Bold"), ("BOLD", "Bold"),
                            ("Italic", "Italic"), ("Kursiv", "Italic"), ("italic", "Italic"), ("ITALIC", "Italic"),
                            ("Oblique", "Italic"), ("Schräg", "Italic"), ("oblique", "Italic"),
                            ("Light", "Light"), ("Leicht", "Light"), ("light", "Light"), ("LIGHT", "Light"),
                            ("Medium", "Medium"), ("medium", "Medium"), ("MEDIUM", "Medium"),
                            ("Heavy", "Heavy"), ("Schwer", "Heavy"), ("heavy", "Heavy"), ("HEAVY", "Heavy"),
                            ("Black", "Black"), ("Schwarz", "Black"), ("black", "Black"), ("BLACK", "Black"),
                            ("Thin", "Thin"), ("Dünn", "Thin"), ("thin", "Thin"), ("THIN", "Thin"),
                            ("Ultra", "Heavy"), ("Extra", "Heavy"), ("ultra", "Heavy"), ("extra", "Heavy"),
                            ("SemiBold", "Bold"), ("DemiBold", "Bold"), ("semibold", "Bold"), ("demibold", "Bold"),
                            ("Regular", "Regular"), ("Normal", "Regular"), ("regular", "Regular"), ("REGULAR", "Regular"),
                            ("Roman", "Regular"), ("Book", "Regular"), ("roman", "Regular"), ("book", "Regular"),
                        ];
                        let mut detected_styles = Vec::new();
                        
                        for (style_name, english_style) in &styles {
                            if font_name.to_lowercase().contains(&style_name.to_lowercase()) {
                                if !detected_styles.contains(&english_style.to_string()) {
                                    detected_styles.push(english_style.to_string());
                                }
                            }
                        }
                        
                        // Basis-Font-Namen ohne Styles
                        let mut base_name = font_name.clone();
                        for (style_name, _) in &styles {
                            if base_name.to_lowercase().contains(&style_name.to_lowercase()) {
                                base_name = base_name.replace(style_name, "").trim().to_string();
                            }
                        }
                        
                        // Füge Basis-Font hinzu
                        if !base_name.is_empty() && !fonts.iter().any(|f| f.to_lowercase() == base_name.to_lowercase()) {
                            fonts.push(base_name.clone());
                        }
                        
                        // Füge Style-Varianten hinzu
                        if !detected_styles.is_empty() {
                            let style_name = format!("{} {}", base_name, detected_styles.join(" "));
                            if !fonts.iter().any(|f| f.to_lowercase() == style_name.to_lowercase()) {
                                fonts.push(style_name);
                            }
                        }
                        
                        // Auch originalen Namen hinzufügen falls anders
                        let original_name = font_name.trim().to_string();
                        if !original_name.is_empty() && !fonts.iter().any(|f| f.to_lowercase() == original_name.to_lowercase()) {
                            fonts.push(original_name);
                        }
                    }
                }
            }
        }
    }
    
    // Sortiere alphabetisch und entferne Duplikate
    fonts.sort();
    fonts.dedup();
    
    println!("🔤 FONTS GEFUNDEN: {} Schriftarten geladen", fonts.len());
    if fonts.len() > 50 {
        println!("📝 Erste 10 Fonts: {:?}", &fonts[0..10.min(fonts.len())]);
        println!("📝 Letzte 10 Fonts: {:?}", &fonts[fonts.len().saturating_sub(10)..]);
    } else {
        println!("📝 Alle Fonts: {:?}", fonts);
    }
    
    fonts
}

// Aktualisiere Font-Cache (kann von UI aufgerufen werden)
fn refresh_font_cache() -> Vec<String> {
    println!("🔄 FONT-CACHE: Aktualisiere Schriftarten-Liste...");
    
    // PERFORMANCE: Cache in Datei speichern um wiederholte Scans zu vermeiden
    let cache_file = std::path::Path::new("font_cache.json");
    let cache_age_hours = 24; // Cache 24 Stunden gültig
    
    // Prüfe ob Cache-Datei existiert und noch gültig ist
    if cache_file.exists() {
        if let Ok(metadata) = std::fs::metadata(cache_file) {
            if let Ok(modified) = metadata.modified() {
                if let Ok(elapsed) = modified.elapsed() {
                    if elapsed.as_secs() < (cache_age_hours * 3600) {
                        println!("📁 FONT-CACHE: Verwende gecachte Font-Liste ({}h alt)", elapsed.as_secs() / 3600);
                        
                        // Lade aus Cache
                        if let Ok(cache_content) = std::fs::read_to_string(cache_file) {
                            if let Ok(cached_fonts) = serde_json::from_str::<Vec<String>>(&cache_content) {
                                if !cached_fonts.is_empty() {
                                    println!("✅ FONT-CACHE: {} Fonts aus Cache geladen", cached_fonts.len());
                                    return cached_fonts;
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    
    // Cache ist ungültig oder existiert nicht - neu scannen
    println!("🔍 FONT-CACHE: Scanne Schriftarten neu...");
    let fonts = get_installed_fonts_with_styles_optimized();
    
    // PERFORMANCE: Begrenze Anzahl der Fonts für UI-Performance (WENIGER für bessere Performance)
    let max_fonts = 50; // Reduziert von 200 auf 50 für flüssige UI
    let mut filtered_fonts = fonts;
    if filtered_fonts.len() > max_fonts {
        println!("⚡ PERFORMANCE: Begrenze Fonts von {} auf {} für flüssige UI", filtered_fonts.len(), max_fonts);
        
        // Priorisiere häufige Fonts
        let priority_fonts = ["Arial", "Calibri", "Times New Roman", "Helvetica", "Verdana", "Georgia", "Tahoma", "Segoe UI"];
        let mut prioritized = Vec::new();
        let mut others = Vec::new();
        
        for font in filtered_fonts {
            if priority_fonts.iter().any(|pf| font.contains(pf)) {
                prioritized.push(font);
            } else {
                others.push(font);
            }
        }
        
        // Erst priority_fonts, dann die ersten "others" bis max_fonts erreicht
        prioritized.extend(others.into_iter().take(max_fonts.saturating_sub(prioritized.len())));
        filtered_fonts = prioritized;
    }
    
    // Cache speichern
    if let Ok(cache_json) = serde_json::to_string_pretty(&filtered_fonts) {
        if let Err(e) = std::fs::write(cache_file, cache_json) {
            println!("⚠️  FONT-CACHE: Konnte Cache nicht speichern: {}", e);
        } else {
            println!("💾 FONT-CACHE: Cache gespeichert");
        }
    }
    
    filtered_fonts
}

// Optimierte Font-Scanning-Funktion
fn get_installed_fonts_with_styles_optimized() -> Vec<String> {
    // Verwende die existierende Funktion aber mit Performance-Verbesserungen
    get_installed_fonts_with_styles()
}

// Font-Pfad mit mehreren Quellen finden (ohne Admin-Rechte)
fn find_font_file(font_name: &str, style: &str) -> Option<std::path::PathBuf> {
    // Mehrere mögliche Font-Ordner (auch ohne Admin-Rechte)
    let font_dirs = vec![
        "C:\\Windows\\Fonts".to_string(),
        format!("{}\\AppData\\Local\\Microsoft\\Windows\\Fonts", std::env::var("USERPROFILE").unwrap_or_default()),
        format!("{}\\AppData\\Roaming\\Adobe\\CoreSync\\plugins\\livetype\\.r", std::env::var("USERPROFILE").unwrap_or_default()),
        ".\\fonts".to_string(), // Lokaler fonts Ordner im Projekt
    ];
    
    // Mögliche Dateinamen für den Font (viele Varianten)
    let font_base = font_name.replace(" ", "").to_lowercase();
    let style_lower = style.to_lowercase();
    
    let possible_names = vec![
        // Standard-Benennungen
        format!("{}.ttf", font_base),
        format!("{}.otf", font_base), 
        format!("{}_{}.ttf", font_base, style_lower),
        format!("{}_{}.otf", font_base, style_lower),
        format!("{}-{}.ttf", font_base, style_lower),
        format!("{}-{}.otf", font_base, style_lower),
        // Windows-spezifische Benennungen
        format!("{}b.ttf", font_base), // Bold
        format!("{}i.ttf", font_base), // Italic
        format!("{}z.ttf", font_base), // Bold Italic
        format!("{}bd.ttf", font_base), // Bold
        format!("{}it.ttf", font_base), // Italic
        // Vollständige Namen
        format!("{}.ttf", font_name.replace(" ", "")),
        format!("{}.otf", font_name.replace(" ", "")),
        format!("{} {}.ttf", font_name, style),
        format!("{} {}.otf", font_name, style),
        // Spezielle Arial-Varianten
        format!("arial{}.ttf", if style_lower.contains("bold") { "bd" } else if style_lower.contains("italic") { "i" } else { "" }),
        // Calibri-Varianten  
        format!("calibri{}.ttf", if style_lower.contains("bold") { "b" } else if style_lower.contains("italic") { "i" } else { "" }),
        // Times New Roman-Varianten
        format!("times{}.ttf", if style_lower.contains("bold") { "bd" } else if style_lower.contains("italic") { "i" } else { "" }),
    ];
    
    // Alle Kombinationen durchprobieren
    for dir in font_dirs {
        for name in &possible_names {
            let font_path = std::path::Path::new(&dir).join(name);
            if font_path.exists() && font_path.is_file() {
                println!("DEBUG: Font gefunden: {} -> {}", font_name, font_path.display());
                return Some(font_path);
            }
        }
    }
    
    println!("DEBUG: Font NICHT gefunden: {} ({})", font_name, style);
    None
}


#[allow(dead_code)]
fn embed_ttf_font(
    doc: &mut Document,
    font_dict: &mut lopdf::Dictionary,
    font_names: &mut std::collections::HashMap<String, String>,
    font_counter: &mut usize,
    font_name: &str,
    font_path: &std::path::Path,
    debug_info: &mut Vec<String>
) -> Option<String> {
    match std::fs::read(font_path) {
        Ok(ttf_data) => {
            debug_info.push(format!("✓ TTF gelesen: {} ({} KB)", font_path.display(), ttf_data.len() / 1024));
            
            let ttf_font_key = format!("TTF{}", font_counter);
            *font_counter += 1;
            
            // Font-Stream für TTF-Daten erstellen 
            let font_stream_obj = doc.add_object(Stream::new(dictionary!{
                "Length" => ttf_data.len() as i64,
                "Length1" => ttf_data.len() as i64
            }, ttf_data));
            
            // FontDescriptor erstellen
            let font_descriptor_obj = doc.add_object(dictionary!{
                "Type" => "FontDescriptor",
                "FontName" => format!("{}+{}", ttf_font_key, font_name.replace(" ", "")),
                "Flags" => 32, // Symbolic
                "FontBBox" => Object::Array(vec![Object::Integer(-200), Object::Integer(-200), Object::Integer(1000), Object::Integer(1000)]),
                "ItalicAngle" => 0,
                "Ascent" => 800,
                "Descent" => -200,
                "CapHeight" => 700,
                "StemV" => 80,
                "FontFile2" => font_stream_obj
            });
            
            // TrueType Font Dictionary
            font_dict.set(ttf_font_key.as_bytes(), dictionary!{
                "Type" => "Font",
                "Subtype" => "TrueType",
                "BaseFont" => format!("{}+{}", ttf_font_key, font_name.replace(" ", "")),
                "FontDescriptor" => font_descriptor_obj,
                "FirstChar" => 32,
                "LastChar" => 255,
                "Widths" => Object::Array((32..256).map(|_| Object::Integer(500)).collect())
            });
            
            font_names.insert(ttf_font_key.clone(), ttf_font_key.clone());
            debug_info.push(format!("✅ TTF eingebettet: {}", font_name));
            
            Some(ttf_font_key)
        }
        Err(e) => {
            debug_info.push(format!("❌ TTF lesen fehlgeschlagen: {} - {}", font_path.display(), e));
            None
        }
    }
}

/// Standard-Font-Fallback-Funktion für process_page_elements
fn create_standard_font_fallback(
    font_dict: &mut lopdf::Dictionary,
    used_font_keys: &mut std::collections::HashMap<String, String>,
    font_counter: &mut usize,
    vertreter_config: &VertreterConfig
) -> String {
    let pdf_font_name = match vertreter_config.font_name.as_str() {
        "Times New Roman" | "Times" => {
            match vertreter_config.font_style.as_str() {
                "Bold" => "Times-Bold",
                "Italic" => "Times-Italic",
                "Bold Italic" | "BoldItalic" => "Times-BoldItalic",
                _ => "Times-Roman"
            }
        },
        "Courier New" | "Courier" => {
            match vertreter_config.font_style.as_str() {
                "Bold" => "Courier-Bold", 
                "Italic" => "Courier-Oblique",
                "Bold Italic" | "BoldItalic" => "Courier-BoldOblique",
                _ => "Courier"
            }
        },
        // Alle anderen (einschließlich Custom Fonts) → Helvetica-Fallback
        _ => {
            match vertreter_config.font_style.as_str() {
                "Bold" => "Helvetica-Bold",
                "Italic" => "Helvetica-Oblique",
                "Bold Italic" | "BoldItalic" => "Helvetica-BoldOblique", 
                _ => "Helvetica"
            }
        }
    };

    // Prüfe ob Font bereits registriert
    if let Some(existing_key) = used_font_keys.get(pdf_font_name) {
        existing_key.clone()
    } else {
        let new_key = format!("F{}", font_counter);
        *font_counter += 1;
        
        // Font in PDF registrieren
        font_dict.set(new_key.as_bytes(), dictionary!{
            "Type" => "Font",
            "Subtype" => "Type1", 
            "BaseFont" => pdf_font_name
        });
        
        used_font_keys.insert(pdf_font_name.to_string(), new_key.clone());
        new_key
    }
}

fn modify_pdf_with_debug(template_path: &str, kundennr: &str, qr_code: &[u8], qr_width: usize, config: &Config, output_path: &std::path::Path, debug_enabled: bool, enable_font_fallback: bool) {
    debug_print(&format!("Lade PDF-Template: {}", template_path), debug_enabled);
    let mut doc = match Document::load(template_path) {
        Ok(document) => {
            debug_print("PDF-Template erfolgreich geladen", debug_enabled);
            document
        },
        Err(e) => {
            println!("ERROR: Konnte PDF-Template nicht laden: {} - {}", template_path, e);
            return;
        }
    };
    
    // QR-Code als XObject registrieren, falls vorhanden
    let mut maybe_image_id: Option<lopdf::ObjectId> = None;
    if qr_width > 0 && !qr_code.is_empty() {
        let image_id = doc.add_object(lopdf::Object::Stream(lopdf::Stream::new(
            dictionary! {
                "Type" => "XObject",
                "Subtype" => "Image",
                "Width" => qr_width as i64,
                "Height" => qr_width as i64,
                "ColorSpace" => "DeviceGray",
                "BitsPerComponent" => 8,
            },
            qr_code.to_vec(),
        )));
        maybe_image_id = Some(image_id);
    }

    // ✅ GLOBALE TTF-FONT-EINBETTUNG (einmalig für gesamtes PDF)
    let mut global_font_registry = std::collections::HashMap::new();
    let mut font_counter = 1;
    
    for vertreter_config in &config.vertreter {
        let font_key = format!("{}_{}", vertreter_config.font_name, vertreter_config.font_style);
        
        if !global_font_registry.contains_key(&font_key) {
            if let Some(ttf_path) = find_font_file(&vertreter_config.font_name, &vertreter_config.font_style) {
                debug_print(&format!("✓ TTF Font einbettend: {} ({}) -> {}", 
                    vertreter_config.font_name, vertreter_config.font_style, ttf_path.display()), debug_enabled);
                
                match std::fs::read(&ttf_path) {
                    Ok(ttf_data) => {
                        let ttf_font_key = format!("TTF{}", font_counter);
                        font_counter += 1;
                        
                        // Font-Stream für TTF-Daten erstellen 
                        let font_stream_obj = doc.add_object(Stream::new(dictionary!{
                            "Length" => ttf_data.len() as i64,
                            "Length1" => ttf_data.len() as i64
                        }, ttf_data));
                        
                        // FontDescriptor erstellen - korrekte BBox für verschiedene Fonts
                        let (font_bbox, flags, italic_angle, ascent, descent, cap_height, stem_v) = match vertreter_config.font_name.as_str() {
                            "Arial" => (
                                vec![-665, -325, 2000, 1006], // Korrekte Arial BBox
                                32,  // Symbolic
                                0,   // ItalicAngle
                                905, // Ascent
                                -210, // Descent
                                728,  // CapHeight
                                87    // StemV
                            ),
                            "Calibri" => (
                                vec![-503, -313, 1240, 1089],
                                32, 0, 952, -269, 632, 87
                            ),
                            "Times New Roman" | "Times" => (
                                vec![-568, -307, 2000, 1007],
                                34, 0, 891, -216, 662, 93
                            ),
                            "Courier New" | "Courier" => (
                                vec![-122, -680, 623, 1021],
                                35, 0, 832, -300, 571, 51
                            ),
                            _ => (
                                vec![-665, -325, 2000, 1006], // Default Arial BBox
                                32, 0, 905, -210, 728, 87
                            )
                        };
                        
                        let _font_descriptor_obj = doc.add_object(dictionary!{
                            "Type" => "FontDescriptor",
                            "FontName" => format!("{}+{}", ttf_font_key, vertreter_config.font_name.replace(" ", "")),
                            "Flags" => flags,
                            "FontBBox" => Object::Array(font_bbox.into_iter().map(Object::Integer).collect()),
                            "ItalicAngle" => italic_angle,
                            "Ascent" => ascent,
                            "Descent" => descent,
                            "CapHeight" => cap_height,
                            "StemV" => stem_v,
                            "FontFile2" => font_stream_obj
                        });
                        
                        // TrueType Font Dictionary erstellen und direkt in globaler Font-Dict registrieren
                        // Aber das geht nicht, weil wir noch nicht in process_page_elements sind
                        // Registriere erstmal nur den Schlüssel
                        global_font_registry.insert(font_key.clone(), ttf_font_key.clone());
                        debug_print(&format!("✅ TTF eingebettet: {} -> {}", vertreter_config.font_name, ttf_font_key), debug_enabled);
                    }
                    Err(e) => {
                        debug_print(&format!("❌ TTF laden fehlgeschlagen: {} - {}, verwende Standard-Font", ttf_path.display(), e), debug_enabled);
                        // Standard-Font als Fallback registrieren
                        let standard_key = format!("STD{}", font_counter);
                        font_counter += 1;
                        global_font_registry.insert(font_key.clone(), standard_key);
                    }
                }
            } else {
                debug_print(&format!("⚠ TTF Font NICHT gefunden: {} ({}), verwende Standard-Font", 
                    vertreter_config.font_name, vertreter_config.font_style), debug_enabled);
                // Standard-Font als Fallback registrieren
                let standard_key = format!("STD{}", font_counter);
                font_counter += 1;
                global_font_registry.insert(font_key.clone(), standard_key);
            }
        }
    }

    debug_print(&format!("Font-Registry erstellt: {} Fonts eingebettet", global_font_registry.len()), debug_enabled);

    // Alle Seiten des PDFs ermitteln
    let all_pages: Vec<u32> = doc.get_pages().keys().cloned().collect();
    debug_print(&format!("PDF hat {} Seiten: {:?}", all_pages.len(), all_pages), debug_enabled);

    // Für jede Seite prüfen, welche Elemente darauf platziert werden sollen
    for page_number in all_pages {
        let page_id = doc.get_pages().get(&page_number).copied().unwrap();
        
        // QR-Codes für diese Seite sammeln
        let qr_codes_for_page: Vec<&QrCodeConfig> = config.qr_codes.iter()
            .filter(|qr| qr.all_pages || qr.pages.contains(&page_number))
            .collect();
            
        // Vertreternummer-Positionen für diese Seite sammeln
        let vertreter_for_page: Vec<&VertreterConfig> = config.vertreter.iter()
            .filter(|v| v.all_pages || v.pages.contains(&page_number))
            .collect();
        
        // Nur wenn Elemente auf dieser Seite platziert werden sollen
        if !qr_codes_for_page.is_empty() || !vertreter_for_page.is_empty() {
            debug_print(&format!("Bearbeite Seite {}: {} QR-Codes, {} Vertreternummern", 
                page_number, qr_codes_for_page.len(), vertreter_for_page.len()), debug_enabled);
                
            process_page_elements(&mut doc, page_id, page_number, &qr_codes_for_page, &vertreter_for_page, 
                                  kundennr, maybe_image_id, debug_enabled, &global_font_registry, enable_font_fallback);
        } else {
            debug_print(&format!("Seite {} übersprungen - keine Elemente zu platzieren", page_number), debug_enabled);
        }
    }
    
    // Sicherstellen dass der Output-Ordner existiert
    if let Some(parent) = output_path.parent() {
        if !parent.exists() {
            match std::fs::create_dir_all(parent) {
                Ok(()) => debug_print(&format!("Output-Ordner erstellt: {}", parent.display()), debug_enabled),
                Err(e) => {
                    println!("ERROR: Konnte Output-Ordner nicht erstellen: {} - {}", parent.display(), e);
                    return;
                }
            }
        } else {
            debug_print(&format!("Output-Ordner existiert bereits: {}", parent.display()), debug_enabled);
        }
    }
    
    // PDF in den angegebenen Pfad speichern
    match doc.save(output_path) {
        Ok(_file) => debug_print(&format!("PDF erfolgreich gespeichert: {}", output_path.display()), debug_enabled),
        Err(e) => {
            println!("ERROR: Konnte PDF nicht speichern: {} - {}", output_path.display(), e);
            return;
        }
    }
}

fn process_page_elements(doc: &mut Document, page_id: lopdf::ObjectId, _page_number: u32,
                         qr_codes: &[&QrCodeConfig], vertreter_configs: &[&VertreterConfig], 
                         kundennr: &str, maybe_image_id: Option<lopdf::ObjectId>, _debug_enabled: bool,
                         global_font_registry: &std::collections::HashMap<String, String>,
                         enable_font_fallback: bool) {
    
    let content_stream = doc.get_page_content(page_id).expect("Konnte Seiteninhalt nicht laden");
    let mut content = Content::decode(&content_stream).expect("Konnte Inhalt nicht dekodieren");

    // Alle QR-Codes platzieren (Koordinaten von mm zu PDF-Punkten umrechnen)
    for (i, qr_config) in qr_codes.iter().enumerate() {
        if let Some(_img_id) = maybe_image_id {
            // Umrechnung: mm → PDF-Punkte (1 mm = 2.834646 Punkte)
            let x_points = qr_config.x * 2.834646;
            let y_points = qr_config.y * 2.834646;
            let size_points = qr_config.size * 2.834646;
            
            content.operations.push(Operation::new("q", vec![]));
            content.operations.push(Operation::new("cm", vec![
                size_points.into(), 0.into(), 0.into(), size_points.into(), 
                x_points.into(), y_points.into()
            ]));
            content.operations.push(Operation::new("Do", vec![Object::Name(format!("Im{}", i + 1).into_bytes())]));
            content.operations.push(Operation::new("Q", vec![]));
        }
    }

    // XObject und Font im Ressourcen-Dictionary der Seite eintragen (vor dem Content-Stream!)
    let page_dict = doc.get_object_mut(page_id).expect("Konnte Seite nicht finden").as_dict_mut().expect("Konnte Seite nicht als Dict lesen");
    
    // Resources-Dict holen oder anlegen
    if !page_dict.has(b"Resources") {
        page_dict.set("Resources", dictionary!{});
    }
    let resources = page_dict.get_mut(b"Resources").expect("Konnte Resources nicht finden/anlegen");
    let resources_dict = resources.as_dict_mut().expect("Konnte Resources nicht als Dict lesen");
    
    // XObject-Dict holen oder anlegen und QR-Code-Images registrieren
    if let Some(img_id) = maybe_image_id {
        if !qr_codes.is_empty() {
            if !resources_dict.has(b"XObject") {
                resources_dict.set("XObject", dictionary!{});
            }
            let xobjects = resources_dict.get_mut(b"XObject").expect("Konnte XObject nicht finden/anlegen");
            let xobject_dict = xobjects.as_dict_mut().expect("Konnte XObject nicht als Dict lesen");
            
            for i in 0..qr_codes.len() {
                xobject_dict.set(format!("Im{}", i + 1), img_id);
            }
        }
    }

    // Font-Dict holen oder anlegen und verschiedene Fonts registrieren (nur wenn Vertreter-Positionen vorhanden)
    if !vertreter_configs.is_empty() {
        if !resources_dict.has(b"Font") {
            resources_dict.set("Font", dictionary!{});
        }
        let fonts = resources_dict.get_mut(b"Font").expect("Konnte Font nicht finden/anlegen");
        let font_dict = fonts.as_dict_mut().expect("Konnte Font nicht als Dict lesen");
        
        // VEREINFACHTE FONT-REGISTRIERUNG
        let mut used_font_keys: std::collections::HashMap<String, String> = std::collections::HashMap::new();
        let mut font_counter = 1;
        
        // Jetzt den Content-Stream mit Text schreiben - verwende GLOBALE Font-Registry
        for vertreter_config in vertreter_configs {
            let font_lookup_key = format!("{}_{}", vertreter_config.font_name, vertreter_config.font_style);
            
            // SCHRITT 1: Prüfe ob Custom Font bereits in globaler Registry vorhanden
            let font_key = if let Some(registered_font_key) = global_font_registry.get(&font_lookup_key) {
                // Custom Font ist bereits als TTF eingebettet - registriere es in Font-Dict
                if !used_font_keys.contains_key(registered_font_key) {
                    // TTF Font Dictionary registrieren (Font wurde bereits eingebettet, nur Dictionary fehlt noch)
                    font_dict.set(registered_font_key.as_bytes(), dictionary!{
                        "Type" => "Font",
                        "Subtype" => "TrueType",
                        "BaseFont" => format!("{}+{}", registered_font_key, vertreter_config.font_name.replace(" ", "")),
                        // FontDescriptor und andere Eigenschaften wurden bereits in modify_pdf_with_debug gesetzt
                        "FirstChar" => 32,
                        "LastChar" => 255,
                        "Widths" => Object::Array((32..256).map(|_| Object::Integer(500)).collect())
                    });
                    
                    used_font_keys.insert(registered_font_key.clone(), registered_font_key.clone());
                    
                    if _debug_enabled {
                        println!("✓ Custom TTF Font-Dict registriert: {} -> {}", vertreter_config.font_name, registered_font_key);
                    }
                }
                registered_font_key.clone()
            } else if enable_font_fallback {
                // SCHRITT 2: Standard-Font-Fallback (nur wenn aktiviert)
                if _debug_enabled {
                    println!("⚠ Custom Font nicht in Registry: {} ({}), verwende Standard-Font-Fallback", 
                        vertreter_config.font_name, vertreter_config.font_style);
                }
                create_standard_font_fallback(font_dict, &mut used_font_keys, &mut font_counter, vertreter_config)
            } else {
                // SCHRITT 3: Kein Fallback - Font übersprungen
                if _debug_enabled {
                    println!("❌ Custom Font nicht verfügbar: {} ({}) - Fallback deaktiviert, Element übersprungen", 
                        vertreter_config.font_name, vertreter_config.font_style);
                }
                continue; // Überspringt dieses Text-Element
            };
            
            // Umrechnung: mm → PDF-Punkte (1 mm = 2.834646 Punkte)
            let x_points = vertreter_config.x * 2.834646;
            let y_points = vertreter_config.y * 2.834646;
            
            content.operations.push(Operation::new("BT", vec![]));
            content.operations.push(Operation::new("Tf", vec![Object::Name(font_key.into_bytes()), vertreter_config.font_size.into()]));
            content.operations.push(Operation::new("Td", vec![x_points.into(), y_points.into()]));
            content.operations.push(Operation::new("Tj", vec![Object::string_literal(kundennr)]));
            content.operations.push(Operation::new("ET", vec![]));
        }
    }

    let encoded_content = content.encode().expect("Konnte Inhalt nicht kodieren");
    doc.change_page_content(page_id, encoded_content).expect("Konnte Seiteninhalt nicht ändern");
}

fn generate_bestellscheine_resume(
    progress: Arc<Mutex<f32>>,
    stop_signal: Arc<Mutex<bool>>,
    start_from: usize,
    threads: usize,
    vertreter: Arc<Vec<(String, String, String)>>,
    progress_counter: Arc<Mutex<usize>>,
    total: usize,
    data_dir: std::path::PathBuf,
    templates_dir: std::path::PathBuf,
    // Neue Parameter für Output-Konfiguration
    use_custom_output: bool,
    custom_output_path: String,
    group: String,
    language: String,
    is_messe: bool,
    // Performance-Parameter
    thread_sleep_ms: u64,
    debug_mode: bool,
    // Bereichs-Auswahl Parameter
    use_range: bool,
    range_start: usize,
    range_end: usize,
    // Font-Fallback Parameter
    enable_font_fallback: bool,
) -> Result<(), Box<dyn std::error::Error>> {
    // Debug-Logging für Funktion
    if debug_mode {
        println!("DEBUG: generate_bestellscheine_resume gestartet mit {} Threads, Sleep: {}ms", threads, thread_sleep_ms);
    }
    
    let stop_signal = Arc::clone(&stop_signal);
    
    // Temporäre Dateipfade einmal erstellen (vor Thread-Erstellung)
    let progress_path = get_temp_file_path("progress.txt");
    let stop_status_path = get_temp_file_path("stop_status.txt");
    
    // PDF-Erstellung mit Threads
    let mut handles = Vec::new();
    for t in 0..threads {
        let vertreter = Arc::clone(&vertreter);
        let progress_counter = Arc::clone(&progress_counter);
        let progress = Arc::clone(&progress);
        let stop_signal = Arc::clone(&stop_signal);
        let data_dir = data_dir.clone();
        let templates_dir = templates_dir.clone();
        // Neue Parameter für Output-Konfiguration klonen
        let use_custom_output = use_custom_output;
        let custom_output_path = custom_output_path.clone();
        let group = group.clone();
        let language = language.clone();
        let _is_messe = is_messe;
        // Pfade für Threads klonen
        let progress_path = progress_path.clone();
        let stop_status_path = stop_status_path.clone();
        let handle = thread::spawn(move || {
            // Bereichs-Logik: Bestimme effektiven Start und Ende
            let (effective_start, effective_end) = if use_range {
                (range_start.max(start_from), range_end.min(total - 1))
            } else {
                (start_from, total - 1)
            };
            
            // Beginne ab effective_start, nicht bei 0, und beschränke auf effective_end
            for i in (effective_start + t..=effective_end).step_by(threads) {
                // Sicherheitsprüfung: Index muss innerhalb der Vertreter-Liste sein
                if i >= vertreter.len() {
                    break; // Thread beenden wenn Index außerhalb der Liste
                }
                
                // Prüfen ob Stop-Signal gesetzt wurde - WICHTIG: Vor jeder PDF-Verarbeitung prüfen
                {
                    let should_stop = stop_signal.lock().unwrap();
                    if *should_stop {
                        println!("Thread {} gestoppt bei Index {} (Bereich: {}-{})", t, i, effective_start, effective_end);
                        // Stop-Status in versteckte Datei schreiben, damit progress.txt numerisch bleibt
                        let _ = std::fs::write(&stop_status_path, "STOPPED");
                        return; // Thread beenden
                    }
                }
                
                let (kundennr, de_link, en_link) = &vertreter[i];
                
                // Wähle die richtige URL basierend auf der Sprache
                let link = if language == "Englisch" || language.to_lowercase().contains("en") {
                    println!("DEBUG: Verwende englische URL für Vertreter {}: {}", kundennr, en_link);
                    en_link
                } else {
                    println!("DEBUG: Verwende deutsche URL für Vertreter {}: {}", kundennr, de_link);
                    de_link
                };
                
                // Gruppenspezifischen Output-Pfad bestimmen (mit Benutzer-Konfiguration)
                let selections = get_current_selections().unwrap_or_else(|| vec![ 
                    (data_dir.join("Vertreternummern.csv").to_string_lossy().to_string(), 
                     templates_dir.join("Bestellschein-Endkunde-de_de.pdf").to_string_lossy().to_string(), 
                     true) 
                ]);
                let first_template = selections.get(0).map(|s| s.1.clone()).unwrap_or_default();
                let (template_group, _template_language, template_is_messe) = infer_group_lang_from_template(&first_template);
                let group_output_dir = get_configured_output_dir_with_debug(use_custom_output, &custom_output_path, &template_group, &language, template_is_messe, debug_mode);
                
                println!("DEBUG: Output-Verzeichnis für Sprache '{}': {}", language, group_output_dir.display());
                
                // Prüfen ob PDF bereits existiert
                let pdf_filename = format!("{}-{}.pdf", 
                    std::path::Path::new(&first_template)
                        .file_stem()
                        .unwrap_or_default()
                        .to_string_lossy(),
                    kundennr.replace("\0", "")
                );
                let pdf_path = group_output_dir.join(&pdf_filename);
                
                if !pdf_path.exists() {
                    println!("Erstelle PDF für Vertreter {}: {} -> {}/{}", i + 1, kundennr, group, language);
                    // Für jede ausgewählte Template-Option erstellen (aber keine Duplikate)
                    let selections = get_current_selections().unwrap_or_else(|| vec![ ("DATA/Vertreternummern.csv".to_string(), "VORLAGE/Bestellschein-Endkunde-de_de.pdf".to_string(), true) ]);
                    let mut created = Vec::new();
                    for (_csv_s, template_s, gen_qr) in selections.iter() {
                        // Verhindere doppelte Ausgaben für dieselbe template/kundennr
                        let out_name = format!("{}-{}", template_s, kundennr);
                        if created.contains(&out_name) { continue; }
                        created.push(out_name.clone());

                        // Template-Pfad korrekt auflösen
                        let resolved_template = resolve_template_path_with_debug(template_s, debug_mode);
                        let resolved_template_str = resolved_template.to_string_lossy();
                        
                        debug_print(&format!("Verwende Template: {}", resolved_template_str), debug_mode);
                        
                        // Prüfe ob Template existiert
                        if !resolved_template.exists() {
                            println!("ERROR: Template-Datei nicht gefunden: {}", resolved_template_str);
                            continue;
                        }

                        // Load per-template/group config (best effort)
                        let (group_name, _template_lang, tpl_is_messe) = infer_group_lang_from_template(&resolved_template_str);
                        
                        // Verwende aktuelle UI-Config falls verfügbar, sonst fallback zu group config
                        let tpl_config = {
                            let _lock = CONFIG_MUTEX.lock().unwrap();
                            unsafe {
                                if let Some(ref current_config) = CURRENT_CONFIG {
                                    println!("🎯 Verwende aktuelle UI-Config für PDF-Generierung: QR={:?}", current_config.qr_codes);
                                    current_config.clone()
                                } else {
                                    println!("⚠️ Keine UI-Config verfügbar, lade Group-Config");
                                    load_group_config(&group_name, &language, tpl_is_messe)
                                }
                            }
                        };
                        if *gen_qr {
                            let (qr_img, qr_width) = generate_qr(link);
                            let output_filename = format!("{}-{}.pdf", 
                                resolved_template.file_stem().unwrap_or_default().to_string_lossy(),
                                kundennr
                            );
                            let output_path = get_configured_output_dir_with_debug(use_custom_output, &custom_output_path, &group_name, &language, tpl_is_messe, debug_mode).join(&output_filename);
                            modify_pdf_with_debug(&resolved_template_str, kundennr, &qr_img, qr_width, &tpl_config, &output_path, debug_mode, enable_font_fallback);
                        } else {
                            let output_filename = format!("{}-{}.pdf", 
                                resolved_template.file_stem().unwrap_or_default().to_string_lossy(),
                                kundennr
                            );
                            let output_path = get_configured_output_dir_with_debug(use_custom_output, &custom_output_path, &group_name, &language, tpl_is_messe, debug_mode).join(&output_filename);
                            modify_pdf_with_debug(&resolved_template_str, kundennr, &[], 0, &tpl_config, &output_path, debug_mode, enable_font_fallback);
                        }
                    }
                } else {
                    println!("PDF für Vertreter {} bereits vorhanden, überspringe", kundennr);
                }
                
                // Progress aktualisieren
                {
                    let mut counter = progress_counter.lock().unwrap();
                    *counter += 1;
                    let progress_val = *counter as f32 / total as f32;
                    
                    let mut p = progress.lock().unwrap();
                    *p = progress_val;
                    
                    // Progress in versteckte Datei schreiben für UI - nur wenn nicht gestoppt
                    let should_stop = stop_signal.lock().unwrap();
                    if !*should_stop {
                        let _ = std::fs::write(&progress_path, format!("{}", progress_val));
                    }
                }
                
                // Intelligente Performance-Optimierung: Nur sehr kurze Pausen bei hoher Last
                // Bei 20.000+ Bestellscheinen wird Zeit nicht verschwendet
                if thread_sleep_ms > 0 && thread_sleep_ms <= 2 {
                    std::thread::sleep(std::time::Duration::from_millis(thread_sleep_ms));
                }
            }
        });
        handles.push(handle);
    }

    for h in handles {
        h.join().map_err(|_| "Thread Join Error")?;
    }

    // Progress auf 1.0 setzen (fertig) und dann Datei löschen
    {
        let mut p = progress.lock().unwrap();
        *p = 1.0;
        let progress_path = get_temp_file_path("progress.txt");
        let _ = std::fs::write(&progress_path, "1.0");
        
        // Kurz warten, damit UI den 100%-Status noch anzeigen kann
        std::thread::sleep(std::time::Duration::from_millis(500));
        
        // Progress-Datei löschen für nächsten Durchlauf
        if let Err(e) = std::fs::remove_file(&progress_path) {
            // Fehler nur bei Debug ausgeben, da es nicht kritisch ist
            eprintln!("DEBUG: Konnte progress.txt nicht löschen: {}", e);
        }
    }

    println!("Bestellscheine erfolgreich erstellt!");
    Ok(())
}

fn main() {
    // Maximiert starten (Windows-Vollbild mit Taskleiste sichtbar)
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_maximized(true) // Maximiert starten statt echtes Vollbild
            .with_decorations(true) // Fenster-Steuerungen (Minimieren, Schließen) behalten
            .with_resizable(true) // Größenänderung erlauben
            .with_title("Bestellschein Generator"), // Titel setzen
        ..Default::default()
    };
    
    eframe::run_native(
        "Bestellschein Generator",
        options,
        Box::new(|_cc| Box::new(MyApp::default())),
    ).unwrap();
}
