#![windows_subsystem = "windows"]

// Bestellscheine mit integrierter UI
use eframe::egui;
use eframe::App;
use std::sync::{Arc, Mutex};
use std::thread;
use std::fs;
use std::sync::atomic::{AtomicBool, Ordering};
use lopdf::{Document, content::{Content, Operation}, dictionary, Object};
use qrcode::QrCode;

// Funktion um PDF-Vorlage zu laden und als Vorschau zu erstellen
// NOTE: aktuell nicht in der schnellen Config-Ansicht verwendet; bleibt für optionalen Full-PDF-Preview erhalten
#[allow(dead_code)]
fn load_actual_template_for_group(group: &str, language: &str, is_messe: bool) -> Option<(String, lopdf::Document)> {
    let exe_dir = std::env::current_exe()
        .ok()
        .and_then(|p| p.parent().map(|p| p.to_path_buf()))
        .unwrap_or_else(|| std::env::current_dir().unwrap_or_default());
    
    // Try to find template using find_best_template
    if let Some(template_name) = find_best_template(group, language, if is_messe { Some("messe") } else { None }) {
        let template_path = exe_dir.join(&template_name);
        if template_path.exists() {
            if let Ok(doc) = Document::load(&template_path) {
                return Some((template_name, doc));
            }
        }
    }
    
    // Fallback templates (verwende VORLAGE/-Ordner)
    let fallback_templates = if is_messe {
        vec![
            format!("VORLAGE/Bestellschein-{}-{}_messe.pdf", group, language.to_lowercase()),
            format!("VORLAGE/Bestellschein-{}-messe.pdf", group),
            "VORLAGE/Bestellschein-messe.pdf".to_string(),
        ]
    } else {
        vec![
            format!("VORLAGE/Bestellschein-{}-{}.pdf", group, language.to_lowercase()),
            format!("VORLAGE/Bestellschein-{}.pdf", group),
            "VORLAGE/Bestellschein.pdf".to_string(),
        ]
    };
    
    for template in &fallback_templates {
        let template_path = exe_dir.join(template);
        if template_path.exists() {
            if let Ok(doc) = Document::load(&template_path) {
                return Some((template.clone(), doc));
            }
        }
    }
    
    None
}

// Funktion um PDF-Vorlage zu laden und als Vorschau zu erstellen  
fn load_template_preview() -> Option<String> {
    // Try to use current group selection
    let exe_dir = std::env::current_exe()
        .ok()
        .and_then(|p| p.parent().map(|p| p.to_path_buf()))
        .unwrap_or_else(|| std::env::current_dir().unwrap_or_default());
    
    // Fallback: Try to find any available template in VORLAGE directory
    let vorlage_dir = exe_dir.join("VORLAGE");
    println!("Suche Template in: {:?}", vorlage_dir);
    
    if let Ok(entries) = std::fs::read_dir(&vorlage_dir) {
        for entry in entries.flatten() {
            let path = entry.path();
            if let Some(ext) = path.extension() {
                if ext == "pdf" {
                    println!("Gefunden: {:?}", path);
                    if let Ok(_doc) = Document::load(&path) {
                        let filename = path.file_name().unwrap().to_string_lossy();
                        return Some(format!("PDF-Vorlage '{}' erfolgreich geladen", filename));
                    }
                }
            }
        }
    }
    
    println!("Keine PDF-Vorlagen gefunden in {:?}", vorlage_dir);
    None
}

// Debug-Logging-Funktion (nur wenn Debug-Modus aktiv)
fn debug_log(message: &str, debug_enabled: bool) {
    if debug_enabled {
        let timestamp = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs();
        
        let log_entry = format!("[{}] {}\n", timestamp, message);
        let log_path = get_temp_file_path("debug.log");
        
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

// Globaler Debug-Flag, ermöglicht Debug-Ausgaben auch in Funktionen ohne lokalen Flag-Parameter
static GLOBAL_DEBUG: Lazy<AtomicBool> = Lazy::new(|| AtomicBool::new(false));

fn debug_print_global(message: &str) {
    if GLOBAL_DEBUG.load(Ordering::Relaxed) {
        println!("DEBUG: {}", message);
        debug_log(message, true);
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
    
    // Unterstütze als Input sowohl UI-Namen ("Deutsch", "Englisch") als auch Sprachcodes ("de_de", "en_us")
    let lang_low = language.to_lowercase();
    let language_folder = if lang_low.starts_with("en") || lang_low.contains("engl") || lang_low == "english" {
        "EN"
    } else {
        "DE"
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

// Robustere Spracherkennung: Liefert einen kanonischen Sprachcode wie "de_de" oder "en_us".
// Reihenfolge: 1) Template-Dateiname (falls angegeben) 2) CSV-Dateiname 3) UI-Auswahl
fn detect_language_code(ui_language: &str, template_path: Option<&str>, csv_path: Option<&str>) -> String {
    // Hilfsfunktion zur Normalisierung einfacher Tokens
    fn norm_token(tok: &str) -> Option<String> {
        let t = tok.to_lowercase();
        // Beispiele: de, de_de, en, en_us, german, deutsch, english, englisch
        if t.starts_with("de") || t.contains("deutsch") || t.contains("german") {
            return Some("de_de".to_string());
        }
        if t.starts_with("en") || t.contains("engl") || t.contains("english") {
            return Some("en_us".to_string());
        }
        None
    }

    // 1) Versuch: aus Template-Dateiname extrahieren (letztes '-' Segment, außer 'messe')
    if let Some(tp) = template_path {
        let stem = std::path::Path::new(tp)
            .file_stem()
            .and_then(|s| s.to_str())
            .unwrap_or(tp);
        let parts: Vec<&str> = stem.split('-').collect();
        if !parts.is_empty() {
            // Suche von hinten: häufig ist letzter Teil der Sprachcode (z.B. de_de, en_us)
            for part in parts.iter().rev() {
                let p = part.trim();
                if p.eq_ignore_ascii_case("messe") { continue; }
                if let Some(n) = norm_token(p) { return n; }
                // Manche Dateien nutzen formate wie "de_de" oder "en_us"
                if p.contains('_') {
                    let tokens: Vec<&str> = p.split('_').collect();
                    if let Some(lang_tok) = tokens.get(0) {
                        if let Some(n2) = norm_token(lang_tok) { return n2; }
                    }
                }
            }
        }
    }

    // 2) Versuch: CSV-Pfad auswerten (z.B. names-en.csv)
    if let Some(cp) = csv_path {
        let fname = std::path::Path::new(cp)
            .file_name()
            .and_then(|s| s.to_str())
            .unwrap_or(cp)
            .to_lowercase();
        if fname.contains("_en") || fname.contains("english") || fname.contains("-en") { return "en_us".to_string(); }
        if fname.contains("_de") || fname.contains("deutsch") || fname.contains("-de") { return "de_de".to_string(); }
    }

    // 3) Fallback: UI-Auswahl
    if let Some(n) = norm_token(ui_language) { return n; }

    // Default
    "de_de".to_string()
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
    /// X-Position in PDF-Punkten (von links)
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
            qr_codes: vec![QrCodeConfig { x: 50.0, y: 50.0, size: 18.0, pages: vec![1], all_pages: false }],
            vertreter: vec![
                VertreterConfig { x: 77.0, y: 80.0, size: 12.0, pages: vec![1], all_pages: false, font_name: "Arial".to_string(), font_size: 12.0, font_style: "Normal".to_string() },
                VertreterConfig { x: 100.0, y: 650.0, size: 10.0, pages: vec![1], all_pages: false, font_name: "Arial".to_string(), font_size: 10.0, font_style: "Normal".to_string() }
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
                // Apo Messe - andere Positionen
                Config {
                    qr_codes: vec![QrCodeConfig { x: 80.0, y: 70.0, size: 22.0, pages: vec![1, 2], all_pages: false }],
                    vertreter: vec![
                        VertreterConfig { x: 120.0, y: 100.0, size: 14.0, pages: vec![1, 2], all_pages: false, font_name: "Arial".to_string(), font_size: 14.0, font_style: "Normal".to_string() },
                        VertreterConfig { x: 150.0, y: 700.0, size: 12.0, pages: vec![1], all_pages: false, font_name: "Arial".to_string(), font_size: 12.0, font_style: "Normal".to_string() }
                    ],
                }
            } else {
                // Apo Normal - optimiert für Apotheken-Formulare
                Config {
                    qr_codes: vec![QrCodeConfig { x: 75.0, y: 60.0, size: 20.0, pages: vec![1], all_pages: false }],
                    vertreter: vec![
                        VertreterConfig { x: 100.0, y: 90.0, size: 14.0, pages: vec![1], all_pages: false, font_name: "Arial".to_string(), font_size: 14.0, font_style: "Normal".to_string() },
                        VertreterConfig { x: 130.0, y: 680.0, size: 12.0, pages: vec![1], all_pages: false, font_name: "Arial".to_string(), font_size: 12.0, font_style: "Normal".to_string() }
                    ],
                }
            }
        },
        "endkunde" | "endnutzer" => {
            if is_messe {
                // Endkunde Messe - angepasst für Messestände
                Config {
                    qr_codes: vec![QrCodeConfig { x: 60.0, y: 80.0, size: 24.0, pages: vec![1], all_pages: false }],
                    vertreter: vec![
                        VertreterConfig { x: 90.0, y: 120.0, size: 12.0, pages: vec![1], all_pages: false, font_name: "Arial".to_string(), font_size: 12.0, font_style: "Normal".to_string() },
                        VertreterConfig { x: 120.0, y: 720.0, size: 10.0, pages: vec![1], all_pages: false, font_name: "Arial".to_string(), font_size: 10.0, font_style: "Normal".to_string() }
                    ],
                }
            } else {
                // Endkunde Normal - Standard-Layout
                Config {
                    qr_codes: vec![QrCodeConfig { x: 50.0, y: 50.0, size: 18.0, pages: vec![1], all_pages: false }],
                    vertreter: vec![
                        VertreterConfig { x: 77.0, y: 80.0, size: 12.0, pages: vec![1], all_pages: false, font_name: "Arial".to_string(), font_size: 12.0, font_style: "Normal".to_string() },
                        VertreterConfig { x: 100.0, y: 650.0, size: 10.0, pages: vec![1], all_pages: false, font_name: "Arial".to_string(), font_size: 10.0, font_style: "Normal".to_string() }
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

    // Defaults setzen wenn nichts gefunden wurde
    if qr_codes.is_empty() {
        qr_codes.push(QrCodeConfig { x: 50.0, y: 50.0, size: 18.0, pages: vec![1], all_pages: false });
    }
    if vertreter.is_empty() {
        vertreter = vec![
            VertreterConfig { x: 77.0, y: 80.0, size: 12.0, pages: vec![1], all_pages: false, font_name: "Arial".to_string(), font_size: 12.0, font_style: "Normal".to_string() },
            VertreterConfig { x: 100.0, y: 650.0, size: 10.0, pages: vec![1], all_pages: false, font_name: "Arial".to_string(), font_size: 10.0, font_style: "Normal".to_string() },
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
    toml.push_str("# Koordinaten sind in PDF-Punkten (1 Punkt = 1/72 Zoll ≈ 0.35mm)\n\n");
    
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
    
    // If we have a previously loaded config path, prefer saving back to it
    if let Some(p) = get_current_config_path() {
        if let Err(e) = std::fs::write(&p, toml.clone()) {
            eprintln!("Konnte gruppenspezifische Config nicht in geladenem Pfad {:?} speichern: {}", p, e);
        } else {
            println!("✅ Gruppenspezifische Config gespeichert (geladene Datei): {:?}", p);
            println!("Config-Werte: QR={:?}, Vertreter={:?}", config.qr_codes, config.vertreter);
            return;
        }
    }

    // Fallback: write to default group filename
    if let Err(e) = std::fs::write(&group_filename, toml) {
        eprintln!("Konnte gruppenspezifische Config nicht speichern: {}", e);
    } else {
        println!("✅ Gruppenspezifische Config gespeichert: {:?}", group_filename);
        println!("Config-Werte: QR={:?}, Vertreter={:?}", 
                 config.qr_codes, config.vertreter);
    }
}

// Helper to explicitly save to a path (used by UI when user chooses "In Datei speichern")
fn save_group_config_to_path(path: &std::path::Path, config: &Config) -> Result<(), std::io::Error> {
    let mut toml = String::new();
    toml.push_str("# Manuell gespeicherte Config\n");
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
    toml.push_str("]\n");
    std::fs::write(path, toml)
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
    // Font-Caching für Performance
    cached_fonts: Vec<String>, // Gecachte Font-Liste
    // Bereichs-Auswahl für Vertreternummern
    use_range_selection: bool,  // Ob Bereichs-Auswahl aktiviert ist
    range_start_index: String,  // Start-Index (0-basiert)
    range_end_index: String,    // End-Index (0-basiert)
    // Resume-Information
    resume_info: Option<(usize, usize, u64)>, // (current_index, total_count, elapsed_seconds)
    // Selected element in the preview: kind ("qr"/"vertreter") and index
    selected_element: Option<(String, usize)>,
    // UI helpers for saving config
    config_save_path: String,
    show_save_as: bool,
    // Full PDF preview state (removed - kept preview lightweight)
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

        // Überprüfe ob Template verfügbar ist
        let _template_loaded = load_template_preview().is_some();

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
            // Font-Caching für Performance
            cached_fonts: refresh_font_cache(), // Einmalig beim Start laden (mit Cache)
            // Bereichs-Auswahl Defaults
            use_range_selection: false,
            range_start_index: String::new(),
            range_end_index: String::new(),
            // Resume-Information (gruppenspezifisch beim Start geladen)
            resume_info: initial_resume_info,
            selected_element: None,
            config_save_path: String::new(),
            show_save_as: false,
            // preview state removed
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

        // Exakte Kombination Gruppe+Code (z.B. bestellschein-endkunde-en_us) besonders belohnen
        let preferred_codes = get_preferred_language_codes(lang);
        if let Some(primary) = preferred_codes.get(0) {
            // primary may be like "en_us" -> prefix "en"
            let prefix = primary.split('_').next().unwrap_or(primary).to_string();
            let group_code_exact1 = format!("{}-{}", group.to_lowercase(), primary.replace(' ', ""));
            let group_code_exact2 = format!("{}-{}", group.to_lowercase(), prefix);
            let group_code_exact3 = format!("{}_{}", group.to_lowercase(), prefix);
            if filename_lower.contains(&group_code_exact1) || filename_lower.contains(&group_code_exact2) || filename_lower.contains(&group_code_exact3) {
                score += 40;
            }
        }
        
        // Sprache passt (verwende robusteren Matcher)
        score += language_match_score(&filename_lower, lang);
        
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
                                score += language_match_score(&filename_lower, lang);
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
    
        // Zuerst vorhandene Templates bevorzugen, dann nach Score (höchster zuerst)
    results.sort_by(|a, b| {
        // a.2 and b.2 sind 'exists' bools
        if a.2 == b.2 {
            b.1.cmp(&a.1)
        } else {
            // existierende Dateien zuerst
            b.2.cmp(&a.2)
        }
    });
    results
}

fn capitalize_first(s: &str) -> String {
    let mut cs = s.chars();
    match cs.next() {
        None => String::new(),
        Some(f) => f.to_uppercase().collect::<String>() + cs.as_str(),
    }
}

// Liefert eine priorisierte Liste von Sprachecodes (z.B. ["en_us","en"]) für die gewünschte Sprache
fn get_preferred_language_codes(request: &str) -> Vec<String> {
    let r = request.to_lowercase();
    if r.contains("en") || r.contains("engl") || r.contains("english") {
        vec!["en_us".to_string(), "en".to_string()]
    } else if r.contains("fr") || r.contains("franz") || r.contains("french") {
        vec!["fr_fr".to_string(), "fr".to_string()]
    } else if r.contains("de") || r.contains("deut") || r.contains("german") || r.contains("deutsch") {
        vec!["de_de".to_string(), "de".to_string()]
    } else if r.contains('_') {
        vec![r]
    } else {
        vec!["de_de".to_string(), "de".to_string()]
    }
}

// Liefert mögliche Varianten eines Sprachkürzels, z.B. für "Deutsch" -> ["de_de","de","DE"]
fn get_language_code_variants(lang: &str) -> Vec<String> {
    let r = lang.to_lowercase();
    if r.contains("en") || r.contains("engl") || r.contains("english") {
        vec!["en_us".to_string(), "en".to_string(), "EN".to_string()]
    } else if r.contains("fr") || r.contains("franz") || r.contains("french") {
        vec!["fr_fr".to_string(), "fr".to_string(), "FR".to_string()]
    } else if r.contains("de") || r.contains("deut") || r.contains("german") || r.contains("deutsch") {
        vec!["de_de".to_string(), "de".to_string(), "DE".to_string()]
    } else if r.len() == 2 {
        // Two-letter code provided
        vec![r.clone(), r.to_uppercase()]
    } else if r.contains('_') {
        vec![r.clone()]
    } else {
        // Fallback: try lower/upper two-letter and full
        let two = r.chars().take(2).collect::<String>();
        vec![format!("{}_{}", two, two), two.clone(), two.to_uppercase()]
    }
}

// Prüfe ob ein Token isoliert im Dateinamen vorkommt (z.B. -en-, _en_, en. oder am Ende/Anfang),
// ohne dass es Teil eines größeren Wortes ist (vermeidet Matches in "endkunde").
fn isolated_token_present(haystack: &str, token: &str) -> bool {
    if token.is_empty() { return false; }
    let hay = haystack;
    let t = token;
    let mut start = 0usize;
    while let Some(idx) = hay[start..].find(t) {
        let pos = start + idx;
        // vorheriges Zeichen
        let prev_ok = if pos == 0 { true } else { !hay.as_bytes()[pos - 1].is_ascii_alphabetic() };
        // nachfolgendes Zeichen
        let after_pos = pos + t.len();
        let next_ok = if after_pos >= hay.len() { true } else { !hay.as_bytes()[after_pos].is_ascii_alphabetic() };
        if prev_ok && next_ok { return true; }
        start = pos + 1;
    }
    false
}

// Berechne Sprach-Matching-Score für einen Dateinamen.
// Stellt sicher, dass exakte Sprachcodes (z.B. "en_us") höher bewertet werden
// und vermeidet falsche Treffer durch einfache Substring-Suchen (z.B. "de" in "default").
fn language_match_score(filename_lower: &str, lang: &str) -> i32 {
    // (moved to top-level) See get_preferred_language_codes

    // Known language codes - extend here when adding new languages
    let known_codes: Vec<&str> = vec!["en_us", "en", "de_de", "de", "fr_fr", "fr"];
    let preferred = get_preferred_language_codes(lang);

    let mut score = 0;

    for code in known_codes.iter() {
        // check for precise tokens: -code, _code, code. or -code-
        let token_precise = filename_lower.contains(&format!("-{}", code))
            || filename_lower.contains(&format!("_{}", code))
            || filename_lower.contains(&format!("{}.", code))
            || filename_lower.contains(&format!("-{}-", code));

        let token_loose = filename_lower.contains(code);

        if preferred.iter().any(|p| p == code) {
            // higher weight for preferred codes; earlier preferred codes get slightly more
            let pos = preferred.iter().position(|c| c == code).unwrap_or(0) as i32;
            let base = 30 - pos * 4; // 30, 26, ... for decreasing preference
            if token_precise { score += base; }
            else if token_loose { score += base / 2; }
        } else {
            // Not preferred: penalize precise other-language tags so they don't tie with preferred
            if token_precise { score -= 8; }
            else if token_loose { score -= 2; }
        }
    }

    score
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
    let lang_variants = get_language_code_variants(language);
    if is_messe {
        // Try variants with '-' and '_' separators and different cases
        for lv in &lang_variants {
            candidates.push(config_dir.join(format!("config_{}-{}_messe.toml", group.to_lowercase(), lv)));
            candidates.push(config_dir.join(format!("config_{}_{}_messe.toml", group.to_lowercase(), lv)));
        }
        candidates.push(config_dir.join(format!("config_{}_messe.toml", group.to_lowercase())));
        candidates.push(config_dir.join("config_messe.toml"));
    }
    for lv in &lang_variants {
        candidates.push(config_dir.join(format!("config_{}-{}.toml", group.to_lowercase(), lv)));
        candidates.push(config_dir.join(format!("config_{}_{}.toml", group.to_lowercase(), lv)));
    }
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
                // Remember the exact file path we loaded so saves write back to the same file
                set_current_config_path(c);
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
        // Prefer using the scored template list so automatic selection respects preferred language codes
        let templates = self.find_templates_with_score(group, lang, country.is_some() && country.unwrap_or("") == "messe");

        if !templates.is_empty() {
            // preferred codes (e.g. en_us, en)
            let preferred = get_preferred_language_codes(lang);

            // 1) try to find existing template that contains preferred code (primary or prefix)
            for (t, _s, exists) in templates.iter() {
                if !*exists { continue; }
                let t_low = t.to_lowercase();
                for p in &preferred {
                    let prefix = p.split('_').next().unwrap_or(p);
                    if t_low.contains(p) || t_low.contains(prefix) {
                        return Some(t.clone());
                    }
                }
            }

            // 2) otherwise return highest-scoring existing template
            for (t, _s, exists) in templates.iter() {
                if *exists { return Some(t.clone()); }
            }

            // 3) fallback to first candidate (even if missing)
            return templates.first().map(|(t, _s, _e)| t.clone());
        }

        // If no scored templates (edge case), fallback to original search
        if self.use_custom_template_dir {
            find_best_template_in_dir(group, lang, country, &self.custom_template_dir)
        } else {
            find_best_template(group, lang, country)
        }
    }

    // Wähle automatisch das beste Template (wie in der UI angezeigt):
    // 1) Vorzugsweise vorhandene Template mit preferred language code
    // 2) sonst erstes vorhandenes Template
    // 3) sonst erstes Kandidaten-Template
    fn select_best_template_auto(&self, group: &str, lang: &str, is_messe: bool) -> Option<String> {
        let templates = self.find_templates_with_score(group, lang, is_messe);
        if templates.is_empty() { return None; }

        let preferred_codes = get_preferred_language_codes(lang);

        // 1) existing template containing preferred code
        for (t, _s, exists) in templates.iter() {
            if !*exists { continue; }
            let t_low = t.to_lowercase();
                for p in &preferred_codes {
                    let prefix = p.split('_').next().unwrap_or(p);
                    if isolated_token_present(&t_low, p) || isolated_token_present(&t_low, prefix) {
                        return Some(t.clone());
                    }
                }
        }

        // 2) first existing template
        for (t, _s, exists) in templates.iter() {
            if *exists { return Some(t.clone()); }
        }

        // 3) first candidate (even if missing)
        templates.first().map(|(t, _s, _e)| t.clone())
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
                                
                                // Höhere Priorität für exakte Matches (Gruppe+Lang) - benutze primären Preferred-Code
                                let preferred_codes = get_preferred_language_codes(lang);
                                if let Some(primary) = preferred_codes.get(0) {
                                    if filename_lower.contains(&format!("{}-{}", group.to_lowercase(), primary.replace(' ', ""))) {
                                        score += 40;
                                    }
                                }
                                
                                results.push((relative_path, score, true)); // exists = true (wir haben es gescannt)
                            }
                        }
                    }
                }
            }
        }
        
    // Zuerst vorhandene Templates bevorzugen, dann nach Score (höchster zuerst)
    results.sort_by(|a, b| {
        if a.2 == b.2 {
            b.1.cmp(&a.1)
        } else {
            b.2.cmp(&a.2)
        }
    });
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

        // Keyboard nudging for selected preview element: Arrow keys move selected element.
        // Use 1 UI-pixel per press as base, Shift => 10 pixels. Convert UI pixels to PDF points
        // using the same scale as the preview: PDF A4 ~595x842 mapped to ui a4 size.
        let a4_width = 350.0_f32;
        let a4_height = 495.0_f32;
        let scale_x = 595.0_f32 / a4_width; // PDF points per UI pixel horizontally
        let scale_y = 842.0_f32 / a4_height; // PDF points per UI pixel vertically

        let ui_move_pixels = if ctx.input(|i| i.modifiers.shift) { 10.0 } else { 1.0 };

        // Movement amounts in PDF points
        let pdf_move_x = ui_move_pixels * scale_x;
        let pdf_move_y = ui_move_pixels * scale_y;

        // Arrow keys: allow continuous movement while key is down
    let left = ctx.input(|i| i.key_down(egui::Key::ArrowLeft));
    let right = ctx.input(|i| i.key_down(egui::Key::ArrowRight));
    let up = ctx.input(|i| i.key_down(egui::Key::ArrowUp));
    let down = ctx.input(|i| i.key_down(egui::Key::ArrowDown));

        if left || right || up || down {
            if let Some((ref kind, idx)) = self.selected_element.clone() {
                let dx = if left { -pdf_move_x } else if right { pdf_move_x } else { 0.0 };
                let dy = if up { pdf_move_y } else if down { -pdf_move_y } else { 0.0 };

                if kind == "qr" {
                    if let Some(qr) = self.config.qr_codes.get_mut(idx) {
                        qr.x = (qr.x + dx).max(0.0).min(595.0 - qr.size);
                        qr.y = (qr.y + dy).max(0.0).min(842.0 - qr.size);
                        // update manual fields for immediate feedback
                        if idx == 0 {
                            self.manual_qr_x = format!("{:.1}", qr.x);
                            self.manual_qr_y = format!("{:.1}", qr.y);
                            self.manual_qr_size = format!("{:.1}", qr.size);
                        }
                    }
                } else if kind == "vertreter" {
                    if let Some(v) = self.config.vertreter.get_mut(idx) {
                        v.x = (v.x + dx).max(0.0).min(595.0 - v.size);
                        v.y = (v.y + dy).max(0.0).min(842.0 - v.size);
                        // update manual vertreter fields only for first
                        if idx == 0 {
                            self.manual_vertreter_x = format!("{:.1}", v.x);
                            self.manual_vertreter_y = format!("{:.1}", v.y);
                            self.manual_vertreter_size = format!("{:.1}", v.size);
                        }
                    }
                }
            }
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
                                debug_print_global(&format!("Verwende gespeicherte Auswahl: {:?}", s));
                                s
                            },
                            None => {
                                debug_print_global("Keine Auswahl gefunden, verwende Standard");
                                get_default_selections()
                            }
                        };
                        
                        let csv_path = selections.get(0).map(|s| s.0.clone()).unwrap_or_else(|| get_default_csv_path("Endkunde"));
                        debug_print_global(&format!("CSV-Pfad: {}", csv_path));
                        
                        // Prüfe ob CSV-Datei existiert
                        if !std::path::Path::new(&csv_path).exists() {
                            self.status_message = format!("FEHLER: CSV-Datei nicht gefunden: {}", csv_path);
                            println!("ERROR: CSV-Datei nicht gefunden: {}", csv_path);
                            return;
                        }
                        
                        let vertreter_vec = match std::panic::catch_unwind(|| read_vertreter(&csv_path)) {
                            Ok(vertreter) => {
                                debug_print_global(&format!("{} Vertreter geladen", vertreter.len()));
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
                                debug_print_global("Ordner erfolgreich ermittelt");
                                debug_print_global(&format!("Data-Dir: {}", dirs.1.display()));
                                debug_print_global(&format!("Templates-Dir: {}", dirs.2.display()));
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
                        if ui.selectable_label(self.selected_group == "Fachkreise", "Fachkreise").clicked() {
                            // Alte Progress-Dateien löschen bei Kategorie-Wechsel
                            clear_progress_files(&self.selected_group, &self.selected_language, self.is_messe);
                            self.selected_group = "Fachkreise".to_string();
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
                            println!("Gruppe geändert zu Fachkreise - Config neu geladen");
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
                            // Wähle automatisch: bevorzuge eine existierende Vorlage, die einen preferred language code enthält
                            let preferred_codes = get_preferred_language_codes(&self.selected_language);
                            let mut chosen: Option<(&String, &i32, &bool)> = None;

                            // 1) Suche existierende Templates mit preferred code (primary or prefix)
                            for (t, s, e) in templates_with_score.iter() {
                                if !*e { continue; }
                                let t_low = t.to_lowercase();
                                let mut matches_pref = false;
                                for p in &preferred_codes {
                                    let prefix = p.split('_').next().unwrap_or(p);
                                    if isolated_token_present(&t_low, p) || isolated_token_present(&t_low, prefix) {
                                        matches_pref = true; break;
                                    }
                                }
                                if matches_pref {
                                    chosen = Some((t, s, e)); break;
                                }
                            }

                            // 2) Falls nicht gefunden: erstes existierendes Template
                            if chosen.is_none() {
                                for (t, s, e) in templates_with_score.iter() {
                                    if *e { chosen = Some((t, s, e)); break; }
                                }
                            }

                            // 3) Falls immer noch nicht gefunden: das erste Kandidat (auch wenn nicht existierend)
                            if chosen.is_none() {
                                if let Some((t, s, e)) = templates_with_score.first() {
                                    chosen = Some((t, s, e));
                                }
                            }

                            if let Some((template, score, exists)) = chosen {
                                let status = if *exists { "✅ gefunden" } else { "❌ fehlt" };
                                ui.label(format!("🤖 Automatische Wahl: {} (Score: {}) - {}", template, score, status));
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
                                // Automatische Erkennung verwenden (jetzt: best-scoring + preferred language)
                                    self.select_best_template_auto(&self.selected_group, &self.selected_language, self.is_messe)
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
                        ui.label(egui::RichText::new("Version 0.3.2").size(12.0));
                        ui.add_space(5.0);
                        ui.label(egui::RichText::new("Programmentwicklung: Alexander Löschke, IT - Abteilung")
                            .strong()
                            .color(egui::Color32::from_rgb(100, 149, 237))); // Cornflower Blue
                            
                        // Versteckter Debug/Performance-Bereich (nur sichtbar bei Ctrl+Shift+D)
                        // Stabile Tastenkombination mit Toggle
                        if ctx.input(|i| i.modifiers.ctrl && i.modifiers.shift && i.key_pressed(egui::Key::D)) {
                            if !self.debug_key_pressed {
                                self.debug_mode = !self.debug_mode; // Toggle Debug-Modus
                                // Setze globalen Debug-Flag damit zentrale Logs sichtbar werden
                                GLOBAL_DEBUG.store(self.debug_mode, Ordering::Relaxed);
                                save_debug_config(self.debug_mode); // Persistieren des Debug-Status
                                self.debug_key_pressed = true;
                            }
                        } else {
                            self.debug_key_pressed = false; // Reset wenn Tasten losgelassen
                        }
                        
                        if self.debug_mode {
                            ui.separator();
                            ui.label(egui::RichText::new("🔧 Erweiterte Einstellungen (Debug-Modus)").size(12.0).color(egui::Color32::RED));
                            
                            if ui.checkbox(&mut self.debug_mode, "Debug-Modus aktivieren (Log-Datei)").clicked() {
                                // Checkbox hat geändert: setze GLOBAL_DEBUG und persist
                                GLOBAL_DEBUG.store(self.debug_mode, Ordering::Relaxed);
                                save_debug_config(self.debug_mode);
                            }
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
                egui::Window::new("Positionen auf DIN A4 konfigurieren")
                    .open(&mut show_config)
                    .resizable(true)
                    .default_size([800.0, 600.0])
                    .show(ctx, |ui| {
                    
                    ui.horizontal(|ui| {
                        // Linke Seite - DIN A4 Darstellung
                        ui.vertical(|ui| {
                            ui.horizontal(|ui| {
                                ui.vertical(|ui| {
                                    ui.label("Ziehen Sie die Elemente an die gewünschte Position:");
                                    if ui.small_button("🔄 Neu laden").clicked() {
                                        self.config = load_group_config(&self.selected_group, &self.selected_language, self.is_messe);
                                        println!("Gruppenspezifische Config manuell neu geladen!");
                                    }
                                });

                                ui.separator();

                                ui.vertical(|ui| {
                                    ui.label("Aktuelle Config");
                                    if let Some(p) = get_current_config_path() {
                                        let s = p.display().to_string();
                                        let short = if s.len() > 48 { format!("...{}", &s[s.len()-45..]) } else { s };
                                        ui.monospace(short);
                                    } else {
                                        ui.monospace("(keine spezifische Datei)");
                                    }
                                });

                                ui.with_layout(egui::Layout::right_to_left(egui::Align::TOP), |ui| {
                                    ui.vertical(|ui| {
                                        ui.label(egui::RichText::new("Speichern").strong());
                                        ui.horizontal(|ui| {
                                            if ui.button("Speichern").clicked() {
                                                // Quick-save: try to save to loaded path, else fallback to group file
                                                save_group_config(&self.selected_group, &self.selected_language, self.is_messe, &self.config);
                                            }
                                            if ui.small_button("⇩").on_hover_text("Speichern unter...").clicked() {
                                                self.show_save_as = true;
                                                // Pre-fill suggestion
                                                if let Some(p) = get_current_config_path() {
                                                    self.config_save_path = p.display().to_string();
                                                } else {
                                                    let suggested = format!("CONFIG/config_{}-{}.toml", self.selected_group.to_lowercase(), self.selected_language.chars().take(2).collect::<String>());
                                                    self.config_save_path = suggested;
                                                }
                                            }
                                        });
                                    });
                                });
                            });
                            
                            // DIN A4 Verhältnis: 210mm x 297mm ≈ 1:1.414
                            let a4_width = 350.0;
                            let a4_height = 495.0; // 350 * 1.414
                            
                            // Bereich für DIN A4 Darstellung reservieren
                            let (a4_rect, _a4_response) = ui.allocate_exact_size(
                                egui::vec2(a4_width, a4_height), 
                                egui::Sense::drag()
                            );
                            
                            // DIN A4 Hintergrund mit erweiterten Vorlagen-Design zeichnen
                            ui.painter().rect_filled(a4_rect, 5.0, egui::Color32::WHITE);
                            ui.painter().rect_stroke(a4_rect, 5.0, egui::Stroke::new(2.0, egui::Color32::BLACK));
                            
                            // Lightweight: Prüfe, ob eine passende Template-Datei vorhanden ist (ohne das PDF zu laden)
                            let (_, _, templates_dir, _, _) = get_release_dirs_with_debug(self.debug_mode);
                            let mut template_found: Option<String> = None;
                            if let Ok(entries) = std::fs::read_dir(&templates_dir) {
                                for e in entries.flatten() {
                                    if let Some(fname) = e.file_name().to_str() {
                                        let fname_l = fname.to_lowercase();
                                        // grobe Suche: enthält Gruppenname und ggf. Sprachkürzel
                                        let grp = self.selected_group.to_lowercase();
                                        let lang_prefix = self.selected_language.chars().take(2).collect::<String>().to_lowercase();
                                        if fname_l.contains(&grp) && (fname_l.contains(&lang_prefix) || fname_l.contains(&"de") || fname_l.contains(&"en")) {
                                            template_found = Some(fname.to_string());
                                            break;
                                        }
                                    }
                                }
                            }
                            let (template_status, template_loaded) = if let Some(ref name) = template_found {
                                (format!("✅ Template vorhanden: {}", name), true)
                            } else {
                                (format!("⚠️ Keine Template-Datei für {} {} (Messe: {}) gefunden", self.selected_group, self.selected_language, self.is_messe), false)
                            };
                            
                            // CSV-Datei-Status überprüfen
                            let csv_path = if self.selected_group == "Apo" { 
                                get_default_csv_path("Apo") 
                            } else { 
                                get_default_csv_path("Endkunde") 
                            };
                            
                            let (_, data_dir, _, _, _) = get_release_dirs_with_debug(self.debug_mode);
                            let full_csv_path = data_dir.join(&csv_path.replace("Data/", ""));
                            let csv_exists = full_csv_path.exists();
                            let csv_status = if csv_exists {
                                format!("✅ CSV-Datei gefunden: {}", csv_path)
                            } else {
                                format!("❌ CSV-Datei fehlt: {} (erwartet: {})", csv_path, full_csv_path.display())
                            };
                            
                            // Lightweight schematic: draw A4 grid in millimeters and overlay configured QR boxes.
                            // This is cheap and avoids loading the full PDF template.
                            let margin = 12.0;
                            let inner_rect = egui::Rect::from_min_size(
                                egui::pos2(a4_rect.left() + margin, a4_rect.top() + margin),
                                egui::vec2(a4_width - 2.0 * margin, a4_height - 2.0 * margin)
                            );

                            // Draw pale background and border
                            ui.painter().rect_filled(inner_rect, 3.0, egui::Color32::from_rgb(250, 250, 250));
                            ui.painter().rect_stroke(inner_rect, 3.0, egui::Stroke::new(1.0, egui::Color32::from_rgb(200, 200, 200)));

                            // A4 in mm: 210 x 297
                            let a4_mm_w = 210.0;
                            let a4_mm_h = 297.0;
                            let mm_to_ui_x = inner_rect.width() / a4_mm_w;
                            let mm_to_ui_y = inner_rect.height() / a4_mm_h;

                            // Draw grid lines every 10 mm (lighter every 10, darker every 50)
                            for mm in (0..=210).step_by(5) {
                                let x = inner_rect.left() + (mm as f32) * mm_to_ui_x;
                                let color = if mm % 50 == 0 { egui::Color32::from_gray(200) } else if mm % 10 == 0 { egui::Color32::from_gray(220) } else { egui::Color32::from_gray(235) };
                                ui.painter().line_segment([egui::pos2(x, inner_rect.top()), egui::pos2(x, inner_rect.bottom())], egui::Stroke::new(1.0, color));
                            }
                            for mm in (0..=297).step_by(5) {
                                let y = inner_rect.top() + (mm as f32) * mm_to_ui_y;
                                let color = if mm % 50 == 0 { egui::Color32::from_gray(200) } else if mm % 10 == 0 { egui::Color32::from_gray(220) } else { egui::Color32::from_gray(235) };
                                ui.painter().line_segment([egui::pos2(inner_rect.left(), y), egui::pos2(inner_rect.right(), y)], egui::Stroke::new(1.0, color));
                            }

                            // Draw simple rulers (top and left) with mm labels every 50 mm
                            for mm in (0..=210).step_by(50) {
                                let x = inner_rect.left() + (mm as f32) * mm_to_ui_x;
                                ui.painter().text(egui::pos2(x + 2.0, inner_rect.top() - 14.0), egui::Align2::LEFT_TOP, format!("{}mm", mm), egui::FontId::proportional(10.0), egui::Color32::from_gray(120));
                            }
                            for mm in (0..=297).step_by(50) {
                                let y = inner_rect.top() + (mm as f32) * mm_to_ui_y;
                                ui.painter().text(egui::pos2(inner_rect.left() - 40.0, y - 6.0), egui::Align2::LEFT_TOP, format!("{}mm", mm), egui::FontId::proportional(10.0), egui::Color32::from_gray(120));
                            }

                            // Overlay configured QR boxes using config coordinates (PDF points -> UI scale).
                            // PDF points: A4 ~ 595 x 842
                            let pdf_w = 595.0_f32;
                            let pdf_h = 842.0_f32;
                            let scale_x = inner_rect.width() / pdf_w;
                            let scale_y = inner_rect.height() / pdf_h;
                            for (i, qr) in self.config.qr_codes.iter().enumerate() {
                                let qr_ui_x = inner_rect.left() + qr.x as f32 * scale_x;
                                let qr_ui_y = inner_rect.top() + qr.y as f32 * scale_y;
                                let size_ui = qr.size as f32 * scale_x;
                                let qr_rect = egui::Rect::from_min_size(egui::pos2(qr_ui_x, qr_ui_y), egui::vec2(size_ui, size_ui));
                                ui.painter().rect_stroke(qr_rect, 2.0, egui::Stroke::new(1.5, egui::Color32::from_rgb(120, 160, 220)));
                                ui.painter().text(egui::pos2(qr_ui_x + 2.0, qr_ui_y + 2.0), egui::Align2::LEFT_TOP, format!("QR{}", i+1), egui::FontId::proportional(10.0), egui::Color32::from_rgb(80,80,80));
                            }
                            // Small header area inside the A4 schematic for status text
                            let header_rect = egui::Rect::from_min_size(
                                inner_rect.min,
                                egui::vec2(inner_rect.width(), 60.0)
                            );
                            ui.painter().rect_filled(header_rect, 3.0, egui::Color32::from_rgb(245, 245, 250));
                            ui.painter().rect_stroke(header_rect, 3.0, egui::Stroke::new(1.0, egui::Color32::from_rgb(200, 200, 200)));
                            ui.painter().text(
                                egui::pos2(header_rect.left() + 10.0, header_rect.top() + 10.0),
                                egui::Align2::LEFT_TOP,
                                &template_status,
                                egui::FontId::proportional(11.0),
                                if template_loaded { egui::Color32::from_rgb(0, 120, 0) } else { egui::Color32::from_rgb(200, 100, 0) },
                            );
                            
                            // CSV-Status anzeigen
                            ui.painter().text(
                                egui::pos2(header_rect.left() + 10.0, header_rect.top() + 25.0),
                                egui::Align2::LEFT_TOP,
                                &csv_status,
                                egui::FontId::proportional(11.0),
                                if csv_exists { egui::Color32::from_rgb(0, 120, 0) } else { egui::Color32::from_rgb(200, 0, 0) },
                            );
                            
                            ui.painter().text(
                                egui::pos2(header_rect.center().x, header_rect.top() + 45.0),
                                egui::Align2::CENTER_TOP,
                                &format!("📄 {} {} Vorlage{}", 
                                        self.selected_group,
                                        self.selected_language,
                                        if self.is_messe { " (Messe)" } else { "" }),
                                egui::FontId::proportional(16.0),
                                egui::Color32::from_rgb(80, 80, 80),
                            );

                            // (Full preview control removed to keep UI simple and performant)
                            
                            // (Removed simulated customer/order form sections to keep preview minimal and unambiguous)
                            // Skalierungsfaktor für die Koordinaten (DIN A4 in PDF ist ca. 595x842 Punkte)
                            let scale_x = 595.0 / a4_width;
                            let scale_y = 842.0 / a4_height;
                            
                            // QR-Codes
                            for (i, qr) in self.config.qr_codes.iter_mut().enumerate() {
                                let qr_display_size = qr.size * 1.5; // Größer für bessere Sichtbarkeit
                                let qr_pos_x = qr.x / scale_x;
                                let qr_pos_y = a4_height - (qr.y / scale_y); // Y-Koordinate umkehren
                                
                                // DEBUG: Position für ersten QR-Code anzeigen
                                if i == 0 {
                                    ui.painter().text(
                                        egui::pos2(a4_rect.left(), a4_rect.bottom() + 25.0),
                                        egui::Align2::LEFT_TOP,
                                        format!("QR1: x={:.1}, y={:.1} -> UI: x={:.1}, y={:.1}", 
                                               qr.x, qr.y, qr_pos_x, qr_pos_y),
                                        egui::FontId::proportional(11.0),
                                        egui::Color32::from_rgb(100, 100, 100),
                                    );
                                }
                                
                                let qr_rect = egui::Rect::from_min_size(
                                    egui::pos2(a4_rect.left() + qr_pos_x, a4_rect.top() + qr_pos_y),
                                    egui::vec2(qr_display_size, qr_display_size)
                                );
                                
                                // QR-Code mit interact_rect für bessere Drag-Detection
                                let qr_id = egui::Id::new(format!("qr_code_drag_{}", i));
                                let qr_response = ui.interact(qr_rect, qr_id, egui::Sense::drag());
                                
                                // QR-Code visuell darstellen
                                // If selected, draw a highlighted border
                                if let Some((ref kind, idx)) = self.selected_element {
                                    if kind == "qr" && idx == i {
                                        ui.painter().rect_stroke(qr_rect.expand(4.0), 3.0, egui::Stroke::new(2.0, egui::Color32::from_rgb(40, 160, 40)));
                                    }
                                }
                                ui.painter().rect_filled(qr_rect, 3.0, egui::Color32::from_rgb(255, 165, 0)); // Orange
                                ui.painter().text(
                                    qr_rect.center(),
                                    egui::Align2::CENTER_CENTER,
                                    format!("📱 Q{}", i + 1),
                                    egui::FontId::default(),
                                    egui::Color32::BLACK,
                                );
                                
                                if qr_response.dragged() {
                                    let delta = qr_response.drag_delta();
                                    qr.x += delta.x * scale_x;
                                    qr.y -= delta.y * scale_y; // Y-Koordinate umkehren
                                    // Grenzen prüfen
                                    qr.x = qr.x.max(0.0).min(595.0 - qr.size);
                                    qr.y = qr.y.max(0.0).min(842.0 - qr.size);
                                    
                                    // Update manual input fields for first QR code
                                    if i == 0 {
                                        self.manual_qr_x = format!("{:.1}", qr.x);
                                        self.manual_qr_y = format!("{:.1}", qr.y);
                                        self.manual_qr_size = format!("{:.1}", qr.size);
                                    }
                                }
                                if qr_response.clicked() {
                                    self.selected_element = Some(("qr".to_string(), i));
                                }
                            }
                            
                            // Draw minimal markers for Vertreternummer positions so users can see and nudge them.
                            for (i, v) in self.config.vertreter.iter_mut().enumerate() {
                                // Use the same transform as above (map PDF points -> A4 UI coordinates)
                                // Compute a rectangular marker that reflects the font size and expected text width.
                                // Approximate width: character_width_factor * font_size * estimated_chars
                                let est_chars = v.font_name.len().max(6) as f32; // rough estimate (at least room for number)
                                let char_width_factor = 0.6; // approximate em-width factor
                                let text_width_pts = v.font_size * char_width_factor * est_chars; // in PDF pts
                                // Convert PDF points to UI using same scale_x/scale_y mapping
                                let marker_w_ui = text_width_pts / scale_x;
                                let marker_h_ui = (v.font_size * 1.2) / scale_y; // little padding
                                let v_pos_x = v.x / scale_x;
                                let v_pos_y = a4_height - (v.y / scale_y); // invert Y like for QR

                                let vert_rect = egui::Rect::from_min_size(
                                    egui::pos2(a4_rect.left() + v_pos_x, a4_rect.top() + v_pos_y),
                                    egui::vec2(marker_w_ui.max(12.0), marker_h_ui.max(10.0)),
                                );

                                // Small interactable area so the user can drag the vertreter marker
                                let vert_id = egui::Id::new(format!("vertreter_drag_{}", i));
                                let vert_resp = ui.interact(vert_rect, vert_id, egui::Sense::drag());

                                // Draw the marker: pale blue rounded rect with a thin border and label
                                ui.painter().rect_filled(vert_rect, 4.0, egui::Color32::from_rgb(200, 220, 255));
                                ui.painter().rect_stroke(vert_rect, 4.0, egui::Stroke::new(1.0, egui::Color32::from_rgb(120, 140, 180)));
                                // Highlight if selected
                                if let Some((ref kind, idx)) = self.selected_element {
                                    if kind == "vertreter" && idx == i {
                                        ui.painter().rect_stroke(vert_rect.expand(4.0), 4.0, egui::Stroke::new(2.0, egui::Color32::from_rgb(40, 160, 40)));
                                    }
                                }
                                ui.painter().text(
                                    egui::pos2(vert_rect.left() + 4.0, vert_rect.center().y - 6.0),
                                    egui::Align2::LEFT_CENTER,
                                    format!("V{} ({})", i + 1, v.font_name),
                                    egui::FontId::proportional((v.font_size / 1.2).max(10.0)),
                                    egui::Color32::from_rgb(40, 40, 40),
                                );

                                if vert_resp.dragged() {
                                    let delta = vert_resp.drag_delta();
                                    v.x += delta.x * scale_x;
                                    v.y -= delta.y * scale_y; // invert Y

                                    // Boundaries (PDF coordinates)
                                    v.x = v.x.max(0.0).min(595.0 - v.size);
                                    v.y = v.y.max(0.0).min(842.0 - v.size);

                                    // If first vertreter, update manual input fields for quick feedback
                                    // No manual text fields for vertreter currently; positions update in config.
                                }
                                if vert_resp.clicked() {
                                    self.selected_element = Some(("vertreter".to_string(), i));
                                }
                            }
                        });

                        // Speichern-unter Modal
                        if self.show_save_as {
                            egui::Window::new("Speichern unter")
                                .collapsible(false)
                                .resizable(false)
                                .anchor(egui::Align2::CENTER_CENTER, egui::vec2(0.0, 0.0))
                                .show(ctx, |ui| {
                                    ui.label("Dateiname (z. B. CONFIG/config_fachkreise-DE.toml):");
                                    ui.text_edit_singleline(&mut self.config_save_path);
                                    ui.horizontal(|ui| {
                                        if ui.button("Speichern").clicked() {
                                            let p = std::path::Path::new(&self.config_save_path);
                                            match save_group_config_to_path(p, &self.config) {
                                                Ok(_) => {
                                                    println!("Config gespeichert nach {:?}", p);
                                                    set_current_config_path(p);
                                                    self.show_save_as = false;
                                                }
                                                Err(e) => {
                                                    eprintln!("Fehler beim Speichern: {}", e);
                                                }
                                            }
                                        }
                                        if ui.button("Durchsuchen").clicked() {
                                            // Open native save dialog
                                            if let Some(path) = rfd::FileDialog::new().set_directory("CONFIG").save_file() {
                                                if let Some(s) = path.to_str() {
                                                    self.config_save_path = s.to_string();
                                                }
                                            }
                                        }
                                        if ui.button("Abbrechen").clicked() {
                                            self.show_save_as = false;
                                        }
                                    });
                                });
                        }
                        
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
                                                .width(150.0)
                                                .show_ui(ui, |ui| {
                                                    // PERFORMANCE: Nur die wichtigsten Fonts zeigen
                                                    
                                                    let common_fonts = ["Arial", "Calibri", "Times New Roman", "Helvetica", "Verdana", "Georgia"];
                                                    let current_font = v.font_name.clone();
                                                    
                                                    // Aktueller Font zuerst (falls nicht in common_fonts)
                                                    if !common_fonts.contains(&current_font.as_str()) {
                                                        ui.selectable_value(&mut v.font_name, current_font.clone(), current_font);
                                                        ui.separator();
                                                    }
                                                    
                                                    // Häufige Fonts - IMMER sichtbar
                                                    ui.label("📌 Häufig:");
                                                    for font in &common_fonts {
                                                        ui.selectable_value(&mut v.font_name, font.to_string(), *font);
                                                    }
                                                    
                                                    ui.separator();
                                                    
                                                    // PERFORMANCE: Nur erste 15 andere Fonts für flüssige Performance
                                                    ui.label("🔤 Weitere:");
                                                    let mut shown_count = 0;
                                                    let max_initial = 15; // Stark reduziert für Performance
                                                    
                                                    for font in &self.cached_fonts {
                                                        // Skip häufige Fonts
                                                        if common_fonts.contains(&font.as_str()) {
                                                            continue;
                                                        }
                                                        
                                                        if shown_count < max_initial {
                                                            ui.selectable_value(&mut v.font_name, font.clone(), font);
                                                            shown_count += 1;
                                                        } else {
                                                            break;
                                                        }
                                                    }
                                                    
                                                    // Info über weitere verfügbare Fonts
                                                    if self.cached_fonts.len() > (common_fonts.len() + max_initial) {
                                                        ui.small(format!("� +{} weitere verfügbar", 
                                                            self.cached_fonts.len() - common_fonts.len() - max_initial));
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
                                            ui.add(egui::Slider::new(&mut v.font_size, 6.0..=48.0).suffix(" pt"));
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
                                    ui.label(format!("QR-Code {}: x={:.1}, y={:.1}, Größe={:.1}, {}", 
                                        i + 1, qr.x, qr.y, qr.size, pages_str));
                                }
                                for (i, pos) in self.config.vertreter.iter().enumerate() {
                                    let pages_str = if pos.all_pages {
                                        "alle Seiten".to_string()
                                    } else {
                                        format!("Seiten: {:?}", pos.pages)
                                    };
                                    ui.label(format!("Vertreternummer {}: x={:.1}, y={:.1}, Größe={:.1}, Font: {} {} ({}pt), {}", 
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
                                            ui.label("X:");
                                            ui.text_edit_singleline(&mut self.manual_qr_x);
                                        });
                                        ui.horizontal(|ui| {
                                            ui.label("Y:");
                                            ui.text_edit_singleline(&mut self.manual_qr_y);
                                        });
                                        ui.horizontal(|ui| {
                                            ui.label("Größe:");
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
                                                println!("QR-Code Position manuell gesetzt: x={}, y={}, size={}", x, y, size);
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
                                            ui.label("X:");
                                            ui.text_edit_singleline(&mut self.manual_vertreter_x);
                                        });
                                        ui.horizontal(|ui| {
                                            ui.label("Y:");
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
                                                println!("Kundennummer Position manuell gesetzt: x={}, y={}", x, y);
                                            }
                                        }
                                    });
                                }
                                
                                ui.separator();
                                ui.label("💡 Tipp: PDF-Koordinaten beginnen links-unten bei (0,0)");
                                ui.label("📐 Referenz: DIN A4 = 595×842 Punkte");
                            });
                            
                            ui.separator();
                            
                            // Speichern/Abbrechen Buttons
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
                                    self.show_config = false;
                                    
                                    // Nach dem Speichern nochmal laden um sicherzustellen dass alles stimmt
                                    let loaded_config = load_group_config(&self.selected_group, &self.selected_language, self.is_messe);
                                    println!("NACH Laden - Geladene gruppenspezifische Config: QR={:?}, Vertreter={:?}", 
                                            loaded_config.qr_codes, loaded_config.vertreter);
                                    self.config = loaded_config;
                                    println!("=== SPEICHERN ABGESCHLOSSEN ===");
                                }
                                if ui.button("❌ Abbrechen").clicked() {
                                    println!("=== ABBRECHEN GEDRÜCKT ===");
                                    // Gruppenspezifische Config komplett neu laden um alle Änderungen zu verwerfen
                                    self.config = load_group_config(&self.selected_group, &self.selected_language, self.is_messe);
                                    println!("Abgebrochen - gruppenspezifische Config zurückgesetzt: QR={:?}, Vertreter={:?}", 
                                            self.config.qr_codes, self.config.vertreter);
                                    self.show_config = false;
                                }
                            });
                        });
                    });
                });
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
    debug_print_global(&format!("Versuche CSV zu lesen: {}", file_path));
    
    let content = match fs::read_to_string(file_path) {
        Ok(content) => {
                debug_print_global(&format!("CSV erfolgreich gelesen, {} Zeichen", content.len()));
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
        debug_print_global(&format!("Header übersprungen: {}", header.trim()));
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

// Track the last-loaded config file path in a safe Mutex for multi-threaded access
use once_cell::sync::Lazy;
static CURRENT_CONFIG_PATH: Lazy<std::sync::Mutex<Option<std::path::PathBuf>>> = Lazy::new(|| std::sync::Mutex::new(None));

// Funktion um die aktuelle Config zu setzen (threadsafe)
fn set_current_config(config: &Config) {
    let _lock = CONFIG_MUTEX.lock().unwrap();
    unsafe {
        CURRENT_CONFIG = Some(config.clone());
    }
    println!("Aktuelle Config gesetzt für PDF-Generierung: QR={:?}, Vertreter={:?}", 
             config.qr_codes, config.vertreter);
}

fn set_current_config_path(path: &std::path::Path) {
    let mut g = CURRENT_CONFIG_PATH.lock().unwrap();
    *g = Some(path.to_path_buf());
    println!("Aktuelle Config-Datei gesetzt: {:?}", path);
}

fn get_current_config_path() -> Option<std::path::PathBuf> {
    let g = CURRENT_CONFIG_PATH.lock().unwrap();
    g.clone()
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
                debug_print_global(&format!("Font gefunden: {} -> {}", font_name, font_path.display()));
                return Some(font_path);
            }
        }
    }
    
    debug_print_global(&format!("Font NICHT gefunden: {} ({})", font_name, style));
    None
}

fn modify_pdf_with_debug(template_path: &str, kundennr: &str, qr_code: &[u8], qr_width: usize, config: &Config, output_path: &std::path::Path, debug_enabled: bool) {
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
                                  kundennr, maybe_image_id, Some(qr_width), debug_enabled);
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
                         kundennr: &str, maybe_image_id: Option<lopdf::ObjectId>, maybe_image_width: Option<usize>, _debug_enabled: bool) {
    
    let content_stream = doc.get_page_content(page_id).expect("Konnte Seiteninhalt nicht laden");
    let mut content = Content::decode(&content_stream).expect("Konnte Inhalt nicht dekodieren");

    // Alle QR-Codes platzieren
    for (i, qr_config) in qr_codes.iter().enumerate() {
        if let Some(_img_id) = maybe_image_id {
            // Compute image-space -> user-space scale so that the drawn image width equals qr_config.size points
            let scale = if let Some(w) = maybe_image_width { qr_config.size / (w as f32) } else { qr_config.size };
            content.operations.push(Operation::new("q", vec![]));
            content.operations.push(Operation::new("cm", vec![
                (scale).into(), 0.into(), 0.into(), (scale).into(), 
                qr_config.x.into(), qr_config.y.into()
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
        
        // Verschiedene Standard-Fonts definieren
        let mut font_counter = 1;
        let mut font_names = std::collections::HashMap::new();
        
        for vertreter_config in vertreter_configs {
            // Erweiterte Font-Erkennung für mehr Schriftarten (inkl. Source Fonts)
            let base_font = match vertreter_config.font_name.as_str() {
                // Times Familie
                "Times New Roman" | "Times" => "Times",
                // Arial/Helvetica Familie
                "Arial" | "Helvetica" => "Helvetica", 
                // Courier Familie
                "Courier New" | "Courier" => "Courier",
                // Adobe/Source Fonts (falls installiert, fallback auf Helvetica)
                name if name.contains("Adobe") || name.contains("Myriad") || name.contains("Minion") || name.contains("Source") => "Helvetica",
                // Candara (fallback auf Helvetica)
                "Candara" => "Helvetica",
                // Calibri (fallback auf Helvetica)
                "Calibri" => "Helvetica", 
                // Weitere Windows-Fonts (fallback auf Helvetica)
                "Verdana" | "Georgia" | "Trebuchet MS" | "Comic Sans MS" | "Impact" | "Lucida Console" | "Tahoma" => "Helvetica",
                // Font-Namen mit Styles (extrahiere Basis-Font)
                name if name.contains("Bold") || name.contains("Italic") || name.contains("Light") || name.contains("Medium") => {
                    if name.contains("Times") {
                        "Times"
                    } else if name.contains("Arial") {
                        "Helvetica"
                    } else if name.contains("Courier") {
                        "Courier"
                    } else {
                        "Helvetica"
                    }
                },
                _ => "Helvetica" // Sicherer Fallback
            };
            
            // Erweiterte Style-Zuordnung
            let pdf_font_name = match (base_font, vertreter_config.font_style.as_str()) {
                // Times Familie
                ("Times", "Bold") => "Times-Bold",
                ("Times", "Italic") => "Times-Italic", 
                ("Times", "Bold Italic") | ("Times", "BoldItalic") => "Times-BoldItalic",
                ("Times", _) => "Times-Roman",
                // Helvetica Familie (Arial → Helvetica Mapping)
                ("Helvetica", "Bold") => "Helvetica-Bold",
                ("Helvetica", "Italic") => "Helvetica-Oblique",
                ("Helvetica", "Bold Italic") | ("Helvetica", "BoldItalic") => "Helvetica-BoldOblique", 
                ("Helvetica", "Light") => "Helvetica", // Light nicht verfügbar, fallback
                ("Helvetica", "Medium") => "Helvetica-Bold", // Medium → Bold mapping
                ("Helvetica", "Heavy") | ("Helvetica", "Black") => "Helvetica-Bold", // Heavy/Black → Bold mapping
                ("Helvetica", "Thin") => "Helvetica", // Thin nicht verfügbar, fallback
                ("Helvetica", _) => "Helvetica",
                // Courier Familie
                ("Courier", "Bold") => "Courier-Bold",
                ("Courier", "Italic") => "Courier-Oblique",
                ("Courier", "Bold Italic") | ("Courier", "BoldItalic") => "Courier-BoldOblique",
                ("Courier", _) => "Courier",
                // Sicherer Fallback
                (_, _) => "Helvetica"
            };
            
            let font_key = format!("F{}", font_counter);
            if !font_names.contains_key(pdf_font_name) {
                font_dict.set(font_key.as_bytes(), dictionary!{
                    "Type" => "Font",
                    "Subtype" => "Type1",
                    "BaseFont" => pdf_font_name
                });
                font_names.insert(pdf_font_name, font_key.clone());
                font_counter += 1;
            }
        }

        // Jetzt den Content-Stream mit Text schreiben
        for vertreter_config in vertreter_configs {
            // NEUE LOGIK: Erst Custom TTF-Font versuchen, dann Fallback auf Standard-PDF-Fonts
            
            if let Some(font_path) = find_font_file(&vertreter_config.font_name, &vertreter_config.font_style) {
                debug_print(&format!("Custom Font gefunden: {} -> {}", vertreter_config.font_name, font_path.display()), _debug_enabled);
                
                // Versuche TTF-Font zu laden und in PDF einzubetten
                if let Ok(font_data) = std::fs::read(&font_path) {
                    let _custom_font_name = format!("CustomFont_{}_{}", 
                        vertreter_config.font_name.replace(" ", ""), 
                        vertreter_config.font_style.replace(" ", ""));
                    
                    // TODO: TTF-Font-Einbettung (kompliziert in lopdf)
                    // Für jetzt: Besseres Standard-Font-Mapping
                    debug_print(&format!("TTF-Font-Daten geladen ({}kb), verwende verbessertes Mapping", font_data.len() / 1024), _debug_enabled);
                } else {
                    debug_print(&format!("Konnte TTF-Font-Datei nicht lesen: {}", font_path.display()), _debug_enabled);
                }
            } else {
                debug_print(&format!("Custom Font NICHT gefunden: {} ({}), verwende Standard-PDF-Font", 
                    vertreter_config.font_name, vertreter_config.font_style), _debug_enabled);
            }
            
            // 2. Fallback: Verbessertes Standard-PDF-Font-Mapping
            let base_font = match vertreter_config.font_name.as_str() {
                // Times Familie - BESSERE ERKENNUNG
                "Times New Roman" | "Times" | "TimesNewRoman" => "Times",
                // Arial/Helvetica Familie - BESSERE ERKENNUNG  
                "Arial" | "Helvetica" | "Arial Unicode MS" => "Helvetica",
                // Courier Familie
                "Courier New" | "Courier" | "CourierNew" => "Courier",
                // Calibri -> Helvetica (ähnlichster Standard-Font)
                "Calibri" => "Helvetica",
                // Verdana -> Helvetica 
                "Verdana" => "Helvetica",
                // Georgia -> Times (Serif-Font)
                "Georgia" => "Times",
                // Adobe/Source Fonts -> Helvetica
                name if name.contains("Adobe") || name.contains("Myriad") || name.contains("Minion") || name.contains("Source") => "Helvetica",
                // Weitere Windows-Fonts
                "Trebuchet MS" | "Comic Sans MS" | "Impact" | "Lucida Console" | "Tahoma" | "Candara" => "Helvetica",
                // Font-Namen mit Styles (extrahiere Basis-Font)
                name if name.contains("Bold") || name.contains("Italic") || name.contains("Light") || name.contains("Medium") => {
                    if name.contains("Times") {
                        "Times"
                    } else if name.contains("Arial") {
                        "Helvetica"
                    } else if name.contains("Courier") {
                        "Courier"
                    } else {
                        "Helvetica"
                    }
                },
                _ => "Helvetica" // Sicherer Fallback
            };
            
            // Erweiterte Style-Zuordnung (identisch mit oben)
            let pdf_font_name = match (base_font, vertreter_config.font_style.as_str()) {
                // Times Familie
                ("Times", "Bold") => "Times-Bold",
                ("Times", "Italic") => "Times-Italic", 
                ("Times", "Bold Italic") | ("Times", "BoldItalic") => "Times-BoldItalic",
                ("Times", _) => "Times-Roman",
                // Helvetica Familie (Arial → Helvetica Mapping)
                ("Helvetica", "Bold") => "Helvetica-Bold",
                ("Helvetica", "Italic") => "Helvetica-Oblique",
                ("Helvetica", "Bold Italic") | ("Helvetica", "BoldItalic") => "Helvetica-BoldOblique", 
                ("Helvetica", "Light") => "Helvetica", // Light nicht verfügbar, fallback
                ("Helvetica", "Medium") => "Helvetica-Bold", // Medium → Bold mapping
                ("Helvetica", "Heavy") | ("Helvetica", "Black") => "Helvetica-Bold", // Heavy/Black → Bold mapping
                ("Helvetica", "Thin") => "Helvetica", // Thin nicht verfügbar, fallback
                ("Helvetica", _) => "Helvetica",
                // Courier Familie
                ("Courier", "Bold") => "Courier-Bold",
                ("Courier", "Italic") => "Courier-Oblique",
                ("Courier", "Bold Italic") | ("Courier", "BoldItalic") => "Courier-BoldOblique",
                ("Courier", _) => "Courier",
                // Sicherer Fallback
                (_, _) => "Helvetica"
            };
            
            let font_key = font_names.get(pdf_font_name).unwrap_or(&"F1".to_string()).clone();
            
            content.operations.push(Operation::new("BT", vec![]));
            content.operations.push(Operation::new("Tf", vec![Object::Name(font_key.into_bytes()), vertreter_config.font_size.into()]));
            content.operations.push(Operation::new("Td", vec![vertreter_config.x.into(), vertreter_config.y.into()]));
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
) -> Result<(), Box<dyn std::error::Error>> {
    // Debug-Logging für Funktion
    if debug_mode {
    debug_print_global(&format!("generate_bestellscheine_resume gestartet mit {} Threads, Sleep: {}ms", threads, thread_sleep_ms));
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
                    debug_print_global(&format!("Verwende englische URL für Vertreter {}: {}", kundennr, en_link));
                    en_link
                } else {
                    debug_print_global(&format!("Verwende deutsche URL für Vertreter {}: {}", kundennr, de_link));
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
                
                debug_print_global(&format!("Output-Verzeichnis für Sprache '{}': {}", language, group_output_dir.display()));
                
                // Prüfen ob PDF bereits existiert
                // Versuche, eine Sprache zu detektieren basierend auf Auswahl/template (für Existenz-Check)
                let detected_for_check = detect_language_code(&language, Some(&first_template), None);
                let pdf_filename = format!("{}-{}-{}.pdf", 
                    std::path::Path::new(&first_template)
                        .file_stem()
                        .unwrap_or_default()
                        .to_string_lossy(),
                    detected_for_check,
                    kundennr.replace("\0", "")
                );
                let pdf_path = group_output_dir.join(&pdf_filename);
                
                if !pdf_path.exists() {
                    println!("Erstelle PDF für Vertreter {}: {} -> {}/{}", i + 1, kundennr, group, language);
                    // Für jede ausgewählte Template-Option erstellen (aber keine Duplikate)
                    let selections = get_current_selections().unwrap_or_else(|| vec![ ("DATA/Vertreternummern.csv".to_string(), "VORLAGE/Bestellschein-Endkunde-de_de.pdf".to_string(), true) ]);
                    // Debug: Liste der verwendeten Selections ausgeben, damit wir nachvollziehen können,
                    // ob die UI-Auswahl oder der Fallback verwendet wird.
                    debug_print_global(&format!("selections for generation (count={}): {:?}", selections.len(), selections));
                    let mut created = Vec::new();
                    for (csv_s, template_s, gen_qr) in selections.iter() {
                        // Verhindere doppelte Ausgaben für dieselbe template/kundennr
                        let out_name = format!("{}-{}", template_s, kundennr);
                        if created.contains(&out_name) { continue; }
                        created.push(out_name.clone());

                        // Template-Pfad korrekt auflösen
                        let resolved_template = resolve_template_path_with_debug(template_s, debug_mode);
                        let resolved_template_str = resolved_template.to_string_lossy();

                        debug_print(&format!("Verwende Template: {}", resolved_template_str), debug_mode);

                        // Debug: Zeige welches template_s wir versuchen zu verwenden und ob die Datei existiert
                        debug_print_global(&format!("template_s='{}' -> resolved='{}' (exists={})", template_s, resolved_template_str, resolved_template.exists()));

                        // Prüfe ob Template existiert
                        if !resolved_template.exists() {
                            println!("ERROR: Template-Datei nicht gefunden: {}", resolved_template_str);
                            continue;
                        }

                        // Bestimme kanonischen Sprachcode (de_de / en_us) anhand Template, CSV oder UI
                        let detected_lang_code = detect_language_code(&language, Some(&resolved_template_str), Some(csv_s));
                        let (group_name, _template_lang, tpl_is_messe) = infer_group_lang_from_template(&resolved_template_str);

                        // Debug: Zeige den detektierten Sprachcode und den Pfad, wohin geschrieben werden soll
                        let tpl_stem = resolved_template.file_stem().unwrap_or_default().to_string_lossy();
                        let output_filename_preview = format!("{}-{}.pdf", tpl_stem, kundennr);
                        let output_path_preview = get_configured_output_dir_with_debug(use_custom_output, &custom_output_path, &group_name, &detected_lang_code, tpl_is_messe, debug_mode).join(&output_filename_preview);
                        debug_print_global(&format!("detected_lang_code='{}', preview_output_path='{}'", detected_lang_code, output_path_preview.display()));

                        // Verwende aktuelle UI-Config falls verfügbar, sonst fallback zu group config (mit detektiertem Sprachcode)
                        let tpl_config = {
                            let _lock = CONFIG_MUTEX.lock().unwrap();
                            unsafe {
                                if let Some(ref current_config) = CURRENT_CONFIG {
                                    println!("🎯 Verwende aktuelle UI-Config für PDF-Generierung: QR={:?}", current_config.qr_codes);
                                    current_config.clone()
                                } else {
                                    println!("⚠️ Keine UI-Config verfügbar, lade Group-Config (detected lang: {})", detected_lang_code);
                                    load_group_config(&group_name, &detected_lang_code, tpl_is_messe)
                                }
                            }
                        };
                        if *gen_qr {
                            let (qr_img, qr_width) = generate_qr(link);
                            // Output-Dateiname jetzt inklusive Sprachcode
                            let tpl_stem = resolved_template.file_stem().unwrap_or_default().to_string_lossy();
                            // Verwende nur Template-Stem + Kundennr als Dateiname
                            let output_filename = format!("{}-{}.pdf", tpl_stem, kundennr);
                            let output_path = get_configured_output_dir_with_debug(use_custom_output, &custom_output_path, &group_name, &detected_lang_code, tpl_is_messe, debug_mode).join(&output_filename);
                            modify_pdf_with_debug(&resolved_template_str, kundennr, &qr_img, qr_width, &tpl_config, &output_path, debug_mode);
                        } else {
                            let tpl_stem = resolved_template.file_stem().unwrap_or_default().to_string_lossy();
                            let output_filename = format!("{}-{}.pdf", tpl_stem, kundennr);
                            let output_path = get_configured_output_dir_with_debug(use_custom_output, &custom_output_path, &group_name, &detected_lang_code, tpl_is_messe, debug_mode).join(&output_filename);
                            modify_pdf_with_debug(&resolved_template_str, kundennr, &[], 0, &tpl_config, &output_path, debug_mode);
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
            debug_print_global(&format!("Konnte progress.txt nicht löschen: {}", e));
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
