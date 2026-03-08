#![cfg_attr(target_os = "windows", windows_subsystem = "windows")]

use std::{
    collections::{BTreeMap, HashMap, HashSet},
    env,
    fs::{self, File},
    io::{self, Write},
    path::{Path, PathBuf},
    process::Command,
    sync::{Arc, Mutex, OnceLock, mpsc},
    thread,
    time::Duration,
};

use anyhow::{Context, Result, anyhow};
use axum::{
    Json, Router,
    extract::{DefaultBodyLimit, Form, Multipart, Path as AxumPath, Query, State},
    http::{StatusCode, header},
    response::{Html, IntoResponse, Response},
    routing::{get, post},
};
use base64::{Engine as _, engine::general_purpose::URL_SAFE_NO_PAD};
use chrono::{DateTime, SecondsFormat, Utc};
use dotenvy::from_path_override;
use lettre::{
    Message, SmtpTransport, Transport,
    message::{Attachment, Mailbox, MultiPart, SinglePart},
    transport::smtp::authentication::Credentials,
};
use lopdf::{Document, Object, ObjectId};
use regex::Regex;
use serde::{Deserialize, Serialize};
use serde_json::{Value, json};
use tera::{Context as TeraContext, Tera};
use tokio::sync::{RwLock, Semaphore};
use tower_http::{services::ServeDir, trace::TraceLayer};
use tracing::{error, info};
use uuid::Uuid;
use zip::{ZipArchive, read::HasZipMetadata};

#[cfg(target_os = "windows")]
use std::os::windows::process::CommandExt;
#[cfg(target_os = "windows")]
use tao::{
    event::Event,
    event_loop::{ControlFlow, EventLoopBuilder},
};
#[cfg(target_os = "windows")]
use tray_icon::{
    Icon, TrayIcon, TrayIconBuilder, TrayIconEvent,
    menu::{
        CheckMenuItem, Icon as MenuIcon, IconMenuItem, Menu, MenuEvent, MenuItem,
        PredefinedMenuItem,
    },
};
#[cfg(target_os = "windows")]
use winreg::{RegKey, enums::HKEY_CURRENT_USER};

const UNGROUPED_TEAM_KEY: &str = "__ungrouped__";
const DEFAULT_RECIPIENT_EMAIL: &str = "roboken.report.tool@gmail.com";
const STANDARD_PRINTER_KEYWORDS: &[&str] = &[
    "microsoft print to pdf",
    "microsoft xps document writer",
    "onenote",
    "send to onenote",
    "fax",
];
const DEFAULT_SERVER_PORT: u16 = 80;
const APP_DISPLAY_NAME: &str = "Report PDF Converter";
const DEFAULT_EMAIL_SENDER: &str = "roboken.report.tool@gmail.com";
const DEFAULT_EMAIL_DISPLAY_NAME: &str = "ロボ研報告書作成ツール";
const GMAIL_SEND_SCOPE: &str = "https://www.googleapis.com/auth/gmail.send";
const GMAIL_SEND_ENDPOINT: &str = "https://gmail.googleapis.com/gmail/v1/users/me/messages/send";
#[cfg(target_os = "windows")]
const CREATE_NO_WINDOW_FLAG: u32 = 0x0800_0000;
#[cfg(target_os = "windows")]
const SERVER_START_TIMEOUT_SECONDS: u64 = 10;
#[cfg(target_os = "windows")]
const WINDOWS_STARTUP_REG_PATH: &str = "Software\\Microsoft\\Windows\\CurrentVersion\\Run";
#[cfg(target_os = "windows")]
const WINDOWS_STARTUP_VALUE_NAME: &str = "ReportPDFConverter";

const WORD_TO_PDF_SCRIPT: &str = r#"
$ErrorActionPreference = 'Stop'
$src = $env:RPC_SOURCE
$dst = $env:RPC_DEST
if ([string]::IsNullOrWhiteSpace($src) -or [string]::IsNullOrWhiteSpace($dst)) {
    throw 'RPC_SOURCE or RPC_DEST is empty.'
}

$word = New-Object -ComObject Word.Application
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($src, $false, $true)
    try {
        $doc.ExportAsFixedFormat($dst, 17)
    } finally {
        if ($null -ne $doc) {
            $doc.Close($false)
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
        }
    }
} finally {
    if ($null -ne $word) {
        $word.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
"#;

const PRINT_SCRIPT: &str = r#"
$ErrorActionPreference = 'Stop'
$file = $env:RPC_FILE
$printer = $env:RPC_PRINTER
if (-not (Test-Path -LiteralPath $file)) {
    throw "file not found: $file"
}
Start-Process -FilePath $file -Verb PrintTo -ArgumentList ('"' + $printer + '"')
Start-Sleep -Seconds 2
"#;
const QUERY_PRINTERS_SCRIPT: &str = r#"
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
Get-Printer | Select-Object Name,Default | ConvertTo-Json -Depth 3
"#;
const MAX_UPLOAD_BODY_BYTES: usize = 256 * 1024 * 1024;

#[derive(Clone)]
struct AppState {
    jobs: Arc<RwLock<HashMap<String, JobState>>>,
    upload_sessions: Arc<RwLock<HashMap<String, UploadSession>>>,
    paths: RuntimePaths,
    tera: Arc<Tera>,
    order_manager: Arc<OrderManager>,
    job_semaphore: Arc<Semaphore>,
}

#[derive(Clone)]
struct RuntimePaths {
    base_dir: PathBuf,
    uploads_dir: PathBuf,
    work_dir: PathBuf,
    order_file: PathBuf,
    faculty_contacts_file: PathBuf,
    templates_dir: PathBuf,
    static_dir: PathBuf,
}

impl RuntimePaths {
    fn prepare(base_dir: PathBuf) -> Result<Self> {
        let data_dir = resolve_data_dir(&base_dir);
        let uploads_dir = data_dir.join("uploads");
        let work_dir = data_dir.join("work");
        let lab_members_dir = data_dir.join("lab_members");
        let order_file = lab_members_dir.join("order.json");
        let faculty_contacts_file = lab_members_dir.join("faculty_contacts.json");
        let bundled_lab_members_dir = base_dir.join("lab_members");
        let bundled_order_file = bundled_lab_members_dir.join("order.json");
        let bundled_faculty_contacts_file = bundled_lab_members_dir.join("faculty_contacts.json");
        let templates_dir = base_dir.join("templates");
        let static_dir = base_dir.join("static");

        fs::create_dir_all(&uploads_dir)?;
        fs::create_dir_all(&work_dir)?;
        fs::create_dir_all(&lab_members_dir)?;
        if !order_file.exists() && bundled_order_file.exists() {
            fs::copy(&bundled_order_file, &order_file)?;
        }
        if !faculty_contacts_file.exists() && bundled_faculty_contacts_file.exists() {
            fs::copy(&bundled_faculty_contacts_file, &faculty_contacts_file)?;
        }

        Ok(Self {
            base_dir,
            uploads_dir,
            work_dir,
            order_file,
            faculty_contacts_file,
            templates_dir,
            static_dir,
        })
    }
}

#[derive(Clone, Serialize, Deserialize, Debug)]
struct MemberRecord {
    name: String,
    #[serde(default, skip_serializing_if = "String::is_empty")]
    email: String,
}

#[derive(Clone, Debug)]
struct OrderPreferences {
    team_sequence: Vec<String>,
    member_sequences: HashMap<String, Vec<MemberRecord>>,
}

impl OrderPreferences {
    fn empty() -> Self {
        Self {
            team_sequence: Vec::new(),
            member_sequences: HashMap::new(),
        }
    }
}

struct OrderManager {
    storage_file: PathBuf,
    lock: Mutex<()>,
}

#[derive(Clone, Serialize, Deserialize, Debug)]
struct ZipEntry {
    identifier: String,
    display_name: String,
    archive_name: String,
    team_name: Option<String>,
    #[serde(default)]
    persons: Vec<String>,
    sanitized_name: Option<String>,
}

#[derive(Clone)]
struct UploadSession {
    entries: HashMap<String, ZipEntry>,
    team_options: Vec<String>,
    zip_filename: String,
}

#[derive(Clone, Serialize)]
struct JobState {
    id: String,
    email: String,
    status: String,
    message: String,
    created_at: DateTime<Utc>,
    updated_at: DateTime<Utc>,
    order: Vec<String>,
    zip_path: PathBuf,
    zip_original_name: String,
    entries: HashMap<String, ZipEntry>,
    team_options: Vec<String>,
    merged_pdf: Option<PathBuf>,
    processing_started_at: Option<DateTime<Utc>>,
    processing_completed_at: Option<DateTime<Utc>>,
    report_number: Option<String>,
    progress_current: usize,
    progress_total: usize,
    email_delivery_status: String,
    manual_email_delivery_status: String,
    manual_email_sent_once: bool,
    conversion_order: Vec<String>,
    conversion_statuses: HashMap<String, String>,
    conversion_threads: HashMap<String, usize>,
    show_conversion_progress: bool,
}

#[derive(Debug)]
struct ApiError {
    status: StatusCode,
    message: String,
}

impl IntoResponse for ApiError {
    fn into_response(self) -> Response {
        (self.status, Json(json!({ "error": self.message }))).into_response()
    }
}

impl ApiError {
    fn bad_request(message: impl Into<String>) -> Self {
        Self {
            status: StatusCode::BAD_REQUEST,
            message: message.into(),
        }
    }
    fn not_found(message: impl Into<String>) -> Self {
        Self {
            status: StatusCode::NOT_FOUND,
            message: message.into(),
        }
    }
    fn internal(message: impl Into<String>) -> Self {
        Self {
            status: StatusCode::INTERNAL_SERVER_ERROR,
            message: message.into(),
        }
    }
}

#[derive(Debug, Deserialize, Default)]
struct IndexQuery {
    message: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct TeamQuery {
    team: Option<String>,
}

#[derive(Debug, Deserialize, Default)]
struct StartForm {
    #[serde(default)]
    job_id: String,
    #[serde(default)]
    order: String,
}

#[derive(Debug, Deserialize)]
struct PrintRequest {
    printer_name: String,
}

fn main() {
    if let Err(err) = run_application() {
        append_startup_log(None, &format!("startup failed: {err:#}"));
    }
}

fn run_application() -> Result<()> {
    init_tracing();

    let base_dir = resolve_base_dir()?;
    load_environment_variables(&base_dir);
    let server_port = resolve_server_port(&base_dir);
    let paths = RuntimePaths::prepare(base_dir)?;
    cleanup_data_directories(&paths).ok();
    let state = build_app_state(paths)?;

    info!(
        "base_dir={}, server_port={server_port}",
        state.paths.base_dir.display()
    );

    #[cfg(target_os = "windows")]
    {
        let initial_server = match start_server_blocking(&state, server_port) {
            Ok(server) => server,
            Err(err) => {
                append_startup_log(Some(&state.paths.base_dir), &format!("{err:#}"));
                return Err(err);
            }
        };
        run_tray_loop(state, server_port, initial_server)
    }

    #[cfg(not(target_os = "windows"))]
    {
        let runtime = tokio::runtime::Builder::new_multi_thread()
            .enable_all()
            .build()
            .context("failed to build tokio runtime")?;
        runtime.block_on(run_http_server(state, server_port, None, None))
    }
}

fn resolve_server_port(base_dir: &Path) -> u16 {
    if let Ok(raw) = env::var("RPC_PORT") {
        if let Some(port) = parse_server_port_value(&raw) {
            return port;
        }
        let trimmed = raw.trim();
        if !trimmed.is_empty() {
            error!(
                "invalid RPC_PORT='{}', fallback to {}",
                trimmed, DEFAULT_SERVER_PORT
            );
        }
    }

    for candidate in env_file_candidates(base_dir, ".env") {
        if let Some(port) = read_rpc_port_from_env_file(&candidate) {
            return port;
        }
    }

    DEFAULT_SERVER_PORT
}

fn parse_server_port_value(raw: &str) -> Option<u16> {
    let trimmed = raw.trim();
    if trimmed.is_empty() {
        return None;
    }
    match trimmed.parse::<u16>() {
        Ok(port) if port > 0 => Some(port),
        _ => None,
    }
}

fn read_rpc_port_from_env_file(path: &Path) -> Option<u16> {
    if !path.exists() {
        return None;
    }

    let bytes = match fs::read(path) {
        Ok(bytes) => bytes,
        Err(err) => {
            error!("failed to read {}: {err}", path.display());
            return None;
        }
    };

    let content = decode_env_file_content(&bytes);
    for line in content.lines() {
        let line = line.trim_start_matches('\u{feff}').trim();
        if line.is_empty() || line.starts_with('#') {
            continue;
        }
        let Some((key, value)) = line.split_once('=') else {
            continue;
        };
        if key.trim() != "RPC_PORT" {
            continue;
        }
        let normalized = normalize_env_file_value(value);

        if let Some(port) = parse_server_port_value(&normalized) {
            return Some(port);
        }

        if !normalized.is_empty() {
            error!("invalid RPC_PORT in {}: '{}'", path.display(), normalized);
        }
        return None;
    }

    None
}

fn normalize_env_file_value(raw: &str) -> String {
    let value_without_comment = raw.split('#').next().unwrap_or_default().trim();
    value_without_comment
        .strip_prefix('"')
        .and_then(|v| v.strip_suffix('"'))
        .or_else(|| {
            value_without_comment
                .strip_prefix('\'')
                .and_then(|v| v.strip_suffix('\''))
        })
        .unwrap_or(value_without_comment)
        .trim()
        .to_string()
}

fn decode_env_file_content(bytes: &[u8]) -> String {
    if bytes.starts_with(&[0xEF, 0xBB, 0xBF]) {
        return String::from_utf8_lossy(&bytes[3..]).into_owned();
    }
    if bytes.starts_with(&[0xFF, 0xFE]) {
        let mut units = Vec::with_capacity((bytes.len().saturating_sub(2)) / 2);
        for chunk in bytes[2..].chunks_exact(2) {
            units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
        }
        return String::from_utf16_lossy(&units);
    }
    if bytes.starts_with(&[0xFE, 0xFF]) {
        let mut units = Vec::with_capacity((bytes.len().saturating_sub(2)) / 2);
        for chunk in bytes[2..].chunks_exact(2) {
            units.push(u16::from_be_bytes([chunk[0], chunk[1]]));
        }
        return String::from_utf16_lossy(&units);
    }
    String::from_utf8_lossy(bytes).into_owned()
}

fn env_file_candidates(base_dir: &Path, file_name: &str) -> Vec<PathBuf> {
    let mut candidates = vec![base_dir.join(file_name)];
    if let Ok(cwd) = env::current_dir() {
        if cwd != base_dir {
            candidates.push(cwd.join(file_name));
        }
    }
    candidates
}

fn build_app_state(paths: RuntimePaths) -> Result<AppState> {
    let template_glob = format!(
        "{}/**/*",
        paths.templates_dir.to_string_lossy().replace('\\', "/")
    );
    let tera = Tera::new(&template_glob).with_context(|| {
        format!(
            "failed to load templates from {}",
            paths.templates_dir.display()
        )
    })?;

    Ok(AppState {
        jobs: Arc::new(RwLock::new(HashMap::new())),
        upload_sessions: Arc::new(RwLock::new(HashMap::new())),
        paths: paths.clone(),
        tera: Arc::new(tera),
        order_manager: Arc::new(OrderManager::new(paths.order_file.clone())),
        job_semaphore: Arc::new(Semaphore::new(1)),
    })
}

#[cfg(target_os = "windows")]
struct RunningServer {
    port: u16,
    shutdown_tx: Option<tokio::sync::oneshot::Sender<()>>,
    join_handle: Option<thread::JoinHandle<()>>,
}

#[cfg(target_os = "windows")]
impl RunningServer {
    fn shutdown(&mut self) {
        if let Some(tx) = self.shutdown_tx.take() {
            let _ = tx.send(());
        }
        if let Some(handle) = self.join_handle.take() {
            let _ = handle.join();
        }
    }
}

#[cfg(target_os = "windows")]
fn spawn_server_in_background(state: AppState, port: u16) -> (mpsc::Receiver<Result<(), String>>, RunningServer) {
    let (startup_tx, startup_rx) = mpsc::channel();
    let base_dir = state.paths.base_dir.clone();
    let (shutdown_tx, shutdown_rx) = tokio::sync::oneshot::channel::<()>();
    let join_handle = thread::spawn(move || {
        let runtime = match tokio::runtime::Builder::new_multi_thread()
            .enable_all()
            .build()
        {
            Ok(runtime) => runtime,
            Err(err) => {
                let message = format!("failed to build tokio runtime: {err}");
                let _ = startup_tx.send(Err(message.clone()));
                append_startup_log(Some(&base_dir), &message);
                return;
            }
        };

        if let Err(err) =
            runtime.block_on(run_http_server(state, port, Some(startup_tx), Some(shutdown_rx)))
        {
            append_startup_log(Some(&base_dir), &format!("server terminated: {err:#}"));
        }
    });
    (
        startup_rx,
        RunningServer {
            port,
            shutdown_tx: Some(shutdown_tx),
            join_handle: Some(join_handle),
        },
    )
}

#[cfg(target_os = "windows")]
fn start_server_blocking(state: &AppState, port: u16) -> Result<RunningServer> {
    let (startup_rx, mut server) = spawn_server_in_background(state.clone(), port);
    match startup_rx.recv_timeout(Duration::from_secs(SERVER_START_TIMEOUT_SECONDS)) {
        Ok(Ok(())) => Ok(server),
        Ok(Err(message)) => {
            server.shutdown();
            Err(anyhow!(message))
        }
        Err(err) => {
            server.shutdown();
            Err(anyhow!("server startup timed out: {err}"))
        }
    }
}

fn build_router(state: AppState) -> Router {
    let app = Router::new()
        .route("/", get(index))
        .route("/prepare", post(prepare_upload))
        .route("/default-order-editor", get(default_order_editor))
        .route("/recipient-selector", get(recipient_selector))
        .route(
            "/default-order-editor/save",
            post(save_default_member_order),
        )
        .route(
            "/default-order-editor/delete",
            post(delete_default_member_order),
        )
        .route("/api/default-order", get(api_default_order))
        .route("/send-email/{job_id}", post(send_email))
        .route("/start", post(start_processing))
        .route("/download/{job_id}", get(download_merged_pdf))
        .route("/api/printers", get(api_printers))
        .route("/print/{job_id}", post(print_pdf))
        .route("/status/{job_id}", get(job_status))
        .route("/health", get(health))
        .nest_service("/static", ServeDir::new(state.paths.static_dir.clone()))
        .layer(DefaultBodyLimit::max(MAX_UPLOAD_BODY_BYTES))
        .with_state(state.clone())
        .layer(TraceLayer::new_for_http());

    app
}

async fn run_http_server(
    state: AppState,
    port: u16,
    startup_tx: Option<mpsc::Sender<Result<(), String>>>,
    shutdown_rx: Option<tokio::sync::oneshot::Receiver<()>>,
) -> Result<()> {
    let app = build_router(state);
    let mut startup_tx = startup_tx;

    let listener = match tokio::net::TcpListener::bind(("0.0.0.0", port)).await {
        Ok(listener) => {
            if let Some(tx) = startup_tx.take() {
                let _ = tx.send(Ok(()));
            }
            listener
        }
        Err(err) => {
            let message = format!("failed to bind port {port}: {err}");
            if let Some(tx) = startup_tx.take() {
                let _ = tx.send(Err(message.clone()));
            }
            return Err(err).context(message);
        }
    };

    info!("server started: http://127.0.0.1:{port}");

    let server = axum::serve(listener, app);
    if let Some(shutdown_rx) = shutdown_rx {
        server
            .with_graceful_shutdown(async move {
                let _ = shutdown_rx.await;
            })
            .await
            .context("server terminated unexpectedly")
    } else {
        server.await.context("server terminated unexpectedly")
    }
}

fn init_tracing() {
    tracing_subscriber::fmt()
        .with_env_filter(
            tracing_subscriber::EnvFilter::try_from_default_env()
                .unwrap_or_else(|_| "info,tower_http=info".into()),
        )
        .init();
}

fn resolve_base_dir() -> Result<PathBuf> {
    if let Ok(value) = env::var("RPC_BASE_DIR") {
        let candidate = PathBuf::from(value);
        if candidate.is_absolute() {
            return Ok(candidate);
        }
        return Ok(env::current_dir()?.join(candidate));
    }
    if cfg!(debug_assertions) {
        return env::current_dir().context("failed to resolve current directory");
    }
    let exe = env::current_exe().context("failed to resolve executable path")?;
    let base = exe
        .parent()
        .ok_or_else(|| anyhow!("executable parent directory is missing"))?;
    Ok(base.to_path_buf())
}

fn resolve_data_dir(_base_dir: &Path) -> PathBuf {
    #[cfg(target_os = "windows")]
    {
        if let Ok(local_app_data) = env::var("LOCALAPPDATA") {
            return PathBuf::from(local_app_data)
                .join("ReportPDFConverter")
                .join("data");
        }
        if let Ok(user_profile) = env::var("USERPROFILE") {
            return PathBuf::from(user_profile)
                .join("AppData")
                .join("Local")
                .join("ReportPDFConverter")
                .join("data");
        }
        return env::temp_dir().join("ReportPDFConverter").join("data");
    }
    #[cfg(not(target_os = "windows"))]
    {
        _base_dir.join("data")
    }
}

fn append_startup_log(base_dir: Option<&Path>, message: &str) {
    let data_dir = match base_dir {
        Some(base_dir) => resolve_data_dir(base_dir),
        None => {
            #[cfg(target_os = "windows")]
            {
                if let Ok(local_app_data) = env::var("LOCALAPPDATA") {
                    PathBuf::from(local_app_data)
                        .join("ReportPDFConverter")
                        .join("data")
                } else {
                    env::temp_dir().join("ReportPDFConverter").join("data")
                }
            }
            #[cfg(not(target_os = "windows"))]
            {
                env::temp_dir().join("report-pdf-converter").join("data")
            }
        }
    };

    if fs::create_dir_all(&data_dir).is_err() {
        return;
    }

    let log_path = data_dir.join("startup.log");
    let line = format!("{} {message}\n", Utc::now().to_rfc3339());
    if let Ok(mut file) = fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(log_path)
    {
        let _ = file.write_all(line.as_bytes());
    }
}

#[cfg(target_os = "windows")]
fn open_application_url(port: u16) {
    let url = format!("http://127.0.0.1:{port}/");
    let _ = Command::new("cmd")
        .creation_flags(CREATE_NO_WINDOW_FLAG)
        .args(["/C", "start", "", &url])
        .spawn();
}

#[cfg(target_os = "windows")]
fn load_tray_icon(paths: &RuntimePaths) -> Result<Icon> {
    let icon_path = paths.static_dir.join("favicon.ico");
    if icon_path.exists() {
        return Icon::from_path(&icon_path, Some((32, 32)))
            .map_err(|err| anyhow!("failed to load tray icon {}: {err}", icon_path.display()));
    }
    Icon::from_resource(1, Some((32, 32)))
        .map_err(|err| anyhow!("failed to load embedded icon resource: {err}"))
}

#[cfg(target_os = "windows")]
enum TrayUserEvent {
    Tray(TrayIconEvent),
    Menu(MenuEvent),
}

#[cfg(target_os = "windows")]
fn windows_startup_enabled() -> bool {
    let hkcu = RegKey::predef(HKEY_CURRENT_USER);
    let Ok(run_key) = hkcu.open_subkey(WINDOWS_STARTUP_REG_PATH) else {
        return false;
    };
    run_key
        .get_value::<String, _>(WINDOWS_STARTUP_VALUE_NAME)
        .map(|v| !v.trim().is_empty())
        .unwrap_or(false)
}

#[cfg(target_os = "windows")]
fn set_windows_startup_enabled(enabled: bool) -> Result<()> {
    let hkcu = RegKey::predef(HKEY_CURRENT_USER);
    let (run_key, _) = hkcu
        .create_subkey(WINDOWS_STARTUP_REG_PATH)
        .context("failed to open startup registry key")?;

    if enabled {
        let exe_path = env::current_exe().context("failed to resolve current executable path")?;
        let command = format!("\"{}\"", exe_path.display());
        run_key
            .set_value(WINDOWS_STARTUP_VALUE_NAME, &command)
            .context("failed to set startup registry value")?;
    } else if let Err(err) = run_key.delete_value(WINDOWS_STARTUP_VALUE_NAME)
        && err.kind() != io::ErrorKind::NotFound
    {
        return Err(err).context("failed to delete startup registry value");
    }

    Ok(())
}

#[cfg(target_os = "windows")]
fn create_status_dot_icon(running: bool) -> Option<MenuIcon> {
    const SIZE: u32 = 16;
    const CENTER: f32 = 7.5;
    const OUTER_RADIUS_SQ: f32 = 42.0;
    const INNER_RADIUS_SQ: f32 = 30.0;
    let (r, g, b) = if running {
        (255u8, 255u8, 255u8)
    } else {
        (0u8, 0u8, 0u8)
    };

    let mut rgba = vec![0u8; (SIZE * SIZE * 4) as usize];
    for y in 0..SIZE {
        for x in 0..SIZE {
            let dx = x as f32 - CENTER;
            let dy = y as f32 - CENTER;
            let d2 = dx * dx + dy * dy;
            if d2 <= OUTER_RADIUS_SQ {
                let idx = ((y * SIZE + x) * 4) as usize;
                if d2 >= INNER_RADIUS_SQ {
                    rgba[idx] = 255;
                    rgba[idx + 1] = 255;
                    rgba[idx + 2] = 255;
                    rgba[idx + 3] = 255;
                } else {
                    rgba[idx] = r;
                    rgba[idx + 1] = g;
                    rgba[idx + 2] = b;
                    rgba[idx + 3] = 255;
                }
            }
        }
    }

    MenuIcon::from_rgba(rgba, SIZE, SIZE).ok()
}

#[cfg(target_os = "windows")]
fn update_tray_ui(
    status_item: &IconMenuItem,
    port_item: &MenuItem,
    url_item: &MenuItem,
    start_item: &MenuItem,
    stop_item: &MenuItem,
    tray_icon: &TrayIcon,
    running: bool,
    port: u16,
) {
    status_item.set_text(if running {
        "サーバー状態: 起動中"
    } else {
        "サーバー状態: 停止中"
    });
    status_item.set_icon(create_status_dot_icon(running));
    port_item.set_text(format!("ポート: {port}"));
    url_item.set_text(format!("URL: http://127.0.0.1:{port}/"));
    start_item.set_enabled(!running);
    stop_item.set_enabled(running);
    let tooltip = if running {
        format!("{APP_DISPLAY_NAME} (ON) http://127.0.0.1:{port}")
    } else {
        format!("{APP_DISPLAY_NAME} (OFF)")
    };
    let _ = tray_icon.set_tooltip(Some(tooltip));
}

#[cfg(target_os = "windows")]
fn run_tray_loop(state: AppState, configured_port: u16, initial_server: RunningServer) -> Result<()> {
    let event_loop = EventLoopBuilder::<TrayUserEvent>::with_user_event().build();

    let tray_proxy = event_loop.create_proxy();
    TrayIconEvent::set_event_handler(Some(move |event| {
        let _ = tray_proxy.send_event(TrayUserEvent::Tray(event));
    }));
    let menu_proxy = event_loop.create_proxy();
    MenuEvent::set_event_handler(Some(move |event| {
        let _ = menu_proxy.send_event(TrayUserEvent::Menu(event));
    }));

    let status_item = IconMenuItem::new(
        "サーバー状態: 起動中",
        true,
        create_status_dot_icon(true),
        None,
    );
    let port_item = MenuItem::new(format!("ポート: {configured_port}"), true, None);
    let url_item = MenuItem::new(format!("URL: http://127.0.0.1:{configured_port}/"), true, None);
    let start_item = MenuItem::new("サーバー開始", true, None);
    let stop_item = MenuItem::new("サーバー停止", true, None);
    let open_item = MenuItem::new("画面を開く", true, None);
    let startup_item = CheckMenuItem::new(
        "Windows起動時に自動起動",
        true,
        windows_startup_enabled(),
        None,
    );
    let quit_item = MenuItem::new("終了", true, None);

    let status_id = status_item.id().clone();
    let port_id = port_item.id().clone();
    let url_id = url_item.id().clone();
    let start_id = start_item.id().clone();
    let stop_id = stop_item.id().clone();
    let open_id = open_item.id().clone();
    let startup_id = startup_item.id().clone();
    let quit_id = quit_item.id().clone();

    let separator1 = PredefinedMenuItem::separator();
    let separator2 = PredefinedMenuItem::separator();
    let separator3 = PredefinedMenuItem::separator();

    let tray_menu = Menu::new();
    tray_menu
        .append_items(&[
            &status_item,
            &port_item,
            &url_item,
            &separator1,
            &start_item,
            &stop_item,
            &separator2,
            &open_item,
            &startup_item,
            &separator3,
            &quit_item,
        ])
        .map_err(|err| anyhow!("failed to create tray menu: {err}"))?;

    // Keep these rows non-interactive while avoiding the grayed-out icon rendering on Windows.
    status_item.set_enabled(false);
    port_item.set_enabled(false);
    url_item.set_enabled(false);

    let tray_icon_image = load_tray_icon(&state.paths)?;
    let tray_icon = TrayIconBuilder::new()
        .with_tooltip(APP_DISPLAY_NAME)
        .with_menu(Box::new(tray_menu))
        .with_icon(tray_icon_image)
        .build()
        .map_err(|err| anyhow!("failed to create tray icon: {err}"))?;
    tray_icon.set_show_menu_on_left_click(true);

    let mut running_server = Some(initial_server);
    let mut active_port = configured_port;
    update_tray_ui(
        &status_item,
        &port_item,
        &url_item,
        &start_item,
        &stop_item,
        &tray_icon,
        true,
        active_port,
    );

    event_loop.run(move |event, _, control_flow| {
        *control_flow = ControlFlow::Wait;
        match event {
            Event::UserEvent(TrayUserEvent::Menu(menu_event)) => {
                if menu_event.id == status_id
                    || menu_event.id == port_id
                    || menu_event.id == url_id
                {
                    return;
                }

                if menu_event.id == start_id {
                    if running_server.is_none() {
                        match start_server_blocking(&state, configured_port) {
                            Ok(server) => {
                                active_port = server.port;
                                running_server = Some(server);
                                update_tray_ui(
                                    &status_item,
                                    &port_item,
                                    &url_item,
                                    &start_item,
                                    &stop_item,
                                    &tray_icon,
                                    true,
                                    active_port,
                                );
                            }
                            Err(err) => {
                                append_startup_log(
                                    Some(&state.paths.base_dir),
                                    &format!("server start failed: {err:#}"),
                                );
                                update_tray_ui(
                                    &status_item,
                                    &port_item,
                                    &url_item,
                                    &start_item,
                                    &stop_item,
                                    &tray_icon,
                                    false,
                                    configured_port,
                                );
                            }
                        }
                    }
                } else if menu_event.id == stop_id {
                    if let Some(mut server) = running_server.take() {
                        server.shutdown();
                    }
                    update_tray_ui(
                        &status_item,
                        &port_item,
                        &url_item,
                        &start_item,
                        &stop_item,
                        &tray_icon,
                        false,
                        configured_port,
                    );
                } else if menu_event.id == open_id {
                    if running_server.is_none() {
                        match start_server_blocking(&state, configured_port) {
                            Ok(server) => {
                                active_port = server.port;
                                running_server = Some(server);
                                update_tray_ui(
                                    &status_item,
                                    &port_item,
                                    &url_item,
                                    &start_item,
                                    &stop_item,
                                    &tray_icon,
                                    true,
                                    active_port,
                                );
                            }
                            Err(err) => {
                                append_startup_log(
                                    Some(&state.paths.base_dir),
                                    &format!("server start failed: {err:#}"),
                                );
                                update_tray_ui(
                                    &status_item,
                                    &port_item,
                                    &url_item,
                                    &start_item,
                                    &stop_item,
                                    &tray_icon,
                                    false,
                                    configured_port,
                                );
                                return;
                            }
                        }
                    }
                    open_application_url(active_port);
                } else if menu_event.id == startup_id {
                    let enabled = startup_item.is_checked();
                    if let Err(err) = set_windows_startup_enabled(enabled) {
                        error!("failed to update Windows startup setting: {err:#}");
                        startup_item.set_checked(!enabled);
                    }
                } else if menu_event.id == quit_id {
                    if let Some(mut server) = running_server.take() {
                        server.shutdown();
                    }
                    *control_flow = ControlFlow::Exit;
                }
            }
            Event::UserEvent(TrayUserEvent::Tray(TrayIconEvent::DoubleClick { .. })) => {
                if running_server.is_none() {
                    match start_server_blocking(&state, configured_port) {
                        Ok(server) => {
                            active_port = server.port;
                            running_server = Some(server);
                            update_tray_ui(
                                &status_item,
                                &port_item,
                                &url_item,
                                &start_item,
                                &stop_item,
                                &tray_icon,
                                true,
                                active_port,
                            );
                        }
                        Err(err) => {
                            append_startup_log(
                                Some(&state.paths.base_dir),
                                &format!("server start failed: {err:#}"),
                            );
                            update_tray_ui(
                                &status_item,
                                &port_item,
                                &url_item,
                                &start_item,
                                &stop_item,
                                &tray_icon,
                                false,
                                configured_port,
                            );
                            return;
                        }
                    }
                }
                open_application_url(active_port);
            }
            _ => {}
        }
        let _ = &tray_icon;
    });
}

fn load_environment_variables(base_dir: &Path) {
    for candidate in env_file_candidates(base_dir, ".env") {
        if candidate.exists() {
            if let Err(err) = from_path_override(&candidate) {
                error!("failed to load {}: {err}", candidate.display());
            }
        }
    }
}

async fn health() -> Json<Value> {
    Json(json!({ "status": "ok" }))
}

async fn index(
    State(state): State<AppState>,
    Query(query): Query<IndexQuery>,
) -> Result<Response, ApiError> {
    let mut messages = Vec::new();
    if let Some(message) = query.message {
        let trimmed = message.trim();
        if !trimmed.is_empty() {
            messages.push(trimmed.to_string());
        }
    }
    render_index(&state, messages)
}

fn render_index(state: &AppState, messages: Vec<String>) -> Result<Response, ApiError> {
    let mut context = TeraContext::new();
    context.insert("messages", &messages);
    render_template(state, "index.html", &context)
}

fn render_template(
    state: &AppState,
    name: &str,
    context: &TeraContext,
) -> Result<Response, ApiError> {
    let body = state
        .tera
        .render(name, context)
        .map_err(|err| ApiError::internal(format!("template render failed ({name}): {err}")))?;
    Ok(Html(body).into_response())
}

fn json_response(status: StatusCode, payload: Value) -> Response {
    (status, Json(payload)).into_response()
}

async fn prepare_upload(
    State(state): State<AppState>,
    mut multipart: Multipart,
) -> Result<Response, ApiError> {
    let mut zip_bytes: Option<Vec<u8>> = None;
    let mut upload_name: Option<String> = None;

    while let Some(field) = multipart
        .next_field()
        .await
        .map_err(|err| ApiError::bad_request(format!("multipart parse error: {err}")))?
    {
        if field.name().unwrap_or_default() != "zip_file" {
            continue;
        }
        upload_name = field.file_name().map(|v| v.to_string());
        let bytes = field
            .bytes()
            .await
            .map_err(|err| ApiError::bad_request(format!("failed to read upload: {err}")))?;
        zip_bytes = Some(bytes.to_vec());
        break;
    }

    let zip_bytes = match zip_bytes {
        Some(v) => v,
        None => {
            return render_index(
                &state,
                vec!["アップロードするZIPファイルを選択してください。".to_string()],
            );
        }
    };

    let job_id = Uuid::new_v4().simple().to_string();
    let zip_path = state.paths.uploads_dir.join(format!("{job_id}.zip"));
    tokio::fs::write(&zip_path, &zip_bytes)
        .await
        .map_err(|err| ApiError::internal(format!("failed to save upload: {err}")))?;

    let original_name = upload_name.clone();
    let zip_path_for_extract = zip_path.clone();
    let entries = tokio::task::spawn_blocking(move || {
        extract_entries(&zip_path_for_extract, original_name.as_deref())
    })
    .await
    .map_err(|err| ApiError::internal(format!("prepare task failed: {err}")))?
    .map_err(|err| ApiError::internal(format!("failed to inspect zip: {err:#}")))?;

    if entries.is_empty() {
        let _ = fs::remove_file(&zip_path);
        return render_index(
            &state,
            vec!["アップロードされたZIPにWordファイルが見つかりませんでした。".to_string()],
        );
    }

    let preferences = state.order_manager.load_preferences();
    let preference_team_options = team_labels_from_preferences(&preferences);
    let entry_map = entries
        .into_iter()
        .map(|entry| (entry.display_name.clone(), entry))
        .collect::<HashMap<_, _>>();

    let (team_sequence, team_entries) = state
        .order_manager
        .initial_layout(entry_map.values().cloned().collect());

    let mut ordered_display_names = Vec::new();
    let mut team_blocks = Vec::<Value>::new();
    for team_key in &team_sequence {
        let items = team_entries.get(team_key).cloned().unwrap_or_default();
        if items.is_empty() {
            continue;
        }
        let mut block_entries = Vec::new();
        for item in items {
            ordered_display_names.push(item.display_name.clone());
            block_entries.push(json!({
                "display_name": item.display_name,
                "team": item.team_name.unwrap_or_default(),
                "persons": item.persons,
            }));
        }
        team_blocks.push(json!({
            "key": team_key,
            "label": team_display_label(team_key),
            "count": block_entries.len(),
            "entries": block_entries,
        }));
    }

    let team_options = team_blocks
        .iter()
        .filter_map(|block| {
            block
                .get("label")
                .and_then(Value::as_str)
                .map(str::to_string)
        })
        .collect::<Vec<_>>();
    let session_team_options = if !preference_team_options.is_empty() {
        preference_team_options
    } else {
        team_options
    };

    let default_member_sequences = preferences
        .member_sequences
        .iter()
        .map(|(key, members)| (key.clone(), extract_member_names(members)))
        .collect::<HashMap<_, _>>();

    {
        let mut sessions = state.upload_sessions.write().await;
        sessions.insert(
            job_id.clone(),
            UploadSession {
                entries: entry_map,
                team_options: session_team_options,
                zip_filename: upload_name.unwrap_or_default(),
            },
        );
    }

    let initial_state = team_blocks
        .iter()
        .map(|block| {
            json!({
                "key": block.get("key").cloned().unwrap_or(Value::String(String::new())),
                "label": block.get("label").cloned().unwrap_or(Value::String(String::new())),
                "entries": block.get("entries").cloned().unwrap_or(Value::Array(Vec::new())),
            })
        })
        .collect::<Vec<_>>();
    let initial_state_json = serde_json::to_string(&initial_state)
        .map_err(|err| ApiError::internal(format!("failed to serialize initial state: {err}")))?;
    let default_member_sequences_json =
        serde_json::to_string(&default_member_sequences).map_err(|err| {
            ApiError::internal(format!(
                "failed to serialize default member sequences: {err}"
            ))
        })?;

    let mut context = TeraContext::new();
    context.insert("job_id", &job_id);
    context.insert("ordered_display_names", &ordered_display_names);
    context.insert("team_blocks", &team_blocks);
    context.insert("initial_state_json", &initial_state_json);
    context.insert(
        "default_member_sequences_json",
        &default_member_sequences_json,
    );
    render_template(&state, "order.html", &context)
}

async fn default_order_editor(
    State(state): State<AppState>,
    Query(query): Query<TeamQuery>,
) -> Result<Response, ApiError> {
    let preferences = state.order_manager.load_preferences();
    let teams = collect_preference_teams(&preferences);
    let available_keys = teams
        .iter()
        .filter_map(|team| team.get("key").and_then(Value::as_str).map(str::to_string))
        .collect::<HashSet<_>>();

    let mut initial_team_key = query.team.and_then(|team| {
        let normalized = normalize_team_key(&team);
        if available_keys.contains(&normalized) {
            Some(normalized)
        } else {
            None
        }
    });
    if initial_team_key.is_none() {
        if let Some(first) = teams
            .first()
            .and_then(|team| team.get("key").and_then(Value::as_str))
        {
            initial_team_key = Some(first.to_string());
        } else {
            initial_team_key = Some(UNGROUPED_TEAM_KEY.to_string());
        }
    }

    let initial_data = json!({
        "teams": teams,
        "team_order": teams.iter().filter_map(|v| v.get("key").and_then(Value::as_str)).collect::<Vec<_>>(),
        "initial_team_key": initial_team_key.unwrap_or_else(|| UNGROUPED_TEAM_KEY.to_string()),
        "ungrouped_key": UNGROUPED_TEAM_KEY,
        "ungrouped_label": team_display_label(UNGROUPED_TEAM_KEY),
    });
    let initial_data_json = serde_json::to_string(&initial_data).map_err(|err| {
        ApiError::internal(format!("failed to serialize default order data: {err}"))
    })?;

    let mut context = TeraContext::new();
    context.insert("initial_data_json", &initial_data_json);
    render_template(&state, "default_order_editor.html", &context)
}

async fn recipient_selector(State(state): State<AppState>) -> Result<Response, ApiError> {
    let preferences = state.order_manager.load_preferences();
    let groups = collect_email_recipient_groups(&preferences, &state.paths.faculty_contacts_file);
    let initial_data_json = serde_json::to_string(&json!({ "groups": groups }))
        .map_err(|err| ApiError::internal(format!("failed to serialize recipient data: {err}")))?;
    let mut context = TeraContext::new();
    context.insert("initial_data_json", &initial_data_json);
    render_template(&state, "recipient_selector.html", &context)
}

async fn save_default_member_order(
    State(state): State<AppState>,
    Json(payload): Json<Value>,
) -> Response {
    let team_key = payload
        .get("team_key")
        .map(value_to_trimmed_string)
        .unwrap_or_default();
    let members_raw = payload.get("members").and_then(Value::as_array).cloned();

    if team_key.is_empty() {
        return json_response(
            StatusCode::BAD_REQUEST,
            json!({ "error": "班名を入力してください。" }),
        );
    }
    let members_raw = match members_raw {
        Some(v) => v,
        None => {
            return json_response(
                StatusCode::BAD_REQUEST,
                json!({ "error": "メンバー情報の形式が正しくありません。" }),
            );
        }
    };

    let normalized_team = normalize_team_key(&team_key);
    let preferences = state.order_manager.load_preferences();
    let existing = preferences
        .member_sequences
        .get(&normalized_team)
        .cloned()
        .unwrap_or_default();
    let cleaned_members = normalize_member_objects_from_values(&members_raw, &existing);

    if let Err(err) = state
        .order_manager
        .save_member_sequence(&normalized_team, &cleaned_members)
    {
        error!("failed to save default member order: {err:#}");
        return json_response(
            StatusCode::INTERNAL_SERVER_ERROR,
            json!({ "error": "保存に失敗しました。" }),
        );
    }

    json_response(
        StatusCode::OK,
        json!({
            "status": "ok",
            "team_key": normalized_team,
            "label": team_display_label(&normalized_team),
            "members": extract_member_names(&cleaned_members),
            "member_details": cleaned_members,
        }),
    )
}

async fn delete_default_member_order(
    State(state): State<AppState>,
    Json(payload): Json<Value>,
) -> Response {
    let team_key = payload
        .get("team_key")
        .map(value_to_trimmed_string)
        .unwrap_or_default();
    if team_key.is_empty() {
        return json_response(
            StatusCode::BAD_REQUEST,
            json!({ "error": "班を指定してください。" }),
        );
    }

    let normalized_team = normalize_team_key(&team_key);
    if let Err(err) = state.order_manager.delete_member_sequence(&normalized_team) {
        error!("failed to delete default member order: {err:#}");
        return json_response(
            StatusCode::INTERNAL_SERVER_ERROR,
            json!({ "error": "削除に失敗しました。" }),
        );
    }

    json_response(
        StatusCode::OK,
        json!({ "status": "ok", "team_key": normalized_team }),
    )
}

async fn api_default_order(State(state): State<AppState>) -> Json<Value> {
    let preferences = state.order_manager.load_preferences();
    let teams = collect_preference_teams(&preferences);
    let member_sequences = preferences
        .member_sequences
        .iter()
        .map(|(key, members)| (key.clone(), extract_member_names(members)))
        .collect::<HashMap<_, _>>();
    Json(json!({
        "team_sequence": preferences.team_sequence,
        "member_sequences": member_sequences,
        "member_details": preferences.member_sequences,
        "teams": teams,
    }))
}

async fn send_email(
    State(state): State<AppState>,
    AxumPath(job_id): AxumPath<String>,
    Json(payload): Json<Value>,
) -> Response {
    if !email_transport_is_configured(&state.paths) {
        return json_response(
            StatusCode::SERVICE_UNAVAILABLE,
            json!({ "error": "メール設定が構成されていません。" }),
        );
    }

    let (mut to_recipients, mut cc_recipients) = parse_recipients(payload.get("recipients"));
    if to_recipients.is_empty() && !cc_recipients.is_empty() {
        to_recipients = cc_recipients.clone();
        cc_recipients.clear();
    }
    if to_recipients.is_empty() {
        return json_response(
            StatusCode::BAD_REQUEST,
            json!({ "error": "宛先を選択してください。" }),
        );
    }

    let confirm_resend = parse_bool_value(payload.get("confirm_resend"));
    let mut subject = payload
        .get("subject")
        .and_then(Value::as_str)
        .unwrap_or_default()
        .trim()
        .to_string();
    let body = payload
        .get("body")
        .and_then(Value::as_str)
        .unwrap_or_default()
        .to_string();

    let attachment_path = {
        let mut jobs = state.jobs.write().await;
        let Some(job) = jobs.get_mut(&job_id) else {
            return json_response(
                StatusCode::BAD_REQUEST,
                json!({ "error": "完了済みのジョブが見つかりません。" }),
            );
        };
        if job.status != "completed" || job.merged_pdf.is_none() {
            return json_response(
                StatusCode::BAD_REQUEST,
                json!({ "error": "完了済みのジョブが見つかりません。" }),
            );
        }
        let current_manual = job.manual_email_delivery_status.trim().to_ascii_lowercase();
        if current_manual == "sending" {
            return json_response(
                StatusCode::CONFLICT,
                json!({
                    "error": "現在メール送信中です。完了後に再度ご確認ください。",
                    "manual_email_delivery_status": "sending",
                    "manual_email_sent_once": job.manual_email_sent_once,
                }),
            );
        }
        if job.manual_email_sent_once && !confirm_resend {
            return json_response(
                StatusCode::CONFLICT,
                json!({
                    "error": "既にメールを送信しています。再送するには「もう一度メールを送る」にチェックしてください。",
                    "requires_resend_confirmation": true,
                    "manual_email_delivery_status": if job.manual_email_delivery_status.is_empty() { "sent" } else { &job.manual_email_delivery_status },
                    "manual_email_sent_once": true,
                }),
            );
        }
        if subject.is_empty() {
            subject = if let Some(number) = &job.report_number {
                format!("第{number}回報告書")
            } else {
                "報告書".to_string()
            };
        }
        job.manual_email_delivery_status = "sending".to_string();
        job.updated_at = Utc::now();
        job.merged_pdf.clone().unwrap_or_default()
    };

    if !attachment_path.exists() {
        return json_response(
            StatusCode::NOT_FOUND,
            json!({ "error": "添付するPDFが見つかりません。" }),
        );
    }

    let to_for_send = to_recipients.clone();
    let cc_for_send = cc_recipients.clone();
    let subject_for_send = subject.clone();
    let body_for_send = body.clone();
    let attachment_for_send = attachment_path.clone();
    let paths_for_send = state.paths.clone();
    let send_result = tokio::task::spawn_blocking(move || {
        send_email_with_attachment(
            &paths_for_send,
            &to_for_send,
            &cc_for_send,
            &subject_for_send,
            &body_for_send,
            &attachment_for_send,
        )
    })
    .await;

    let send_error = match send_result {
        Ok(Ok(())) => None,
        Ok(Err(err)) => Some(format!("{err:#}")),
        Err(err) => Some(format!("メール送信タスクの実行に失敗しました: {err}")),
    };
    let delivery_status = if send_error.is_some() {
        "failed"
    } else {
        "sent"
    }
    .to_string();

    {
        let mut jobs = state.jobs.write().await;
        if let Some(job) = jobs.get_mut(&job_id) {
            job.manual_email_delivery_status = delivery_status.clone();
            if delivery_status == "sent" {
                job.manual_email_sent_once = true;
            }
            job.updated_at = Utc::now();
        }
    }

    if let Some(err_message) = send_error {
        return json_response(
            StatusCode::BAD_GATEWAY,
            json!({
                "status": "failed",
                "error": format!("メール送信に失敗しました: {err_message}"),
                "email_delivery_status": "failed",
                "manual_email_delivery_status": "failed",
                "manual_email_sent_once": false,
                "to": to_recipients,
                "cc": cc_recipients,
            }),
        );
    }

    json_response(
        StatusCode::OK,
        json!({
            "status": "sent",
            "email_delivery_status": "sent",
            "manual_email_delivery_status": "sent",
            "manual_email_sent_once": true,
            "failures": [],
            "to": to_recipients,
            "cc": cc_recipients,
        }),
    )
}

async fn start_processing(
    State(state): State<AppState>,
    Form(form): Form<StartForm>,
) -> Result<Response, ApiError> {
    let job_id = form.job_id.trim().to_string();
    let order_data = form.order.trim().to_string();

    if job_id.is_empty() || order_data.is_empty() {
        return render_index(
            &state,
            vec!["アップロード情報を確認できませんでした。もう一度お試しください。".to_string()],
        );
    }

    let order = order_data
        .split('|')
        .map(str::trim)
        .filter(|v| !v.is_empty())
        .map(str::to_string)
        .collect::<Vec<_>>();
    if order.is_empty() {
        return render_index(
            &state,
            vec!["ドキュメントの並び順が空です。もう一度お試しください。".to_string()],
        );
    }

    let session = {
        let mut sessions = state.upload_sessions.write().await;
        sessions.remove(&job_id)
    };
    let session = match session {
        Some(v) => v,
        None => {
            return render_index(
                &state,
                vec![
                    "アップロード情報を確認できませんでした。もう一度お試しください。".to_string(),
                ],
            );
        }
    };

    if session.entries.is_empty() {
        return render_index(
            &state,
            vec!["処理対象のデータを取得できませんでした。もう一度お試しください。".to_string()],
        );
    }

    let mut team_options = session.team_options.clone();
    if team_options.is_empty() {
        let preferences = state.order_manager.load_preferences();
        team_options = team_labels_from_preferences(&preferences);
    }
    if team_options.is_empty() {
        let mut labels = session
            .entries
            .values()
            .map(|entry| {
                let team_key =
                    normalize_team_key(entry.team_name.as_deref().unwrap_or(UNGROUPED_TEAM_KEY));
                team_display_label(&team_key)
            })
            .collect::<Vec<_>>();
        labels.sort();
        labels.dedup();
        team_options = labels;
    }

    let now = Utc::now();
    let zip_path = state.paths.uploads_dir.join(format!("{job_id}.zip"));
    let zip_original_name = if session.zip_filename.trim().is_empty() {
        zip_path
            .file_name()
            .and_then(|v| v.to_str())
            .unwrap_or("upload.zip")
            .to_string()
    } else {
        session.zip_filename.clone()
    };

    let job_state = JobState {
        id: job_id.clone(),
        email: DEFAULT_RECIPIENT_EMAIL.to_string(),
        status: "queued".to_string(),
        message: "他のユーザの処理が終わるまで待機しています。".to_string(),
        created_at: now,
        updated_at: now,
        order,
        zip_path,
        zip_original_name,
        entries: session.entries.clone(),
        team_options,
        merged_pdf: None,
        processing_started_at: None,
        processing_completed_at: None,
        report_number: None,
        progress_current: 0,
        progress_total: session.entries.len(),
        email_delivery_status: "pending".to_string(),
        manual_email_delivery_status: String::new(),
        manual_email_sent_once: false,
        conversion_order: Vec::new(),
        conversion_statuses: HashMap::new(),
        conversion_threads: HashMap::new(),
        show_conversion_progress: false,
    };

    {
        let mut jobs = state.jobs.write().await;
        jobs.insert(job_id.clone(), job_state);
    }

    let state_for_task = state.clone();
    let job_id_for_task = job_id.clone();
    tokio::spawn(async move {
        process_job(state_for_task, job_id_for_task).await;
    });

    let mut context = TeraContext::new();
    context.insert("job_id", &job_id);
    render_template(&state, "status.html", &context)
}

async fn download_merged_pdf(
    State(state): State<AppState>,
    AxumPath(job_id): AxumPath<String>,
) -> Result<Response, ApiError> {
    let file_path = {
        let jobs = state.jobs.read().await;
        let job = jobs
            .get(&job_id)
            .ok_or_else(|| ApiError::not_found("ジョブが見つかりません。"))?;
        if job.status != "completed" {
            return Err(ApiError::bad_request("ジョブがまだ完了していません。"));
        }
        job.merged_pdf
            .clone()
            .ok_or_else(|| ApiError::not_found("結合済みPDFが見つかりません。"))?
    };

    let bytes = tokio::fs::read(&file_path)
        .await
        .map_err(|_| ApiError::not_found("PDFファイルが見つかりません。"))?;
    let file_name = file_path
        .file_name()
        .and_then(|v| v.to_str())
        .unwrap_or("report.pdf")
        .to_string();
    let disposition = format!("attachment; filename=\"{file_name}\"");
    Ok((
        [
            (header::CONTENT_TYPE, "application/pdf"),
            (header::CONTENT_DISPOSITION, disposition.as_str()),
        ],
        bytes,
    )
        .into_response())
}

async fn api_printers() -> Result<Json<Value>, ApiError> {
    let printers = tokio::task::spawn_blocking(query_printers)
        .await
        .map_err(|err| ApiError::internal(format!("printer task failed: {err}")))?
        .map_err(|err| ApiError::internal(format!("printer query failed: {err:#}")))?;
    Ok(Json(json!({ "printers": printers })))
}

async fn print_pdf(
    State(state): State<AppState>,
    AxumPath(job_id): AxumPath<String>,
    Json(payload): Json<PrintRequest>,
) -> Response {
    let printer_name = payload.printer_name.trim().to_string();
    if printer_name.is_empty() {
        return json_response(
            StatusCode::BAD_REQUEST,
            json!({ "error": "プリンタを選択してください。" }),
        );
    }
    if is_standard_printer(&printer_name) {
        return json_response(
            StatusCode::BAD_REQUEST,
            json!({ "error": "指定したプリンタは使用できません。" }),
        );
    }

    let pdf_path = {
        let jobs = state.jobs.read().await;
        let Some(job) = jobs.get(&job_id) else {
            return json_response(
                StatusCode::BAD_REQUEST,
                json!({ "error": "ジョブが完了していません。" }),
            );
        };
        if job.status != "completed" || job.merged_pdf.is_none() {
            return json_response(
                StatusCode::BAD_REQUEST,
                json!({ "error": "ジョブが完了していません。" }),
            );
        }
        job.merged_pdf.clone().unwrap_or_default()
    };

    if !pdf_path.exists() {
        return json_response(
            StatusCode::NOT_FOUND,
            json!({ "error": "印刷対象のPDFを見つけられませんでした。" }),
        );
    }

    let printers = match tokio::task::spawn_blocking(query_printers).await {
        Ok(Ok(list)) => list,
        _ => {
            return json_response(
                StatusCode::INTERNAL_SERVER_ERROR,
                json!({ "error": "プリンタ情報の取得に失敗しました。" }),
            );
        }
    };
    let lookup = printers
        .iter()
        .filter_map(|item| {
            let name = item.get("name").and_then(Value::as_str)?.trim().to_string();
            if name.is_empty() {
                None
            } else {
                Some((name.to_ascii_lowercase(), name))
            }
        })
        .collect::<HashMap<_, _>>();
    let canonical = match lookup.get(&printer_name.to_ascii_lowercase()) {
        Some(name) => name.clone(),
        None => {
            return json_response(
                StatusCode::BAD_REQUEST,
                json!({ "error": "指定したプリンタが見つかりません。" }),
            );
        }
    };

    let pdf_for_print = pdf_path.clone();
    let printer_for_print = canonical.clone();
    let result =
        tokio::task::spawn_blocking(move || run_print_command(&pdf_for_print, &printer_for_print))
            .await;

    match result {
        Ok(Ok(())) => json_response(
            StatusCode::OK,
            json!({ "status": "ok", "printer": canonical }),
        ),
        Ok(Err(err)) => json_response(
            StatusCode::INTERNAL_SERVER_ERROR,
            json!({ "error": format!("印刷に失敗しました: {err}") }),
        ),
        Err(_) => json_response(
            StatusCode::INTERNAL_SERVER_ERROR,
            json!({ "error": "印刷の送信に失敗しました。" }),
        ),
    }
}

async fn job_status(State(state): State<AppState>, AxumPath(job_id): AxumPath<String>) -> Response {
    let (status, payload) = {
        let jobs = state.jobs.read().await;
        match jobs.get(&job_id) {
            Some(job) => (StatusCode::OK, job.to_dict()),
            None => (
                StatusCode::NOT_FOUND,
                json!({ "error": "ジョブが見つかりません。" }),
            ),
        }
    };

    let mut response = json_response(status, payload);
    let headers = response.headers_mut();
    headers.insert(
        header::CACHE_CONTROL,
        header::HeaderValue::from_static("no-store, no-cache, must-revalidate, max-age=0"),
    );
    headers.insert(header::PRAGMA, header::HeaderValue::from_static("no-cache"));
    headers.insert(header::EXPIRES, header::HeaderValue::from_static("0"));
    headers.insert(
        header::HeaderName::from_static("x-accel-buffering"),
        header::HeaderValue::from_static("no"),
    );
    headers.insert(
        header::CONNECTION,
        header::HeaderValue::from_static("close"),
    );
    response
}

async fn process_job(state: AppState, job_id: String) {
    let permit = match state.job_semaphore.acquire().await {
        Ok(permit) => permit,
        Err(err) => {
            error!("failed to acquire queue permit: {err}");
            with_job_mut_blocking(&state, &job_id, |job| {
                job.status = "failed".to_string();
                job.message = "処理キューの取得に失敗しました。".to_string();
                job.processing_completed_at = Some(Utc::now());
                job.show_conversion_progress = false;
            });
            return;
        }
    };

    let state_for_block = state.clone();
    let job_id_for_block = job_id.clone();
    let result = tokio::task::spawn_blocking(move || {
        run_process_job_sync(&state_for_block, &job_id_for_block)
    })
    .await;
    match result {
        Ok(Ok(())) => {}
        Ok(Err(err)) => {
            error!("job failed: {err:#}");
            with_job_mut_blocking(&state, &job_id, |job| {
                job.status = "failed".to_string();
                job.message = format!("エラーが発生しました: {err}");
                job.processing_completed_at = Some(Utc::now());
                job.show_conversion_progress = false;
            });
        }
        Err(err) => {
            error!("job task join failed: {err}");
            with_job_mut_blocking(&state, &job_id, |job| {
                job.status = "failed".to_string();
                job.message = "処理タスクが異常終了しました。".to_string();
                job.processing_completed_at = Some(Utc::now());
                job.show_conversion_progress = false;
            });
        }
    }
    drop(permit);
}

fn run_process_job_sync(state: &AppState, job_id: &str) -> Result<()> {
    let job = {
        let jobs = state.jobs.blocking_read();
        jobs.get(job_id).cloned()
    };
    let Some(job) = job else {
        return Ok(());
    };

    let result = (|| -> Result<()> {
        update_job_blocking(
            state,
            job_id,
            Some("running"),
            Some("ZIPファイルを展開しています…"),
            None,
            None,
            None,
            None,
            None,
        );

        let work_root = state.paths.work_dir.join(job_id);
        let extract_dir = work_root.join("extracted");
        let pdf_dir = work_root.join("pdf");
        if work_root.exists() {
            fs::remove_dir_all(&work_root)?;
        }
        fs::create_dir_all(&extract_dir)?;
        fs::create_dir_all(&pdf_dir)?;

        extract_zip_smart(&job.zip_path, &extract_dir)?;

        let mut entries = job.entries.clone();
        apply_team_prefixes(&extract_dir, &mut entries);
        let ordered_entries = job
            .order
            .iter()
            .filter_map(|name| entries.get(name).cloned())
            .collect::<Vec<_>>();
        if ordered_entries.is_empty() {
            return Err(anyhow!("処理対象のドキュメントが見つかりませんでした。"));
        }

        initialize_conversion_progress(state, job_id, &ordered_entries);
        update_job_blocking(
            state,
            job_id,
            None,
            Some("PDF変換の準備をしています…"),
            None,
            None,
            Some(ordered_entries.len()),
            None,
            None,
        );

        let report_number = determine_report_number(&job.zip_original_name, &ordered_entries);
        with_job_mut_blocking(state, job_id, |job_mut| {
            job_mut.report_number = Some(report_number.clone());
        });

        update_job_blocking(
            state,
            job_id,
            None,
            Some("WordファイルをPDFに変換中・・・"),
            None,
            None,
            None,
            None,
            None,
        );
        let pdf_paths =
            convert_entries_parallel(state, job_id, &ordered_entries, &extract_dir, &pdf_dir)?;

        let merged_path = work_root.join(format!("第{report_number}回報告書.pdf"));
        update_job_blocking(
            state,
            job_id,
            None,
            Some("PDFファイルを結合しています…"),
            None,
            None,
            None,
            None,
            None,
        );
        merge_pdfs(&pdf_paths, &merged_path)?;

        let recipient_email = job.email.trim().to_string();
        let auto_send_enabled =
            email_transport_is_configured(&state.paths) && !recipient_email.is_empty();

        update_job_blocking(
            state,
            job_id,
            Some("completed"),
            Some("PDFの結合が完了しました。"),
            Some(merged_path.clone()),
            None,
            None,
            Some(if auto_send_enabled {
                "sending".to_string()
            } else {
                String::new()
            }),
            Some(false),
        );

        if auto_send_enabled {
            let state_clone = state.clone();
            let job_id_owned = job_id.to_string();
            let subject = format!("第{report_number}回報告書");
            let paths_for_send = state.paths.clone();
            thread::spawn(move || {
                let send_result = send_email_with_attachment(
                    &paths_for_send,
                    &[recipient_email],
                    &[],
                    &subject,
                    "",
                    &merged_path,
                );
                if send_result.is_ok() {
                    update_job_blocking(
                        &state_clone,
                        &job_id_owned,
                        None,
                        None,
                        None,
                        None,
                        None,
                        Some("sent".to_string()),
                        None,
                    );
                } else {
                    update_job_blocking(
                        &state_clone,
                        &job_id_owned,
                        None,
                        None,
                        None,
                        None,
                        None,
                        Some("failed".to_string()),
                        None,
                    );
                }
                schedule_delayed_cleanup(state_clone, 600);
            });
        } else {
            schedule_delayed_cleanup(state.clone(), 600);
        }
        Ok(())
    })();

    if let Err(err) = fs::remove_file(&job.zip_path) {
        if err.kind() != io::ErrorKind::NotFound {
            error!("failed to remove upload {}: {err}", job.zip_path.display());
        }
    }

    if let Err(err) = &result {
        update_job_blocking(
            state,
            job_id,
            Some("failed"),
            Some(&format!("エラーが発生しました: {err}")),
            None,
            None,
            None,
            None,
            Some(false),
        );
    }
    result
}

fn convert_entries_parallel(
    state: &AppState,
    job_id: &str,
    ordered_entries: &[ZipEntry],
    extract_dir: &Path,
    pdf_dir: &Path,
) -> Result<Vec<PathBuf>> {
    let mut handles = Vec::new();

    for (index, entry) in ordered_entries.iter().cloned().enumerate() {
        let state_clone = state.clone();
        let job_id_owned = job_id.to_string();
        let extract_dir = extract_dir.to_path_buf();
        let pdf_dir = pdf_dir.to_path_buf();

        handles.push(thread::spawn(move || -> Result<(usize, PathBuf)> {
            update_conversion_status(&state_clone, &job_id_owned, &entry.display_name, "running");
            let source_path = extract_dir.join(archive_name_to_path(&entry.archive_name));
            if !source_path.exists() {
                update_conversion_status(
                    &state_clone,
                    &job_id_owned,
                    &entry.display_name,
                    "failed",
                );
                return Err(anyhow!(
                    "展開後のファイルが見つかりません: {}",
                    entry.archive_name
                ));
            }

            let base = entry
                .sanitized_name
                .clone()
                .unwrap_or_else(|| entry.display_name.clone());
            let stem = Path::new(&base)
                .file_stem()
                .and_then(|v| v.to_str())
                .map(sanitize_file_stem)
                .unwrap_or_else(|| format!("document_{}", index + 1));
            let pdf_path = pdf_dir.join(format!("{:03}_{}.pdf", index + 1, stem));
            convert_word_to_pdf(&source_path, &pdf_path)?;

            update_conversion_status(
                &state_clone,
                &job_id_owned,
                &entry.display_name,
                "completed",
            );
            update_job_blocking(
                &state_clone,
                &job_id_owned,
                None,
                None,
                None,
                Some(1),
                None,
                None,
                None,
            );
            Ok((index, pdf_path))
        }));
    }

    let mut pdf_paths = vec![PathBuf::new(); ordered_entries.len()];
    let mut errors = Vec::new();
    for handle in handles {
        match handle.join() {
            Ok(Ok((index, path))) => pdf_paths[index] = path,
            Ok(Err(err)) => errors.push(err.to_string()),
            Err(_) => errors.push("PDF変換タスクが異常終了しました。".to_string()),
        }
    }

    if !errors.is_empty() {
        return Err(anyhow!(errors.join(" / ")));
    }
    if pdf_paths.iter().any(|p| p.as_os_str().is_empty()) {
        return Err(anyhow!("PDF変換結果の整合性に問題があります。"));
    }
    Ok(pdf_paths)
}

fn with_job_mut_blocking<F>(state: &AppState, job_id: &str, update: F)
where
    F: FnOnce(&mut JobState),
{
    let mut jobs = state.jobs.blocking_write();
    if let Some(job) = jobs.get_mut(job_id) {
        update(job);
        job.updated_at = Utc::now();
    }
}

fn update_conversion_status(state: &AppState, job_id: &str, display_name: &str, status: &str) {
    with_job_mut_blocking(state, job_id, |job| {
        if !job.conversion_order.iter().any(|name| name == display_name) {
            job.conversion_order.push(display_name.to_string());
        }
        job.conversion_statuses
            .insert(display_name.to_string(), status.to_string());
        if !job.conversion_threads.contains_key(display_name) {
            let next = job.conversion_threads.len() + 1;
            job.conversion_threads
                .insert(display_name.to_string(), next);
        }
    });
}

fn initialize_conversion_progress(state: &AppState, job_id: &str, entries: &[ZipEntry]) {
    with_job_mut_blocking(state, job_id, |job| {
        job.conversion_order = entries
            .iter()
            .map(|entry| entry.display_name.clone())
            .collect();
        job.conversion_statuses = entries
            .iter()
            .map(|entry| (entry.display_name.clone(), "queued".to_string()))
            .collect();
        job.conversion_threads = entries
            .iter()
            .enumerate()
            .map(|(index, entry)| (entry.display_name.clone(), index + 1))
            .collect();
        job.show_conversion_progress = true;
    });
}

#[allow(clippy::too_many_arguments)]
fn update_job_blocking(
    state: &AppState,
    job_id: &str,
    status: Option<&str>,
    message: Option<&str>,
    merged_pdf: Option<PathBuf>,
    progress_increment: Option<usize>,
    progress_total: Option<usize>,
    email_delivery_status: Option<String>,
    show_conversion_progress: Option<bool>,
) {
    with_job_mut_blocking(state, job_id, |job| {
        let now = Utc::now();
        if let Some(status) = status {
            job.status = status.to_string();
            if status == "running" && job.processing_started_at.is_none() {
                job.processing_started_at = Some(now);
            }
            if status == "completed" || status == "failed" {
                job.processing_completed_at = Some(now);
            }
        }
        if let Some(message) = message {
            job.message = message.to_string();
        }
        if let Some(path) = merged_pdf {
            job.merged_pdf = Some(path);
        }
        if let Some(total) = progress_total {
            job.progress_total = total;
            if total == 0 {
                job.progress_current = 0;
            } else if job.progress_current > total {
                job.progress_current = total;
            }
        }
        if let Some(increment) = progress_increment {
            job.progress_current = job.progress_current.saturating_add(increment);
            if job.progress_total > 0 {
                job.progress_current = job.progress_current.min(job.progress_total);
            }
        }
        if let Some(status) = email_delivery_status {
            job.email_delivery_status = status;
        }
        if let Some(show) = show_conversion_progress {
            job.show_conversion_progress = show;
        }
        if status == Some("completed") && job.progress_total > 0 {
            job.progress_current = job.progress_total;
            job.show_conversion_progress = false;
        }
        if status == Some("failed") {
            job.show_conversion_progress = false;
        }
    });
}

fn schedule_delayed_cleanup(state: AppState, delay_seconds: u64) {
    tokio::spawn(async move {
        tokio::time::sleep(Duration::from_secs(delay_seconds)).await;
        let active_jobs = {
            let jobs = state.jobs.read().await;
            jobs.values()
                .any(|job| job.status != "completed" && job.status != "failed")
        };
        if active_jobs {
            schedule_delayed_cleanup(state.clone(), delay_seconds);
            return;
        }
        if let Err(err) = cleanup_data_directories(&state.paths) {
            error!("cleanup failed: {err:#}");
        }
    });
}

impl JobState {
    fn to_dict(&self) -> Value {
        let elapsed_seconds = match (self.processing_started_at, self.processing_completed_at) {
            (Some(start), Some(end)) => Some((end - start).num_milliseconds() as f64 / 1000.0),
            _ => None,
        };
        let elapsed_display = elapsed_seconds.map(format_elapsed);
        let progress_percent = if self.progress_total > 0 {
            let ratio = self.progress_current as f64 / self.progress_total as f64;
            Some((ratio.clamp(0.0, 1.0) * 100.0).round() as i64)
        } else {
            None
        };

        json!({
            "id": self.id,
            "email": self.email,
            "status": self.status,
            "message": self.message,
            "created_at": self.created_at.to_rfc3339(),
            "updated_at": self.updated_at.to_rfc3339(),
            "team_options": self.team_options,
            "report_number": self.report_number,
            "final_pdf_name": self.merged_pdf.as_ref().and_then(|p| p.file_name()).and_then(|v| v.to_str()).map(str::to_string),
            "elapsed_seconds": elapsed_seconds,
            "elapsed_display": elapsed_display,
            "progress_current": self.progress_current,
            "progress_total": self.progress_total,
            "progress_percent": progress_percent,
            "email_delivery_status": self.email_delivery_status,
            "manual_email_delivery_status": self.manual_email_delivery_status,
            "manual_email_sent_once": self.manual_email_sent_once,
            "conversion_progress": self.conversion_order.iter().map(|name| {
                json!({
                    "display_name": name,
                    "status": self.conversion_statuses.get(name).cloned().unwrap_or_else(|| "queued".to_string()),
                    "thread": self.conversion_threads.get(name).copied(),
                })
            }).collect::<Vec<_>>(),
            "show_conversion_progress": self.show_conversion_progress,
        })
    }
}

impl OrderManager {
    fn new(storage_file: PathBuf) -> Self {
        Self {
            storage_file,
            lock: Mutex::new(()),
        }
    }

    fn load_preferences(&self) -> OrderPreferences {
        if !self.storage_file.exists() {
            return OrderPreferences::empty();
        }
        let raw = match fs::read_to_string(&self.storage_file) {
            Ok(v) => v,
            Err(_) => return OrderPreferences::empty(),
        };
        let parsed: Value = match serde_json::from_str(&raw) {
            Ok(v) => v,
            Err(_) => return OrderPreferences::empty(),
        };
        if parsed.is_object() {
            return order_preferences_from_dict(&parsed);
        }
        if let Some(items) = parsed.as_array() {
            let list = items
                .iter()
                .filter_map(Value::as_str)
                .map(str::to_string)
                .collect::<Vec<_>>();
            return order_preferences_from_legacy_list(&list);
        }
        OrderPreferences::empty()
    }

    fn save_member_sequence(&self, team_key: &str, members: &[MemberRecord]) -> Result<()> {
        let _guard = self.lock.lock().expect("order lock poisoned");
        let normalized_team = normalize_team_key(team_key);
        let mut preferences = self.load_preferences();
        let existing = preferences
            .member_sequences
            .get(&normalized_team)
            .cloned()
            .unwrap_or_default();
        let cleaned = normalize_member_objects_from_members(members, &existing);

        if cleaned.is_empty() {
            preferences.member_sequences.remove(&normalized_team);
            preferences
                .team_sequence
                .retain(|key| key != &normalized_team);
        } else {
            if !preferences.team_sequence.contains(&normalized_team) {
                preferences.team_sequence.push(normalized_team.clone());
            }
            preferences
                .member_sequences
                .insert(normalized_team, cleaned);
        }
        self.write_preferences(&preferences)
    }

    fn delete_member_sequence(&self, team_key: &str) -> Result<()> {
        let _guard = self.lock.lock().expect("order lock poisoned");
        let normalized_team = normalize_team_key(team_key);
        let mut preferences = self.load_preferences();
        preferences.member_sequences.remove(&normalized_team);
        preferences
            .team_sequence
            .retain(|key| key != &normalized_team);
        self.write_preferences(&preferences)
    }

    fn write_preferences(&self, preferences: &OrderPreferences) -> Result<()> {
        if let Some(parent) = self.storage_file.parent() {
            fs::create_dir_all(parent)?;
        }
        let member_sequences = preferences
            .member_sequences
            .iter()
            .map(|(k, v)| (k.clone(), json!(v)))
            .collect::<serde_json::Map<String, Value>>();
        let payload = json!({
            "team_sequence": preferences.team_sequence,
            "member_sequences": member_sequences,
        });
        let text = serde_json::to_string_pretty(&payload)?;
        fs::write(&self.storage_file, text)?;
        Ok(())
    }

    fn initial_layout(
        &self,
        entries: Vec<ZipEntry>,
    ) -> (Vec<String>, HashMap<String, Vec<ZipEntry>>) {
        let preferences = self.load_preferences();
        let mut team_map: HashMap<String, Vec<ZipEntry>> = HashMap::new();
        let mut appearance = Vec::new();

        for entry in entries {
            let team_key =
                normalize_team_key(entry.team_name.as_deref().unwrap_or(UNGROUPED_TEAM_KEY));
            team_map.entry(team_key.clone()).or_default().push(entry);
            if !appearance.contains(&team_key) {
                appearance.push(team_key);
            }
        }

        let mut sequence = Vec::new();
        for key in &preferences.team_sequence {
            if team_map.contains_key(key) && !sequence.contains(key) {
                sequence.push(key.clone());
            }
        }
        for key in appearance {
            if !sequence.contains(&key) {
                sequence.push(key);
            }
        }

        let mut ordered = HashMap::new();
        for key in &sequence {
            let members = preferences
                .member_sequences
                .get(key)
                .cloned()
                .unwrap_or_default();
            let member_order = extract_member_names(&members);
            let items = team_map.get(key).cloned().unwrap_or_default();
            ordered.insert(key.clone(), sort_team_entries(items, &member_order));
        }
        (sequence, ordered)
    }
}

fn value_to_trimmed_string(value: &Value) -> String {
    match value {
        Value::String(s) => s.trim().to_string(),
        _ => value.to_string().trim_matches('"').trim().to_string(),
    }
}

fn parse_bool_value(value: Option<&Value>) -> bool {
    match value {
        Some(Value::Bool(v)) => *v,
        Some(Value::Number(n)) => n.as_i64().unwrap_or(0) != 0,
        Some(Value::String(s)) => {
            let normalized = s.trim().to_ascii_lowercase();
            matches!(normalized.as_str(), "1" | "true" | "yes" | "on")
        }
        _ => false,
    }
}

fn parse_recipients(raw: Option<&Value>) -> (Vec<String>, Vec<String>) {
    let mut to = Vec::new();
    let mut cc = Vec::new();

    if let Some(Value::Object(map)) = raw {
        if let Some(to_raw) = map.get("to").and_then(Value::as_array) {
            for item in to_raw {
                if let Some(email) = item.as_str() {
                    to.push(email.to_string());
                }
            }
        }
        if let Some(cc_raw) = map.get("cc").and_then(Value::as_array) {
            for item in cc_raw {
                if let Some(email) = item.as_str() {
                    cc.push(email.to_string());
                }
            }
        }
    } else if let Some(Value::Array(items)) = raw {
        for item in items {
            if let Some(obj) = item.as_object() {
                let email = obj
                    .get("email")
                    .and_then(Value::as_str)
                    .unwrap_or_default()
                    .trim()
                    .to_string();
                if email.is_empty() {
                    continue;
                }
                let preferred = obj
                    .get("preferred")
                    .and_then(Value::as_str)
                    .unwrap_or_default()
                    .trim()
                    .to_ascii_lowercase();
                if preferred == "to" {
                    to.push(email);
                } else {
                    cc.push(email);
                }
            } else if let Some(email) = item.as_str() {
                cc.push(email.to_string());
            }
        }
    }
    (clean_email_list(&to), clean_email_list(&cc))
}

fn clean_email_list(items: &[String]) -> Vec<String> {
    let mut out = Vec::new();
    let mut seen = HashSet::new();
    for item in items {
        let email = item.trim();
        if email.is_empty() {
            continue;
        }
        let key = email.to_ascii_lowercase();
        if seen.contains(&key) {
            continue;
        }
        seen.insert(key);
        out.push(email.to_string());
    }
    out
}

fn normalize_team_key(value: &str) -> String {
    let candidate = value.trim();
    if candidate.is_empty() || candidate == UNGROUPED_TEAM_KEY || candidate == "班なし" {
        UNGROUPED_TEAM_KEY.to_string()
    } else {
        candidate.to_string()
    }
}

fn team_display_label(team_key: &str) -> String {
    if team_key == UNGROUPED_TEAM_KEY {
        "班なし".to_string()
    } else {
        team_key.to_string()
    }
}

fn team_labels_from_preferences(preferences: &OrderPreferences) -> Vec<String> {
    let mut labels = Vec::new();
    let mut seen = HashSet::new();
    for key in &preferences.team_sequence {
        let label = team_display_label(key);
        if seen.insert(label.clone()) {
            labels.push(label);
        }
    }
    for key in preferences.member_sequences.keys() {
        let label = team_display_label(key);
        if seen.insert(label.clone()) {
            labels.push(label);
        }
    }
    labels
}

fn collect_preference_teams(preferences: &OrderPreferences) -> Vec<Value> {
    let mut teams = Vec::new();
    let mut seen = HashSet::new();

    let mut append_team = |team_key: &str| {
        if !seen.insert(team_key.to_string()) {
            return;
        }
        let members = preferences
            .member_sequences
            .get(team_key)
            .cloned()
            .unwrap_or_default();
        teams.push(json!({
            "key": team_key,
            "label": team_display_label(team_key),
            "members": normalize_member_objects_from_members(&members, &members),
        }));
    };

    for key in &preferences.team_sequence {
        append_team(key);
    }
    for key in preferences.member_sequences.keys() {
        append_team(key);
    }
    teams
}

fn load_faculty_contacts(file_path: &Path) -> Vec<Value> {
    if !file_path.exists() {
        return Vec::new();
    }
    let raw = match fs::read_to_string(file_path) {
        Ok(v) => v,
        Err(_) => return Vec::new(),
    };
    let parsed: Value = match serde_json::from_str(&raw) {
        Ok(v) => v,
        Err(_) => return Vec::new(),
    };
    let list = if let Some(arr) = parsed.get("members").and_then(Value::as_array) {
        arr.clone()
    } else if let Some(arr) = parsed.as_array() {
        arr.clone()
    } else {
        Vec::new()
    };

    let mut out = Vec::new();
    let mut seen = HashSet::new();
    for item in list {
        let Some(obj) = item.as_object() else {
            continue;
        };
        let email = obj
            .get("email")
            .and_then(Value::as_str)
            .unwrap_or_default()
            .trim()
            .to_string();
        if email.is_empty() {
            continue;
        }
        let key = email.to_ascii_lowercase();
        if seen.contains(&key) {
            continue;
        }
        seen.insert(key);
        let name = obj
            .get("name")
            .and_then(Value::as_str)
            .unwrap_or(&email)
            .trim()
            .to_string();
        let preferred = obj
            .get("preferred")
            .and_then(Value::as_str)
            .unwrap_or_default()
            .trim()
            .to_ascii_lowercase();
        let mut contact = json!({
            "name": if name.is_empty() { email.clone() } else { name },
            "email": email,
        });
        if (preferred == "to" || preferred == "cc") && contact.is_object() {
            if let Some(map) = contact.as_object_mut() {
                map.insert("preferred".to_string(), Value::String(preferred));
            }
        }
        out.push(contact);
    }
    out
}

fn collect_email_recipient_groups(
    preferences: &OrderPreferences,
    faculty_file: &Path,
) -> Vec<Value> {
    let mut team_order = Vec::new();
    for key in &preferences.team_sequence {
        if !team_order.contains(key) {
            team_order.push(key.clone());
        }
    }
    for key in preferences.member_sequences.keys() {
        if !team_order.contains(key) {
            team_order.push(key.clone());
        }
    }

    let mut groups = Vec::new();
    for key in team_order {
        let members = preferences
            .member_sequences
            .get(&key)
            .cloned()
            .unwrap_or_default();
        let mut member_values = Vec::new();
        let mut seen = HashSet::new();
        for member in members {
            let email = member.email.trim().to_string();
            if email.is_empty() {
                continue;
            }
            let normalized = email.to_ascii_lowercase();
            if seen.contains(&normalized) {
                continue;
            }
            seen.insert(normalized);
            member_values.push(json!({
                "name": if member.name.trim().is_empty() { email.clone() } else { member.name },
                "email": email,
            }));
        }
        if member_values.is_empty() {
            continue;
        }
        groups.push(json!({
            "key": key,
            "label": team_display_label(&key),
            "members": member_values,
        }));
    }

    let faculty_contacts = load_faculty_contacts(faculty_file);
    if !faculty_contacts.is_empty() {
        groups.insert(
            0,
            json!({
                "key": "__faculty__",
                "label": "教員・秘書・Dr",
                "members": faculty_contacts,
            }),
        );
    }
    groups
}

fn order_preferences_from_dict(payload: &Value) -> OrderPreferences {
    let Some(map) = payload.as_object() else {
        return OrderPreferences::empty();
    };

    let mut team_sequence = Vec::new();
    if let Some(items) = map.get("team_sequence").and_then(Value::as_array) {
        for item in items {
            if let Some(team) = item.as_str() {
                let normalized = normalize_team_key(team);
                if !team_sequence.contains(&normalized) {
                    team_sequence.push(normalized);
                }
            }
        }
    }

    let mut member_sequences = HashMap::new();
    if let Some(raw_members) = map.get("member_sequences").and_then(Value::as_object) {
        for (raw_key, raw_value) in raw_members {
            let Some(raw_list) = raw_value.as_array() else {
                continue;
            };
            let team_key = normalize_team_key(raw_key);
            let members = normalize_member_objects_from_values(raw_list, &[]);
            if !members.is_empty() {
                member_sequences.insert(team_key, members);
            }
        }
    }

    if !member_sequences.contains_key(UNGROUPED_TEAM_KEY) {
        team_sequence.retain(|team| team != UNGROUPED_TEAM_KEY);
    }
    OrderPreferences {
        team_sequence,
        member_sequences,
    }
}

fn order_preferences_from_legacy_list(items: &[String]) -> OrderPreferences {
    let mut team_sequence = Vec::new();
    let mut member_sequences: HashMap<String, Vec<MemberRecord>> = HashMap::new();
    for name in items {
        let mut team_key = UNGROUPED_TEAM_KEY.to_string();
        let mut stripped = name.clone();
        if name.starts_with('[') && name.contains(']') {
            if let Some((prefix, remainder)) = name.split_once(']') {
                let inner = prefix.trim_start_matches('[').trim();
                if !inner.is_empty() {
                    team_key = inner.to_string();
                }
                stripped = remainder.trim().to_string();
            }
        }
        team_key = normalize_team_key(&team_key);
        if !team_sequence.contains(&team_key) {
            team_sequence.push(team_key.clone());
        }
        if !stripped.trim().is_empty() {
            member_sequences
                .entry(team_key)
                .or_default()
                .push(MemberRecord {
                    name: stripped.trim().to_string(),
                    email: String::new(),
                });
        }
    }
    for members in member_sequences.values_mut() {
        *members = normalize_member_objects_from_members(members, &[]);
    }
    OrderPreferences {
        team_sequence,
        member_sequences,
    }
}

fn extract_member_names(members: &[MemberRecord]) -> Vec<String> {
    let mut names = Vec::new();
    for member in members {
        let name = member.name.trim();
        if !name.is_empty() && !names.contains(&name.to_string()) {
            names.push(name.to_string());
        }
    }
    names
}

fn normalize_member_objects_from_members(
    members: &[MemberRecord],
    existing: &[MemberRecord],
) -> Vec<MemberRecord> {
    let values = members
        .iter()
        .map(|member| json!({ "name": member.name, "email": member.email }))
        .collect::<Vec<_>>();
    normalize_member_objects_from_values(&values, existing)
}

fn normalize_member_objects_from_values(
    members: &[Value],
    existing: &[MemberRecord],
) -> Vec<MemberRecord> {
    let lookup = existing
        .iter()
        .filter_map(|member| {
            let key = member_name_key(&member.name);
            if key.is_empty() {
                None
            } else {
                Some((key, member.clone()))
            }
        })
        .collect::<HashMap<_, _>>();

    let mut out = Vec::new();
    let mut seen = HashSet::new();
    for value in members {
        let (raw_name, raw_email) = match value {
            Value::String(name) => (name.trim().to_string(), String::new()),
            Value::Object(map) => (
                map.get("name")
                    .and_then(Value::as_str)
                    .unwrap_or_default()
                    .trim()
                    .to_string(),
                map.get("email")
                    .and_then(Value::as_str)
                    .unwrap_or_default()
                    .trim()
                    .to_string(),
            ),
            _ => (String::new(), String::new()),
        };
        if raw_name.is_empty() {
            continue;
        }
        let key = member_name_key(&raw_name);
        if key.is_empty() || seen.contains(&key) {
            continue;
        }
        seen.insert(key.clone());
        let base = lookup.get(&key).cloned().unwrap_or(MemberRecord {
            name: raw_name.clone(),
            email: String::new(),
        });
        out.push(MemberRecord {
            name: if base.name.trim().is_empty() {
                raw_name
            } else {
                base.name
            },
            email: if raw_email.is_empty() {
                base.email
            } else {
                raw_email
            },
        });
    }
    out
}

fn member_name_key(name: &str) -> String {
    normalize_person_token(name)
}

fn sort_team_entries(mut items: Vec<ZipEntry>, member_order: &[String]) -> Vec<ZipEntry> {
    if items.is_empty() {
        return items;
    }
    let mut by_name = items.clone();
    by_name.sort_by(|a, b| {
        a.display_name
            .to_lowercase()
            .cmp(&b.display_name.to_lowercase())
    });
    let fallback = by_name
        .iter()
        .enumerate()
        .map(|(idx, entry)| (entry.identifier.clone(), idx))
        .collect::<HashMap<_, _>>();

    items.sort_by(|a, b| {
        let a_idx = find_member_order_index(member_order, &a.persons);
        let b_idx = find_member_order_index(member_order, &b.persons);
        let a_key = if let Some(idx) = a_idx {
            (0usize, idx)
        } else {
            (
                1usize,
                fallback.get(&a.identifier).copied().unwrap_or(usize::MAX),
            )
        };
        let b_key = if let Some(idx) = b_idx {
            (0usize, idx)
        } else {
            (
                1usize,
                fallback.get(&b.identifier).copied().unwrap_or(usize::MAX),
            )
        };
        a_key.cmp(&b_key)
    });
    items
}

fn find_member_order_index(member_order: &[String], persons: &[String]) -> Option<usize> {
    if member_order.is_empty() || persons.is_empty() {
        return None;
    }
    let normalized_order = member_order
        .iter()
        .enumerate()
        .filter_map(|(idx, name)| {
            let normalized = normalize_person_token(name);
            if normalized.is_empty() {
                None
            } else {
                Some((normalized, idx))
            }
        })
        .collect::<Vec<_>>();
    if normalized_order.is_empty() {
        return None;
    }

    let mut best: Option<usize> = None;
    for person in persons {
        let normalized_person = normalize_person_token(person);
        if normalized_person.is_empty() {
            continue;
        }
        for (normalized_name, order_idx) in &normalized_order {
            if normalized_person == *normalized_name
                || normalized_person.contains(normalized_name)
                || normalized_name.contains(&normalized_person)
            {
                if best.is_none() || Some(*order_idx) < best {
                    best = Some(*order_idx);
                }
                if best == Some(0) {
                    return best;
                }
                break;
            }
        }
    }
    best
}

fn normalize_person_token(value: &str) -> String {
    static RE: OnceLock<Regex> = OnceLock::new();
    let re = RE.get_or_init(|| Regex::new(r"[\s・･\.,，、．｡]+").expect("invalid regex"));
    re.replace_all(value, "").to_lowercase().trim().to_string()
}

fn cleanup_data_directories(paths: &RuntimePaths) -> Result<()> {
    for directory in [&paths.uploads_dir, &paths.work_dir] {
        if !directory.exists() {
            continue;
        }
        for entry in fs::read_dir(directory)? {
            let path = entry?.path();
            if path.is_dir() {
                fs::remove_dir_all(&path)?;
            } else if let Err(err) = fs::remove_file(&path) {
                if err.kind() != io::ErrorKind::NotFound {
                    return Err(err.into());
                }
            }
        }
    }
    Ok(())
}

fn determine_zip_name_strategy(zip_path: &Path) -> Result<bool> {
    let file = File::open(zip_path)?;
    let mut archive = ZipArchive::new(file)?;
    let mut has_utf8 = false;
    for idx in 0..archive.len() {
        let file = archive.by_index(idx)?;
        if file.get_metadata().is_utf8 {
            has_utf8 = true;
            break;
        }
    }
    Ok(has_utf8)
}

fn decode_name_cp932(bytes: &[u8]) -> String {
    let (decoded, _, _) = encoding_rs::SHIFT_JIS.decode(bytes);
    decoded.into_owned()
}

fn sanitize_zip_path(name: &str) -> Option<PathBuf> {
    let normalized = name.replace('\\', "/");
    let mut out = PathBuf::new();
    for part in normalized.split('/') {
        let part = part.trim();
        if part.is_empty() || part == "." {
            continue;
        }
        if part == ".." || part.contains(':') {
            return None;
        }
        out.push(part);
    }
    if out.as_os_str().is_empty() {
        None
    } else {
        Some(out)
    }
}

fn extract_zip_smart(zip_path: &Path, extract_to: &Path) -> Result<()> {
    let has_utf8 = determine_zip_name_strategy(zip_path).unwrap_or(true);
    let file = File::open(zip_path)?;
    let mut archive = ZipArchive::new(file)?;

    for idx in 0..archive.len() {
        let mut entry = archive.by_index(idx)?;
        let decoded_name = if has_utf8 {
            entry.name().to_string()
        } else {
            decode_name_cp932(entry.name_raw())
        };
        let Some(rel_path) = sanitize_zip_path(&decoded_name) else {
            continue;
        };
        let out_path = extract_to.join(rel_path);

        if entry.is_dir() {
            fs::create_dir_all(&out_path)?;
            continue;
        }
        if let Some(parent) = out_path.parent() {
            fs::create_dir_all(parent)?;
        }
        let mut out_file = File::create(&out_path)?;
        io::copy(&mut entry, &mut out_file)?;
    }
    Ok(())
}

fn extract_entries(zip_path: &Path, original_name: Option<&str>) -> Result<Vec<ZipEntry>> {
    let has_utf8 = determine_zip_name_strategy(zip_path).unwrap_or(true);
    let file = File::open(zip_path)?;
    let mut archive = ZipArchive::new(file)?;

    let mut paths = Vec::<PathBuf>::new();
    for idx in 0..archive.len() {
        let entry = archive.by_index(idx)?;
        if entry.is_dir() {
            continue;
        }
        let decoded_name = if has_utf8 {
            entry.name().to_string()
        } else {
            decode_name_cp932(entry.name_raw())
        };
        let Some(rel_path) = sanitize_zip_path(&decoded_name) else {
            continue;
        };
        let ext = rel_path
            .extension()
            .and_then(|v| v.to_str())
            .map(|v| v.to_ascii_lowercase())
            .unwrap_or_default();
        if ext == "doc" || ext == "docx" {
            paths.push(rel_path);
        }
    }

    if paths.is_empty() {
        return Ok(Vec::new());
    }
    let team_level = infer_team_directory_level(&paths);
    let default_team_name = if team_level.is_none() {
        Some(
            original_name
                .and_then(|v| Path::new(v).file_stem())
                .and_then(|v| v.to_str())
                .unwrap_or_else(|| {
                    zip_path
                        .file_stem()
                        .and_then(|v| v.to_str())
                        .unwrap_or_default()
                })
                .to_string(),
        )
    } else {
        None
    };

    let mut duplicate_counter = HashMap::new();
    let mut entries = Vec::new();
    for rel_path in &paths {
        let directories = rel_path
            .parent()
            .map(|parent| {
                parent
                    .components()
                    .filter_map(|c| c.as_os_str().to_str().map(str::to_string))
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default();
        let team_name = if let Some(level) = team_level {
            if directories.len() > level {
                Some(directories[level].clone())
            } else {
                None
            }
        } else {
            default_team_name.clone()
        };
        let file_name = rel_path
            .file_name()
            .and_then(|v| v.to_str())
            .unwrap_or_default();
        let sanitized_name = sanitize_report_filename(file_name);
        let display_name = build_display_name(
            &sanitized_name,
            team_name.as_deref(),
            &mut duplicate_counter,
        );
        entries.push(ZipEntry {
            identifier: Uuid::new_v4().to_string(),
            display_name,
            archive_name: path_to_posix(rel_path),
            team_name,
            persons: extract_person_names(&sanitized_name),
            sanitized_name: Some(sanitized_name),
        });
    }
    Ok(entries)
}

fn infer_team_directory_level(paths: &[PathBuf]) -> Option<usize> {
    let directories = paths
        .iter()
        .map(|path| {
            path.parent()
                .map(|parent| {
                    parent
                        .components()
                        .filter_map(|component| component.as_os_str().to_str().map(str::to_string))
                        .collect::<Vec<_>>()
                })
                .unwrap_or_default()
        })
        .collect::<Vec<_>>();
    let max_depth = directories.iter().map(Vec::len).max().unwrap_or(0);
    for level in 0..max_depth {
        let names = directories
            .iter()
            .filter_map(|parts| parts.get(level).cloned())
            .collect::<HashSet<_>>();
        if names.len() > 1 {
            return Some(level);
        }
    }
    None
}

fn path_to_posix(path: &Path) -> String {
    path.components()
        .filter_map(|c| c.as_os_str().to_str().map(str::to_string))
        .collect::<Vec<_>>()
        .join("/")
}

fn archive_name_to_path(archive_name: &str) -> PathBuf {
    let mut out = PathBuf::new();
    for part in archive_name.split('/') {
        if !part.trim().is_empty() {
            out.push(part);
        }
    }
    out
}

fn apply_team_prefixes(extract_dir: &Path, entries: &mut HashMap<String, ZipEntry>) {
    for entry in entries.values_mut() {
        if entry.team_name.is_none() {
            continue;
        }
        let source = extract_dir.join(archive_name_to_path(&entry.archive_name));
        if !source.exists() {
            continue;
        }
        let target = source.with_file_name(&entry.display_name);
        if target != source {
            if let Some(parent) = target.parent() {
                let _ = fs::create_dir_all(parent);
            }
            if fs::rename(&source, &target).is_err() {
                continue;
            }
        }
        if let Ok(relative) = target.strip_prefix(extract_dir) {
            entry.archive_name = path_to_posix(relative);
        }
    }
}

fn build_display_name(
    sanitized_name: &str,
    team_name: Option<&str>,
    duplicate_counter: &mut HashMap<String, usize>,
) -> String {
    let base = Path::new(sanitized_name)
        .file_name()
        .and_then(|v| v.to_str())
        .unwrap_or(sanitized_name)
        .to_string();
    let prefixed = if let Some(team_name) = team_name {
        let prefix = format!("[{team_name}] ");
        if base.starts_with(&prefix) {
            base
        } else {
            format!("{prefix}{base}")
        }
    } else {
        base
    };
    let key = prefixed.to_lowercase();
    let count = duplicate_counter.entry(key).or_insert(0);
    *count += 1;
    if *count == 1 {
        prefixed
    } else {
        append_duplicate_suffix(&prefixed, *count)
    }
}

fn append_duplicate_suffix(base_name: &str, counter: usize) -> String {
    let path = Path::new(base_name);
    let stem = path
        .file_stem()
        .and_then(|v| v.to_str())
        .unwrap_or(base_name);
    let ext = path
        .extension()
        .and_then(|v| v.to_str())
        .unwrap_or_default();
    if ext.is_empty() {
        format!("{stem} ({counter})")
    } else {
        format!("{stem} ({counter}).{ext}")
    }
}

fn sanitize_report_filename(original_name: &str) -> String {
    let path = Path::new(original_name);
    let suffix = path
        .extension()
        .and_then(|v| v.to_str())
        .map(|v| format!(".{v}"))
        .unwrap_or_default();
    let mut stem = path
        .file_stem()
        .and_then(|v| v.to_str())
        .unwrap_or_default()
        .to_string();

    static PUNCT_RE: OnceLock<Regex> = OnceLock::new();
    static SPACE_RE: OnceLock<Regex> = OnceLock::new();
    let punct_re = PUNCT_RE.get_or_init(|| Regex::new(r"[･\.,，、．｡]+").expect("invalid regex"));
    let space_re = SPACE_RE.get_or_init(|| Regex::new(r"[₋_＿\s]+").expect("invalid regex"));

    stem = punct_re.replace_all(&stem, "・").to_string();
    stem = space_re.replace_all(&stem, " ").to_string();
    stem = stem.replace("報告会", "報告書");
    stem = stem.split_whitespace().collect::<Vec<_>>().join(" ");
    if stem.contains('回') && !stem.contains("報告書") {
        if let Some(pos) = stem.find('回') {
            let insert_pos = pos + "回".len();
            stem = format!("{}報告書{}", &stem[..insert_pos], &stem[insert_pos..]);
        }
    }
    if !stem.contains("報告書") {
        stem = stem.replace("回 ", "回報告書 ");
    }
    format!("{}{}", stem.trim(), suffix)
}

fn extract_person_names(sanitized_name: &str) -> Vec<String> {
    let stem = Path::new(sanitized_name)
        .file_stem()
        .and_then(|v| v.to_str())
        .unwrap_or_default();
    let remainder = if let Some((_, right)) = stem.split_once("報告書") {
        right.trim()
    } else {
        stem.trim()
    };
    if remainder.is_empty() {
        return Vec::new();
    }
    static RE: OnceLock<Regex> = OnceLock::new();
    let re = RE.get_or_init(|| Regex::new(r"[・･\.,，、．｡\s]+").expect("invalid regex"));
    re.split(remainder)
        .map(str::trim)
        .filter(|v| !v.is_empty())
        .map(str::to_string)
        .collect()
}

fn determine_report_number(zip_original_name: &str, entries: &[ZipEntry]) -> String {
    if let Some(number) = extract_report_number_from_name(Some(zip_original_name)) {
        return number;
    }
    let mut counts: HashMap<String, usize> = HashMap::new();
    for entry in entries {
        let candidate = extract_report_number_from_name(entry.sanitized_name.as_deref())
            .or_else(|| extract_report_number_from_name(Some(&entry.display_name)));
        if let Some(value) = candidate {
            *counts.entry(value).or_insert(0) += 1;
        }
    }
    if !counts.is_empty() {
        let mut items = counts.into_iter().collect::<Vec<_>>();
        items.sort_by(|(a_val, a_count), (b_val, b_count)| {
            let a_num = a_val.parse::<i64>().unwrap_or(-1);
            let b_num = b_val.parse::<i64>().unwrap_or(-1);
            b_count.cmp(a_count).then_with(|| b_num.cmp(&a_num))
        });
        if let Some((best, _)) = items.first() {
            return best.clone();
        }
    }
    "1".to_string()
}

fn extract_report_number_from_name(name: Option<&str>) -> Option<String> {
    let name = name?;
    let stem = Path::new(name)
        .file_stem()
        .and_then(|v| v.to_str())
        .unwrap_or(name);
    static RE: OnceLock<Regex> = OnceLock::new();
    let re = RE.get_or_init(|| Regex::new(r"第\s*(\d{1,3})\s*回").expect("invalid regex"));
    if let Some(captures) = re.captures(stem)
        && let Some(number) = captures.get(1)
    {
        return Some(number.as_str().to_string());
    }
    let digits = stem
        .chars()
        .filter(|ch| ch.is_ascii_digit())
        .collect::<String>();
    if digits.is_empty() {
        None
    } else {
        Some(digits)
    }
}

fn format_elapsed(total_seconds: f64) -> String {
    let secs = total_seconds.round() as i64;
    let hours = secs / 3600;
    let minutes = (secs % 3600) / 60;
    let seconds = secs % 60;
    let mut parts = Vec::new();
    if hours > 0 {
        parts.push(format!("{hours}時間"));
    }
    if minutes > 0 || hours > 0 {
        parts.push(format!("{minutes}分"));
    }
    parts.push(format!("{seconds}秒"));
    parts.join("")
}

fn is_standard_printer(name: &str) -> bool {
    let normalized = name.trim().to_ascii_lowercase();
    if normalized.is_empty() {
        return true;
    }
    STANDARD_PRINTER_KEYWORDS
        .iter()
        .any(|keyword| normalized.contains(keyword))
}

fn powershell_command() -> Command {
    let mut command = Command::new("powershell.exe");
    #[cfg(target_os = "windows")]
    {
        command.creation_flags(CREATE_NO_WINDOW_FLAG);
    }
    command.arg("-NoProfile").arg("-NonInteractive");
    command
}

fn decode_powershell_output(bytes: &[u8]) -> String {
    if bytes.is_empty() {
        return String::new();
    }
    if bytes.starts_with(&[0xFF, 0xFE]) {
        let (decoded, _, _) = encoding_rs::UTF_16LE.decode(&bytes[2..]);
        return decoded.into_owned();
    }
    if bytes.starts_with(&[0xFE, 0xFF]) {
        let (decoded, _, _) = encoding_rs::UTF_16BE.decode(&bytes[2..]);
        return decoded.into_owned();
    }
    if bytes.starts_with(&[0xEF, 0xBB, 0xBF]) {
        if let Ok(text) = String::from_utf8(bytes[3..].to_vec()) {
            return text;
        }
    }
    if let Ok(text) = String::from_utf8(bytes.to_vec()) {
        return text;
    }
    let (decoded_sjis, _, _) = encoding_rs::SHIFT_JIS.decode(bytes);
    decoded_sjis.into_owned()
}

fn query_printers() -> Result<Vec<Value>> {
    let output = powershell_command()
        .arg("-Command")
        .arg(QUERY_PRINTERS_SCRIPT)
        .output()
        .context("failed to run powershell Get-Printer")?;

    if !output.status.success() {
        return Err(anyhow!(
            "Get-Printer command failed: {}",
            decode_powershell_output(&output.stderr).trim()
        ));
    }

    let stdout = decode_powershell_output(&output.stdout).trim().to_string();
    if stdout.is_empty() {
        return Ok(Vec::new());
    }
    let parsed: Value =
        serde_json::from_str(&stdout).with_context(|| format!("invalid printer JSON: {stdout}"))?;

    let mut printers = Vec::new();
    if let Some(arr) = parsed.as_array() {
        for item in arr {
            if let Some(printer) = parse_printer(item) {
                printers.push(printer);
            }
        }
    } else if parsed.is_object() {
        if let Some(printer) = parse_printer(&parsed) {
            printers.push(printer);
        }
    }

    printers.retain(|item| {
        item.get("name")
            .and_then(Value::as_str)
            .map(|name| !is_standard_printer(name))
            .unwrap_or(false)
    });
    printers.sort_by(|a, b| {
        let a_default = a
            .get("is_default")
            .and_then(Value::as_bool)
            .unwrap_or(false);
        let b_default = b
            .get("is_default")
            .and_then(Value::as_bool)
            .unwrap_or(false);
        let a_name = a
            .get("name")
            .and_then(Value::as_str)
            .unwrap_or_default()
            .to_ascii_lowercase();
        let b_name = b
            .get("name")
            .and_then(Value::as_str)
            .unwrap_or_default()
            .to_ascii_lowercase();
        b_default.cmp(&a_default).then_with(|| a_name.cmp(&b_name))
    });
    Ok(printers)
}

fn parse_printer(value: &Value) -> Option<Value> {
    let name = value.get("Name")?.as_str()?.trim().to_string();
    if name.is_empty() {
        return None;
    }
    let is_default = value
        .get("Default")
        .and_then(Value::as_bool)
        .unwrap_or(false);
    Some(json!({
        "name": name,
        "is_default": is_default,
        "color_supported": true,
        "duplex_supported": true,
        "warnings": [],
    }))
}

#[derive(Clone, Debug)]
struct GmailApiConfig {
    sender: String,
    display_name: String,
    credentials_json: PathBuf,
    token_json: PathBuf,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct GmailTokenFile {
    #[serde(default)]
    token: String,
    #[serde(default)]
    access_token: String,
    #[serde(default)]
    refresh_token: String,
    #[serde(default)]
    token_uri: String,
    #[serde(default)]
    client_id: String,
    #[serde(default)]
    client_secret: String,
    #[serde(default)]
    scopes: Vec<String>,
    #[serde(default)]
    scope: String,
    #[serde(default)]
    expiry: String,
    #[serde(flatten, default)]
    extra: HashMap<String, Value>,
}

#[derive(Debug, Clone)]
struct GmailClientSecret {
    client_id: String,
    client_secret: String,
    token_uri: String,
}

#[derive(Debug, Deserialize)]
struct GoogleTokenRefreshResponse {
    access_token: String,
    expires_in: Option<i64>,
    scope: Option<String>,
    token_type: Option<String>,
}

fn run_print_command(pdf_path: &Path, printer_name: &str) -> Result<()> {
    let output = powershell_command()
        .arg("-ExecutionPolicy")
        .arg("Bypass")
        .arg("-Command")
        .arg(PRINT_SCRIPT)
        .env("RPC_FILE", pdf_path)
        .env("RPC_PRINTER", printer_name)
        .output()
        .context("failed to launch print script")?;
    if !output.status.success() {
        return Err(anyhow!(
            "print command failed (code={:?}): stderr={}, stdout={}",
            output.status.code(),
            decode_powershell_output(&output.stderr).trim(),
            decode_powershell_output(&output.stdout).trim()
        ));
    }
    Ok(())
}

fn resolve_path_from_env(var_name: &str, base_dir: &Path) -> Option<PathBuf> {
    let raw = resolve_env_var(var_name, base_dir)?;
    let candidate = PathBuf::from(raw);
    if candidate.is_absolute() {
        Some(candidate)
    } else {
        Some(base_dir.join(candidate))
    }
}

fn resolve_env_var(var_name: &str, base_dir: &Path) -> Option<String> {
    if let Ok(raw) = env::var(var_name) {
        let trimmed = raw.trim();
        if !trimmed.is_empty() {
            return Some(trimmed.to_string());
        }
    }

    for candidate in env_file_candidates(base_dir, ".env") {
        if let Some(value) = read_env_var_from_file(&candidate, var_name) {
            return Some(value);
        }
    }
    None
}

fn read_env_var_from_file(path: &Path, var_name: &str) -> Option<String> {
    if !path.exists() {
        return None;
    }

    let bytes = fs::read(path).ok()?;
    let content = decode_env_file_content(&bytes);

    for line in content.lines() {
        let line = line.trim_start_matches('\u{feff}').trim();
        if line.is_empty() || line.starts_with('#') {
            continue;
        }
        let Some((key, value)) = line.split_once('=') else {
            continue;
        };
        if key.trim() != var_name {
            continue;
        }
        let normalized = normalize_env_file_value(value);
        if normalized.is_empty() {
            return None;
        }
        return Some(normalized);
    }
    None
}

fn find_default_gmail_credentials_json(paths: &RuntimePaths) -> Option<PathBuf> {
    let mut candidates = vec![
        paths
            .base_dir
            .join("gmail_api_credentials")
            .join("credentials.json"),
    ];
    if let Some(parent) = paths.base_dir.parent() {
        candidates.push(
            parent
                .join("Python版")
                .join("gmail_api_credentials")
                .join("credentials.json"),
        );
        candidates.push(
            parent
                .join("gmail_api_credentials")
                .join("credentials.json"),
        );
    }
    if let Ok(cwd) = env::current_dir() {
        candidates.push(cwd.join("gmail_api_credentials").join("credentials.json"));
        candidates.push(
            cwd.join("Python版")
                .join("gmail_api_credentials")
                .join("credentials.json"),
        );
    }

    for candidate in candidates {
        if candidate.exists() {
            return Some(candidate);
        }
    }
    None
}

fn find_default_gmail_token_json(paths: &RuntimePaths, credentials_json: &Path) -> Option<PathBuf> {
    let mut candidates = Vec::new();
    if let Some(parent) = credentials_json.parent() {
        candidates.push(parent.join("token.json"));
    }
    candidates.push(
        paths
            .base_dir
            .join("gmail_api_credentials")
            .join("token.json"),
    );
    if let Some(parent) = paths.base_dir.parent() {
        candidates.push(
            parent
                .join("Python版")
                .join("gmail_api_credentials")
                .join("token.json"),
        );
        candidates.push(parent.join("gmail_api_credentials").join("token.json"));
    }
    if let Ok(cwd) = env::current_dir() {
        candidates.push(cwd.join("gmail_api_credentials").join("token.json"));
        candidates.push(
            cwd.join("Python版")
                .join("gmail_api_credentials")
                .join("token.json"),
        );
    }
    for candidate in candidates {
        if candidate.exists() {
            return Some(candidate);
        }
    }
    None
}

fn copy_if_missing(source: &Path, destination: &Path) {
    if destination.exists() || !source.exists() {
        return;
    }
    if let Some(parent) = destination.parent()
        && let Err(err) = fs::create_dir_all(parent)
    {
        error!(
            "failed to create destination directory for gmail auth files: {} ({err})",
            parent.display()
        );
        return;
    }
    if let Err(err) = fs::copy(source, destination) {
        error!(
            "failed to copy gmail auth file: {} -> {} ({err})",
            source.display(),
            destination.display()
        );
    }
}

fn resolve_gmail_api_config(paths: &RuntimePaths) -> GmailApiConfig {
    let runtime_credentials_dir = resolve_data_dir(&paths.base_dir).join("gmail_api_credentials");
    let runtime_credentials_json = runtime_credentials_dir.join("credentials.json");
    let runtime_token_json = runtime_credentials_dir.join("token.json");

    let credentials_json = if let Some(path) =
        resolve_path_from_env("GMAIL_CREDENTIALS_JSON", &paths.base_dir)
    {
        path
    } else {
        if !runtime_credentials_json.exists()
            && let Some(source_credentials) = find_default_gmail_credentials_json(paths)
        {
            copy_if_missing(&source_credentials, &runtime_credentials_json);
        }
        if runtime_credentials_json.exists() {
            runtime_credentials_json.clone()
        } else {
            find_default_gmail_credentials_json(paths).unwrap_or(runtime_credentials_json.clone())
        }
    };

    let token_json = if let Some(path) = resolve_path_from_env("GMAIL_TOKEN_JSON", &paths.base_dir)
    {
        path
    } else {
        if !runtime_token_json.exists()
            && let Some(source_token) = find_default_gmail_token_json(paths, &credentials_json)
        {
            copy_if_missing(&source_token, &runtime_token_json);
        }
        if runtime_token_json.exists() {
            runtime_token_json.clone()
        } else {
            find_default_gmail_token_json(paths, &credentials_json)
                .unwrap_or(runtime_token_json.clone())
        }
    };

    GmailApiConfig {
        sender: env::var("EMAIL_SENDER").unwrap_or_else(|_| DEFAULT_EMAIL_SENDER.to_string()),
        display_name: env::var("EMAIL_DISPLAY_NAME")
            .unwrap_or_else(|_| DEFAULT_EMAIL_DISPLAY_NAME.to_string()),
        credentials_json,
        token_json,
    }
}

fn gmail_api_is_configured(paths: &RuntimePaths) -> bool {
    let config = resolve_gmail_api_config(paths);
    !config.sender.trim().is_empty()
        && config.credentials_json.exists()
        && config.token_json.exists()
}

fn smtp_is_configured() -> bool {
    fn non_empty_env(name: &str) -> bool {
        env::var(name)
            .map(|value| !value.trim().is_empty())
            .unwrap_or(false)
    }
    non_empty_env("SMTP_HOST") && non_empty_env("SMTP_USER") && non_empty_env("SMTP_PASS")
}

fn email_transport_is_configured(paths: &RuntimePaths) -> bool {
    gmail_api_is_configured(paths) || smtp_is_configured()
}

fn build_email_message(
    from_name: &str,
    from_addr: &str,
    to: &[String],
    cc: &[String],
    subject: &str,
    body: &str,
    attachment_path: &Path,
) -> Result<Message> {
    let from_mailbox = Mailbox::new(
        Some(from_name.to_string()),
        from_addr
            .parse()
            .with_context(|| format!("invalid from address: {from_addr}"))?,
    );
    let mut builder = Message::builder().from(from_mailbox);
    for addr in to {
        builder = builder.to(addr
            .parse()
            .with_context(|| format!("invalid to address: {addr}"))?);
    }
    for addr in cc {
        builder = builder.cc(addr
            .parse()
            .with_context(|| format!("invalid cc address: {addr}"))?);
    }

    let file_name = attachment_path
        .file_name()
        .and_then(|s| s.to_str())
        .unwrap_or("report.pdf")
        .to_string();
    let attachment_bytes = fs::read(attachment_path)?;
    let content_type = "application/pdf".parse()?;

    let message = builder.subject(subject).multipart(
        MultiPart::mixed()
            .singlepart(SinglePart::plain(body.to_string()))
            .singlepart(Attachment::new(file_name).body(attachment_bytes, content_type)),
    )?;
    Ok(message)
}

fn load_gmail_client_secret(credentials_path: &Path) -> Result<GmailClientSecret> {
    let raw = fs::read_to_string(credentials_path)
        .with_context(|| format!("failed to read {}", credentials_path.display()))?;
    let parsed: Value = serde_json::from_str(&raw)
        .with_context(|| format!("failed to parse {}", credentials_path.display()))?;
    let secret_node = parsed
        .get("installed")
        .or_else(|| parsed.get("web"))
        .unwrap_or(&parsed);

    let client_id = secret_node
        .get("client_id")
        .and_then(Value::as_str)
        .unwrap_or_default()
        .trim()
        .to_string();
    let client_secret = secret_node
        .get("client_secret")
        .and_then(Value::as_str)
        .unwrap_or_default()
        .trim()
        .to_string();
    let token_uri = secret_node
        .get("token_uri")
        .and_then(Value::as_str)
        .unwrap_or("https://oauth2.googleapis.com/token")
        .trim()
        .to_string();

    if client_id.is_empty() || client_secret.is_empty() {
        return Err(anyhow!(
            "Gmail credentials.json に client_id/client_secret がありません。"
        ));
    }
    Ok(GmailClientSecret {
        client_id,
        client_secret,
        token_uri,
    })
}

fn load_gmail_token(token_path: &Path) -> Result<GmailTokenFile> {
    let raw = fs::read_to_string(token_path)
        .with_context(|| format!("failed to read {}", token_path.display()))?;
    let token = serde_json::from_str::<GmailTokenFile>(&raw)
        .with_context(|| format!("failed to parse {}", token_path.display()))?;
    Ok(token)
}

fn save_gmail_token(token_path: &Path, token: &GmailTokenFile) -> Result<()> {
    if let Some(parent) = token_path.parent() {
        fs::create_dir_all(parent)?;
    }
    let text = serde_json::to_string_pretty(token)?;
    fs::write(token_path, text)
        .with_context(|| format!("failed to write {}", token_path.display()))?;
    Ok(())
}

fn token_expired(expiry: &str) -> bool {
    let trimmed = expiry.trim();
    if trimmed.is_empty() {
        return true;
    }
    match DateTime::parse_from_rfc3339(trimmed) {
        Ok(expiry_dt) => {
            expiry_dt.with_timezone(&Utc) <= Utc::now() + chrono::Duration::seconds(60)
        }
        Err(_) => true,
    }
}

fn refresh_gmail_access_token(
    token: &mut GmailTokenFile,
    client_secret: &GmailClientSecret,
    token_path: &Path,
) -> Result<String> {
    let refresh_token = token.refresh_token.trim().to_string();
    if refresh_token.is_empty() {
        return Err(anyhow!(
            "token.json に refresh_token がありません。Python版で一度認証を完了してください。"
        ));
    }
    let token_uri = if token.token_uri.trim().is_empty() {
        client_secret.token_uri.clone()
    } else {
        token.token_uri.clone()
    };

    let http_client = reqwest::blocking::Client::builder()
        .timeout(Duration::from_secs(30))
        .build()?;
    let response = http_client
        .post(&token_uri)
        .form(&[
            ("client_id", client_secret.client_id.as_str()),
            ("client_secret", client_secret.client_secret.as_str()),
            ("refresh_token", refresh_token.as_str()),
            ("grant_type", "refresh_token"),
        ])
        .send()
        .context("failed to refresh Gmail access token")?;

    if !response.status().is_success() {
        let status = response.status();
        let body = response.text().unwrap_or_default();
        return Err(anyhow!(
            "Gmail token refresh failed (status={}): {}",
            status,
            body
        ));
    }

    let refreshed: GoogleTokenRefreshResponse = response
        .json()
        .context("invalid Gmail token refresh response")?;
    if refreshed.access_token.trim().is_empty() {
        return Err(anyhow!(
            "Gmail token refresh response did not include access_token"
        ));
    }

    let access_token = refreshed.access_token.trim().to_string();
    token.token = access_token.clone();
    token.access_token = access_token.clone();
    token.token_uri = token_uri;
    if token.client_id.trim().is_empty() {
        token.client_id = client_secret.client_id.clone();
    }
    if token.client_secret.trim().is_empty() {
        token.client_secret = client_secret.client_secret.clone();
    }

    if let Some(scope) = refreshed
        .scope
        .as_ref()
        .map(|v| v.trim())
        .filter(|v| !v.is_empty())
    {
        token.scope = scope.to_string();
        token.scopes = scope.split_whitespace().map(str::to_string).collect();
    } else if token.scopes.is_empty() {
        token.scopes = vec![GMAIL_SEND_SCOPE.to_string()];
        token.scope = GMAIL_SEND_SCOPE.to_string();
    }
    if let Some(token_type) = refreshed
        .token_type
        .as_ref()
        .map(|v| v.trim())
        .filter(|v| !v.is_empty())
    {
        token.extra.insert(
            "token_type".to_string(),
            Value::String(token_type.to_string()),
        );
    }

    let expires_in = refreshed.expires_in.unwrap_or(3600).max(60);
    let expiry = Utc::now() + chrono::Duration::seconds(expires_in);
    token.expiry = expiry.to_rfc3339_opts(SecondsFormat::Secs, true);
    save_gmail_token(token_path, token)?;
    Ok(access_token)
}

fn ensure_gmail_access_token(
    token: &mut GmailTokenFile,
    client_secret: &GmailClientSecret,
    token_path: &Path,
) -> Result<String> {
    let current = if token.token.trim().is_empty() {
        token.access_token.trim().to_string()
    } else {
        token.token.trim().to_string()
    };
    if !current.is_empty() && !token_expired(&token.expiry) {
        return Ok(current);
    }
    refresh_gmail_access_token(token, client_secret, token_path)
}

fn send_email_via_gmail_api(
    config: &GmailApiConfig,
    to: &[String],
    cc: &[String],
    subject: &str,
    body: &str,
    attachment_path: &Path,
) -> Result<()> {
    let client_secret = load_gmail_client_secret(&config.credentials_json)?;
    let mut token = load_gmail_token(&config.token_json)?;
    let mut access_token =
        ensure_gmail_access_token(&mut token, &client_secret, &config.token_json)?;

    let message = build_email_message(
        &config.display_name,
        &config.sender,
        to,
        cc,
        subject,
        body,
        attachment_path,
    )?;
    let encoded_message = URL_SAFE_NO_PAD.encode(message.formatted());
    let payload = json!({ "raw": encoded_message });

    let http_client = reqwest::blocking::Client::builder()
        .timeout(Duration::from_secs(30))
        .build()?;

    for attempt in 1..=2 {
        let response = http_client
            .post(GMAIL_SEND_ENDPOINT)
            .bearer_auth(&access_token)
            .json(&payload)
            .send()
            .context("failed to call Gmail send API")?;

        if response.status().is_success() {
            return Ok(());
        }

        let status = response.status();
        let body_text = response.text().unwrap_or_default();
        let should_retry = attempt == 1 && (status.as_u16() == 401 || status.as_u16() == 403);
        if should_retry {
            access_token =
                refresh_gmail_access_token(&mut token, &client_secret, &config.token_json)?;
            continue;
        }

        return Err(anyhow!(
            "Gmail API send failed (status={}): {}",
            status,
            body_text
        ));
    }
    Err(anyhow!("Gmail API send failed"))
}

fn send_email_via_smtp(
    to: &[String],
    cc: &[String],
    subject: &str,
    body: &str,
    attachment_path: &Path,
) -> Result<()> {
    let host = env::var("SMTP_HOST").context("SMTP_HOST is not set")?;
    let port: u16 = env::var("SMTP_PORT")
        .ok()
        .and_then(|v| v.parse().ok())
        .unwrap_or(587);
    let user = env::var("SMTP_USER").context("SMTP_USER is not set")?;
    let pass = env::var("SMTP_PASS").context("SMTP_PASS is not set")?;
    let from_addr = env::var("SMTP_FROM").unwrap_or_else(|_| user.clone());
    let from_name =
        env::var("SMTP_FROM_NAME").unwrap_or_else(|_| "Report PDF Converter".to_string());
    let message = build_email_message(
        &from_name,
        &from_addr,
        to,
        cc,
        subject,
        body,
        attachment_path,
    )?;

    let creds = Credentials::new(user, pass);
    let mailer = SmtpTransport::relay(&host)?
        .port(port)
        .credentials(creds)
        .build();
    mailer.send(&message).context("SMTP send failed")?;
    Ok(())
}

fn send_email_with_attachment(
    paths: &RuntimePaths,
    to: &[String],
    cc: &[String],
    subject: &str,
    body: &str,
    attachment_path: &Path,
) -> Result<()> {
    if gmail_api_is_configured(paths) {
        let config = resolve_gmail_api_config(paths);
        match send_email_via_gmail_api(&config, to, cc, subject, body, attachment_path) {
            Ok(()) => return Ok(()),
            Err(err) => {
                if smtp_is_configured() {
                    info!("gmail api send failed, fallback to smtp: {err:#}");
                } else {
                    return Err(err);
                }
            }
        }
    }
    if smtp_is_configured() {
        return send_email_via_smtp(to, cc, subject, body, attachment_path);
    }
    Err(anyhow!(
        "メール設定が構成されていません。gmail_api_credentials か SMTP を設定してください。"
    ))
}

fn convert_word_to_pdf(source: &Path, destination: &Path) -> Result<()> {
    if let Some(parent) = destination.parent() {
        fs::create_dir_all(parent)?;
    }
    let mut last_error = String::new();
    for attempt in 1..=3 {
        let output = powershell_command()
            .arg("-ExecutionPolicy")
            .arg("Bypass")
            .arg("-Command")
            .arg(WORD_TO_PDF_SCRIPT)
            .env("RPC_SOURCE", source)
            .env("RPC_DEST", destination)
            .output()?;
        if output.status.success() && destination.exists() {
            return Ok(());
        }
        last_error = format!(
            "attempt={attempt}, status={:?}, stderr={}, stdout={}",
            output.status.code(),
            decode_powershell_output(&output.stderr).trim(),
            decode_powershell_output(&output.stdout).trim()
        );
        thread::sleep(Duration::from_secs(attempt as u64));
    }
    Err(anyhow!(
        "Word変換に3回失敗しました: {} -> {} ({})",
        source.display(),
        destination.display(),
        last_error
    ))
}

fn sanitize_file_stem(value: &str) -> String {
    let mut out = String::new();
    for ch in value.chars() {
        if ch.is_ascii_alphanumeric() || matches!(ch, '-' | '_' | '.' | '(' | ')') {
            out.push(ch);
        } else {
            out.push('_');
        }
    }
    let trimmed = out.trim_matches('_');
    if trimmed.is_empty() {
        "document".to_string()
    } else {
        trimmed.to_string()
    }
}

fn merge_pdfs(pdf_paths: &[PathBuf], output_path: &Path) -> Result<()> {
    if pdf_paths.is_empty() {
        return Err(anyhow!("merge input is empty"));
    }

    let mut max_id: u32 = 1;
    let mut pages: BTreeMap<ObjectId, Object> = BTreeMap::new();
    let mut objects: BTreeMap<ObjectId, Object> = BTreeMap::new();
    let mut merged = Document::with_version("1.5");

    for path in pdf_paths {
        let mut doc = Document::load(path)?;
        doc.renumber_objects_with(max_id);
        max_id = doc.max_id + 1;
        for (_, object_id) in doc.get_pages() {
            let page = doc.get_object(object_id)?.to_owned();
            pages.insert(object_id, page);
        }
        objects.extend(doc.objects);
    }

    let mut catalog_object: Option<(ObjectId, Object)> = None;
    let mut pages_object: Option<(ObjectId, Object)> = None;
    for (object_id, object) in &objects {
        match object.type_name().unwrap_or(b"") {
            b"Catalog" => {
                let root_id = catalog_object
                    .as_ref()
                    .map(|(id, _)| *id)
                    .unwrap_or(*object_id);
                catalog_object = Some((root_id, object.clone()));
            }
            b"Pages" => {
                if let Ok(dict) = object.as_dict() {
                    let mut merged_dict = dict.clone();
                    if let Some((_, existing_obj)) = pages_object.as_ref()
                        && let Ok(existing_dict) = existing_obj.as_dict()
                    {
                        merged_dict.extend(existing_dict);
                    }
                    let root_id = pages_object
                        .as_ref()
                        .map(|(id, _)| *id)
                        .unwrap_or(*object_id);
                    pages_object = Some((root_id, Object::Dictionary(merged_dict)));
                }
            }
            b"Page" | b"Outlines" | b"Outline" => {}
            _ => {
                merged.objects.insert(*object_id, object.clone());
            }
        }
    }

    let pages_object =
        pages_object.ok_or_else(|| anyhow!("PDF内のPagesルートを取得できませんでした。"))?;
    for (object_id, object) in &pages {
        if let Ok(dict) = object.as_dict() {
            let mut dict = dict.clone();
            dict.set("Parent", pages_object.0);
            merged.objects.insert(*object_id, Object::Dictionary(dict));
        }
    }

    let catalog_object =
        catalog_object.ok_or_else(|| anyhow!("PDF内のCatalogルートを取得できませんでした。"))?;
    if let Ok(dict) = pages_object.1.as_dict() {
        let mut dict = dict.clone();
        dict.set("Count", pages.len() as u32);
        dict.set(
            "Kids",
            pages
                .keys()
                .copied()
                .map(Object::Reference)
                .collect::<Vec<_>>(),
        );
        merged
            .objects
            .insert(pages_object.0, Object::Dictionary(dict));
    }
    if let Ok(dict) = catalog_object.1.as_dict() {
        let mut dict = dict.clone();
        dict.set("Pages", pages_object.0);
        dict.remove(b"Outlines");
        merged
            .objects
            .insert(catalog_object.0, Object::Dictionary(dict));
    }

    merged.trailer.set("Root", catalog_object.0);
    merged.max_id = merged.objects.len() as u32;
    merged.renumber_objects();
    merged.adjust_zero_pages();
    merged.compress();

    if let Some(parent) = output_path.parent() {
        fs::create_dir_all(parent)?;
    }
    merged.save(output_path)?;
    Ok(())
}
