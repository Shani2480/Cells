//! AI requests for Cells' AI tools (formula from words, explain a formula, suggest a chart, summarize data), with the
//! providers and keys Inkspire uses (`docs/ai.md`). Keys stay in this process: the webview only learns whether one is
//! there. They live in the system keyring under Inkspire's service name, or, where there is no keyring, in Inkspire's
//! owner-only fallback file, so a key entered in either app works in both. Ported from Inkspire's `ai.rs` and
//! `secrets.rs`, without the HTML handling (Cells only wants plain text) and without Inkspire's built-in engine.

use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::sync::{Arc, LazyLock, Mutex, OnceLock};
use std::time::Duration;

use futures_util::StreamExt;
use serde::{Deserialize, Serialize};
use tauri::{Emitter, Manager};

/// Inkspire's keyring service; the account is the provider ("openai", "anthropic", "compatible").
const SERVICE: &str = "com.aiword.app.ai";
/// Inkspire's app identifier: its fallback key file is `<config dir>/com.aiword.app/ai-keys.json`.
const INKSPIRE_ID: &str = "com.aiword.app";
const KEY_FIELDS: [&str; 3] = ["openai", "anthropic", "compatible"];

const OPENAI_URL: &str = "https://api.openai.com/v1/chat/completions";
const OPENAI_MODELS_URL: &str = "https://api.openai.com/v1/models";
const ANTHROPIC_URL: &str = "https://api.anthropic.com/v1/messages";
const DEFAULT_OPENAI_MODEL: &str = "gpt-5-mini";
const DEFAULT_ANTHROPIC_MODEL: &str = "claude-sonnet-5-5";
const DEFAULT_OLLAMA_MODEL: &str = "llama3.2";
const DEFAULT_COMPATIBLE_MODEL: &str = "default";
const MAX_PROMPT_CHARS: usize = 8_000;
const MAX_SYSTEM_CHARS: usize = 8_000;
const MAX_CONTEXT_CHARS: usize = 24_000;
const MAX_OUTPUT_CHARS: usize = 40_000;
/// A small local model can fall into a loop; this stops it after a long page.
const MAX_REPLY_TOKENS: u32 = 2048;

#[derive(Debug, Clone, Copy, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "lowercase")]
pub enum AiProvider {
    Openai,
    Anthropic,
    Ollama,
    /// Any OpenAI-compatible endpoint (LM Studio, llama.cpp, vLLM, a company gateway): `base_url`, and a key when it
    /// wants one.
    Compatible,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProviderConfig {
    provider: AiProvider,
    #[serde(default)]
    model: Option<String>,
    #[serde(default)]
    base_url: Option<String>,
}

// ── Keys ──────────────────────────────────────────────────────────────────────────────────────────────────────────

fn keyring_read(field: &str) -> Result<Option<String>, String> {
    let entry = keyring::Entry::new(SERVICE, field).map_err(|err| err.to_string())?;
    match entry.get_password() {
        Ok(value) => Ok(Some(value)),
        Err(keyring::Error::NoEntry) => Ok(None),
        Err(err) => Err(err.to_string()),
    }
}

fn keyring_write(field: &str, value: &str) -> Result<(), String> {
    let entry = keyring::Entry::new(SERVICE, field).map_err(|err| err.to_string())?;
    if value.is_empty() {
        return match entry.delete_credential() {
            Ok(()) | Err(keyring::Error::NoEntry) => Ok(()),
            Err(err) => Err(err.to_string()),
        };
    }
    entry.set_password(value).map_err(|err| err.to_string())
}

fn key_file(app: &tauri::AppHandle) -> Result<PathBuf, String> {
    let base = app
        .path()
        .config_dir()
        .map_err(|_| "Could not open the settings folder.".to_string())?;
    Ok(base.join(INKSPIRE_ID).join("ai-keys.json"))
}

fn read_key_file(path: &Path) -> HashMap<String, String> {
    std::fs::read_to_string(path)
        .ok()
        .and_then(|text| serde_json::from_str(&text).ok())
        .unwrap_or_default()
}

fn write_key_file(path: &Path, keys: &HashMap<String, String>) -> Result<(), String> {
    let failed = || "Could not store the key.".to_string();
    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent).map_err(|_| failed())?;
        #[cfg(unix)]
        {
            use std::os::unix::fs::PermissionsExt;
            let _ = std::fs::set_permissions(parent, std::fs::Permissions::from_mode(0o700));
        }
    }
    let text = serde_json::to_string(keys).map_err(|_| failed())?;
    #[cfg(unix)]
    {
        use std::io::Write;
        use std::os::unix::fs::{OpenOptionsExt, PermissionsExt};
        let mut file = std::fs::OpenOptions::new()
            .write(true)
            .create(true)
            .truncate(true)
            .mode(0o600)
            .open(path)
            .map_err(|_| failed())?;
        file.write_all(text.as_bytes()).map_err(|_| failed())?;
        let _ = std::fs::set_permissions(path, std::fs::Permissions::from_mode(0o600));
        Ok(())
    }
    #[cfg(not(unix))]
    {
        std::fs::write(path, text).map_err(|_| failed())
    }
}

/// A saved key: the keyring's, or else the fallback file's. Empty when there is none.
fn load_key(app: &tauri::AppHandle, field: &str) -> String {
    if let Ok(Some(value)) = keyring_read(field)
        && !value.trim().is_empty()
    {
        return value.trim().to_string();
    }
    key_file(app)
        .map(|path| read_key_file(&path))
        .unwrap_or_default()
        .get(field)
        .map(|value| value.trim().to_string())
        .unwrap_or_default()
}

/// Saves a key (an empty one removes it) into the keyring, or the owner-only file when there is no keyring.
/// Returns where it went: "keyring", "file" or "none".
fn save_key(app: &tauri::AppHandle, field: &str, value: &str) -> Result<&'static str, String> {
    let path = key_file(app)?;
    if keyring_write(field, value).is_ok() {
        // A copy left in the file must not outlive the key.
        let mut keys = read_key_file(&path);
        if keys.remove(field).is_some() {
            if keys.values().all(|v| v.trim().is_empty()) {
                let _ = std::fs::remove_file(&path);
            } else {
                let _ = write_key_file(&path, &keys);
            }
        }
        return Ok(if value.is_empty() { "none" } else { "keyring" });
    }
    let mut keys = read_key_file(&path);
    keys.insert(field.to_string(), value.to_string());
    write_key_file(&path, &keys)?;
    Ok(if value.is_empty() { "none" } else { "file" })
}

#[derive(Serialize)]
pub struct AiKeyStatus {
    /// Per provider that takes a key: "saved", "environment" (OPENAI_API_KEY / ANTHROPIC_API_KEY) or "missing".
    openai: String,
    anthropic: String,
    compatible: String,
    /// Where keys are kept: "keyring", or "file" when there is no keyring to use.
    storage: String,
}

fn key_source(saved: &str, env_name: &str) -> &'static str {
    if !saved.trim().is_empty() {
        "saved"
    } else if !env_name.is_empty()
        && std::env::var(env_name)
            .map(|value| !value.trim().is_empty())
            .unwrap_or(false)
    {
        "environment"
    } else {
        "missing"
    }
}

#[tauri::command]
pub fn ai_key_status(app: tauri::AppHandle) -> AiKeyStatus {
    let source = |field: &str, env: &str| key_source(&load_key(&app, field), env).to_string();
    AiKeyStatus {
        openai: source("openai", "OPENAI_API_KEY"),
        anthropic: source("anthropic", "ANTHROPIC_API_KEY"),
        compatible: source("compatible", ""),
        storage: if keyring_read("probe").is_ok() {
            "keyring"
        } else {
            "file"
        }
        .to_string(),
    }
}

#[tauri::command]
pub fn ai_save_key(app: tauri::AppHandle, provider: String, key: String) -> Result<String, String> {
    let field = provider.trim();
    if !KEY_FIELDS.contains(&field) {
        return Err("Choose a provider that takes a key.".into());
    }
    let key = key.trim();
    if key.chars().count() > 400 || key.chars().any(|ch| ch.is_control()) {
        return Err("That key cannot be saved.".into());
    }
    save_key(&app, field, key).map(str::to_string)
}

fn api_key(
    app: &tauri::AppHandle,
    field: &str,
    env_name: &str,
    provider: &str,
) -> Result<String, String> {
    let saved = load_key(app, field);
    if !saved.is_empty() {
        return Ok(saved);
    }
    let key = std::env::var(env_name)
        .unwrap_or_default()
        .trim()
        .to_string();
    if key.is_empty() {
        return Err(format!("Add a {provider} key in File › Options › AI."));
    }
    Ok(key)
}

// ── Requests ──────────────────────────────────────────────────────────────────────────────────────────────────────

/// Requests that may still be stopped, by id.
static RUNNING: LazyLock<Mutex<HashMap<String, Arc<tokio::sync::Notify>>>> =
    LazyLock::new(|| Mutex::new(HashMap::new()));

#[derive(Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct AiChunk {
    request_id: String,
    text: String,
}

#[tauri::command]
pub fn ai_cancel(request_id: String) {
    if let Some(stop) = RUNNING
        .lock()
        .ok()
        .and_then(|map| map.get(&request_id).cloned())
    {
        // A stored permit: a stop sent before the request waits for it still counts.
        stop.notify_one();
    }
}

/// Sends one request and streams the reply as `ai-chunk` events to the window; resolves with the whole reply.
#[tauri::command]
pub async fn ai_generate(
    app: tauri::AppHandle,
    request_id: String,
    system: String,
    prompt: String,
    context: String,
    provider_config: ProviderConfig,
) -> Result<String, String> {
    let request_id = request_id.trim().to_string();
    if request_id.is_empty() || request_id.len() > 80 {
        return Err("The request has no id.".into());
    }
    let prompt = prompt.trim();
    if prompt.is_empty() {
        return Err("Enter a request.".into());
    }
    if prompt.chars().count() > MAX_PROMPT_CHARS || system.chars().count() > MAX_SYSTEM_CHARS {
        return Err("The request is too long.".into());
    }
    let context = clip_chars(context.trim(), MAX_CONTEXT_CHARS);
    let stop = Arc::new(tokio::sync::Notify::new());
    if let Ok(mut map) = RUNNING.lock() {
        map.insert(request_id.clone(), stop.clone());
    }
    let result = tokio::select! {
        result = run(&app, &request_id, &provider_config, system.trim(), prompt, &context) => result,
        _ = stop.notified() => Err("Stopped.".into()),
    };
    if let Ok(mut map) = RUNNING.lock() {
        map.remove(&request_id);
    }
    result
}

async fn run(
    app: &tauri::AppHandle,
    request_id: &str,
    config: &ProviderConfig,
    system: &str,
    prompt: &str,
    context: &str,
) -> Result<String, String> {
    let response = send_request(app, config, system, &user_message(prompt, context)).await?;
    read_reply(response, config.provider, |delta| {
        let _ = app.emit_to(
            "main",
            "ai-chunk",
            AiChunk {
                request_id: request_id.to_string(),
                text: delta.to_string(),
            },
        );
    })
    .await
}

/// Reads a streamed reply, handing each piece to `on_text` as it comes; resolves with the whole reply.
async fn read_reply(
    response: reqwest::Response,
    provider: AiProvider,
    mut on_text: impl FnMut(&str),
) -> Result<String, String> {
    let mut parser = StreamParser::new(provider);
    let mut collected = String::new();
    let mut stream = response.bytes_stream();
    let mut add = |delta: String, collected: &mut String| -> Result<(), String> {
        if collected.chars().count() + delta.chars().count() > MAX_OUTPUT_CHARS {
            return Err("The reply was longer than Cells can use.".into());
        }
        on_text(&delta);
        collected.push_str(&delta);
        Ok(())
    };
    while let Some(item) = stream.next().await {
        let bytes = item.map_err(http_failure)?;
        for delta in parser.push(&bytes)? {
            add(delta, &mut collected)?;
        }
    }
    if let Some(tail) = parser.finish()? {
        add(tail, &mut collected)?;
    }
    let text = strip_fences(collected.trim()).trim().to_string();
    if text.is_empty() {
        return Err("The model returned no text.".into());
    }
    Ok(text)
}

async fn send_request(
    app: &tauri::AppHandle,
    config: &ProviderConfig,
    system: &str,
    message: &str,
) -> Result<reqwest::Response, String> {
    let builder = match config.provider {
        AiProvider::Openai => {
            let model = clean_model(config.model.as_deref(), DEFAULT_OPENAI_MODEL)?;
            let key = api_key(app, "openai", "OPENAI_API_KEY", "OpenAI")?;
            http()
                .post(OPENAI_URL)
                .bearer_auth(key)
                .header("accept", "text/event-stream")
                .json(&serde_json::json!({
                    "model": model,
                    "stream": true,
                    "messages": [
                        { "role": "system", "content": system },
                        { "role": "user", "content": message }
                    ]
                }))
        }
        AiProvider::Anthropic => {
            let model = clean_model(config.model.as_deref(), DEFAULT_ANTHROPIC_MODEL)?;
            let key = api_key(app, "anthropic", "ANTHROPIC_API_KEY", "Anthropic")?;
            http()
                .post(ANTHROPIC_URL)
                .header("x-api-key", key)
                .header("anthropic-version", "2023-06-01")
                .header("accept", "text/event-stream")
                .json(&serde_json::json!({
                    "model": model,
                    "max_tokens": MAX_REPLY_TOKENS,
                    "stream": true,
                    "system": system,
                    "messages": [{ "role": "user", "content": message }]
                }))
        }
        AiProvider::Ollama => {
            let model = clean_model(config.model.as_deref(), DEFAULT_OLLAMA_MODEL)?;
            http()
                .post(format!(
                    "{}/api/chat",
                    ollama_base(config.base_url.as_deref())?
                ))
                .json(&serde_json::json!({
                    "model": model,
                    "stream": true,
                    "options": { "num_predict": MAX_REPLY_TOKENS },
                    "messages": [
                        { "role": "system", "content": system },
                        { "role": "user", "content": message }
                    ]
                }))
        }
        AiProvider::Compatible => {
            let model = clean_model(config.model.as_deref(), DEFAULT_COMPATIBLE_MODEL)?;
            let url = format!(
                "{}/chat/completions",
                compatible_base(config.base_url.as_deref())?
            );
            let mut builder = http().post(url).header("accept", "text/event-stream");
            let key = load_key(app, "compatible");
            if !key.is_empty() {
                builder = builder.bearer_auth(key);
            }
            builder.json(&chat_body(&model, system, message))
        }
    };
    let response = builder.send().await.map_err(http_failure)?;
    if response.status().is_success() {
        return Ok(response);
    }
    let status = response.status();
    let body = response.text().await.unwrap_or_default();
    Err(provider_message(status, &body, config.provider))
}

/// Checks the provider answers (and, for Ollama, has the model), without spending much.
#[tauri::command]
pub async fn ai_test(
    app: tauri::AppHandle,
    provider_config: ProviderConfig,
) -> Result<String, String> {
    let config = provider_config;
    let request =
        match config.provider {
            AiProvider::Openai => {
                clean_model(config.model.as_deref(), DEFAULT_OPENAI_MODEL)?;
                http().get(OPENAI_MODELS_URL).bearer_auth(api_key(
                    &app,
                    "openai",
                    "OPENAI_API_KEY",
                    "OpenAI",
                )?)
            }
            AiProvider::Anthropic => {
                let model = clean_model(config.model.as_deref(), DEFAULT_ANTHROPIC_MODEL)?;
                http()
                .post(ANTHROPIC_URL)
                .header("x-api-key", api_key(&app, "anthropic", "ANTHROPIC_API_KEY", "Anthropic")?)
                .header("anthropic-version", "2023-06-01")
                .json(&serde_json::json!({
                    "model": model,
                    "max_tokens": 16,
                    "messages": [{ "role": "user", "content": "Reply with the single word ok." }]
                }))
            }
            AiProvider::Ollama => http().get(format!(
                "{}/api/tags",
                ollama_base(config.base_url.as_deref())?
            )),
            AiProvider::Compatible => {
                let mut request = http().get(format!(
                    "{}/models",
                    compatible_base(config.base_url.as_deref())?
                ));
                let key = load_key(&app, "compatible");
                if !key.is_empty() {
                    request = request.bearer_auth(key);
                }
                request
            }
        };
    let response = request.send().await.map_err(http_failure)?;
    let status = response.status();
    let body = response.text().await.unwrap_or_default();
    if !status.is_success() {
        return Err(provider_message(status, &body, config.provider));
    }
    match config.provider {
        AiProvider::Ollama => {
            let model = clean_model(config.model.as_deref(), DEFAULT_OLLAMA_MODEL)?;
            if !ollama_has_model(&body, &model) {
                return Err(format!(
                    "Connected to Ollama, but {model} is not installed. Pull it in Ollama, then try again."
                ));
            }
            Ok(format!("Connected to Ollama. Model ready: {model}."))
        }
        AiProvider::Openai => Ok("Connected to OpenAI.".into()),
        AiProvider::Anthropic => Ok("Connected to Anthropic.".into()),
        AiProvider::Compatible => Ok("Connected to the endpoint.".into()),
    }
}

fn http() -> &'static reqwest::Client {
    static CLIENT: OnceLock<reqwest::Client> = OnceLock::new();
    CLIENT.get_or_init(|| {
        reqwest::Client::builder()
            .connect_timeout(Duration::from_secs(10))
            .read_timeout(Duration::from_secs(120))
            .redirect(reqwest::redirect::Policy::none())
            .build()
            .expect("HTTP client")
    })
}

fn chat_body(model: &str, system: &str, message: &str) -> serde_json::Value {
    serde_json::json!({
        "model": model,
        "stream": true,
        "max_tokens": MAX_REPLY_TOKENS,
        "messages": [
            { "role": "system", "content": system },
            { "role": "user", "content": message }
        ]
    })
}

fn user_message(prompt: &str, context: &str) -> String {
    if context.is_empty() {
        format!("Task:\n{prompt}")
    } else {
        format!("Task:\n{prompt}\n\nWorkbook context:\n{context}")
    }
}

fn clean_model(model: Option<&str>, default: &str) -> Result<String, String> {
    let model = model.unwrap_or("").trim();
    let model = if model.is_empty() { default } else { model };
    let valid = model.len() <= 80
        && model
            .chars()
            .all(|ch| ch.is_ascii_alphanumeric() || matches!(ch, '.' | '_' | '-' | ':' | '/'));
    if !valid {
        return Err("The model name contains unsupported characters.".into());
    }
    Ok(model.to_string())
}

/// An OpenAI-compatible endpoint's base: https anywhere, or http on this computer or the local network.
fn compatible_base(base_url: Option<&str>) -> Result<String, String> {
    let raw = base_url.unwrap_or("").trim().trim_end_matches('/');
    if raw.is_empty() {
        return Err("Enter the endpoint's address (for example http://127.0.0.1:1234/v1).".into());
    }
    let raw = raw.strip_suffix("/chat/completions").unwrap_or(raw);
    let url = reqwest::Url::parse(raw)
        .map_err(|_| "The endpoint address is not a valid URL.".to_string())?;
    let host = url.host_str().unwrap_or("").to_string();
    match url.scheme() {
        "https" => {}
        "http" if is_local_host(&host) => {}
        "http" => return Err("A remote endpoint has to use https.".into()),
        _ => return Err("The endpoint has to use http or https.".into()),
    }
    if !url.username().is_empty() || url.password().is_some() {
        return Err("Leave the key out of the address; enter it in the key field.".into());
    }
    Ok(raw.to_string())
}

fn is_local_host(host: &str) -> bool {
    if matches!(host, "localhost" | "[::1]" | "::1")
        || host.ends_with(".local")
        || host.ends_with(".localhost")
    {
        return true;
    }
    match host.parse::<std::net::Ipv4Addr>() {
        Ok(ip) => ip.is_loopback() || ip.is_private() || ip.is_link_local(),
        Err(_) => false,
    }
}

fn ollama_base(base_url: Option<&str>) -> Result<String, String> {
    let base = base_url.unwrap_or("").trim().trim_end_matches('/');
    let base = if base.is_empty() {
        "http://127.0.0.1:11434"
    } else {
        base
    };
    let url = reqwest::Url::parse(base)
        .map_err(|_| "The Ollama address is not a valid URL.".to_string())?;
    if url.scheme() != "http" && url.scheme() != "https" {
        return Err("Ollama must use http or https.".into());
    }
    if !matches!(
        url.host_str().unwrap_or(""),
        "localhost" | "127.0.0.1" | "[::1]" | "::1"
    ) {
        return Err("Ollama is limited to this computer.".into());
    }
    Ok(base.to_string())
}

fn ollama_has_model(body: &str, model: &str) -> bool {
    let Ok(value) = serde_json::from_str::<serde_json::Value>(body) else {
        return false;
    };
    let needle = model.to_ascii_lowercase();
    value
        .get("models")
        .and_then(|models| models.as_array())
        .map(|models| {
            models.iter().any(|item| {
                item.get("name")
                    .and_then(|name| name.as_str())
                    .map(|name| {
                        let name = name.to_ascii_lowercase();
                        name == needle || name.starts_with(&format!("{needle}:"))
                    })
                    .unwrap_or(false)
            })
        })
        .unwrap_or(false)
}

fn http_failure(err: reqwest::Error) -> String {
    if err.is_connect() {
        "Could not reach the AI provider. For Ollama or LM Studio, start it and try again.".into()
    } else if err.is_timeout() {
        "The AI provider took too long to respond.".into()
    } else {
        "The AI request failed.".into()
    }
}

const TOO_LONG: &str = "The data is too big for this model. Select fewer cells, or choose a model with a larger context.";

fn provider_message(status: reqwest::StatusCode, body: &str, provider: AiProvider) -> String {
    if status.as_u16() == 401 || status.as_u16() == 403 {
        return match provider {
            AiProvider::Openai => "OpenAI rejected the API key.".into(),
            AiProvider::Anthropic => "Anthropic rejected the API key.".into(),
            AiProvider::Ollama => "Ollama rejected the request.".into(),
            AiProvider::Compatible => "The endpoint rejected the API key.".into(),
        };
    }
    if status.as_u16() == 429 {
        return "The AI provider is rate limiting requests (429).".into();
    }
    if provider == AiProvider::Ollama && status.as_u16() == 404 {
        return "That Ollama model is not installed. Pull it in Ollama, then try again.".into();
    }
    if let Ok(value) = serde_json::from_str::<serde_json::Value>(body) {
        let message = value
            .get("error")
            .and_then(|error| {
                error
                    .get("message")
                    .and_then(|m| m.as_str())
                    .or_else(|| error.as_str())
            })
            .unwrap_or("");
        let lower = message.to_lowercase();
        if (lower.contains("context") && (lower.contains("exceed") || lower.contains("maximum")))
            || lower.contains("prompt is too long")
        {
            return TOO_LONG.into();
        }
        if !message.is_empty() {
            return redact(message);
        }
    }
    format!("The AI provider returned status {}.", status.as_u16())
}

/// Keeps anything that looks like a key out of a message shown to the person.
fn redact(message: &str) -> String {
    let mut out = String::new();
    let mut rest = message;
    while let Some(start) = rest.find("sk-") {
        out.push_str(&rest[..start]);
        out.push_str("[redacted]");
        let after = &rest[start + 3..];
        let skip = after
            .chars()
            .take_while(|ch| ch.is_ascii_alphanumeric() || *ch == '-' || *ch == '_')
            .count();
        rest = &after[skip..];
    }
    out.push_str(rest);
    out.chars().take(300).collect()
}

/// Reads a streamed reply line by line, keeping bytes until a line is whole, so a character split between network
/// chunks (é, Hebrew, an emoji) decodes cleanly.
struct StreamParser {
    provider: AiProvider,
    buffer: Vec<u8>,
}

impl StreamParser {
    fn new(provider: AiProvider) -> Self {
        Self {
            provider,
            buffer: Vec::new(),
        }
    }

    fn push(&mut self, chunk: impl AsRef<[u8]>) -> Result<Vec<String>, String> {
        self.buffer.extend_from_slice(chunk.as_ref());
        let mut deltas = Vec::new();
        while let Some(index) = self.buffer.iter().position(|byte| *byte == b'\n') {
            let line: Vec<u8> = self.buffer.drain(..=index).collect();
            let line = String::from_utf8_lossy(&line[..index]);
            if let Some(text) = self.consume_line(line.trim_end_matches('\r'))? {
                deltas.push(text);
            }
        }
        Ok(deltas)
    }

    fn finish(&mut self) -> Result<Option<String>, String> {
        let line = std::mem::take(&mut self.buffer);
        let line = String::from_utf8_lossy(&line);
        let line = line.trim();
        if line.is_empty() {
            return Ok(None);
        }
        self.consume_line(line)
    }

    fn consume_line(&self, line: &str) -> Result<Option<String>, String> {
        let line = line.trim();
        if line.is_empty() || line.starts_with(':') || line.starts_with("event:") {
            return Ok(None);
        }
        if self.provider == AiProvider::Ollama {
            return ollama_delta(line);
        }
        let Some(data) = line.strip_prefix("data:") else {
            return Ok(None);
        };
        let data = data.trim();
        if data == "[DONE]" {
            return Ok(None);
        }
        if self.provider == AiProvider::Anthropic {
            anthropic_delta(data)
        } else {
            openai_delta(data)
        }
    }
}

fn unreadable() -> String {
    "The AI provider returned a reply Cells could not read.".into()
}

fn openai_delta(data: &str) -> Result<Option<String>, String> {
    let value: serde_json::Value = serde_json::from_str(data).map_err(|_| unreadable())?;
    if let Some(message) = value.pointer("/error/message").and_then(|m| m.as_str()) {
        return Err(redact(message));
    }
    Ok(value
        .pointer("/choices/0/delta/content")
        .and_then(|content| content.as_str())
        .filter(|text| !text.is_empty())
        .map(str::to_string))
}

fn anthropic_delta(data: &str) -> Result<Option<String>, String> {
    let value: serde_json::Value = serde_json::from_str(data).map_err(|_| unreadable())?;
    if value.get("type").and_then(|kind| kind.as_str()) == Some("error") {
        let message = value
            .pointer("/error/message")
            .and_then(|m| m.as_str())
            .unwrap_or("Anthropic rejected the request.");
        return Err(redact(message));
    }
    Ok(value
        .pointer("/delta/text")
        .and_then(|text| text.as_str())
        .filter(|text| !text.is_empty())
        .map(str::to_string))
}

fn ollama_delta(data: &str) -> Result<Option<String>, String> {
    let value: serde_json::Value = serde_json::from_str(data).map_err(|_| unreadable())?;
    if let Some(message) = value.get("error").and_then(|error| error.as_str()) {
        return Err(redact(message));
    }
    Ok(value
        .pointer("/message/content")
        .and_then(|content| content.as_str())
        .filter(|text| !text.is_empty())
        .map(str::to_string))
}

/// A reply wrapped in a Markdown code fence (```json … ```), without the fence.
fn strip_fences(value: &str) -> &str {
    let trimmed = value.trim();
    let Some(rest) = trimmed.strip_prefix("```") else {
        return trimmed;
    };
    let rest = rest.trim_start_matches(|ch: char| ch.is_ascii_alphanumeric());
    let rest = rest.trim_start_matches(['\r', '\n']);
    rest.strip_suffix("```").unwrap_or(rest).trim()
}

fn clip_chars(value: &str, max: usize) -> String {
    value.chars().take(max).collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn stream_lines_from_each_provider() {
        let mut p = StreamParser::new(AiProvider::Openai);
        let line = "data: {\"choices\":[{\"delta\":{\"content\":\"=SUM(\"}}]}\n";
        assert_eq!(p.push(line).unwrap(), vec!["=SUM(".to_string()]);
        assert_eq!(p.push("data: [DONE]\n").unwrap(), Vec::<String>::new());
        let mut p = StreamParser::new(AiProvider::Anthropic);
        assert_eq!(
            p.push(
                "event: x\ndata: {\"type\":\"content_block_delta\",\"delta\":{\"text\":\"B2\"}}\n"
            )
            .unwrap(),
            vec!["B2".to_string()]
        );
        assert!(
            p.push("data: {\"type\":\"error\",\"error\":{\"message\":\"bad sk-abc123\"}}\n")
                .is_err()
        );
        let mut p = StreamParser::new(AiProvider::Ollama);
        assert_eq!(
            p.push("{\"message\":{\"content\":\"ok\"}}").unwrap(),
            Vec::<String>::new()
        );
        assert_eq!(p.finish().unwrap().as_deref(), Some("ok"));
    }

    #[test]
    fn a_character_split_between_chunks_survives() {
        let mut p = StreamParser::new(AiProvider::Openai);
        let line = "data: {\"choices\":[{\"delta\":{\"content\":\"Prüfung\"}}]}\n".as_bytes();
        let cut = line.iter().position(|b| *b == 0xc3).unwrap() + 1;
        assert!(p.push(&line[..cut]).unwrap().is_empty());
        assert_eq!(p.push(&line[cut..]).unwrap(), vec!["Prüfung".to_string()]);
    }

    #[test]
    fn addresses_models_and_messages() {
        assert!(compatible_base(Some("http://127.0.0.1:1234/v1/")).is_ok());
        assert!(compatible_base(Some("http://192.168.1.5:8080/v1")).is_ok());
        assert!(compatible_base(Some("http://example.com/v1")).is_err());
        assert!(compatible_base(Some("https://user:pw@example.com/v1")).is_err());
        assert!(ollama_base(Some("http://example.com:11434")).is_err());
        assert_eq!(ollama_base(None).unwrap(), "http://127.0.0.1:11434");
        assert_eq!(clean_model(Some(" "), "x").unwrap(), "x");
        assert!(clean_model(Some("bad model"), "x").is_err());
        assert!(ollama_has_model(
            r#"{"models":[{"name":"llama3.2:latest"}]}"#,
            "llama3.2"
        ));
        assert_eq!(redact("key sk-abc_DEF-1 wrong"), "key [redacted] wrong");
        assert_eq!(strip_fences("```json\n{\"a\":1}\n```"), "{\"a\":1}");
        assert_eq!(key_source(" ", ""), "missing");
        assert_eq!(key_source("k", ""), "saved");
    }

    #[test]
    fn the_key_file_is_owner_only_and_round_trips() {
        let dir = std::env::temp_dir().join(format!("cells-keys-{}", std::process::id()));
        let path = dir.join(INKSPIRE_ID).join("ai-keys.json");
        let mut keys = HashMap::new();
        keys.insert("openai".to_string(), "sk-test".to_string());
        write_key_file(&path, &keys).unwrap();
        assert_eq!(
            read_key_file(&path).get("openai").map(String::as_str),
            Some("sk-test")
        );
        #[cfg(unix)]
        {
            use std::os::unix::fs::PermissionsExt;
            assert_eq!(
                std::fs::metadata(&path).unwrap().permissions().mode() & 0o777,
                0o600
            );
        }
        let _ = std::fs::remove_dir_all(&dir);
    }

    /// A fake OpenAI-compatible endpoint on this computer that streams its reply in pieces, one of them cutting a
    /// character in two: the whole request path (reqwest's stream, the parser, the fence) reads it back.
    #[test]
    fn a_streamed_reply_from_a_local_endpoint() {
        use std::io::{Read, Write};
        let listener = std::net::TcpListener::bind("127.0.0.1:0").unwrap();
        let address = listener.local_addr().unwrap();
        let server = std::thread::spawn(move || {
            let (mut socket, _) = listener.accept().unwrap();
            let mut request = [0u8; 4096];
            let _ = socket.read(&mut request).unwrap();
            let body = "data: {\"choices\":[{\"delta\":{\"content\":\"```json\\n{\\\"formula\\\": \"}}]}\n\n\
                        data: {\"choices\":[{\"delta\":{\"content\":\"\\\"=SUM(B1:B2)\\\", \\\"explanation\\\": \\\"Prüft\\\"}\\n```\"}}]}\n\n\
                        data: [DONE]\n\n";
            let bytes = body.as_bytes();
            let cut = bytes.iter().position(|b| *b == 0xc3).unwrap() + 1;
            let head =
                "HTTP/1.1 200 OK\r\ncontent-type: text/event-stream\r\nconnection: close\r\n\r\n";
            socket.write_all(head.as_bytes()).unwrap();
            socket.write_all(&bytes[..cut]).unwrap();
            socket.flush().unwrap();
            std::thread::sleep(std::time::Duration::from_millis(50));
            socket.write_all(&bytes[cut..]).unwrap();
        });
        let runtime = tokio::runtime::Builder::new_current_thread()
            .enable_all()
            .build()
            .unwrap();
        let mut pieces = Vec::new();
        let reply = runtime.block_on(async {
            let base = compatible_base(Some(&format!("http://{address}/v1"))).unwrap();
            let response = http()
                .post(format!("{base}/chat/completions"))
                .json(&chat_body("default", "system", "task"))
                .send()
                .await
                .unwrap();
            read_reply(response, AiProvider::Compatible, |t| {
                pieces.push(t.to_string())
            })
            .await
        });
        server.join().unwrap();
        assert_eq!(
            reply.unwrap(),
            "{\"formula\": \"=SUM(B1:B2)\", \"explanation\": \"Prüft\"}"
        );
        assert_eq!(pieces.len(), 2);
    }
}
