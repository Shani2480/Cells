//! Data › Get Data › From Web: the file at a web address, fetched here rather than in the webview (which a site's
//! cross-origin rules would stop). Only http and https; at most 50 MB.

use serde::Serialize;

const LIMIT: usize = 50 * 1024 * 1024;

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Fetched {
    bytes: Vec<u8>,
    content_type: String,
}

#[tauri::command]
pub async fn fetch_url(url: String) -> Result<Fetched, String> {
    let parsed = reqwest::Url::parse(&url).map_err(|_| "That isn't a web address.".to_string())?;
    if parsed.scheme() != "http" && parsed.scheme() != "https" {
        return Err("Only http and https addresses can be read.".into());
    }
    let client = reqwest::Client::builder()
        .user_agent("Cells")
        .build()
        .map_err(|e| e.to_string())?;
    let response = client
        .get(parsed)
        .send()
        .await
        .map_err(|e| format!("Couldn't reach the address: {e}"))?;
    let status = response.status();
    if !status.is_success() {
        return Err(format!("The site answered {status}."));
    }
    if response
        .content_length()
        .is_some_and(|n| n as usize > LIMIT)
    {
        return Err("The file is larger than 50 MB.".into());
    }
    let content_type = response
        .headers()
        .get(reqwest::header::CONTENT_TYPE)
        .and_then(|v| v.to_str().ok())
        .unwrap_or("")
        .to_string();
    let bytes = response.bytes().await.map_err(|e| e.to_string())?;
    if bytes.len() > LIMIT {
        return Err("The file is larger than 50 MB.".into());
    }
    Ok(Fetched {
        bytes: bytes.to_vec(),
        content_type,
    })
}
