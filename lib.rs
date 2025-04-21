use serde::Deserialize;
use std::process::{Command, Stdio};
use std::io::Write;
use tauri::command;

#[derive(Deserialize)]
struct QueryPayload {
  db_path: String,
  query:   String,
}

#[command]
async fn query_duckdb(payload: QueryPayload) -> Result<Vec<serde_json::Value>, String> {
  let mut child = Command::new("python3")
    .arg("scripts/query_duckdb.py")
    .stdin(Stdio::piped())
    .stdout(Stdio::piped())
    .spawn()
    .map_err(|e| e.to_string())?;

  // stdin に JSON を書き込む
  {
    let stdin = child.stdin.as_mut().ok_or("failed to open stdin")?;
    let input = serde_json::to_vec(&payload).map_err(|e| e.to_string())?;
    stdin.write_all(&input).map_err(|e| e.to_string())?;
  }

  let output = child.wait_with_output().map_err(|e| e.to_string())?;
  if !output.status.success() {
    return Err(format!("python error: {:?}", output));
  }

  // Vec<Value> に直接パース
  let rows = serde_json::from_slice(&output.stdout)
    .map_err(|e| format!("invalid json: {}", e))?;
  Ok(rows)
}