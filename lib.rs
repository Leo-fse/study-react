use serde::{Deserialize, Serialize};
use std::process::{Command, Stdio};
use std::io::Write;
use tauri::command;

#[derive(Deserialize)]
struct QueryPayload {
  db_path: String,
  query: String,
}

#[derive(Serialize)]
struct DbRow(serde_json::Value); // 柔軟に受け取れるよう Value

#[command]
async fn query_duckdb(payload: QueryPayload) -> Result<Vec<DbRow>, String> {
  // Python スクリプトを起動
  let mut child = Command::new("python3") // または埋め込み Python の実行ファイル
    .arg("scripts/query_duckdb.py")
    .stdin(Stdio::piped())
    .stdout(Stdio::piped())
    .spawn()
    .map_err(|e| format!("failed to spawn python: {}", e))?;

  // stdin に JSON を書き込む
  {
    let stdin = child.stdin.as_mut().ok_or("failed to open stdin")?;
    let input = serde_json::to_vec(&payload).map_err(|e| e.to_string())?;
    stdin.write_all(&input).map_err(|e| e.to_string())?;
  }

  // stdout を読み取り
  let output = child.wait_with_output().map_err(|e| e.to_string())?;
  if !output.status.success() {
    return Err(format!("python error: {:?}", output));
  }

  // JSON をパースして返却
  let rows: Vec<DbRow> = serde_json::from_slice(&output.stdout)
    .map_err(|e| format!("invalid json: {}", e))?;
  Ok(rows)
}

fn main() {
  tauri::Builder::default()
    .invoke_handler(tauri::generate_handler![query_duckdb])
    .run(tauri::generate_context!())
    .expect("error while running tauri application");
}