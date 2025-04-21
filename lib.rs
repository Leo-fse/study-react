use std::process::{Command, Stdio};
use std::io::Write;
use tauri::command;

#[command]
pub async fn run_python_duckdb_query(db_path: String, query: String) -> Result<String, String> {
    let mut child = Command::new("python3") // Windowsなら "python.exe"
        .arg("python_scripts/query_duckdb.py")
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .map_err(|e| format!("Failed to start Python: {}", e))?;

    let input_json = serde_json::json!({
        "db_path": db_path,
        "query": query
    });

    if let Some(mut stdin) = child.stdin.take() {
        stdin
            .write_all(input_json.to_string().as_bytes())
            .map_err(|e| format!("Failed to write to stdin: {}", e))?;
    }

    let output = child
        .wait_with_output()
        .map_err(|e| format!("Failed to read output: {}", e))?;

    let stdout_str = String::from_utf8_lossy(&output.stdout).to_string();
    let stderr_str = String::from_utf8_lossy(&output.stderr).to_string();

    if output.status.success() {
        Ok(stdout_str)
    } else {
        if stdout_str.trim().starts_with("{") {
            // PythonがエラーをJSON形式で返してるとき
            Err(stdout_str)
        } else {
            // それ以外（謎の爆発）
            Err(format!("Python failed.\nstderr: {}\nstdout: {}", stderr_str, stdout_str))
        }
    }
}