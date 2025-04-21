use std::process::{Command, Stdio};
use std::io::Write;
use tauri::command;

#[command]
fn run_python_duckdb_query(db_path: String, query: String) -> Result<String, String> {
    let mut child = Command::new("python3")
        .arg("query_duckdb.py")
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .spawn()
        .map_err(|e| format!("Failed to start Python: {}", e))?;

    let input = serde_json::json!({
        "db_path": db_path,
        "query": query
    });

    if let Some(stdin) = child.stdin.as_mut() {
        stdin
            .write_all(input.to_string().as_bytes())
            .map_err(|e| format!("Failed to write to stdin: {}", e))?;
    } else {
        return Err("Could not open stdin".to_string());
    }

    let output = child
        .wait_with_output()
        .map_err(|e| format!("Failed to read output: {}", e))?;

    if output.status.success() {
        Ok(String::from_utf8_lossy(&output.stdout).to_string())
    } else {
        Err(String::from_utf8_lossy(&output.stderr).to_string())
    }
}