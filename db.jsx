"use client";

import { useEffect, useState } from "react";
import { invoke } from "@tauri-apps/api/tauri";

export default function ShowDuckDB() {
  const [data, setData] = useState(null);

  useEffect(() => {
    const fetchData = async () => {
      try {
        const result = await invoke("run_python_duckdb_query", {
          dbPath: "python_scripts/class_data.duckdb",
          query: "SELECT * FROM CLASS_DATA",
        });

        const parsed = typeof result === "string" ? JSON.parse(result) : result;
        setData(parsed);
      } catch (error) {
        console.error("Promise rejected! Here's the full error:", error);

        if (error == null) {
          console.error("Error was null or undefined.");
        } else if (typeof error === "string") {
          if (error.startsWith("{")) {
            try {
              const parsed = JSON.parse(error);
              console.error("Parsed JSON error:", parsed);
            } catch {
              console.error("Could not parse error as JSON:", error);
            }
          } else {
            console.error("Plain string error:", error);
          }
        } else {
          console.error("Unknown error format:", error);
        }
      }
    };

    fetchData();
  }, []);

  return (
    <div>
      <h1>Query Result:</h1>
      <pre>{JSON.stringify(data, null, 2)}</pre>
    </div>
  );
}