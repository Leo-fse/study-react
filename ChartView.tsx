'use client';
import { useEffect, useState } from "react";

export default function ChartView() {
  const [files, setFiles] = useState<string[]>([]);
  const [selected, setSelected] = useState<string>("");

  useEffect(() => {
    fetch("/api/chart-list")
      .then((res) => res.json())
      .then((data) => {
        setFiles(data.files);
        setSelected(data.files[0]);
      });
  }, []);

  return (
    <div style={{ padding: "1rem" }}>
      <select value={selected} onChange={(e) => setSelected(e.target.value)}>
        {files.map((file) => (
          <option key={file} value={file}>
            {file}
          </option>
        ))}
      </select>
      <div style={{ marginTop: "1rem", height: "80vh" }}>
        <iframe
          src={`/chart/${selected}`}
          width="100%"
          height="100%"
          style={{ border: "none" }}
        />
      </div>
    </div>
  );
}