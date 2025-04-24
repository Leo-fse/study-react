// pages/chart-view.tsx
import { useState } from "react";
import path from "path";
import fs from "fs";

type Props = {
  files: string[];
};

export async function getStaticProps() {
  const chartDir = path.join(process.cwd(), "public", "chart");
  const files = fs
    .readdirSync(chartDir)
    .filter((file) => file.endsWith(".html"));

  return { props: { files } };
}

export default function ChartViewer({ files }: Props) {
  const [selected, setSelected] = useState(files[0] || "");

  return (
    <div style={{ padding: "1rem" }}>
      <label>
        表示するグラフ：
        <select
          value={selected}
          onChange={(e) => setSelected(e.target.value)}
          style={{ marginLeft: "1rem" }}
        >
          {files.map((file) => (
            <option key={file} value={file}>
              {file}
            </option>
          ))}
        </select>
      </label>

      <div style={{ marginTop: "1rem", height: "80vh" }}>
        {selected && (
          <iframe
            src={`/chart/${selected}`}
            width="100%"
            height="100%"
            style={{ border: "1px solid #ccc" }}
            title="Bokeh Chart"
          />
        )}
      </div>
    </div>
  );
}