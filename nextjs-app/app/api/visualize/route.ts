import { NextRequest } from "next/server";
import fs from "fs";
import os from "os";
import path from "path";
import { spawn } from "child_process";

export const runtime = "nodejs";

function runPython(args: string[]): Promise<{ stdout: string; stderr: string }> {
  return new Promise((resolve, reject) => {
    const child = spawn("python3", args, { stdio: ["ignore", "pipe", "pipe"] });

    let stdout = "";
    let stderr = "";

    child.stdout.on("data", (d) => {
      stdout += d.toString();
    });
    child.stderr.on("data", (d) => {
      stderr += d.toString();
    });

    child.on("close", (code) => {
      if (code === 0) {
        resolve({ stdout, stderr });
      } else {
        reject(new Error(stderr || stdout || `Python exited with code ${code}`));
      }
    });
  });
}

export async function POST(req: NextRequest) {
  const formData = await req.formData();
  const file = formData.get("file");
  const fileType = formData.get("fileType");

  if (!(file instanceof File)) {
    return new Response("file is required", { status: 400 });
  }

  const bytes = Buffer.from(await file.arrayBuffer());
  const workdir = fs.mkdtempSync(path.join(os.tmpdir(), "epoc-"));
  const inputPath = path.join(workdir, file.name);
  fs.writeFileSync(inputPath, bytes);

  // 環境に応じて配置場所を調整してください
  const scriptPath = path.join(process.cwd(), "..", "python", "epoc_auto_visualizer.py");

  const args = [scriptPath, inputPath];
  if (fileType && typeof fileType === "string") {
    args.push("--type", fileType);
  }

  try {
    const { stdout } = await runPython(args);
    const outputLine = stdout.split("\n").find((line) => line.startsWith("output="));
    const jsonOutputLine = stdout.split("\n").find((line) => line.startsWith("json_output="));

    if (!outputLine || !jsonOutputLine) {
      return new Response("output paths not found", { status: 500 });
    }

    const outputPath = outputLine.replace(/^output=/, "").trim();
    const jsonPath = jsonOutputLine.replace(/^json_output=/, "").trim();

    const data = fs.readFileSync(outputPath);
    const jsonData = fs.readFileSync(jsonPath, "utf-8");
    const filename = path.basename(outputPath);

    // Encode excel as base64 to send alongside JSON
    const base64Excel = data.toString("base64");
    const dashboardData = JSON.parse(jsonData);

    return new Response(
      JSON.stringify({
        filename,
        excelBase64: base64Excel,
        dashboardData,
      }),
      {
        status: 200,
        headers: {
          "content-type": "application/json",
        },
      }
    );
  } catch (err) {
    const message = err instanceof Error ? err.message : "unknown error";
    return new Response(message, { status: 500 });
  }
}
