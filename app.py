#!/usr/bin/env python3
from __future__ import annotations

import argparse
import html
import json
import os
import posixpath
import re
import socket
import subprocess
import sys
import threading
import time
import unicodedata
import webbrowser
import zipfile
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse


APP_TITLE = "Mac to Windows ZIP Cleaner"
DEFAULT_PORT = 8765
WINDOWS_INVALID_CHARS_RE = re.compile(r'[<>:"|?*\x00-\x1f]')
WINDOWS_RESERVED_NAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}
MACOS_METADATA_NAMES = {".DS_Store", "__MACOSX"}


def default_output_dir() -> Path:
    desktop = Path.home() / "Desktop"
    localized_desktop = desktop / "\ub370\uc2a4\ud06c\ud0d1"
    return localized_desktop if localized_desktop.exists() else desktop


def nfc(text: str) -> str:
    return unicodedata.normalize("NFC", text)


def should_skip_arcname(arcname: str) -> bool:
    parts = [p for p in arcname.split("/") if p]
    if not parts:
        return True
    return any(part in MACOS_METADATA_NAMES or part.startswith("._") for part in parts)


def windows_safe_part(part: str) -> str:
    part = WINDOWS_INVALID_CHARS_RE.sub("_", nfc(part))
    part = part.rstrip(" .")
    if not part:
        part = "_"

    stem = part.split(".", 1)[0].upper()
    if stem in WINDOWS_RESERVED_NAMES:
        part = f"_{part}"
    return part


def safe_arcname(name: str) -> str:
    name = nfc(name).replace("\\", "/").replace("\x00", "")
    name = re.sub(r"^[A-Za-z]:(?=/|$)", "", name)
    name = posixpath.normpath(name).lstrip("/")
    parts = []
    for part in name.split("/"):
        if part in {"", ".", ".."}:
            continue
        parts.append(windows_safe_part(part))
    return "/".join(parts)


def safe_zip_filename(name: str) -> str:
    name = nfc(name.strip()) or "windows_compatible.zip"
    name = name.replace("/", "_").replace("\\", "_").replace("\x00", "")
    name = windows_safe_part(name)
    if not name.lower().endswith(".zip"):
        name += ".zip"
    return name


def unique_output_path(path: Path, overwrite: bool) -> Path:
    if overwrite or not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    index = 1
    while True:
        candidate = parent / f"{stem}_{index}{suffix}"
        if not candidate.exists():
            return candidate
        index += 1


def unique_arcname(arcname: str, used: set[str]) -> str:
    if arcname not in used:
        used.add(arcname)
        return arcname

    parent, base = posixpath.split(arcname)
    stem, ext = posixpath.splitext(base)
    index = 1
    while True:
        candidate_base = f"{stem}_{index}{ext}"
        candidate = posixpath.join(parent, candidate_base) if parent else candidate_base
        if candidate not in used:
            used.add(candidate)
            return candidate
        index += 1


def write_bytes_to_zip(zf: zipfile.ZipFile, arcname: str, data: bytes) -> None:
    info = zipfile.ZipInfo(arcname)
    info.date_time = time.localtime()[:6]
    info.compress_type = zipfile.ZIP_DEFLATED
    info.flag_bits |= 0x800
    zf.writestr(info, data)


def write_directory_to_zip(zf: zipfile.ZipFile, arcname: str, used: set[str]) -> None:
    arcname = safe_arcname(arcname).rstrip("/") + "/"
    if should_skip_arcname(arcname) or arcname in used:
        return

    used.add(arcname)
    info = zipfile.ZipInfo(arcname)
    info.date_time = time.localtime()[:6]
    info.external_attr = 0o40755 << 16
    info.flag_bits |= 0x800
    zf.writestr(info, b"")


def add_path_to_zip(zf: zipfile.ZipFile, source: Path, used: set[str]) -> int:
    source = source.expanduser().resolve()
    count = 0

    if source.is_file():
        arcname = safe_arcname(source.name)
        if not should_skip_arcname(arcname):
            zf.write(source, unique_arcname(arcname, used))
            count += 1
        return count

    if not source.is_dir():
        raise FileNotFoundError(f"Input path was not found: {source}")

    root_name = safe_arcname(source.name)
    write_directory_to_zip(zf, root_name, used)
    for path in source.rglob("*"):
        rel = path.relative_to(source).as_posix()
        arcname = safe_arcname(f"{root_name}/{rel}")
        if should_skip_arcname(arcname):
            continue
        if path.is_dir():
            write_directory_to_zip(zf, arcname, used)
            continue
        if not path.is_file():
            continue
        zf.write(path, unique_arcname(arcname, used))
        count += 1
    return count


def build_zip_from_paths(source_path: str, output_dir: str, zip_name: str, overwrite: bool) -> tuple[Path, int]:
    if not source_path.strip():
        raise ValueError("Select a file or folder to compress.")

    source = Path(nfc(source_path)).expanduser()
    out_dir = Path(nfc(output_dir)).expanduser()
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = unique_output_path(out_dir / safe_zip_filename(zip_name), overwrite)

    used: set[str] = set()
    with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        count = add_path_to_zip(zf, source, used)
    return out_path, count


def parse_multipart(content_type: str, body: bytes) -> list[tuple[str, str | None, bytes]]:
    raw = (
        f"Content-Type: {content_type}\r\n"
        "MIME-Version: 1.0\r\n\r\n"
    ).encode("utf-8") + body
    message = BytesParser(policy=default).parsebytes(raw)
    items: list[tuple[str, str | None, bytes]] = []
    for part in message.iter_parts():
        if part.get_content_disposition() != "form-data":
            continue
        name = part.get_param("name", header="content-disposition")
        filename = part.get_filename()
        payload = part.get_payload(decode=True) or b""
        items.append((name or "", filename, payload))
    return items


def field_text(payload: bytes) -> str:
    return payload.decode("utf-8", errors="replace")


def build_zip_from_upload(items: list[tuple[str, str | None, bytes]]) -> tuple[Path, int]:
    fields: dict[str, str] = {}
    files: list[tuple[str, bytes]] = []

    for name, filename, payload in items:
        if filename is None:
            fields[name] = field_text(payload)
        elif name == "files":
            files.append((filename, payload))

    if not files:
        raise ValueError("No files were selected.")

    out_dir = Path(nfc(fields.get("output_dir", str(default_output_dir())))).expanduser()
    out_dir.mkdir(parents=True, exist_ok=True)
    zip_name = fields.get("zip_name", "windows_compatible.zip")
    overwrite = fields.get("overwrite") == "true"
    out_path = unique_output_path(out_dir / safe_zip_filename(zip_name), overwrite)

    used: set[str] = set()
    count = 0
    with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for filename, payload in files:
            arcname = safe_arcname(filename)
            if should_skip_arcname(arcname):
                continue
            write_bytes_to_zip(zf, unique_arcname(arcname, used), payload)
            count += 1

    return out_path, count


def json_bytes(data: dict) -> bytes:
    return json.dumps(data, ensure_ascii=False).encode("utf-8")


def choose_path_with_osascript(kind: str) -> str:
    prompts = {
        "source_file": "Select a file to compress.",
        "source_folder": "Select a folder to compress.",
        "output_folder": "Select the output folder for the ZIP file.",
    }
    prompt = prompts.get(kind, "Select a path.")
    if kind == "source_file":
        command = f'POSIX path of (choose file with prompt "{prompt}")'
    else:
        command = f'POSIX path of (choose folder with prompt "{prompt}")'

    completed = subprocess.run(
        ["osascript", "-e", command],
        capture_output=True,
        text=True,
        check=False,
    )
    if completed.returncode != 0:
        if "User canceled" in completed.stderr:
            return ""
        raise RuntimeError(completed.stderr.strip() or "Could not open the path picker.")
    return completed.stdout.strip()


def choose_path_with_tk(kind: str) -> str:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        if kind == "source_file":
            return filedialog.askopenfilename(title="Select a file to compress")
        if kind in {"source_folder", "output_folder"}:
            title = "Select the ZIP output folder" if kind == "output_folder" else "Select a folder to compress"
            return filedialog.askdirectory(title=title)
        raise ValueError("Unknown picker type.")
    finally:
        root.destroy()


def choose_local_path(kind: str) -> str:
    if kind not in {"source_file", "source_folder", "output_folder"}:
        raise ValueError("Unknown picker type.")
    if sys.platform == "darwin":
        return choose_path_with_osascript(kind)
    return choose_path_with_tk(kind)


def html_page() -> str:
    output = "~/Desktop"
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{APP_TITLE}</title>
  <style>
    :root {{
      --bg: #f6f1e7;
      --ink: #1f2933;
      --card: #fffaf0;
      --line: #d9cdb8;
      --accent: #1f6f5b;
      --accent2: #d97941;
      --muted: #6b7280;
    }}
    body {{
      margin: 0;
      background: radial-gradient(circle at 20% 10%, #ffe0b9 0, transparent 32%),
                  linear-gradient(135deg, #f6f1e7, #e7f0ea);
      color: var(--ink);
      font-family: "Apple SD Gothic Neo", "Noto Sans KR", sans-serif;
    }}
    main {{
      max-width: 900px;
      margin: 48px auto;
      padding: 0 20px 48px;
    }}
    h1 {{
      margin: 0 0 8px;
      font-size: 36px;
      letter-spacing: -0.04em;
    }}
    .lead {{
      margin: 0 0 28px;
      color: var(--muted);
      line-height: 1.6;
    }}
    .grid {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 18px;
    }}
    .card {{
      background: color-mix(in srgb, var(--card) 92%, white);
      border: 1px solid var(--line);
      border-radius: 22px;
      padding: 22px;
      box-shadow: 0 18px 40px rgba(72, 53, 27, 0.12);
    }}
    .card.full {{
      grid-column: 1 / -1;
    }}
    h2 {{
      margin: 0 0 12px;
      font-size: 20px;
    }}
    label {{
      display: block;
      margin: 14px 0 6px;
      font-weight: 700;
    }}
    input[type="text"] {{
      width: 100%;
      box-sizing: border-box;
      border: 1px solid var(--line);
      border-radius: 12px;
      padding: 12px 13px;
      font-size: 15px;
      background: white;
    }}
    input[type="file"] {{
      width: 100%;
      padding: 12px;
      box-sizing: border-box;
      border: 1px dashed var(--line);
      border-radius: 12px;
      background: #fffdf8;
    }}
    button {{
      border: 0;
      border-radius: 14px;
      padding: 13px 18px;
      font-weight: 800;
      font-size: 15px;
      color: white;
      background: var(--accent);
      cursor: pointer;
    }}
    button.secondary {{
      background: var(--accent2);
    }}
    button:disabled {{
      opacity: 0.6;
      cursor: wait;
    }}
    .row {{
      display: flex;
      gap: 10px;
      align-items: center;
      flex-wrap: wrap;
      margin-top: 16px;
    }}
    .path-row {{
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto auto;
      gap: 10px;
      align-items: center;
    }}
    .path-row.output {{
      grid-template-columns: minmax(0, 1fr) auto;
    }}
    .path-row button {{
      white-space: nowrap;
      padding-left: 15px;
      padding-right: 15px;
    }}
    .hint {{
      color: var(--muted);
      font-size: 13px;
      line-height: 1.5;
    }}
    .result {{
      white-space: pre-wrap;
      border-radius: 16px;
      padding: 14px;
      background: #0f172a;
      color: #d1fae5;
      min-height: 46px;
      line-height: 1.5;
    }}
    @media (max-width: 760px) {{
      .grid {{ grid-template-columns: 1fr; }}
      .path-row,
      .path-row.output {{ grid-template-columns: 1fr; }}
      h1 {{ font-size: 30px; }}
    }}
  </style>
</head>
<body>
<main>
  <h1>Mac to Windows ZIP Cleaner</h1>
  <p class="lead">Create Windows-friendly ZIP archives by normalizing internal file names and removing common macOS metadata.</p>

  <div class="grid">
    <section class="card">
      <h2>1. Source</h2>
      <label>File or folder path</label>
      <div class="path-row">
        <input id="sourcePath" type="text" placeholder="~/project-folder">
        <button id="chooseFileButton" type="button" class="secondary">Choose File</button>
        <button id="chooseFolderButton" type="button" class="secondary">Choose Folder</button>
      </div>
      <p class="hint">Use the picker buttons or paste a path manually.</p>
    </section>

    <section class="card">
      <h2>2. Output</h2>
      <label>ZIP output folder</label>
      <div class="path-row output">
        <input id="outputDir" type="text" value="{output}">
        <button id="chooseOutputButton" type="button" class="secondary">Choose Output Folder</button>
      </div>
      <label>ZIP file name</label>
      <input id="zipName" type="text" value="package.zip">
      <div class="row">
        <label style="margin:0;font-weight:600;">
          <input id="overwrite" type="checkbox"> Overwrite if the file already exists
        </label>
      </div>
    </section>

    <section class="card full">
      <h2>3. Create ZIP</h2>
      <p class="hint">Compress the selected file or folder. `.DS_Store` and `__MACOSX` entries are skipped automatically.</p>
      <div class="row">
        <button id="zipButton">Create ZIP</button>
      </div>
    </section>

    <section class="card full">
      <h2>Result</h2>
      <div id="result" class="result">Ready</div>
    </section>
  </div>
</main>

<script>
const result = document.getElementById("result");
const zipButton = document.getElementById("zipButton");
const chooseFileButton = document.getElementById("chooseFileButton");
const chooseFolderButton = document.getElementById("chooseFolderButton");
const chooseOutputButton = document.getElementById("chooseOutputButton");

function setBusy(busy) {{
  zipButton.disabled = busy;
  chooseFileButton.disabled = busy;
  chooseFolderButton.disabled = busy;
  chooseOutputButton.disabled = busy;
}}

function maskHomePath(path) {{
  return String(path)
    .replace(/^\\/Users\\/[^/]+(?=\\/|$)/, "~")
    .replace(/^[A-Za-z]:\\\\Users\\\\[^\\\\]+(?=\\\\|$)/, "~");
}}

function show(data) {{
  if (data.ok) {{
    result.textContent = "Done\\nFiles: " + data.count + "\\nSaved to: " + maskHomePath(data.path);
  }} else {{
    result.textContent = "Error\\n" + data.error;
  }}
}}

async function choosePath(kind, targetId) {{
  setBusy(true);
  result.textContent = "Opening picker...";
  try {{
    const res = await fetch("/api/choose?kind=" + encodeURIComponent(kind));
    const data = await res.json();
    if (data.ok) {{
      if (data.path) {{
        const safePath = maskHomePath(data.path);
        document.getElementById(targetId).value = safePath;
        result.textContent = "Selected\\n" + safePath;
      }} else {{
        result.textContent = "Selection canceled";
      }}
    }} else {{
      show(data);
    }}
  }} catch (err) {{
    show({{ok: false, error: String(err)}});
  }} finally {{
    setBusy(false);
  }}
}}

async function createZip() {{
  const payload = {{
    source_path: document.getElementById("sourcePath").value,
    output_dir: document.getElementById("outputDir").value,
    zip_name: document.getElementById("zipName").value,
    overwrite: document.getElementById("overwrite").checked
  }};
  if (!payload.source_path.trim()) {{
    show({{ok: false, error: "Enter a path to compress."}});
    return;
  }}
  setBusy(true);
  result.textContent = "Creating ZIP...";
  try {{
    const res = await fetch("/api/path", {{
      method: "POST",
      headers: {{"Content-Type": "application/json"}},
      body: JSON.stringify(payload)
    }});
    show(await res.json());
  }} catch (err) {{
    show({{ok: false, error: String(err)}});
  }} finally {{
    setBusy(false);
  }}
}}

chooseFileButton.addEventListener("click", () => choosePath("source_file", "sourcePath"));
chooseFolderButton.addEventListener("click", () => choosePath("source_folder", "sourcePath"));
chooseOutputButton.addEventListener("click", () => choosePath("output_folder", "outputDir"));
zipButton.addEventListener("click", createZip);
</script>
</body>
</html>"""


class Handler(BaseHTTPRequestHandler):
    server_version = "MacToWindowsZipCleaner/1.0"

    def log_message(self, fmt: str, *args) -> None:
        print("%s - %s" % (self.address_string(), fmt % args))

    def send_bytes(self, status: int, content_type: str, body: bytes) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_json(self, data: dict, status: int = 200) -> None:
        self.send_bytes(status, "application/json; charset=utf-8", json_bytes(data))

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = unquote(parsed.path)
        if path == "/":
            self.send_bytes(200, "text/html; charset=utf-8", html_page().encode("utf-8"))
            return
        if path == "/api/choose":
            try:
                query = parse_qs(parsed.query)
                kind = query.get("kind", [""])[0]
                selected = choose_local_path(kind)
                self.send_json({"ok": True, "path": selected})
            except Exception as exc:
                self.send_json({"ok": False, "error": str(exc)}, 400)
            return
        self.send_bytes(404, "text/plain; charset=utf-8", "Not found".encode("utf-8"))

    def do_POST(self) -> None:
        try:
            length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(length)

            if self.path == "/api/upload":
                content_type = self.headers.get("Content-Type", "")
                items = parse_multipart(content_type, body)
                out_path, count = build_zip_from_upload(items)
                self.send_json({"ok": True, "path": str(out_path), "count": count})
                return

            if self.path == "/api/path":
                payload = json.loads(body.decode("utf-8"))
                out_path, count = build_zip_from_paths(
                    payload.get("source_path", ""),
                    payload.get("output_dir", str(default_output_dir())),
                    payload.get("zip_name", "windows_compatible.zip"),
                    bool(payload.get("overwrite", False)),
                )
                self.send_json({"ok": True, "path": str(out_path), "count": count})
                return

            self.send_json({"ok": False, "error": "Unknown request."}, 404)
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, 400)


def choose_port(preferred: int) -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        try:
            sock.bind(("127.0.0.1", preferred))
            return preferred
        except OSError:
            sock.bind(("127.0.0.1", 0))
            return int(sock.getsockname()[1])


def main() -> None:
    parser = argparse.ArgumentParser(description="Create Windows-friendly ZIP archives from a local web UI.")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    args = parser.parse_args()

    port = choose_port(args.port)
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = f"http://127.0.0.1:{port}/"

    if not os.environ.get("NO_BROWSER"):
        threading.Timer(0.5, lambda: webbrowser.open(url)).start()
    print(f"{APP_TITLE} running at: {url}")
    print("Press Ctrl+C in this terminal to stop.")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down.")


if __name__ == "__main__":
    main()
