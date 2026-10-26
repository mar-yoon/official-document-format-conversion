"""Build a standalone desktop download on the current operating system.

Install PyInstaller in a temporary build environment first. PyInstaller cannot
cross-compile: run this file on macOS for .app, on Windows for .exe.
"""
from __future__ import annotations

import os
import platform
import plistlib
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
NAME = "公文格式助手"
data_separator = os.pathsep
system = sys.platform
if system not in ("darwin", "win32"):
    raise SystemExit("桌面安装包目前支持 macOS 和 Windows")
BINARY_NAME = "GongwenFormatter" if system == "darwin" else NAME

icon = ROOT / "assets" / ("app.ico" if system == "win32" else "app.icns")
if not icon.exists():
    raise SystemExit("缺少应用图标，请先运行 scripts/create_icon.py")

command = [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "--windowed",
           "--name", BINARY_NAME, "--distpath", str(ROOT / "dist"),
           "--workpath", str(ROOT / ".build" / "pyinstaller"),
           "--specpath", str(ROOT / ".build"),
           "--add-data", f"{ROOT / 'templates'}{data_separator}templates",
           "--icon", str(icon)]
if system == "win32":
    command.append("--onefile")
else:
    command.append("--onedir")
command.append(str(ROOT / "desktop.py"))
build_env = os.environ.copy()
build_env["PYINSTALLER_CONFIG_DIR"] = str(ROOT / ".build" / "pyinstaller-cache")
subprocess.run(command, cwd=ROOT, env=build_env, check=True)

if system == "darwin":
    app = ROOT / "dist" / f"{BINARY_NAME}.app"
    info = app / "Contents" / "Info.plist"
    with info.open("rb") as stream:
        settings = plistlib.load(stream)
    settings["CFBundleDisplayName"] = NAME
    settings["CFBundleName"] = NAME
    with info.open("wb") as stream:
        plistlib.dump(settings, stream)
    subprocess.run(["codesign", "--force", "--deep", "--sign", "-", str(app)], check=True)
    architecture = platform.machine()
    archive = ROOT / "dist" / f"{NAME}-macOS-{architecture}.zip"
    # ditto retains app bundle permissions and symlinks.
    subprocess.run(["ditto", "-c", "-k", "--sequesterRsrc", "--keepParent", str(app), str(archive)], check=True)
    import zipfile
    with zipfile.ZipFile(archive, "a") as zf:
        zf.write(ROOT / "下载使用说明.md", "README.txt")
else:
    import zipfile
    exe = ROOT / "dist" / f"{NAME}.exe"
    archive = ROOT / "dist" / f"{NAME}-Windows.zip"
    with zipfile.ZipFile(archive, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.write(exe, exe.name)
        zf.write(ROOT / "下载使用说明.md", "使用说明.md")
print(f"可分发文件：{archive}")
