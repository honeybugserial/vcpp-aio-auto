#!/usr/bin/env python3
# Visual C++ Redistributable AIO downloader + installer (TechPowerUp package)

import os
import sys
import re
import time
import random
import ctypes
import zipfile
import shutil
import subprocess
import platform
import logging
import argparse

from datetime import datetime
from pathlib import Path

import requests
from rich.console import Console
from rich.panel import Panel
from rich.rule import Rule
from tqdm import tqdm

import warnings
warnings.filterwarnings(
    "ignore",
    category=UserWarning,
    message="pkg_resources is deprecated"
)

import ctypes.wintypes as wt
import pyfiglet

# ------------------------------
# CONFIG
# ------------------------------
US_MIRRORS = ["16", "24", "26", "11", "12", "21", "19", "3", "20"]
TPU_URL = "https://www.techpowerup.com/download/visual-c-redistributable-runtime-package-all-in-one/"

# ------------------------------
# Args
# ------------------------------
def parse_args():
    parser = argparse.ArgumentParser(
        description="Download and install all Visual C++ Redistributable runtimes (TechPowerUp AIO).",
        formatter_class=argparse.RawTextHelpFormatter,
    )

    parser.add_argument(
        "--auto-accept",
        action="store_true",
        help="Run non-interactively (do not prompt for confirmation).",
    )

    parser.add_argument(
        "--dry-run",
        action="store_true",
        help=(
            "Do not execute installers.\n"
            "Still extracts and enumerates installers.\n"
            "Useful for testing output and offline verification."
        ),
    )

    parser.add_argument(
        "--preserve-download",
        action="store_true",
        help=(
            "Preserve the downloaded runtime ZIP instead of deleting it\n"
            "after installation. Local ZIPs are always preserved."
        ),
    )

    return parser.parse_args()

# ------------------------------
# Determine EXE/script directory
# ------------------------------
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent

LOG_DIR = BASE_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

# ------------------------------
# Logging setup
# ------------------------------
console = Console(highlight=False, soft_wrap=False)

def setup_logger():
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_path = LOG_DIR / f"vcredist_install_{ts}.log"

    logger = logging.getLogger("vcredist")
    logger.setLevel(logging.INFO)

    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setFormatter(logging.Formatter(
        "%(asctime)s %(levelname)-8s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    ))

    logger.addHandler(fh)
    return logger, log_path

logger, LOG_PATH = setup_logger()

# ------------------------------
# Console log helpers
# ------------------------------
def info(msg):
    logger.info(msg)
    console.print(f"[cyan][INFO][/cyan]     {msg}")

def success(msg):
    logger.info(f"SUCCESS: {msg}")
    console.print(f"[green][SUCCESS][/green]  {msg}")

def warn(msg):
    logger.warning(msg)
    console.print(f"[yellow][WARNING][/yellow]  {msg}")

def error(msg):
    logger.error(msg)
    console.print(f"[red][ERROR][/red]    {msg}")

def fatal(msg, code=1):
    error(msg)
    sys.exit(code)

def file_fmt(name: str) -> str:
    return f"[bold magenta]{name}[/bold magenta]"

def set_console_size(width=1500, height=600):
    hwnd = ctypes.windll.kernel32.GetConsoleWindow()
    if not hwnd:
        return  # no console

    user32 = ctypes.windll.user32
    user32.SetProcessDPIAware()  # handle high-DPI screens

    # --- Screen resolution ---
    screen_w = user32.GetSystemMetrics(0)  # SM_CXSCREEN
    screen_h = user32.GetSystemMetrics(1)  # SM_CYSCREEN

    # --- Compute centered position ---
    x = int((screen_w - width) / 2)
    y = int((screen_h - height) / 2)

    # --- Move and resize ---
    user32.MoveWindow(hwnd, x, y, width, height, True)

# ------------------------------
# Find local vcredist zip (opt)
# ------------------------------
def find_local_vcredist_zip() -> Path | None:
    zips = sorted(
        BASE_DIR.glob("Visual-C-Runtimes-All-in-One-*.zip"),
        key=lambda p: p.name,
    )
    return zips[-1] if zips else None

# ------------------------------
# Grab Latest TPU Download ID
# ------------------------------
def get_latest_tpu_id() -> str:
    r = requests.get(TPU_URL, timeout=15)
    r.raise_for_status()
    for pat in (
        r'name="id"\s+value="(\d+)"',
        r'download_id\s*=\s*(\d+)',
        r'"id"\s*:\s*"(\d+)"',
    ):
        m = re.search(pat, r.text)
        if m:
            return m.group(1)
    fatal("Unable to locate latest TechPowerUp download ID.")

# ------------------------------
# Download ZIP
# ------------------------------
def download_vcredist() -> Path:
    console.print()
    console.print(Rule("Downloading Visual C++ Runtimes"))

    mirror = random.choice(US_MIRRORS)
    info(f"Using US mirror {mirror}")

    payload = {"id": get_latest_tpu_id(), "server_id": mirror}
    r = requests.post(TPU_URL, data=payload, allow_redirects=False, timeout=15)

    if "Location" not in r.headers:
        fatal("TechPowerUp did not return a redirect")

    url = r.headers["Location"]
    filename = url.split("/")[-1]
    path = BASE_DIR / filename

    success(f"New package found: {file_fmt(filename)}")

    s = requests.get(url, stream=True, timeout=20)
    s.raise_for_status()

    total = int(s.headers.get("Content-Length", 0))

    pbar = None
    if total:
        pbar = tqdm(
            total=total,
            unit="B",
            unit_scale=True,
            ncols=70,
            leave=False,
            file=sys.stderr,
        )

    with open(path, "wb") as f:
        for chunk in s.iter_content(16384):
            if chunk:
                f.write(chunk)
                if pbar:
                    pbar.update(len(chunk))

    if pbar:
        pbar.close()
        # hard reset the line PowerShell thinks it's still on
        sys.stderr.write("\r\033[2K\n")
        sys.stderr.flush()

    success("Download complete")
    console.print()
    return path


# ------------------------------
# Extract ZIP
# ------------------------------
def extract_zip(zip_path: Path) -> Path:
    console.print(Rule("Extracting Package"))

    out_dir = zip_path.with_suffix("")
    if out_dir.exists():
        shutil.rmtree(out_dir, ignore_errors=True)

    with zipfile.ZipFile(zip_path) as z:
        z.extractall(out_dir)

    success(f"Extracted to {out_dir}")
    console.print()
    return out_dir

# ------------------------------
# Install VC++ redistributables
# ------------------------------
def run_vcredists(out_dir: Path, dry_run: bool):
    console.print(Rule("Installing VC++ Runtimes"))

    files = sorted(out_dir.rglob("*.exe"))
    if not files:
        fatal("No redistributable installers found")

    is_x64 = platform.architecture()[0] == "64bit"

    known_versions = ["2005","2008","2010","2012","2013","2015","2017","2019","2022"]
    switches_by_ver = {
        "2005": "/q",
        "2008": "/qb",
        "2010": "/passive /norestart",
        "2012": "/passive /norestart",
        "2013": "/passive /norestart",
        "2015": "/passive /norestart",
        "2017": "/passive /norestart",
        "2019": "/passive /norestart",
        "2022": "/passive /norestart",
    }

    def classify(name):
        lower = name.lower()
        ver = next((v for v in known_versions if v in lower), None)
        arch = "x64" if "x64" in lower else "x86"
        return ver, arch

    for exe in files:
        ver, arch = classify(exe.name)

        if arch == "x64" and not is_x64:
            info(f"Skipping {file_fmt(exe.name)} (x64 on 32-bit OS)")
            continue

        switches = switches_by_ver.get(ver, "/passive /norestart")
        info(f"Installing {file_fmt(exe.name)} ({arch})...")

        if dry_run:
            warn("Dry-run: installer not executed")
            console.print()
            continue

        rc = subprocess.run(
            [str(exe), *switches.split()],
            cwd=str(exe.parent),
            shell=False,
        ).returncode

        if rc in (0, 3010):
            success(f"{exe.name} installed (exit code {rc})")
        else:
            error(f"{exe.name} failed (exit code {rc})")


        console.print()

# ------------------------------
# Cleanup
# ------------------------------
def cleanup(zip_path: Path, extract_dir: Path, delete_zip: bool):
    console.print()
    console.print(Rule("Cleanup"))

    if delete_zip:
        if zip_path.exists():
            try:
                zip_path.unlink()
                success(f"Deleted ZIP: {zip_path.name}")
            except Exception as e:
                warn(f"Could not delete ZIP ({e})")
        else:
            warn(f"ZIP not found: {zip_path.name}")
    else:
        info(f"Preserving local ZIP: {zip_path.name}")

    if extract_dir.exists():
        try:
            shutil.rmtree(extract_dir)
            success(f"Deleted extracted directory: {extract_dir}")
        except Exception as e:
            warn(f"Failed to delete extracted directory ({e})")
    else:
        warn(f"Extracted directory not found: {extract_dir}")


# ------------------------------
# MAIN
# ------------------------------
def main():
    set_console_size(1200, 600)
    args = parse_args()

    auto = args.auto_accept
    dry_run = args.dry_run
    preserve_download = args.preserve_download

    #console.print(Panel("Visual C++ Runtime Auto-Downloader", style="cyan"))
    print(f"     __   __  _______  _______    _       _       ______    _______  ______   ___   _______  _______  ______    ___   _______            ")
    print(f"    |  | |  ||       ||       | _| |_   _| |_    |    _ |  |       ||      | |   | |       ||       ||    _ |  |   | |  _    |           ")
    print(f"    |  |_|  ||       ||    _  ||_   _| |_   _|   |   | ||  |    ___||  _    ||   | |  _____||_     _||   | ||  |   | | |_|   |           ")
    print(f"    |       ||       ||   |_| |  |_|     |_|     |   |_||_ |   |___ | | |   ||   | | |_____   |   |  |   |_||_ |   | |       |           ")
    print(f"    |       ||      _||    ___|                  |    __  ||    ___|| |_|   ||   | |_____  |  |   |  |    __  ||   | |  _   |            ")
    print(f"     |     | |     |_ |   |                      |   |  | ||   |___ |       ||   |  _____| |  |   |  |   |  | ||   | | |_|   |           ")
    print(f"      |___|  |_______||___|                      |___|  |_||_______||______| |___| |_______|  |___|  |___|  |_||___| |_______|           ")
    print(f" _______  ___      ___        _______  _______    ___   __    _  _______  _______  _______  ___      ___      _______  ______    _______ ")
    print(f"|   _   ||   |    |   |      |       ||       |  |   | |  |  | ||       ||       ||   _   ||   |    |   |    |       ||    _ |  |       |")
    print(f"|  |_|  ||   |    |   |      |_     _||   _   |  |   | |   |_| ||  _____||_     _||  |_|  ||   |    |   |    |    ___||   | ||  |  _____|")
    print(f"|       ||   |    |   |        |   |  |  | |  |  |   | |       || |_____   |   |  |       ||   |    |   |    |   |___ |   |_||_ | |_____ ")
    print(f"|       ||   |___ |   |___     |   |  |  |_|  |  |   | |  _    ||_____  |  |   |  |       ||   |___ |   |___ |    ___||    __  ||_____  |")
    print(f"|   _   ||       ||       |    |   |  |       |  |   | | | |   | _____| |  |   |  |   _   ||       ||       ||   |___ |   |  | | _____| |")
    print(f"|__| |__||_______||_______|    |___|  |_______|  |___| |_|  |__||_______|  |___|  |__| |__||_______||_______||_______||___|  |_||_______|")
    console.print(
        "[bold green]Automated Visual C++ Redistributable AIO Downloadering and Installer[/bold green]\n"
    )

    if not auto:
        if console.input("Proceed with download and installation? (Y/N): ").strip().lower() not in ("y","yes"):
            print("User Abortion.")
            time.sleep(1.4)
            sys.exit(0)
    console.print()
    
    local_zip = find_local_vcredist_zip()

    if local_zip:
        success(f"Using local package: {file_fmt(local_zip.name)}")
        zip_file = local_zip
        delete_zip = False
    else:
        if dry_run:
            info("Dry-run: no local package found, downloading once to seed cache")
        zip_file = download_vcredist()
        delete_zip = not preserve_download


    out_dir = extract_zip(zip_file)
    run_vcredists(out_dir, dry_run)
    cleanup(zip_file, out_dir, delete_zip)
    info(f"Log written to {LOG_PATH}")
    console.print()

    console.print(Rule("COMPLETED"))

    console.print(Panel("\nInstallation completed successfully.", style="green\n"))

    if not auto:
        input("\nPress BUTTONS to exit.\n\n")

if __name__ == "__main__":
    main()
