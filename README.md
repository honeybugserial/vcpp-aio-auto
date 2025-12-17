# Visual C++ Redistributable AIO Auto-Installer

Automated downloader and installer for **all Microsoft Visual C++ Redistributable runtimes (2005–2022)** using the official **TechPowerUp Visual C++ Runtime Package (AIO)**.

This script is designed to be **drop-in**, **portable**, and **hands-off**, with logging, silent installs, and sane defaults.

---
## Download

- **Compiled binary:** [vcpp-aio-auto.exe](https://archive.org/download/vcpp-aio-auto/vcpp-aio-auto.exe)


## Features

- Downloads the **latest** TechPowerUp VC++ AIO package automatically
- Supports **2005 → 2022** redistributables
- Handles **x86 / x64** correctly based on OS architecture
- Silent / passive installs with no forced reboots
- Uses **local ZIP if present** (offline-friendly)
- Optional **non-interactive mode**
- **Dry-run mode** for testing / verification

---

## Requirements

- Windows
- Python **3.9+**
- Internet access (unless using a local ZIP)

---

## Usage

### Basic (interactive)
```bash
python vcpp-aio-auto.py
```

You will be prompted before download and installation.

---

### Fully automatic (no prompts)
```bash
python vcpp-aio-auto.py --auto-accept
```

---

### Dry-run (no installers executed)
```bash
python vcpp-aio-auto.py --dry-run
```

- Downloads (if needed)
- Extracts
- Enumerates installers
- **Does not execute anything**

---

### Preserve downloaded ZIP
```bash
python vcpp-aio-auto.py --preserve-download
```

By default, downloaded ZIPs are deleted after install unless this flag is set.

---

## Offline / Local Package Support

If a file matching:
```
Visual-C-Runtimes-All-in-One-*.zip
```
exists in the same directory as the script, it will be **used automatically** and **no download will occur**.


---


## Notes

- x64 redistributables are skipped automatically on 32-bit Windows
- Install switches are version-appropriate (`/q`, `/qb`, `/passive /norestart`)
- Uses official TechPowerUp mirrors only

---

## Disclaimer

No binaries are modified, repackaged, or redistributed by this script.

---

## License

Maybe
