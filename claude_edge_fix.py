"""
claude_edge_fix.py
Diagnostic and repair tool for the Claude browser extension in Microsoft Edge.

Usage:
    py claude_edge_fix.py            # diagnostics only
    py claude_edge_fix.py --fix      # diagnostics + auto-fix

Log:  C:\Users\jaa15\OneDrive\PYProjects\Logs\claude_edge_fix_YYYYMMDD_HHMMSS.log
Repo: https://github.com/unsungIMG/ClaudeEdgeFix
"""

import argparse
import json
import os
import subprocess
import winreg
from datetime import datetime
from pathlib import Path

LOG_DIR = Path(r"C:\Users\jaa15\OneDrive\PYProjects\Logs")
DATESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")
LOG_PATH = LOG_DIR / f"claude_edge_fix_{DATESTAMP}.log"

NMH_NAME_EXT  = "com.anthropic.claude_browser_extension"
NMH_NAME_CODE = "com.anthropic.claude_code_browser_extension"
REG_EDGE_BASE   = r"SOFTWARE\Microsoft\Edge\NativeMessagingHosts"
REG_CHROME_BASE = r"SOFTWARE\Google\Chrome\NativeMessagingHosts"

CLAUDE_EXT_IDS = {
    "dihbgbndebgnbjfmelmegjepbnkhlgni",
    "fcoeoabgfenejglbffodgkkbkcdhcgfn",
    "dngcpimnedloihjnnfngkgjoidhnaolf",
    "npdliodflgoeipmhfodjpjbflbmclnid",
}

CWS_CLAUDE_URL = (
    "https://chromewebstore.google.com/detail/claude/"
    "npdliodflgoeipmhfodjpjbflbmclnid"
)

EDGE_USER_DATA   = Path(os.environ["LOCALAPPDATA"]) / "Microsoft" / "Edge" / "User Data"
CHROME_USER_DATA = Path(os.environ["LOCALAPPDATA"]) / "Google" / "Chrome" / "User Data"
APPDATA_CLAUDE   = Path(os.environ["APPDATA"]) / "Claude"
LOCAL_CLAUDE     = Path(os.environ["LOCALAPPDATA"]) / "Claude"

_log_lines: list[str] = []

def _emit(msg: str) -> None:
    ts = datetime.now().strftime("%H:%M:%S")
    _log_lines.append(f"[{ts}] {msg}")
    print(msg)

def flush_log() -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    LOG_PATH.write_text("\n".join(_log_lines), encoding="utf-8")

def header(title: str) -> None:
    bar = "-" * 60
    _emit(f"\n{bar}\n  {title}\n{bar}")

def ok(label: str, detail: str = "")   -> None: _emit(f"  OK    {label}" + (f"  ({detail})" if detail else ""))
def fail(label: str, detail: str = "") -> None: _emit(f"  FAIL  {label}" + (f"  ({detail})" if detail else ""))
def warn(label: str, detail: str = "") -> None: _emit(f"  WARN  {label}" + (f"  ({detail})" if detail else ""))
def info(label: str, detail: str = "") -> None: _emit(f"  INFO  {label}" + (f"  -> {detail}" if detail else ""))

def reg_read(hive: int, key_path: str, value_name: str = "") -> str | None:
    try:
        with winreg.OpenKey(hive, key_path) as k:
            val, _ = winreg.QueryValueEx(k, value_name)
            return val
    except OSError:
        return None

def reg_write(hive: int, key_path: str, value_name: str, value: str) -> bool:
    try:
        winreg.CreateKey(hive, key_path)
        with winreg.OpenKey(hive, key_path, 0, winreg.KEY_SET_VALUE) as k:
            winreg.SetValueEx(k, value_name, 0, winreg.REG_SZ, value)
        return True
    except OSError as e:
        warn("Registry write failed", str(e))
        return False

def find_nmh_manifest() -> Path | None:
    for reg_base in (REG_EDGE_BASE, REG_CHROME_BASE):
        val = reg_read(winreg.HKEY_CURRENT_USER, rf"{reg_base}\{NMH_NAME_EXT}")
        if val:
            p = Path(val)
            if p.exists():
                return p
    fallback = APPDATA_CLAUDE / "ChromeNativeHost" / f"{NMH_NAME_EXT}.json"
    return fallback if fallback.exists() else None

def get_allowed_origins(manifest_path: Path) -> set[str]:
    try:
        data = json.loads(manifest_path.read_text(encoding="utf-8"))
        ids = set()
        for o in data.get("allowed_origins", []):
            ids.add(o.replace("chrome-extension://", "").strip("/"))
        return ids
    except Exception:
        return set()

def scan_profiles(user_data: Path, ext_ids: set[str]) -> list[tuple[str, str]]:
    hits: list[tuple[str, str]] = []
    if not user_data.exists():
        return hits
    for profile_dir in user_data.iterdir():
        if not profile_dir.is_dir():
            continue
        ext_base = profile_dir / "Extensions"
        if not ext_base.exists():
            continue
        for installed_id in ext_base.iterdir():
            if installed_id.name in ext_ids:
                hits.append((profile_dir.name, installed_id.name))
    return hits

def process_running(name: str) -> bool:
    try:
        out = subprocess.check_output(
            ["tasklist", "/FI", f"IMAGENAME eq {name}", "/NH"],
            text=True, stderr=subprocess.DEVNULL
        )
        return name.lower() in out.lower()
    except Exception:
        return False

def run_diagnostics() -> dict:
    header("PHASE 1 -- DIAGNOSTICS")
    results: dict = {}
    total = 10
    ext_ids = set(CLAUDE_EXT_IDS)

    _emit(f"\nProcessing 1/{total}: Edge NMH registry ({NMH_NAME_EXT})")
    val = reg_read(winreg.HKEY_CURRENT_USER, rf"{REG_EDGE_BASE}\{NMH_NAME_EXT}")
    if val:  ok("Edge NMH registry key present", val);   results["edge_nmh_reg"] = True
    else:    fail("Edge NMH registry key MISSING");      results["edge_nmh_reg"] = False

    _emit(f"\nProcessing 2/{total}: Edge NMH registry ({NMH_NAME_CODE})")
    val = reg_read(winreg.HKEY_CURRENT_USER, rf"{REG_EDGE_BASE}\{NMH_NAME_CODE}")
    if val:  ok("Edge NMH (Claude Code) present", val);                              results["edge_code_nmh_reg"] = True
    else:    warn("Edge NMH (Claude Code) missing -- only needed for Claude Code");  results["edge_code_nmh_reg"] = False

    _emit(f"\nProcessing 3/{total}: Chrome NMH registry (reference)")
    val = reg_read(winreg.HKEY_CURRENT_USER, rf"{REG_CHROME_BASE}\{NMH_NAME_EXT}")
    if val:  ok("Chrome NMH registry present", val);  results["chrome_nmh_reg"] = True
    else:    warn("Chrome NMH registry absent");       results["chrome_nmh_reg"] = False

    _emit(f"\nProcessing 4/{total}: NMH JSON manifest file")
    manifest_path = find_nmh_manifest()
    if manifest_path:
        ok("NMH JSON manifest found", str(manifest_path))
        results["manifest_exists"] = True
        results["_manifest_path"] = str(manifest_path)
    else:
        fail("NMH JSON manifest NOT FOUND")
        results["manifest_exists"] = False
        results["_manifest_path"] = None

    _emit(f"\nProcessing 5/{total}: NMH binary path in manifest")
    if manifest_path:
        try:
            data = json.loads(manifest_path.read_text(encoding="utf-8"))
            bin_path = Path(data.get("path", ""))
            if bin_path.exists():
                ok("Native host binary exists", str(bin_path))
                results["manifest_binary"] = True
            else:
                fail("Native host binary MISSING", str(bin_path))
                results["manifest_binary"] = False
            origins = get_allowed_origins(manifest_path)
            if origins:
                ext_ids.update(origins)
                info("allowed_origins harvested", ", ".join(sorted(origins)))
        except Exception as e:
            fail("Could not parse manifest", str(e))
            results["manifest_binary"] = False
    else:
        warn("Skipping binary check -- manifest not found")
        results["manifest_binary"] = None

    _emit(f"\nProcessing 6/{total}: Claude extension in Edge profiles")
    edge_hits = scan_profiles(EDGE_USER_DATA, ext_ids)
    if edge_hits:
        for p, e in edge_hits: ok(f"Claude extension found in Edge/{p}", e)
        results["edge_ext_installed"] = True
    else:
        fail("Claude extension NOT installed in any Edge profile")
        results["edge_ext_installed"] = False

    _emit(f"\nProcessing 7/{total}: Claude extension in Chrome profiles (reference)")
    chrome_hits = scan_profiles(CHROME_USER_DATA, ext_ids)
    if chrome_hits:
        for p, e in chrome_hits: ok(f"Claude extension found in Chrome/{p}", e)
        results["chrome_ext_installed"] = True
    else:
        warn("Claude extension not found in Chrome profiles")
        results["chrome_ext_installed"] = False

    _emit(f"\nProcessing 8/{total}: Claude Desktop installation")
    claude_paths = [
        LOCAL_CLAUDE / "Claude.exe",
        Path(os.environ["LOCALAPPDATA"]) / "AnthropicClaude" / "Claude.exe",
    ]
    found = next((p for p in claude_paths if p.exists()), None)
    if found:  ok("Claude Desktop installed", str(found));  results["claude_desktop"] = True
    else:      warn("Claude Desktop exe not found");        results["claude_desktop"] = False

    _emit(f"\nProcessing 9/{total}: Claude Desktop process")
    if process_running("Claude.exe"):  ok("Claude Desktop is running");                            results["claude_running"] = True
    else:                              warn("Claude Desktop NOT running -- NMH needs it active");  results["claude_running"] = False

    _emit(f"\nProcessing 10/{total}: Edge browser process")
    if process_running("msedge.exe"):  ok("Microsoft Edge is running");   results["edge_running"] = True
    else:                              warn("Microsoft Edge not running"); results["edge_running"] = False

    results["_ext_ids"] = ext_ids
    return results

def print_summary(results: dict) -> str:
    header("PHASE 2 -- SUMMARY")
    issues = []
    if not results.get("edge_nmh_reg"):         issues.append("nmh");  fail("Edge NMH registry key missing")
    if not results.get("manifest_exists"):       issues.append("mani"); fail("NMH JSON manifest file missing")
    if results.get("manifest_binary") is False:  issues.append("bin");  fail("Native host binary missing")
    if not results.get("edge_ext_installed"):    issues.append("ext");  fail("Claude extension NOT installed in Edge  <- PRIMARY ISSUE")
    if not results.get("claude_running"):        issues.append("proc"); warn("Claude Desktop is not running")
    if not issues:
        ok("All checks passed -- if still broken, restart Edge")
        return "clean"
    _emit("\n  Root cause diagnosis:")
    if "ext" in issues:
        _emit("  -> Extension not installed in Edge.")
        _emit(f"     Install from: {CWS_CLAUDE_URL}")
    if "nmh" in issues:
        _emit("  -> Edge NMH registry entry missing. Run with --fix.")
    if "mani" in issues or "bin" in issues:
        _emit("  -> NMH manifest/binary missing. Reinstall Claude Desktop.")
    if "proc" in issues:
        _emit("  -> Start Claude Desktop before connecting the extension.")
    return "issues"

def run_fixes(results: dict) -> None:
    header("PHASE 3 -- AUTO-FIX")
    fixed_any = False
    total = 3

    _emit(f"\nProcessing 1/{total}: Edge NMH registry")
    if not results.get("edge_nmh_reg"):
        src = results.get("_manifest_path") or reg_read(
            winreg.HKEY_CURRENT_USER, rf"{REG_CHROME_BASE}\{NMH_NAME_EXT}")
        if src:
            if reg_write(winreg.HKEY_CURRENT_USER, rf"{REG_EDGE_BASE}\{NMH_NAME_EXT}", "", src):
                ok("Created Edge NMH registry key", src); fixed_any = True
            else:
                fail("Could not write Edge NMH registry key")
        else:
            fail("No manifest source -- reinstall Claude Desktop")
    else:
        info("Edge NMH already present -- skip")

    _emit(f"\nProcessing 2/{total}: Edge NMH (Claude Code)")
    if not results.get("edge_code_nmh_reg"):
        src = reg_read(winreg.HKEY_CURRENT_USER, rf"{REG_CHROME_BASE}\{NMH_NAME_CODE}")
        if src:
            if reg_write(winreg.HKEY_CURRENT_USER, rf"{REG_EDGE_BASE}\{NMH_NAME_CODE}", "", src):
                ok("Created Edge NMH (Claude Code) key", src); fixed_any = True
        else:
            info("Claude Code NMH not in Chrome -- skip")
    else:
        info("Edge NMH (Claude Code) already present -- skip")

    _emit(f"\nProcessing 3/{total}: Extension installation in Edge")
    if not results.get("edge_ext_installed"):
        _emit("\n  Opening Edge to Chrome Web Store install page...")
        try:
            subprocess.Popen(
                ["cmd", "/c", "start", "msedge", CWS_CLAUDE_URL],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
            )
            ok("Opened Edge -> Chrome Web Store")
            _emit("  ACTION: click 'Add to Chrome' in Edge, then sign in with")
            _emit("  the same account used for Claude Desktop.")
            fixed_any = True
        except Exception as e:
            fail("Could not open Edge", str(e))
    else:
        info("Extension already installed -- skip")

    _emit("")
    if fixed_any: ok("Fixes applied -- restart Edge + Claude Desktop, then reconnect extension")
    else:         info("No fixes needed or all skipped")

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Diagnose and repair Claude extension in Microsoft Edge")
    parser.add_argument("--fix", action="store_true", help="Apply auto-fixes")
    args = parser.parse_args()

    _emit("=" * 62)
    _emit("  Claude Extension -- Edge Diagnostic & Repair Tool")
    _emit(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    _emit("=" * 62)

    results = run_diagnostics()
    status  = print_summary(results)

    if args.fix:
        run_fixes(results)
    elif status == "issues":
        _emit("\n  Re-run with --fix to apply auto-fixes:")
        _emit("    py claude_edge_fix.py --fix")

    flush_log()
    _emit(f"\n  Log -> {LOG_PATH}\n")

if __name__ == "__main__":
    main()
