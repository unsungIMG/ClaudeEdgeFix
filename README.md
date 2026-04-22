# ClaudeEdgeFix

Diagnostic and repair tool for the **Claude browser extension in Microsoft Edge**.

Edge is officially supported but requires separate NMH registration and manual
extension installation -- neither of which Claude Desktop performs automatically.

## Checks

| # | Check | Notes |
|---|-------|-------|
| D1 | Edge NMH registry - claude_browser_extension | Core requirement |
| D2 | Edge NMH registry - claude_code_browser_extension | Claude Code only |
| D3 | Chrome NMH registry | Reference baseline |
| D4 | NMH JSON manifest exists | |
| D5 | NMH binary (chrome-native-host.exe) | |
| D6 | Claude extension in Edge profiles | **Most common failure** |
| D7 | Claude extension in Chrome profiles | Reference |
| D8 | Claude Desktop installed | |
| D9 | Claude Desktop running | Must be running for NMH |
| D10 | Edge process running | |

## Fixes

| Fix | Action |
|-----|--------|
| F1 | Create Edge NMH registry key (copy from Chrome entry) |
| F2 | Create Edge NMH (Claude Code) registry key |
| F3 | Open Edge to Chrome Web Store Claude extension install page |

## Usage

```bash
# Git Bash
py claude_edge_fix.py          # diagnostics only
py claude_edge_fix.py --fix    # diagnostics + auto-fix
```

## After --fix

1. Click **Add to Chrome** on the Web Store page that opens in Edge
2. Sign in with the same account as Claude Desktop
3. Restart Edge + Claude Desktop
4. Verify: Claude Desktop -> Cowork tab, or check extension icon

## Known Edge limitations (April 2026)

- Side panel closes on tab switch (upstream Anthropic bug)
- Cowork dialog hardcodes "Chrome" text even when connected via Edge

## Requirements

Python 3.10+, stdlib only (`winreg`, `subprocess`, `pathlib`, `json`)
