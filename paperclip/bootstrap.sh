#!/usr/bin/env bash
# Paperclip AI bootstrap for the Equity Markets blog.
# Run this ON YOUR LOCAL MACHINE (not in a cloud sandbox), ideally driven by a
# local Claude Code session: "run paperclip/bootstrap.sh --preflight, then --install".
#
# Usage:
#   bash paperclip/bootstrap.sh --preflight   # check prerequisites, change nothing
#   bash paperclip/bootstrap.sh --install     # verified download + install + onboard
set -euo pipefail

INSTALL_URL="https://paperclip.ing/install.sh"
CHECKSUM_URL="https://paperclip.ing/install.sh.sha256"
WORKDIR="$(mktemp -d)"
trap 'rm -rf "$WORKDIR"' EXIT

say()  { printf '\n==> %s\n' "$*"; }
fail() { printf 'ERROR: %s\n' "$*" >&2; exit 1; }

preflight() {
  say "Preflight checks"
  local ok=1

  case "$(uname -s)" in
    Darwin|Linux) echo "  OS: $(uname -s) ✓" ;;
    *) echo "  OS: $(uname -s) — needs WSL2 on Windows"; ok=0 ;;
  esac

  if command -v node >/dev/null 2>&1; then
    local major; major="$(node -p 'process.versions.node.split(".")[0]')"
    if [ "$major" -ge 20 ]; then echo "  Node $(node -v) ✓"
    else echo "  Node $(node -v) — need >= 20"; ok=0; fi
  else echo "  Node: not found — need >= 20"; ok=0; fi

  if command -v pnpm >/dev/null 2>&1; then echo "  pnpm $(pnpm -v) ✓"
  else echo "  pnpm: not found (installer may handle it; 'npm i -g pnpm' to be safe)"; fi

  if command -v claude >/dev/null 2>&1; then echo "  claude CLI ✓"
  else echo "  claude CLI: not found — install Claude Code first"; ok=0; fi

  if command -v codex >/dev/null 2>&1; then echo "  codex CLI ✓"
  else echo "  codex CLI: not found — needed for the Image Producer role"; ok=0; fi

  local avail_kb; avail_kb="$(df -Pk "$HOME" | awk 'NR==2 {print $4}')"
  if [ "$avail_kb" -gt 5242880 ]; then echo "  Disk: $((avail_kb / 1048576)) GB free ✓"
  else echo "  Disk: less than 5 GB free in \$HOME"; ok=0; fi

  [ "$ok" -eq 1 ] || fail "Preflight failed — fix the items above and re-run."
  say "Preflight passed"
}

install() {
  preflight

  say "Downloading installer and checksum"
  curl -fsSL -o "$WORKDIR/install.sh" "$INSTALL_URL"
  curl -fsSL -o "$WORKDIR/install.sh.sha256" "$CHECKSUM_URL"

  say "Verifying SHA-256 checksum"
  ( cd "$WORKDIR" && sha256sum -c install.sh.sha256 ) \
    || fail "Checksum verification FAILED — do not proceed. Re-download or check paperclip.ing."

  say "Running Paperclip installer"
  # Telemetry off by default for this install; flip if you decide otherwise.
  PAPERCLIP_TELEMETRY=0 bash "$WORKDIR/install.sh"

  say "Onboarding"
  command -v paperclipai >/dev/null 2>&1 || fail "paperclipai CLI not on PATH after install — open a new shell and run 'paperclipai onboard'."
  paperclipai onboard

  say "Recording installed version"
  paperclipai --version > "$(dirname "$0")/VERSION" 2>/dev/null || true

  say "Done. Next steps (see paperclip/PLAN.md, Phases 2-3):"
  echo "  1. Open the dashboard (localhost) and confirm it loads. Keep it loopback-only."
  echo "  2. Connect the Claude Code and Codex adapters; smoke-test each."
  echo "  3. Create the org from paperclip/agents/*.md with budgets + heartbeats per PLAN.md §5-6."
  echo "  4. Set the human-approval gate on the Publisher role before anything else runs."
}

case "${1:-}" in
  --preflight) preflight ;;
  --install)   install ;;
  *) echo "Usage: $0 --preflight | --install"; exit 2 ;;
esac
