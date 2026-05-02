#!/usr/bin/env bash
# =============================================================================
# macos_dev_cleanup.sh
# Semi-interactive macOS developer storage cleanup
#
# Scans common large, re-creatable data buckets produced by Xcode, simulators,
# VS Code / Cursor / Windsurf, .NET, Gradle, Android SDK, Flutter, JetBrains,
# Docker, and Homebrew. Shows size and impact for each item before asking.
#
# Usage:
#   ./macos_dev_cleanup.sh [--analyze] [--dry-run] [--yes]
#
#   --analyze   Report sizes and impact only — nothing deleted
#   --dry-run   Show every rm command that would run — nothing deleted
#   --yes       Non-interactive: auto-confirm all prompts (dangerous)
#
# Requirements: bash 3.2+, macOS 11+, python3 (for stale workspace pruning)
# =============================================================================

set -uo pipefail

DRY_RUN=0
ASSUME_YES=0
ANALYZE_ONLY=0

for arg in "$@"; do
  case "$arg" in
    --dry-run)   DRY_RUN=1 ;;
    --yes)       ASSUME_YES=1 ;;
    --analyze)   ANALYZE_ONLY=1 ;;
    -h|--help)
      sed -n '3,14p' "$0" | sed 's/^# \{0,1\}//'
      exit 0 ;;
    *)
      echo "Unknown option: $arg  (use --help)" >&2; exit 1 ;;
  esac
done

# ── parallel scan cache ──────────────────────────────────────────────────────
# All du calls run as background jobs. path_kb() polls the result file so
# item() gets an instant answer once the job finishes. In --analyze mode we
# wait for all jobs before printing; in interactive mode the user's reading
# time lets background scans complete naturally.

SCAN_DIR=""

_init_scan_cache() {
  SCAN_DIR=$(mktemp -d /tmp/macos_cleanup_XXXXXX)
  trap 'rm -rf "$SCAN_DIR"' EXIT INT TERM
}

# Stable short key from path (md5 is available on macOS 10.x+)
_path_key() { printf '%s' "$1" | md5; }

# Enqueue a background du scan; no-op if already queued.
prefork_scan() {
  local p="$1"
  [[ -z "$SCAN_DIR" ]] && _init_scan_cache
  local key result marker
  key=$(_path_key "$p")
  result="$SCAN_DIR/$key.result"
  marker="$SCAN_DIR/$key.running"
  [[ -f "$result" || -f "$marker" ]] && return
  touch "$marker"
  (
    local kb=0
    [[ -e "$p" ]] && kb=$(du -sk "$p" 2>/dev/null | awk '{print $1+0}')
    echo "$kb" > "$result"
    rm -f "$marker"
  ) &
}

# Pre-enqueue all statically-known paths so scans overlap with script init.
_prefork_all_known() {
  local paths
  paths=(
    "$HOME/Library/Developer/Xcode/DerivedData"
    "$HOME/Library/Developer/Xcode/Archives"
    "$HOME/Library/Developer/Xcode/iOS DeviceSupport"
    "$HOME/Library/Developer/Xcode/watchOS DeviceSupport"
    "$HOME/Library/Developer/CoreSimulator/Devices"
    "$HOME/Library/Developer/CoreSimulator/Caches"
    "$HOME/Library/Application Support/Code/User/workspaceStorage"
    "$HOME/Library/Application Support/Code/User/History"
    "$HOME/Library/Application Support/Code/CachedExtensionVSIXs"
    "$HOME/Library/Application Support/Code/logs"
    "$HOME/Library/Application Support/Code/Crashpad"
    "$HOME/Library/Application Support/Code/blob_storage"
    "$HOME/Library/Application Support/Code - Insiders/User/workspaceStorage"
    "$HOME/Library/Application Support/Code - Insiders/User/History"
    "$HOME/Library/Application Support/Cursor/User/workspaceStorage"
    "$HOME/Library/Application Support/Cursor/User/History"
    "$HOME/Library/Application Support/Windsurf/User/workspaceStorage"
    "$HOME/Library/Application Support/Windsurf/User/History"
    "$HOME/.dotnet/packs"
    "$HOME/.dotnet/templates"
    "$HOME/.nuget/packages"
    "$HOME/.gradle/caches"
    "$HOME/.gradle/wrapper/dists"
    "$HOME/Library/Android"
    "$HOME/Library/Android/sdk"
    "$HOME/Android/Sdk"
    "$HOME/.android/avd"
    "$HOME/.pub-cache"
    "$HOME/.dartServer"
    "$HOME/Library/Caches/JetBrains"
    "$HOME/Library/Logs/JetBrains"
    "$HOME/Library/Application Support/MobileSync/Backup"
    "$HOME/Library/Caches"
    "$HOME/Library/Logs"
    "$HOME/.Trash"
    "$HOME/.npm"
    "$HOME/Library/Caches/pip"
    "$HOME/Library/Caches/Yarn"
    "$HOME/.homebrew"
    "$HOME/.brew"
    "/usr/local"
    "$HOME/.vscode"
  )
  for p in "${paths[@]}"; do prefork_scan "$p"; done
}

# ── helpers ───────────────────────────────────────────────────────────────────

human_size() {
  awk -v n="$1" 'BEGIN {
    u[1]="KB"; u[2]="MB"; u[3]="GB"; u[4]="TB"
    v=n+0; i=1
    while (v>=1024 && i<4) { v=v/1024; i++ }
    printf "%.2f %s", v, u[i]
  }'
}

# Returns size in KB. Uses cached result from a background scan when available;
# falls back to inline du if the path was not pre-enqueued.
path_kb() {
  local p="$1"
  [[ -e "$p" ]] || { echo 0; return; }

  if [[ -n "$SCAN_DIR" ]]; then
    local key result marker
    key=$(_path_key "$p")
    result="$SCAN_DIR/$key.result"
    marker="$SCAN_DIR/$key.running"
    # Wait up to 30s for a running background job
    if [[ -f "$marker" || -f "$result" ]]; then
      local i=0
      while [[ ! -f "$result" ]] && (( i < 600 )); do
        sleep 0.05; (( i++ ))
      done
      [[ -f "$result" ]] && { cat "$result"; return; }
    fi
  fi

  # Inline fallback (dynamically-discovered path not pre-enqueued)
  du -sk "$p" 2>/dev/null | awk '{print $1+0}'
}

confirm() {
  [[ "$ASSUME_YES" -eq 1 ]] && return 0
  local ans
  read -r -p "  $1 [y/N] " ans
  [[ "$ans" =~ ^[Yy]$ ]]
}

do_remove_contents() {
  if [[ "$DRY_RUN" -eq 1 ]]; then
    echo "  DRY RUN: rm -rf ${1}/*"
    return
  fi
  find "$1" -mindepth 1 -maxdepth 1 -exec rm -rf {} + 2>/dev/null || true
}

do_remove_path() {
  if [[ "$DRY_RUN" -eq 1 ]]; then
    echo "  DRY RUN: rm -rf $1"
    return
  fi
  rm -rf "$1"
}

# item LABEL  MODE(path|contents)  PATH  WHAT  IMPACT
item() {
  local label="$1" mode="$2" path="$3" desc="${4:-}" impact="${5:-}"
  local kb
  kb=$(path_kb "$path")
  echo ""
  echo "  [$label]"
  echo "  Path   : $path"
  [[ -n "$desc"   ]] && echo "  What   : $desc"
  [[ -n "$impact" ]] && echo "  Impact : $impact"
  if [[ "$kb" -eq 0 ]]; then
    echo "  Size   : empty / not found — skipping"
    return
  fi
  echo "  Size   : $(human_size "$kb")"
  [[ "$ANALYZE_ONLY" -eq 1 ]] && return
  confirm "Delete?" || { echo "  Skipped."; return; }
  case "$mode" in
    contents) do_remove_contents "$path" ;;
    path)     do_remove_path "$path" ;;
  esac
  echo "  Done."
}

# cmd_item LABEL  PREVIEW_CMD  CLEANUP_CMD  WHAT  IMPACT
cmd_item() {
  local label="$1" preview="$2" cleanup="$3" desc="${4:-}" impact="${5:-}"
  echo ""
  echo "  [$label]"
  [[ -n "$desc"   ]] && echo "  What   : $desc"
  [[ -n "$impact" ]] && echo "  Impact : $impact"
  [[ -n "$preview" ]] && eval "$preview" 2>/dev/null || true
  [[ "$ANALYZE_ONLY" -eq 1 ]] && return
  confirm "Run cleanup?" || { echo "  Skipped."; return; }
  if [[ "$DRY_RUN" -eq 1 ]]; then
    echo "  DRY RUN: $cleanup"
    return
  fi
  eval "$cleanup" 2>&1 | tail -8
  echo "  Done."
}

# ── VS Code variant detection ─────────────────────────────────────────────────
# Handles VS Code stable, Insiders, Cursor, and Windsurf — all Electron-based
# editors that store workspace state in the same structure.
# Prints one path per line so callers can use 'while IFS= read -r' safely.
vscode_variants() {
  local -a dirs=(
    "$HOME/Library/Application Support/Code"
    "$HOME/Library/Application Support/Code - Insiders"
    "$HOME/Library/Application Support/Cursor"
    "$HOME/Library/Application Support/Windsurf"
  )
  for d in "${dirs[@]}"; do
    [[ -d "$d" ]] && printf '%s\n' "$d"
  done
}

# ── stale VS Code workspaceStorage pruning ────────────────────────────────────
# One python3 invocation scans every workspace.json at once — vastly faster
# than spawning python3 once per folder (avoids N interpreter startups).
prune_stale_workspace_storage() {
  local base="$1/User/workspaceStorage"
  [[ -d "$base" ]] || return

    # Write the python script to a temp file — more portable than a heredoc
    # inside $() which trips up bash 3.2 in some configurations.
    local py_script="$SCAN_DIR/find_stale_ws.py"
    printf '%s\n' \
      'import json, os, sys' \
      'base = sys.argv[1]' \
      'try:' \
      '    entries = os.listdir(base)' \
      'except OSError:' \
      '    sys.exit(0)' \
      'for name in sorted(entries):' \
      '    d = os.path.join(base, name)' \
      '    if not os.path.isdir(d):' \
      '        continue' \
      '    wj = os.path.join(d, "workspace.json")' \
      '    if not os.path.exists(wj):' \
      '        print(d)' \
      '        continue' \
      '    try:' \
      '        data = json.load(open(wj))' \
      '        folder = data.get("folder", "").replace("file://", "").rstrip("/")' \
      '        if folder and not os.path.isdir(folder):' \
      '            print(d)' \
      '    except Exception:' \
      '        pass' \
      > "$py_script"
    local stale_list
    stale_list=$(python3 "$py_script" "$base" 2>/dev/null || true)

  local freed=0
  while IFS= read -r dir; do
    [[ -z "$dir" ]] && continue
    local kb
    kb=$(du -sk "$dir" 2>/dev/null | awk '{print $1+0}')
    freed=$((freed + kb))
    if [[ "$DRY_RUN" -eq 1 ]]; then
      echo "  DRY RUN: rm -rf $dir"
    else
      rm -rf "$dir"
    fi
  done <<< "$stale_list"

  echo "  Reclaimed approx $(human_size "$freed") from stale workspace entries."
}

# ── Homebrew stale install detection ─────────────────────────────────────────
# Checks common Homebrew prefix locations and identifies which are not active.
check_stale_homebrews() {
  local active_prefix=""
  if command -v brew >/dev/null 2>&1; then
    active_prefix="$(brew --prefix 2>/dev/null || true)"
  fi

  local -a candidates=("$HOME/.homebrew" "$HOME/.brew" "/usr/local")
  for prefix in "${candidates[@]}"; do
    [[ -f "$prefix/bin/brew" ]] || continue
    [[ "$prefix" == "$active_prefix" ]] && continue
    item "Stale Homebrew install ($prefix)" path \
      "$prefix" \
      "A Homebrew installation not on your active PATH (active prefix: ${active_prefix:-none})." \
      "SAFE. Active Homebrew tools will continue to work. Removes only this unused installation."
  done
}

# ── Android SDK location detection ───────────────────────────────────────────
find_android_sdk() {
  local -a candidates=(
    "$HOME/Library/Android/sdk"
    "$HOME/Library/Android"
    "$HOME/Android/Sdk"
    "$HOME/android-sdk"
  )
  # Also honour ANDROID_HOME / ANDROID_SDK_ROOT env vars
  for var in ANDROID_HOME ANDROID_SDK_ROOT; do
    local val="${!var:-}"
    [[ -n "$val" ]] && candidates=("$val" "${candidates[@]}")
  done
  for p in "${candidates[@]}"; do
    [[ -d "$p" ]] && { echo "$p"; return; }
  done
  echo ""
}

# ── Flutter SDK location ──────────────────────────────────────────────────────
# Returns the directory of the Flutter SDK currently on PATH (if any).
active_flutter_sdk() {
  local flutter_bin
  flutter_bin="$(command -v flutter 2>/dev/null || true)"
  [[ -z "$flutter_bin" ]] && { echo ""; return; }
  # flutter binary lives at <sdk>/bin/flutter
  dirname "$(dirname "$flutter_bin")"
}

# =============================================================================
# MAIN
# =============================================================================

# Kick off all background du scans immediately so they run in parallel
# while the header and early sections are being processed/read.
_init_scan_cache
_prefork_all_known

echo ""
echo "macOS Developer Storage Cleanup  ($(date '+%Y-%m-%d'))"
[[ "$ANALYZE_ONLY" -eq 1 ]] && echo "Mode : analyze only — nothing will be deleted"
[[ "$DRY_RUN"      -eq 1 ]] && echo "Mode : dry run — no files will be deleted"
echo "======================================================================"

# In analyze/dry-run mode there is no user interaction to buy scan time,
# so wait for all background jobs before the first item prints.
if [[ "$ANALYZE_ONLY" -eq 1 || "$DRY_RUN" -eq 1 ]]; then
  wait
fi

# ── 1. Xcode / Simulator ─────────────────────────────────────────────────────
echo ""
echo "── 1. XCODE / SIMULATOR ─────────────────────────────────────────────"

item "Xcode DerivedData" path \
  "$HOME/Library/Developer/Xcode/DerivedData" \
  "Intermediate build products for every Xcode project ever built." \
  "SAFE. Rebuilt automatically on next build (adds a few minutes compile time)."

item "Xcode Archives" path \
  "$HOME/Library/Developer/Xcode/Archives" \
  "Archived .xcarchive bundles used to export/upload IPAs." \
  "PERMANENT loss of those archives. Keep if you may need to re-export; delete if all builds are in App Store Connect."

item "Xcode iOS DeviceSupport" path \
  "$HOME/Library/Developer/Xcode/iOS DeviceSupport" \
  "Per-device debug symbol packages downloaded when an iOS device is first connected." \
  "SAFE. Re-downloaded automatically the next time that device is connected to Xcode."

item "Xcode watchOS DeviceSupport" path \
  "$HOME/Library/Developer/Xcode/watchOS DeviceSupport" \
  "Same as above but for Apple Watch devices." \
  "SAFE. Re-downloaded automatically."

item "CoreSimulator Devices" path \
  "$HOME/Library/Developer/CoreSimulator/Devices" \
  "All installed iOS/macOS/watchOS simulator runtimes and their per-app data." \
  "SAFE. Runtimes are re-downloadable via Xcode > Settings > Platforms. Loses simulator app data."

item "CoreSimulator Caches" contents \
  "$HOME/Library/Developer/CoreSimulator/Caches" \
  "Simulator disk image caches used to speed up device creation." \
  "SAFE. Rebuilt automatically on next simulator launch."

# ── 2. VS Code / Cursor / Windsurf ───────────────────────────────────────────
echo ""
echo "── 2. VS CODE / CURSOR / WINDSURF ───────────────────────────────────"

while IFS= read -r base; do
  [[ -z "$base" ]] && continue
  app_name="$(basename "$base")"

  item "$app_name workspaceStorage (ALL)" contents \
    "$base/User/workspaceStorage" \
    "Per-workspace indexes, AI chat history, extension state, search indexes. One folder per workspace ever opened." \
    "Loses all AI chat history and workspace search indexes for ALL workspaces. Editor rebuilds indexes on reopen."

  item "$app_name Edit History" contents \
    "$base/User/History" \
    "Local edit history — lets you recover previous file versions via the Timeline panel." \
    "Loses ability to recover old file versions via Timeline. No impact on current files."

  item "$app_name Cached VSIXs" path \
    "$base/CachedExtensionVSIXs" \
    "Downloaded .vsix extension installers kept as an offline install cache." \
    "SAFE. Re-downloaded from the marketplace on next install. Only affects offline extension installs."

  item "$app_name logs" contents \
    "$base/logs" \
    "Diagnostic and extension logs written continuously by the editor." \
    "SAFE. No functional impact."

  item "$app_name Crashpad" contents \
    "$base/Crashpad" \
    "Crash report dumps pending upload to the vendor." \
    "SAFE. No functional impact."

  item "$app_name blob_storage" contents \
    "$base/blob_storage" \
    "Internal Electron/Chromium blob and IndexedDB storage used by some extensions." \
    "LOW RISK. May reset some extension UI state; extensions re-populate on next launch."

  if [[ "$ANALYZE_ONLY" -ne 1 ]]; then
    echo ""
    echo "  [$app_name workspaceStorage — stale entries only (safer)]"
    echo "  What   : Removes only entries whose workspace folder no longer exists on disk."
    echo "  Impact : Targeted — keeps history for active projects, reclaims orphaned data only."
    if confirm "Prune orphaned workspace entries for $app_name?"; then
      prune_stale_workspace_storage "$base"
    fi
  fi
done < <(vscode_variants)

# ── 3. .NET / NuGet / Gradle ─────────────────────────────────────────────────
echo ""
echo "── 3. .NET / NUGET / GRADLE ─────────────────────────────────────────"

if [[ -d "$HOME/.dotnet" ]]; then
  item ".NET SDK packs" path \
    "$HOME/.dotnet/packs" \
    "Workload SDK packs (Android, iOS, MAUI runtimes) installed via 'dotnet workload install'." \
    "SAFE if unused packs are cleaned first with 'dotnet workload clean'. Full delete forces re-download on next build."

  item ".NET templates" path \
    "$HOME/.dotnet/templates" \
    "Project template packages installed via 'dotnet new install'." \
    "SAFE. Re-installed on demand with 'dotnet new install <template-id>'."

  if command -v dotnet >/dev/null 2>&1; then
    cmd_item ".NET unused workload packs" \
      "dotnet workload list 2>/dev/null | head -12 || true" \
      "dotnet workload clean" \
      "Removes SDK packs no longer referenced by any installed .NET SDK version." \
      "SAFE. Only removes genuinely unused packs; active workloads are preserved."
  fi
fi

item "NuGet package cache" path \
  "$HOME/.nuget/packages" \
  "Global NuGet package cache — all packages ever restored across all .NET projects." \
  "SAFE. Re-downloaded from NuGet.org on next 'dotnet restore' (adds time)."

item "Gradle caches" path \
  "$HOME/.gradle/caches" \
  "Gradle dependency and build artifact cache shared across all Android/Java projects." \
  "SAFE. Re-downloaded on next Gradle build (can take significant time on large projects)."

item "Gradle wrapper distributions" path \
  "$HOME/.gradle/wrapper/dists" \
  "Downloaded Gradle distribution ZIPs, one per Gradle version used across projects." \
  "SAFE. Re-downloaded when you next build a project requiring that Gradle version."

# ── 4. Android SDK ───────────────────────────────────────────────────────────
echo ""
echo "── 4. ANDROID SDK ───────────────────────────────────────────────────"

ANDROID_SDK="$(find_android_sdk)"
if [[ -n "$ANDROID_SDK" ]]; then
  item "Android SDK" path \
    "$ANDROID_SDK" \
    "Full Android SDK (platforms, build-tools, emulator images) managed by Android Studio or Flutter." \
    "DESTRUCTIVE. Breaks all Android builds until re-installed via the Android Studio SDK Manager."

  item "Android AVDs" path \
    "$HOME/.android/avd" \
    "Android Virtual Device definitions and their emulator disk images." \
    "SAFE if you don't use Android emulators. AVDs are re-creatable in Android Studio / AVD Manager."
fi

# ── 5. Flutter / Dart ────────────────────────────────────────────────────────
echo ""
echo "── 5. FLUTTER / DART ────────────────────────────────────────────────"

FLUTTER_SDK="$(active_flutter_sdk)"
if [[ -n "$FLUTTER_SDK" ]]; then
  echo ""
  echo "  [Active Flutter SDK]"
  echo "  Path   : $FLUTTER_SDK  (on PATH — NOT offered for deletion)"
  echo "  Size   : $(human_size "$(path_kb "$FLUTTER_SDK")")"
fi

item "Flutter pub cache" path \
  "$HOME/.pub-cache" \
  "Downloaded Dart/Flutter packages for all projects (analogous to npm's global cache)." \
  "SAFE. Re-downloaded from pub.dev on next 'flutter pub get' in each project (adds time)."

item "Dart analysis server cache" path \
  "$HOME/.dartServer" \
  "Type indexes and semantic data cached by the Dart analysis server across all projects." \
  "SAFE. Rebuilt automatically when a Dart/Flutter project is opened in the editor."

# ── 6. JetBrains IDEs ────────────────────────────────────────────────────────
echo ""
echo "── 6. JETBRAINS IDEs ────────────────────────────────────────────────"

# JetBrains stores caches and logs under versioned directories
JB_BASE="$HOME/Library/Caches/JetBrains"
if [[ -d "$JB_BASE" ]]; then
  item "JetBrains IDE caches" contents \
    "$JB_BASE" \
    "Local caches for all JetBrains IDEs (IntelliJ, Android Studio, GoLand, etc.)." \
    "SAFE. Rebuilt on next IDE launch (causes a slow index warm-up for large projects)."
fi

JB_LOG_BASE="$HOME/Library/Logs/JetBrains"
if [[ -d "$JB_LOG_BASE" ]]; then
  item "JetBrains IDE logs" contents \
    "$JB_LOG_BASE" \
    "Diagnostic logs written by JetBrains IDEs." \
    "SAFE. No functional impact."
fi

# ── 7. General caches / backups ───────────────────────────────────────────────
echo ""
echo "── 7. GENERAL CACHES / BACKUPS ──────────────────────────────────────"

item "iOS / iPadOS device backups" path \
  "$HOME/Library/Application Support/MobileSync/Backup" \
  "Full local Finder/iTunes backups of connected iOS devices." \
  "PERMANENT loss of those backups. Only delete if you rely on iCloud Backup or have a recent backup elsewhere."

item "User Caches" contents \
  "$HOME/Library/Caches" \
  "App-managed caches (Safari, Xcode, Spotlight, and many others)." \
  "SAFE. All apps rebuild caches on next use. May cause slower first launches."

item "User Logs" contents \
  "$HOME/Library/Logs" \
  "Diagnostic logs written by macOS and user-space apps." \
  "SAFE. No functional impact. Console.app log history will appear empty."

item "Trash" contents \
  "$HOME/.Trash" \
  "Files you have already moved to Trash but not permanently deleted." \
  "PERMANENT. Equivalent to 'Empty Trash'. Review manually if unsure of contents."

item "npm cache" path \
  "$HOME/.npm" \
  "npm package download cache shared across all Node.js projects." \
  "SAFE. Re-downloaded from the registry on next 'npm install'."

item "pip cache" path \
  "$HOME/Library/Caches/pip" \
  "pip wheel and HTTP caches shared across all Python environments." \
  "SAFE. Re-downloaded on next 'pip install'."

item "yarn cache" path \
  "$HOME/Library/Caches/Yarn" \
  "Yarn v1 package download cache." \
  "SAFE. Re-downloaded on next 'yarn install'."

# ── 8. Homebrew / Docker / Time Machine ──────────────────────────────────────
echo ""
echo "── 8. BREW / DOCKER / TIME MACHINE ──────────────────────────────────"

check_stale_homebrews

if command -v brew >/dev/null 2>&1; then
  BREW_CACHE="$(brew --cache 2>/dev/null || true)"
  [[ -n "$BREW_CACHE" ]] && item "Homebrew download cache" contents \
    "$BREW_CACHE" \
    "Cached formula and cask download tarballs kept by Homebrew after install." \
    "SAFE. Re-downloaded from the source on next 'brew install' for that formula."
fi

if command -v docker >/dev/null 2>&1; then
  cmd_item "Docker: images, volumes, build cache" \
    "docker system df 2>/dev/null" \
    "docker system prune -a --volumes -f" \
    "All Docker images, stopped containers, unused volumes, and build layer cache." \
    "DESTRUCTIVE for volumes containing persistent data. Images are re-pullable; named volume data is NOT recoverable."
fi

if command -v tmutil >/dev/null 2>&1; then
  cmd_item "Time Machine local snapshots" \
    "tmutil listlocalsnapshots / 2>/dev/null" \
    'for s in $(tmutil listlocalsnapshots / 2>/dev/null | grep "com.apple.TimeMachine" | sed "s/.*\.//"); do sudo tmutil deletelocalsnapshots "$s"; done' \
    "APFS local snapshots stored on the boot drive, created automatically by Time Machine." \
    "SAFE if you have a working external Time Machine drive. Without it, these are your only local restore points."
fi

# ─────────────────────────────────────────────────────────────────────────────
echo ""
echo "======================================================================"
if [[ "$ANALYZE_ONLY" -eq 1 || "$DRY_RUN" -eq 1 ]]; then
  echo "Finished (nothing deleted)."
else
  echo "Finished."
fi
