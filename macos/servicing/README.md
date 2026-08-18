# macos/servicing

Semi-interactive macOS developer storage cleanup tool.

## What it does

- Scans common large, re-creatable developer storage buckets
- Shows size, purpose, and impact before each cleanup action
- Prompts per item (safe by default)
- Supports analyze-only and dry-run modes

Coverage includes:

- Xcode DerivedData, Archives, DeviceSupport, Simulator data
- System-level simulator runtimes and caches (/Library/Developer/CoreSimulator),
  orphaned simulator devices
- VS Code / Code Insiders / Cursor / Windsurf workspace data and logs
- .NET workload packs, NuGet, Gradle caches
- Android SDK/AVD, Flutter/Dart caches
- JetBrains caches/logs
- XDG cache (~/.cache), leftover code-sign clones (Edge/Teams updater bug)
- Per-repo build artifacts under a code root (node_modules, build, .dart_tool,
  Pods, target, .build, DerivedData) — context-sensitive names like "build"
  are only matched next to their manifest (pubspec.yaml, Cargo.toml, Podfile…)
- Homebrew cache, stale Homebrew installs, Docker prune, Time Machine snapshots

## Usage

Run from this folder or with full path:

./macos_dev_cleanup.sh --analyze
./macos_dev_cleanup.sh --dry-run
./macos_dev_cleanup.sh
./macos_dev_cleanup.sh --yes
./macos_dev_cleanup.sh --yes --aggressive
./macos_dev_cleanup.sh --code-root ~/src   # repo scan location (default: ~/Documents/Git)

## Safety model

- SAFE: re-creatable caches and logs
- LOW RISK: may reset tool/editor state
- PERMANENT: user backup/history deletion
- DESTRUCTIVE: can remove runtime/data that must be reinstalled

Plain `--yes` remains conservative: disruptive/destructive operations such as
simulator/runtime removal, Android SDK deletion, Docker volume pruning, and
all-repo build-artifact cleanup are skipped unless `--aggressive` is also set.
Build-sensitive cleanup is skipped while Flutter, Xcode, Gradle, CocoaPods,
Swift, Cargo, .NET, or JavaScript builds are active. Use
`--force-active-builds` only when deliberately overriding that protection.
Unattended runs skip build-sensitive cleanup even when no build process is
currently visible, avoiding races between sequential build commands. Combining
`--yes --aggressive --force-active-builds` is the explicit opt-in for that
behavior. The blanket `~/Library/Caches` cleanup is also treated as aggressive
and build-sensitive because it includes CocoaPods and other developer caches.
Conservative unattended cleanup retains package caches, Gradle distributions,
simulators/AVDs, IDE indexes, XDG caches, and extension state; it automatically
removes only diagnostic logs, crash data, downloaded extension installers, and
stale workspace entries.

## Notes

- Designed for macOS developer machines
- Requires bash and python3
- For Docker, Time Machine, and system-level simulator cache operations,
  extra privileges (sudo) may be requested
- Prints total space freed and remaining free space at the end of each run
