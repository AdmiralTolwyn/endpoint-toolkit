# macos/servicing

Semi-interactive macOS developer storage cleanup tool.

## What it does

- Scans common large, re-creatable developer storage buckets
- Shows size, purpose, and impact before each cleanup action
- Prompts per item (safe by default)
- Supports analyze-only and dry-run modes

Coverage includes:

- Xcode DerivedData, Archives, DeviceSupport, Simulator data
- VS Code / Code Insiders / Cursor / Windsurf workspace data and logs
- .NET workload packs, NuGet, Gradle caches
- Android SDK/AVD, Flutter/Dart caches
- JetBrains caches/logs
- Homebrew cache, stale Homebrew installs, Docker prune, Time Machine snapshots

## Usage

Run from this folder or with full path:

./macos_dev_cleanup.sh --analyze
./macos_dev_cleanup.sh --dry-run
./macos_dev_cleanup.sh
./macos_dev_cleanup.sh --yes

## Safety model

- SAFE: re-creatable caches and logs
- LOW RISK: may reset tool/editor state
- PERMANENT: user backup/history deletion
- DESTRUCTIVE: can remove runtime/data that must be reinstalled

## Notes

- Designed for macOS developer machines
- Requires bash and python3
- For Docker and Time Machine operations, extra privileges may be requested
