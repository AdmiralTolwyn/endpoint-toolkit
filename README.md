# Endpoint Toolkit

A collection of scripts, templates, and tools for managing Windows endpoints at scale — covering Azure Virtual Desktop image builds, session host lifecycle, and day-to-day operational tasks.

## Repository Structure

```
avd/
├── bicep/          # Bicep templates for AVD session host deployment
│   ├── modules/    # Reusable modules (session hosts, image templates)
│   └── main-*.bicep
├── pipelines/      # Azure DevOps YAML pipelines
└── scripts/        # PowerShell scripts used by pipelines

devops/
└── aib-task-v1-patched/   # Patched Azure Image Builder DevOps task (v2)

tools/              # Standalone PowerShell/WPF utilities
windows/            # Windows OS-level fixes and helpers
```

## Getting Started

Most pipeline files use `<YOURVALUE>` placeholders — search for `<YOUR` and replace with your environment-specific values before use.

## Requirements

- PowerShell 5.1+
- Azure CLI / Az PowerShell modules (for AVD scripts and pipelines)
- Windows 11 (for WPF-based tools)

## License

MIT
