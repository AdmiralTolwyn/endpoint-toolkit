# Endpoint Toolkit — Agent Instructions

Operational tooling for Windows endpoints — Azure Virtual Desktop image
builds, Intune remediations, Bicep deployments, security baselines, and
W365 assessments. PowerShell + Bicep + a few WPF utilities. Repo structure
is documented in `README.md`. Read it first.

## Operational impact is the product
- Most scripts in this repo **run with elevated privileges on production
  endpoints, image-bake VMs, or tenant-wide via Intune**. A casual edit
  can brick a fleet.
- Treat every script as production until proven otherwise. If a code path
  could change behavior on a real device, flag it before changing it.

## Scope discipline
- Only modify code/files directly related to the asked task. No drive-by
  PowerShell style fixes, parameter renames, or "improvements" to scripts
  you weren't asked about.
- If you notice something worth fixing elsewhere, mention it at the end.
  Don't touch it unless asked.
- Before changing Bicep template parameters, pipeline YAML, or the
  signature of any parameterized script that's referenced from
  `avd/pipelines/`, stop and describe what's about to change. Wait for
  confirmation. Renaming a parameter silently breaks every pipeline call.

## Simplest solution first
- Implement the simplest thing that works. Native PowerShell cmdlets over
  .NET reflection; built-in modules over third-party where possible.
- Do not add comment-based help blocks, `[CmdletBinding()]` attributes,
  or extra parameter validation to scripts you didn't change.

## Uncertainty and assumptions
- Flag uncertainty before answering. "I'm not sure" beats a confident guess
  — especially for ADMX semantics, Intune CSP behavior, Windows version-
  specific registry paths, and Bicep API versions.
- If a fact can be verified by reading the script, the Bicep module, or
  the docs in `docs/`, read it. Don't infer from filename.

## Protected surfaces (do not modify without explicit permission)
- `avd/bicep/` modules — referenced by pipelines and other modules; API
  version changes need careful review.
- `avd/customizer/ConfigurationFiles/` — bundled VDOT JSON; intentionally
  in-repo so customizer runs don't depend on network fetch.
- `avd/pipelines/*.yml` — Azure DevOps YAML; broken syntax fails the
  pipeline.
- `tools/BaselineAssessor/` — 263 security checks; the checklist is the
  contract.
- `tools/PolicyPilot/`, `tools/W365Assessor/` — Graph permissions are
  least-privilege; do not request additional scopes without flagging.
- `intune/bitlocker/` Proactive Remediation pair — output format is
  consumed by Intune; do not change the exit-code contract.

## Destructive / external actions
- Scripts that touch BitLocker, hardware speculation mitigations, Secure
  Boot, FSLogix profile containers, or session host drain — these are
  fleet-wide if misrun. Never invoke from this repo's terminal "to test"
  without the user explicitly asking in the current message.
- `git push`, `git push --force`, `git reset --hard`, branch deletion,
  `rm -rf` — confirm in the current message before running.
- No `az login` / `Connect-MgGraph` / pipeline triggers unless explicitly
  asked.

## Change summary
After any non-trivial edit, end with:
- **Files changed:** `<list>`
- **Files intentionally not touched:** `<list, if relevant>`
- **Follow-up needed:** `<decisions, blast radius, or test gaps>`
