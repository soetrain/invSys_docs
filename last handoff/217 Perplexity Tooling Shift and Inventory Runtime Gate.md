# Handoff 217: Perplexity tooling shift and Inventory runtime gate

Date: 2026-11-03 Pacific. Predecessor: handoff 216.

## Status

The user authorized Perplexity to take implementation work from Codex, selected
`C:\Users\Justin\repos` as the workspace, and approved publication of two fork
review PRs against the existing working branches. No merge is authorized.

A source-tooling compatibility defect was reproduced, fixed, and tested.
**4be-A remains incomplete, focused packaged Inventory GREEN remains pending,
and the Excel shutdown cause remains unresolved.** Excel was not launched.
This shift does not extend the previous runtime increment or reuse stale
desktop-monitor evidence.

## Published code and bounded scope

- [Tooling review PR6](https://github.com/soetrain/invSys_fork/pull/6):
  `justinwj:work/perplexity-json-literal-types-20261103` into
  `soetrain:work/d8-inv-edit-receipts`.
- Commit: `67a1ac3c91a224b3923015a80003cbd65b1e8441`.
- Original code branch/base: `2c9c3b775b737484243a3a5780dc97a269436091`.
- Original [code PR4](https://github.com/soetrain/invSys_fork/pull/4) and
  [docs PR3](https://github.com/soetrain/invSys_docs/pull/3) remain the owning
  work streams; the new PRs are deliberately stacked on their working branches,
  not targeted at `main`, under the user's specific publication approval.
- No VBA, XLAM, architecture, Plan, Controls, registry, or Excel setting changes.
  Continue to resolve authoritative documents through their CURRENT pointers.

On PowerShell 7.6.9, the diagnostic JSON readers turned ISO timestamp strings
into date objects before schema validation. This caused valid string-valued
timestamps to fail. The two readers now use `ConvertFrom-Json -DateKind String`
where that parameter exists (PowerShell 7.5+). The Windows PowerShell 5.1 path
is unchanged, and no schema validation rules were relaxed.

Changed code-repository files:

- `tools/inventory-vba-surface.ps1`
- `tools/validate-json-contract.ps1`
- `tests/tooling/Test-JsonLiteralTypes.ps1`
- `tests/integration/plan022_slice4be_admin_inventory_entry_results.md`

## Evidence

- New regression on PowerShell 7.6.9: RED 5 PASS / 7 FAIL, then GREEN
  12 PASS / 0 FAIL.
- Same regression on X1 Windows PowerShell 5.1.26100.9711:
  12 PASS / 0 FAIL.
- Existing static-tool contract suite: 10 PASS on each host, including all
  four byte-for-byte deterministic-output checks.
- Five Inventory source-guard suites: 36 PASS on the X1.
- Architecture primitive-boundary Python tests: 4 PASS on the source snapshot.
- Full source scan: completed after the fix, with all baseline metric fields
  equal to the retained `inventory-edit-creation-static-02` report.
  There are 391 components, 6,742 procedures, 146,183 lines, 9 literal and
  45 unresolved Application.Run sites, and 191 duplicate-body candidates.
- All five candidate XLAM hashes matched the original package pins at both
  preflight and postcheck. Excel process count was zero.

These are tooling and source checks, not packaged product behavioral GREEN.
The existing Inventory audit contains the detailed evidence entry.

Local evidence in `invSys_edit_receipts`:
`reports/runtime/perplexity-inventory-preflight-20261104/`.
The directory uses the UTC date. `verified/` contains the full scanner reports
and isolated-workspace RED/GREEN logs; the original source snapshot and
publication bundle are also retained in the parent directory. Synthetic
regression fixtures are under `reports/runtime/json-literal-types/`.

The verified snapshot was built from the clean local checkout, not substituted
from a remote clone. It contained source, tests, tooling, and three text
specifications, but no workbooks, credentials, or Windows settings.

## Local state and execution boundaries

The implementation worktree is still at
`C:\Users\Justin\repos\invSys_edit_receipts`, now on the dedicated review branch
`work/perplexity-json-literal-types-20261103`. The original
`work/d8-inv-edit-receipts` branch was left at its original commit.

The docs worktree is on `work/perplexity-handoff217`, branched from
`53d12b60aa58ac0acc37f838db22562f3cf66c20`. Its original
`work/excel-shutdown-timing` branch is unchanged. Only this handoff and the
`last handoff/CURRENT.md` pointer belong to the docs review commit.
The pre-existing handoff-067 edit and guidance-023/024/025 changes remain
uncommitted and must not be staged, discarded, or treated as this shift's work.
The separate `invSys_fork` worktree's pre-existing AGENTS edit was not touched.

Selecting the workspace fixed folder-command access. However, commands execute
as `PerplexitySandbox`, not the interactive Windows user. Per-user setting
writes are restricted and subprocesses do not survive the command call.
The full local static scan also hit the desktop executor's 120-second limit.
Consequently, do not run the existing long-lived Excel controller through this
lane or assume its HKCU/desktop observations represent the interactive user.

The Git LFS shell hook wrapper could not create its signal pipe in this
environment. The corresponding native `git lfs post-checkout` and
`git lfs post-commit` commands were run directly; no hooks were disabled.
Publication uses exact local commit bundles and the authenticated GitHub
connection, not a second implementation of the patch.

## Next product action

Retain handoff 216's candidate and complete RED evidence. The focused
QuantityPhase baseline has 152 report rows representing 140 unique checks,
with consistent duplicate outcomes: preserve 131 prior PASS and correct nine
RED checks. Run the fresh monitored focused GREEN in an appropriate normal-user
Windows execution context, then the broader Inventory slice if it passes.

Do not repeat already-complete RED work, rebuild unchanged XLAMs merely for
this tooling patch, skip desktop preflight or cleanup, alter Excel security
settings to fit the sandbox, or claim acceptance from these source-only tests.
Perplexity can continue source implementation, diagnostics, static tests, and
review work while the runtime execution lane is resolved. No recurring shift
scheduler or unattended relay was installed.
