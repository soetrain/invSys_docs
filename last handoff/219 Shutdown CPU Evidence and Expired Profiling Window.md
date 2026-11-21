# Handoff 219: Shutdown CPU evidence and expired profiling window

Date: 2026-11-03 Pacific. Predecessor: handoff 218. This is the session-end continuation record; earlier execution requests remain historical, not reusable authorization.

## Goal and release outcome

Complete Slice 4be-A under Architecture v4.11 and Plan 022, with Excel shutdown diagnosis as the immediate blocker. Shutdown cause, Inventory GREEN, early Receiving B0, and overall 4be-A acceptance remain unresolved; no release qualification or package promotion is claimed ([final readiness result](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6623012152)).

The approved D18-REPLAY-01 contract already requires early Receiving B0 before A closes; broader B and Release 1 acceptance remain separate. This diagnostic increment changes no architecture, plan contract, control behavior, or acceptance status; do not reopen the approved replay contract based on older guidance.

## Current verified state

Last verified: November 3, approximately 9:02 PM Pacific, by reviewing Codex's final report and the remote review branches. Native preservation results below are Codex's reported checks; Perplexity did not launch Excel or a profiler during this follow-up ([final report](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6623012152)).

- **Runtime:** paused, Excel closed, no new workload or CPU-stack recording dispatched. The user-opened elevated shell remained idle after the five read-only WPR checks; no session was changed ([final report](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6623012152)).
- **Code review:** draft [PR7](https://github.com/soetrain/invSys_fork/pull/7), branch `work/perplexity-shutdown-table-isolation`, diagnostic commit `62b3da4dbf7e39e0cf626e44f8949f1952a7b03b`, targeting `work/excel-shutdown-timing`, not `main`.
- **Source pins:** main fixture checkout `6d247543c557cca92fb6324f484b6ec2087e4527`; worker/helpers `44557070b2f1c3d7d744e1f8040315a837933a8f`; all 535 manifest entries plus observer/profile remained unchanged ([final report](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6623012152)).
- **Docs review:** this handoff and `last handoff/CURRENT.md` advance draft [PR5](https://github.com/soetrain/invSys_docs/pull/5) on `work/perplexity-handoff218`, from parent `4f2497b720bf1463d6fa1042e1c7dfcd3605238f`. The remote handoff commit is the PR head; the X1 docs checkout was last reported at that parent and has not been pulled by this cloud-only documentation update ([pre-wrap local state](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6622925020)).
- **Preservation:** unrelated main AGENTS and docs changes/untracked files were preserved by Codex. This wrap-up edits only the new handoff and CURRENT in a separate cloud checkout; it neither touches the Windows working trees nor merges either PR ([reported native preservation](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6623012152)).

## Decisions and constraints

- **Execution lane:** Perplexity owns implementation/review; Codex executes tightly scoped normal-Windows-session tests and returns failures without independently changing implementation. Only WPR may use the explicitly user-opened elevated shell; Excel/controller/worker/observer remain normal-user ([approved request](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6622812245)).
- **Expired permission:** Justin approved one conditional workload and one 60-second system-wide CPU-stack capture, latest dispatch 9:00 PM Pacific, with wrap-up by 10:00 PM. No controller was dispatched, but the unconsumed authorization expired at `2026-11-04T05:00:00Z`; a future run requires a new explicit window and resolved readiness, not an edited old date gate ([approval](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6622812245), [expiry result](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6623012152)).
- **Trace privacy/security:** raw ETL, session listings, full paths, command lines, and raw exports stay local. No global WPR cancellation, security-session interference, additional installation, broad capture, or assisted Excel teardown is authorized ([approved request](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6622812245)).
- **Symbols:** Microsoft public-symbol downloads were approved for this bounded analysis, with local cache and process-local configuration. None occurred; this is not a standing authorization for unrelated downloads or a future workload ([approval](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6622812245), [final report](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6623012152)).
- **D13:** this wrap-up changes documentation/pointers only, not runtime or architectural behavior. Validate the diff, pointer target, hashes, and remote PR state; do not manufacture a product RED/GREEN for a handoff.

## Evidence and traceability

| Evidence | Verified conclusion and limitation |
|---|---|
| Paired baseline/OwnedTableTraversal | Finalizer wait fell from 275.974 s to 0.028463 s, but total duration stayed 710.050 versus 709.231 s. Variant Excel exited 278.374 s after finalizer return. No meaningful overall shutdown improvement ([paired result](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6620924659)). |
| First passive observer | One run took 729.43 s with 290.22 s finalizer wait. Live-log sharing failures and premature managed exit detection left gaps; do not use those missing intervals as zero CPU or proof of exit ([v1 result](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6621503293)). |
| Corrected native observer v2 | Finalizer 251.144 s, controller 684.957 s. During roughly 250.47 s of finalizer waiting, Excel used 247.828 s kernel plus 0.328 s user CPU, approximately 0.991 core; worker used 0.031 s user CPU. Continuous CPU coverage localized activity to Excel while the worker was mostly idle, not to an identified internal component ([v2 result](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6622021455)). |
| Regression/preservation | Completed diagnostics retained 78 ordered assertions, 50 PASS / 28 known FAIL, original 535 inputs, natural exits, and reported clear desktop/crash audits. These are preserved diagnostic results, not Inventory GREEN ([paired result](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6620924659), [v2 result](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6622021455)). |
| Analyzer readiness | Microsoft WPA package 11.8.423.0 installed; WPA/WPAExporter file versions 11.8.423.12582 and valid Microsoft signatures verified. Empty analyzer opened/closed normally and exporter help succeeded. XPerf was absent; actual trace-stack attribution remains untested ([installation result](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6622682357)). |
| Final profiling readiness | All five elevated WPR probes succeeded and said not recording, including the documented `MSFT_MpPerformanceRecording` instance. They did not map or clear the other retained ETW session. No workload or recording occurred ([final result](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6623012152)). |

The unresolved defect remains: long natural Excel shutdown, with measured kernel CPU activity, but unknown responsible stack/component. Protecting evidence still requires unchanged packaged outcomes, exact process identity, native natural exits, settings/pin preservation, and delayed fault audits; a future CPU-stack trace would supplement these, not replace acceptance evidence ([profiling recipe](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6622184307)).

## Do Not Repeat

- **Do not promote the Table variant:** a shortened finalizer wait did not shorten total runtime ([paired result](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6620924659)).
- **Do not reuse observer v1 or infer native exit from managed HasExited:** use pinned v2 with explicit per-counter availability and shared log reading; v1's gaps limit its conclusions ([corrected observer request](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6621801697)).
- **Do not repeat default WPR probes as global clearance:** neither default nor one known named-instance negative maps all other ETW sessions. The completed elevated checks added no ownership/compatibility determination ([final report](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6623012152)).
- **Do not blame or disable Defender:** the session label, CLR loader provider, PowerShell provider, and all-bits provider masks establish neither sampled-CPU capture nor causality for Excel's delay ([final report](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6623012152)).
- **Do not run trial captures or more counter-only workloads to bypass readiness:** retain the ambiguity gate and expired-window refusal. Do not substitute the broader built-in CPU profile ([approved scope](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6622812245)).

## Assumptions to re-verify

Recheck native checkouts, preserved dirty/untracked files, Excel/process state, all frozen hashes, active trace-session scope, analyzer availability, and available time before proposing another run. The X1 docs checkout must be safely synchronized to this remote handoff without overwriting unrelated edits; do not assume a cloud push updated Windows.

No raw trace exists from the blocked attempt, no symbol coverage has been demonstrated, and the idle elevated shell is not proof of recording readiness. Reconfirm any future execution window and relevant capture/symbol/elevation permissions rather than treating this handoff as authorization ([final result](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6623012152)).

## Open questions and blockers

The retained session's initiating owner/purpose and relevant kernel/stack scope remain unknown. Its reported user-event providers include CLR LoaderKeyword `0x8`, PowerShell, and unresolved GUID `B1C7A9E1-2E44-4C7F-9B30-6FA52F8A7711`; the available output does not expose a definitive kernel/system collector mapping or sampled-profile/stack configuration ([sanitized provider report](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6623012152)).

That readiness ambiguity is separate from the unknown Excel shutdown cause. No evidence currently identifies Defender, COM table traversal, GC waiting itself, or any particular driver/function as the root cause ([v2 measurement](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6622021455), [final readiness result](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6623012152)).

## Immediate next action

In a new bounded read-only task, determine whether registered provider/session metadata can establish the unresolved session's actual scope and compatibility; return a supported conclusion or precise remaining unknown before requesting any new execution window. Do not repeat the already negative WPR probes or start a recording as a compatibility test.

## Critical references

- **Authority:** `0 plan docs/xlam_invSys/CURRENT_SPEC.md` → `invSys-Design-v4.11.md`, D13 and D18-REPLAY-01; `expert guidance docs/CURRENT.md` → `022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md`; `invSys-Controls-v1.md`. These remain unchanged by this wrap-up.
- **Frozen manifest:** the authoritative original is `invSys_fork/reports/runtime/perplexity-shutdown-isolation-source/frozen-input-pins.json`, SHA256 `F2D2768D8B9B430E2B0D823243E32F75C973E42778893A8C6832D6B16C174B73` ([approved pins](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6622812245)). The versioned companion `last handoff/218 Shutdown Input Pins.json` has the same parsed records but different serialized bytes: its cloud-checkout SHA256 is `989631D82E0F0569C2E42D7E5F1B6B36D7956A7D57E9A662C44F18A9C9C69869`. Wrap-up checks confirmed parsed equality and normalized-text equality; never substitute the companion for the original, normalize runtime files, or regenerate pins to make a mismatch pass.
- **Observer/profile:** `Observe-ShutdownProcessLifetime-v2.ps1`, SHA256 `EB756EB93B65F279D64E508DDC9409E30BC39B9F0FDAC149FD6C1F8C2187BF90`; `InvSysShutdownCpuStacks.wprp`, SHA256 `C3EEC9A18EF47940778C8D158F3B793E27FBAF6961671F839D9683D5E6794373`, in the same local support folder ([approved pins](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6622812245)).
- **Local readiness evidence:** relative to `invSys_fork`, `reports/runtime/excel-shutdown-cpu-profile/d20e0cdb91994063848135c81aeb54de/`, including `session-21.txt`, `session-summary.json`, elevated status outputs, and `followup-hashes.json`; keep raw contents local ([evidence inventory](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6623012152)).
- **Reusable v2 runtime evidence:** `reports/runtime/excel-shutdown-timing/bd3bd2e2902049f69cbbd39991aef9e6`, observer request folder `reports/runtime/excel-shutdown-lifetime-v2/6f747395f80045ea9663fed0a5089032`, and assertion report `reports/runtime/slice4be-admin-inventory-entry/27ea2041f9d141e7a47dc33990a5b666/red.json` ([v2 report](https://github.com/soetrain/invSys_fork/pull/7#issuecomment-6622021455)).
