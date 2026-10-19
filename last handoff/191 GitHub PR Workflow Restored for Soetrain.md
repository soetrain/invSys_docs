# GitHub PR workflow restored for soetrain

## Goal and release outcome

Restore the user's contribution workflow across chats on this computer only:
pull from soetrain, publish feature branches through PRs to soetrain, and preserve
all pending4be-A work.
This housekeeping advances continuity; it does not establish runtime acceptance.

## Current verified state

Last verified2026-10-16 Pacific.

- Code branch `codex/aggregation-authority`, local commit2643a9c0; its AGENTS.md
  workflow change is641383b8. Draft [code PR3](https://github.com/soetrain/invSys_fork/pull/3)
  targets `soetrain/invSys_fork:main` from the matching justinwj feature branch.
- Docs branch `codex/tighten-d8-agg-01`,6fa84fb before this plan/handoff commit.
  Draft [docs PR2](https://github.com/soetrain/invSys_docs/pull/2) targets
  `soetrain/invSys_docs:main` from the matching justinwj feature branch.
- Both `git pull --ff-only origin main` commands reported already up to date.
  Both working branches contain soetrain's latest main. Local main refs match
  origin/main: codeff5db426, docsafcf5d2.
- Docs' old divergent main, with three local-only commits, is preserved locally as
  `archive/main-before-soetrain-sync-764b96d5`; no commits were discarded.
- No Excel runtime, package build, deployment or catalog implementation change
  was performed. The current Core input prototype remains incomplete and unbuilt.
  Final pushes and PR-head equality are required closure checks for this session.

## Decisions and constraints

- User explicitly confirmed the workflow for this computer only. `origin` is the
  matching soetrain repo;
  `fork` is the authenticated contributor's matching justinwj repo. Publishing to
  those contributor branches and submitting upstream PRs is authorized. Do not
  repeatedly ask for fork-destination approval in later chats.
- GitHub CLI default repos are explicitly soetrain. Both repos use
  `remote.pushDefault=fork`, current feature `pushRemote=fork`, `push.default=current`,
  `pull.rebase=false`, `pull.ff=only`. Feature branches track their fork branches;
  fetch origin and merge origin/main explicitly to receive upstream changes.
- Keep unfinished work in draft PRs; soetrain reviews/merges and pulls from his
  own repositories. No direct upstream-main push, force-push, auto-merge or PR merge.
- Both root AGENTS.md files persist the workflow. Code AGENTS.md supplies shared
  project instructions; docs AGENTS.md points to it and repeats the Git workflow.
- This resolves the prior Git publication block. The separate automatic-review
  rejection of the catalog write/save patch remains; this request did not authorize
  that product edit. D8-INV-01's architectural approval remains operative.

## Evidence and traceability

- Native Git fetches/pulls completed for both origins and forks. Native Git status
  was clean before publishing; both initial branch pushes succeeded.
- GitHub API confirms both PRs open/draft, base repository soetrain, base main,
  head repository justinwj and the expected feature branch, auto_merge null.
- Inspected housekeeping diffs, checked whitespace and local pointers. No D13 RED
  is required because no runtime/architectural contract changed.
- All runtime evidence and unfinished gates remain as recorded in handoff190 and
  the maintained Inventory receipt: validated product68949ab0, input108 checks88/20,
  Config baseline50 checks49/1. Current HEAD is not a qualified build.

## Do Not Repeat

Do not push only to a contributor branch and omit the PR. Do not let GitHub's fork
ancestry choose the PR base implicitly. Do not treat an unqualified pull from a
fork-tracking branch as a pull from soetrain. Earlier handoffs' pending Git-push
permission statements are historical and superseded by this user confirmation.

## Assumptions to Re-verify

At resume, fetch upstream, inspect both native Git statuses, PR states and exact
published heads. Local Git/CLI settings are not cloned; do not apply this computer's
PR configuration to another machine. Follow that machine's user-authorized workflow.
Recheck Excel closure before any product build.

## Open questions and blockers

Git publication is authorized and PR destinations are established. The catalog
patch permission, incomplete input/write/result integration and broader4be-A/
Release1 gates from190 remain open; no new product acceptance is implied.

## Immediate next action

After resolving the separate catalog-edit permission, finish the captured-input/write/result integration and require the retained108-check packaged boundary matrix GREEN on a fresh isolated candidate before broader regression gates.

## Critical references

- Both root AGENTS.md files; Plan022; v4.11 D8-INV-01; Controls1.647.
- Handoff `190 D8-INV-01 Config Metadata Diagnosis and Review Block.md` for runtime
  provenance and remaining product work; its Git-block statements are superseded.
- Code `tests/proposals/d8_inv_01_catalog_write_review.md` and adjacent unapplied
  patch; `tests/integration/plan022_slice4be_admin_inventory_entry_results.md` and
  `plan022_slice4be_remaining_acceptance.md`.
- Code PR3 and docs PR2 linked above; no merges or main pushes are authorized.
