# Polyrepo Migration Notes (zebulon-vsto)

> This document records the context immediately after this repository was split out from the `kimhono97/jym_ppt` monorepo, along with the remaining post-split work.
> **This file will be deleted once every checklist item is complete.** The split itself will be retained only in git history and a one-line banner in `readme.md`.

## Origin and Split Date

- Source monorepo: [`kimhono97/jym_ppt`](https://github.com/kimhono97/jym_ppt)
- Source path: `utils/ZebulonVSTO/`
- Split date: 2026-05-14
- Split method: `git filter-repo --path utils/ZebulonVSTO/ --path-rename utils/ZebulonVSTO/:`

## Why We Split

- `jym_ppt` had grown into a bloated monorepo mixing lyrics data with multiple code projects.
- Team members are collaborators on `jym_ppt` solely to edit lyrics, but they were also being exposed to the code folders.
- Vercel Hobby plan has a constraint: pushes by non-owner collaborators do not trigger deployments. → The code repositories needed to be owned/pushed exclusively by the owner.
- Target end state: `jym_ppt` becomes a lyrics-only data repository (eventually restored to private), and each code project runs independently in its own polyrepo, deployed by the owner.

## What Was Preserved / What Was Lost

- ✅ Preserved: 6 commits that touched `utils/ZebulonVSTO/`, with full file-level change history.
- ⚠️ Lost (intentional): 1 merge commit dropped during rebase (default rebase behavior; **verified no code content was lost**).
- ⚠️ Lost (intentional): All commits from lyrics and other code folders; original commit SHAs (rewritten).
- 📌 Added: GitHub's auto-generated `Initial commit (LICENSE)` sits at the bottom of history (folded in via rebase).

## State Immediately After Split

- Default branch: `main` (LICENSE init + 6 VSTO commits)
- No pre-split `CLAUDE.md` existed in the monorepo for this project (this repo is starting its agent guide from scratch).
- New `AGENTS.md` (this project's agent guide) and thin `CLAUDE.md` / `GEMINI.md` pointers are in place.

## Post-Split Checklist

### A. Build and Runtime Verification

- [ ] Open `ZebulonVSTO.sln` in Visual Studio 2022 (Community v17.5.2+) and confirm it builds cleanly.
- [ ] Verify the required workloads and packages are installed (Office/SharePoint development, VSTO, .NET Framework 4.7.2/4.8).
- [ ] Confirm NuGet packages restore correctly (`Microsoft.Bcl.AsyncInterfaces`, `System.Text.Json`, etc. — see `readme.md`).
- [ ] Confirm the add-in works in PowerPoint 2013+ (Sync features: `alert`, `select`, `showslide`, `hideslide`).

### B. Documentation and Metadata

- [ ] Review whether `readme.md` covers enough build/run guidance. Augment if needed (deployment method, debugging tips, etc.). **Note: `readme.md` is user-facing → keep in Korean.**
- [ ] (Optional) Create the first release/tag to mark the split point (e.g., `v1.0-polyrepo`).
- [ ] (Optional) Review `.gitignore` for VS artifacts (`bin/`, `obj/`, `*.user`).

### C. Wrap-Up (only after every item above is `[x]`)

- [ ] Delete this `MIGRATION_NOTES.md`.
- [ ] Remove the migration banner from `readme.md` (one-line block at the very top).
- [ ] Rewrite `AGENTS.md` as a regular project guide (see the "Wrap-Up" section in `AGENTS.md` for the exact instructions).
- [ ] Keep the thin `CLAUDE.md` / `GEMINI.md` pointers in place for multi-agent discoverability.
