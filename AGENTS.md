# Agent Instructions

## Current State: Just Split from Monorepo

This repository was just extracted from the `kimhono97/jym_ppt` monorepo.
**Before any regular development work, complete the post-split migration tasks first.**

This project is a Visual Studio Tools for Office (VSTO) PowerPoint add-in that runs standalone. It had essentially no external dependencies in the monorepo either, so the post-split work is light — mainly "verify the build environment and tidy the docs." It is much lighter than the other two polyrepos (`zebulon-provider`, `zebulon-exporter`).

## Session Start: Reading Order

1. **Read `MIGRATION_NOTES.md` first** — split context, what was preserved/lost, post-split checklist.
2. Work through the checklist items in order. Mark each as `- [x]` in `MIGRATION_NOTES.md` when complete.
3. For external tool verification (Visual Studio build, PowerPoint add-in behavior), surface the action to the user and wait for the result before proceeding.
4. **Do not modify this `AGENTS.md` until the migration is complete** — preserves the migration flow.
5. Keep code changes scoped to what the current checklist item requires. No incidental refactors.
6. When in doubt, ask the user before acting.

## Reference Files

- `CLAUDE.md`, `GEMINI.md`: Thin pointers to this file (multi-agent discoverability — see AGENTS.md convention at https://agents.md).

## Documentation Language Policy

Apply this when creating or editing documentation in this repository.

- **Agent-facing documents** (instructions and machine-read context): write in **English** by default.
  - Examples: `AGENTS.md`, `MIGRATION_NOTES.md`, the thin `CLAUDE.md` / `GEMINI.md` pointers, any file primarily consumed by AI coding agents.
  - Rationale: token efficiency and clearer agent context comprehension.
- **User-facing documents** (read by humans on the team): write in **Korean** by default.
  - Examples: `README.md` / `readme.md`, design and architecture notes intended for team members, end-user-facing docs.
- When a single document serves both audiences, prefer English with a brief Korean summary at the top if needed.
- Code comments: follow the existing convention in the file. Default to English for new code unless a file already follows another convention.
- Commit messages: English (consistent with existing history).

## Wrap-Up (only after every checklist item in MIGRATION_NOTES.md is `[x]`)

When the migration is fully complete:

1. **Delete `MIGRATION_NOTES.md`**.
2. **Remove the migration banner from `readme.md`** (the one-line block at the very top).
3. **Rewrite this `AGENTS.md` as a regular project guide**, covering:
   - Project overview (a PowerPoint VSTO add-in providing UDP-based sync, etc.)
   - Build / debug instructions (Visual Studio 2022, NuGet restore, F5 debug, etc.)
   - Key directories and files (`ZebulonVSTO/` solution structure)
   - Deployment (VSTO deployment manifest, ClickOnce, etc., if applicable)
   - Dependencies (NuGet packages, Office and .NET Framework versions)
   - **Carry over the Documentation Language Policy section above into the new AGENTS.md.**
4. **Keep `CLAUDE.md` and `GEMINI.md` as thin pointers** to `AGENTS.md` for multi-agent discoverability.
5. Finalize with a single clear commit, e.g.:
   `chore: complete polyrepo migration, regenerate AGENTS.md`

## Cautions

- Do not accidentally commit VSTO build artifacts (`bin/`, `obj/`, `*.user`).
- The build environment must exactly match (VS workloads, .NET Framework version) for the add-in to behave correctly.
