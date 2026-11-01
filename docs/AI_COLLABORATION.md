# AI Collaboration Protocol — Claude Code + Codex

This repository is shared between Claude Code and OpenAI Codex. GitHub is the source of truth for committed code and handoffs, not either assistant's chat history.

## Mandatory start-of-task routine
1. Fetch the remote, check the working tree and branch. Never overwrite uncommitted work.
2. Read `AGENTS.md`, `CLAUDE.md` if present, `docs/AI_HANDOFF.md`, README, relevant project docs, current issues and PRs.
3. Inspect actual code and tests; do not treat an earlier handoff as proof that work is implemented.
4. Announce the task, affected files, potential conflicts, and plan before major changes.

## Working rules
- One task per branch (e.g. `codex/<topic>` or `claude/<topic>`), based on an up-to-date default branch. Different assistants working simultaneously must use separate branches and worktrees/directories.
- Do not both edit the same working tree, branch, database, generated files or deployment at the same time.
- Do not force-push or rewrite shared history, discard local edits, or run destructive migrations.
- Commit and push completed steps. Use PRs for review and merging; never presume a push means a PR is merged.
- Run relevant checks; distinguish executed tests from untested assumptions.
- Keep secrets, credentials, patient/client data, production databases and private exports out of Git.
- Follow existing repo-specific instructions; this protocol supplements rather than replaces them.
- If ownership, data safety, interface contracts or status is unclear, document the blocker and avoid destructive actions.

## Mandatory end-of-task handoff
Update `docs/AI_HANDOFF.md` in the same branch/PR as the work:
- Date, assistant, branch, last commit SHA/PR link.
- Goal and what changed, with file paths.
- What was actually tested (exact commands and results).
- Known bugs, limitations, unresolved decisions and next smallest action.
- Any local-only changes or files NOT in Git that another agent must obtain.
- How to start/stop the project and safely continue, if known.
Prefer append-only dated handoff entries; keep the latest status at the top and do not erase historical notes.

## Next assistant
Review the last handoff, verify the branch/PR and commit, sync, inspect diffs and test results, then resume. If the prior assistant has unpushed/uncommitted work, it is NOT available through GitHub; request a handoff or preserve that workspace before proceeding.

## Security
No direct production deployment, outbound vendor orders, irreversible data edits, or secrets exposure without explicit task authorization.
