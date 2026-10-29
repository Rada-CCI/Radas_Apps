# CLAUDE.md

Guidance for Claude when working in this repo.

## Paid GitHub resources: ask first

GitHub Actions minutes, Codespaces, Packages, Git LFS storage and bandwidth, and Copilot requests are shared by the whole organization (in October 2026 automatic Windows builds and Cloudflare deploys used up the month's Actions minutes). Never use, trigger or set up any of them, including adding or changing `.github/workflows` files, unless Brandon asks for that specific use in the conversation. `.claude/settings.json` makes Claude ask before the GitHub tools and commands that use them. Plain git (commits, pushes, files in the repo) costs nothing and is fine.

## Builds

Builds never run on GitHub (Brandon, 2026-10-29), so they use none of the organization's shared Actions minutes. A build that needs Windows is made by hand on Brandon's PC: in `windows-configurator\`, `build_gui.ps1` (PyInstaller) for `CCI_New_PC_Setup_v<VERSION>.exe`, `build.ps1` (ps2exe) for `configure-windows.exe`, then `package_release.ps1` for the release zip. The `release.yml` workflow that built and released on a `v*` tag was removed on 2026-10-29. Don't move it to GitHub Actions or any other cloud runner, even if that looks easier; a Claude session runs on Linux and can't make it, so say what needs building and let Brandon run it.

## Read-only repos

PublicDocumentation and every Stagehand and Spikemark repo are read only on GitHub (Brandon, 2026-10-29): read them as sources, and experimenting on a local branch in a clone is fine, but never push anything to them (no branches, commits, tags, pull requests, issues, comments, settings or workflow files on GitHub). Changes they need go to Brandon as a note (in Rada-Training-and-Docs: `REVIEW-NOTES.md` or `ACTION-ITEMS.md`).
