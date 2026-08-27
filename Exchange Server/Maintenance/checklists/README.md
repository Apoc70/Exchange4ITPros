# Checklists

This folder contains operational checklists for Exchange Server environments. Each checklist is provided in both German (`.de.md`) and English (`.en.md`) and is meant to be copied per event, not edited in place, so that a history of completed checks is preserved over time.

## Available checklists

| Checklist | German | English | Purpose |
|---|---|---|---|
| Post-Maintenance | [post-maintenance-checkliste.de.md](./post-maintenance-checkliste.de.md) | [post-maintenance-checklist.en.md](./post-maintenance-checklist.en.md) | Verify that an Exchange environment is fully restored after a maintenance window, CU, security update, hotfix, or third-party update, before officially closing it out. |

## How to use these checklists

1. Copy the relevant checklist file into your own tracking location (e.g. a per-event folder, a wiki page, or an internal ticket) rather than editing the template directly in this repo.
2. Fill in the record table at the top (update type, version/KB, affected servers or DAG, date, performed by).
3. Work through the checkboxes in order. Not every item will apply to every environment; adapt the list to your actual protocols, DAG size, and hybrid configuration.
4. Only mark the underlying maintenance event as closed once every relevant item has been confirmed, not once services simply report as running again.

## Adding a new checklist

When adding a new checklist to this folder:

- Provide both a German (`*.de.md`) and an English (`*.en.md`) version.
- Use GitHub-flavored Markdown checkboxes (`- [ ]`) so items can be tracked visually when rendered on GitHub.
- Include a short record table at the top for the context the checklist was run in (date, version, scope), following the pattern used in the post-maintenance checklist.
- Add an entry to the table above so the checklist is discoverable from this README.

## Related content

These checklists accompany the ENow blog series on Exchange monitoring. The August 2026 article on partial failures discusses why active, checklist-driven verification after maintenance matters more than relying on passive monitoring alone.
