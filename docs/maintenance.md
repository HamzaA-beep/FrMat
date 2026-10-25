# FrMat maintenance

FrMat is intentionally lightweight, but a few good habits help keep it stable over time.

## Regular maintenance

- keep Python and package versions aligned with the current dependency list
- clean out old temporary workspace data when needed
- verify exports against the original source data after major changes
- keep backups of important datasets before applying transformations

## Workspace hygiene

FrMat stores working data in a local SQLite workspace file. It is useful to periodically review the local workspace and remove stale or unnecessary data if disk usage becomes large.

## Upgrades

If you update the app or the dependency set:

1. back up your data files
2. recreate the virtual environment if needed
3. reinstall dependencies via `requirements.txt`
4. test import, filter, and export flows again

## Summary

Because FrMat is local-first, maintenance is mostly about keeping the environment clean, the dependencies valid, and the underlying data safe.
