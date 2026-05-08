This folder stores test-only VBA modules for validating VbaTimSort.

Important:
- If your local XVBA config.json sets vba_folder to vba-files, modules under vba-files/test are also part of the XVBA synchronization target. This config file is typically created locally and is not tracked in this repository.
- Keep these modules out of release workbooks and package outputs intended for end users.
- Use this folder only in verification or benchmark workbooks.

Recommended usage:
- Production modules: vba-files root
- Test modules: vba-files/test
- Before release, verify that test modules are not included in the distributed workbook
