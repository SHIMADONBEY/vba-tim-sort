This folder stores test-only VBA modules for validating VbaTimSort.

Important:
- Because config.json points vba_folder to vba-files, modules under vba-files/test are also part of the XVBA synchronization target.
- Keep these modules out of release workbooks and package outputs intended for end users.
- Use this folder only in verification or benchmark workbooks.

Recommended usage:
- Production modules: vba-files root
- Test modules: vba-files/test
- Before release, verify that test modules are not included in the distributed workbook
