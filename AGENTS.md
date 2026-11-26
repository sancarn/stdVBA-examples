# Agent guidance for stdVBA

These instructions apply throughout this repository. Read [contribution.md](contribution.md) before changing code, and follow the existing style of the affected class/module.

## Style and dependencies

- Keep classes usable by importing a single class wherever possible. Avoid introducing dependencies that force users to import a chain of additional classes or modules. The contribution guide permits shared generic interfaces such as `stdICallable`; keep other dependencies optional where possible.
- Follow the contribution guide's class-first design, `std`/`stdI` naming, predeclared instances, `Create` constructors, `protInit` initialisers, and `TThis` state conventions. Keep code usable across Office applications where feasible.
- Keep `changelog.md` current for every PR as required below.

## Changelog required for every PR

- Every PR submitted to this repository must include an update to `changelog.md` in the same PR covering what changed. Add or revise entries before opening the PR, and update them if the PR scope changes. This requirement also applies to small changes and documentation, tests, tooling, CI, or maintenance-only PRs; use `DEV` or `NOTE` for those where appropriate.
- Review the full diff against the PR base, not just the latest commit or commit titles. Record each distinct user-visible feature and bug fix; group related implementation/follow-up commits rather than listing merges, generated outputs, or formatting commits separately.
- Follow the changelog's dated bullet format, identify the affected class/module/tool, use its documented tags (`BREAKING`, `FIX`, `FEATURE`, `WIP`, `DEV`, `DEPRECATED`, `NOTE`), and include the PR/issue reference when available. Keep entries in chronological order and avoid duplicating existing entries.
- Explicitly check compatibility: removed/renamed members and parameters (including named arguments), required arguments and argument order, return types/array bounds, enum values, defaults, key comparison, callback contracts, error behavior, and platform/dependency changes. A bug fix can also be breaking; add a separate `BREAKING` entry when existing callers must change.
- Every `BREAKING` entry must describe the old and new behavior, who is affected, and the required migration, with a short before/after example where useful. Mark experimental additions as `WIP`; do not present incomplete platform support or experimental APIs as fully supported features.
- Before submitting the PR, verify the changelog against the final diff, preserve historical entries, and check CRLF bytes and encoding. Summarize breaking changes in the PR description (or explicitly state that none were identified). Do not treat changelog coverage as optional because contribution.md says features and fixes are ideally logged.

## Line endings: check before importing

- All text files created or edited must use CRLF (`\r\n`) line endings, with no lone LF or lone CR. Preserve the file's encoding. Do not convert binary files such as workbooks or `.frx` resources.
- Before importing any VBA source into the VBE, inspect the actual bytes of every file to be imported, including generated modules and test helpers. Normalize text to CRLF before import; do not wait for a test failure or CI to fix it.
- The current `.gitattributes` defaults most text files to LF and disables normalization for `.bas`, `.cls`, and `.frm`. Do not assume checkout, an editor, a generator, or Git has produced CRLF. Recheck after tools rewrite files and before finishing changes.
- CI's `.github/workflows/enforce_crlf.yml` covers `.bas`, `.frm`, and `.cls`; it does not replace the local pre-import check. Avoid unrelated repository-wide line-ending changes.

## Office test environment and VBA project access

- For Excel tests, create one hidden, agent-owned Excel instance for the test run. Keep its COM reference across imports, tests and retries. Use disposable workbooks for cases that need fresh VBA state, and restore application-level settings changed by each case. Restart Excel only when a test needs a fresh process, a startup setting has changed, or the instance cannot recover. Leave the user's Excel instances alone.
- Keep Excel and the VBE hidden during setup, execution and cleanup. Import through the workbook's VBProject and run test entry points by workbook-qualified name with Application.Run. Avoid activating workbooks, selecting code panes or opening the VBE to inspect results. Use a separate hidden instance for other Office hosts where supported.
- Programmatic access to the VBA project object model may be disabled. Before importing modules, check access for the actual host application and installed Office version (Excel, Word, PowerPoint, etc.).
- If access is disabled, locate that application's version-specific Trust Center registry setting (`AccessVBOM`) and temporarily enable it for testing as needed. Determine the correct application/version/key rather than assuming Excel or a fixed Office version. Record whether the value existed, its type, and its original value before changing it. Respect managed policy restrictions.
- Apply the setting before starting the test instance, restarting only the agent-owned instance if necessary. Restore the exact original registry state in cleanup, including removing a value that was originally absent. Do not leave VBA project access enabled as a side effect of testing.

## Error trapping and disruptive dialogs

- Before tests, record the VBE's current **Tools > Options > General > Error Trapping** setting. Options include **Break on All Errors**, **Break in Class Module**, and **Break on Unhandled Errors**.
- Use **Break on Unhandled Errors** for tests so errors intentionally handled by VBA do not interrupt execution. The other settings can cause breakpoints and error dialogs even when code contains an error handler.
- Restore the exact original error-trapping setting in cleanup, including on failure. Do not assume the original was **Break in Class Module** or leave the user's debugging preference changed.
- Start a hidden dialog monitor before potentially blocking Office calls. Run it independently of the COM thread so a modal dialog cannot block the monitor. Identify the test instance by its window handle and process ID. If a compile or runtime error appears, capture its text if available and dismiss it through the recognised dialog's own control without activating it or sending global keystrokes.
- On Windows, track the most recent foreground window outside the test instance, starting before Excel launches. If test-owned UI takes the foreground unexpectedly, dismiss the dialog and hide any exposed test window without activating it. If the test instance still owns the foreground, make one best-effort `SetForegroundWindow` call for the previously active window after checking that it still exists, is visible and belongs to the same process as before. If the user has switched windows or Windows declines the request, leave focus alone. Avoid repeated minimise or focus calls during polling or cleanup.
- Treat the dismissed error as a failed import/compile/test, diagnose and fix it, and rerun the affected checks. Dismissing a dialog is not a passing test. Use bounded waits so automation cannot hang indefinitely.
- Use failure-safe cleanup to restore settings and close only agent-owned test documents and instances. Report any setting that could not be restored and any checks that could not run.

## Preserve generated documentation

- Documentation comments are structured input to `Tools/VBDocsGen`, not disposable prose. Preserve descriptions, tagged comments, examples, remarks, and their association with the correct member. Avoid unrelated rewrites or deletions.
- When changing a signature or behavior, update the corresponding documentation in the same change: parameter names/types/defaults, return values, examples, and other affected tags. Add tags for new parameters and members; remove documentation only when the underlying API or documented behavior is intentionally removed.
- The generated artifact is **`docs.json`**. `.github/workflows/vba_docs_gen.yml` currently runs on pushes and generates it using `node ./Tools/VBDocsGen/main.js`.
- For changes to documented VBA code, run that generator and compare the resulting `docs.json` against the baseline before opening a PR. Check for unexpected removed classes, members, parameters, tags, descriptions, and examples, as well as stale signatures. Investigate and fix unintended losses in the source comments or generator input; do not accept a large deletion merely because generation succeeded.
- Keep generated changes limited to the intended API/documentation changes, verify their line endings, and report documentation generation or comparison failures rather than silently skipping them.
