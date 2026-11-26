# Changelog

- 2026-11-26 `GenericUpdater` FIX - Verify requested update values by ID after a write timeout and report ambiguous creates.
- 2026-11-26 `stdSharepointList` FIX - Read actual multipart operation results and preserve the extra PATCH used by the SharePoint workaround.
- 2026-11-26 `GenericUpdater` DEV - Add a two-user SharePoint batch experiment with individual controls.
- 2026-11-26 `Repository` DEV - Adopt the stdVBA agent guidance, including reuse of one hidden Excel instance per test run and focus recovery after VBA error dialogs.

- 2026-11-26 `GenericUpdater` FEATURE - Recover item IDs after create timeouts using an optional `__create_recovery_key__` text field.
- 2026-11-26 `GenericUpdater` FIX - Exclude rows with unresolved person or multi-person fields and submit the remaining creates and updates. A rejected row leaves its ID available for a later valid update.
- 2026-11-26 `stdSharepointAuthenticator` FEATURE - Prefill the Microsoft sign-in account field from the current classic Outlook profile.
- 2026-11-26 `GenericUpdater` FIX - Clear the output sheet at the start of every upload.
- 2026-11-26 `stdSharepointAuthenticator` BREAKING - Projects that import the form separately could previously compile it with the older stdWebView alone for browser hosting. When replacing the form, also replace stdWebView.cls and import stdOutlook.cls from this example's libs folder. The distributed workbook includes both classes.
- 2026-11-26 `GenericUpdater` FEATURE - Resolve distinct email addresses in SharePoint batches and report lookup errors on the affected rows.
- 2026-11-26 `GenericUpdater` NOTE - Add a Documentation sheet to the workbook covering setup, cell values, field types, timeouts and the create recovery key.
