# Generic SharePoint updater

Range-driven Create/Update tool for arbitrary SharePoint lists. Point an Excel table at a site + list title, fill rows, and batch-write in chunks of 500.

## Files

- `src/GenericUpdater.bas` — core API
- `src/mMain.bas` — demo macros (`mainRunGenericUpdater`, `mainFetchFields`)
- `Generic Sharepoint Updator.xlsm` — workbook with `DataToUpdate` ListObject, `SITE_URL`, `LIST_NAME`

## Quick start

1. Set named ranges `SITE_URL` and `LIST_NAME` on `shInput`.
2. Run `mainFetchFields` to ensure `Type`, `ID`, and writable field columns exist on `DataToUpdate`.
3. Fill rows (`Type` = `Create` or `Update`).
4. Run `mainRunGenericUpdater`. Results dump to `shOutput!A1`.

```vb
Dim results As Variant: results = GenericUpdater.UpdateList(SiteURL, ListTitle, PatchData)
Call GenericUpdater.DumpResults(results, Dest)

Dim added As Long: added = GenericUpdater.EnsureTableFields(SiteURL, ListTitle, Table)
```

## Reserved columns

| Column | Meaning |
| --- | --- |
| `Type` | `Create` or `Update` (case-insensitive). If the column is missing, every row is treated as `Update`. |
| `ID` | Required for `Update` (positive SharePoint item id). Ignored for `Create`. On successful Create, the new SharePoint Id is written into the results `ID` column. |

All other headers must match a writable list field by **Title** or **InternalName** (case-insensitive). Unknown headers fail **every** row and skip all writes.

## Cell encoding (all field columns)

| Cell value | Meaning |
| --- | --- |
| Empty / blank | **Skip** — omit from the payload. Update leaves the existing SharePoint value. Create omits the field. |
| `null` (trim, case-insensitive) | **Clear** — send JSON `null` (person/lookup clear `FieldId`; multi-person/multi-choice send empty results). |
| Any other value | Validate and include in the write. |

Safety notes:

- Blank is intentionally non-destructive so unfinished cells do not wipe SharePoint data.
- Required fields cannot be cleared with `null`.
- On **Create**, required fields cannot be blank (they would be omitted).

## Row `Type` behaviour

### Update

- Requires a numeric positive `ID`.
- Duplicate IDs in the same run: first row may update; later duplicates fail validation.
- Only non-blank, non-`null` cells are PATCHed; blank cells leave existing values alone.

### Create

- `ID` is ignored (leave blank).
- Uses `BatchItemsCreate`.
- On Success, results `ID` is filled with the new item id.
- Required SharePoint fields must be provided (not blank, not `null`).

## Field type special cases

Values below assume the column was mapped from list schema (e.g. via `EnsureTableFields` / `FieldsFetchSchema`) so `stdSharepointList` transforms apply.

### Text / Note

- Sent as a string.
- If SharePoint reports `MaxLength`, values longer than that fail validation.

### Number / Currency / Integer

- Must be numeric in Excel.
- Sent as a number.

### Boolean

Accepted cell text (case-insensitive): `TRUE` / `FALSE`, `1` / `0`, `Yes` / `No`.

### DateTime

- Excel dates or parseable date strings.
- Transformed to SharePoint datetime format by `stdSharepointList`.

### Choice

- Must exactly match one of the list’s choice values (when choices are returned by schema).
- Match is case-insensitive for dictionary lookup.

### MultiChoice

- Split on `;` or `,`.
- Each non-empty part must be a valid choice (when choices are known).
- Example: `Red;Blue` or `Red,Blue`.

### Person

Accepted cell values:

| Input | Behaviour |
| --- | --- |
| Email, e.g. `alex@contoso.com` | Resolved via SharePoint `ensureuser` to a site user id, then written as `FieldNameId`. |
| Numeric SharePoint user id | Used directly as `FieldNameId`. |

Not accepted by GenericUpdater validation:

- Display name only (no `@`)
- Arbitrary login strings without `@` (unless numeric)

The email must be resolvable in that tenant/site. Unknown emails fail at resolve/`ensureuser` time (row or batch error).

Create and Update use the same person resolution path.

### MultiPerson

- Same rules as Person per entry.
- Split on `;` or `,`.
- Example: `alex@contoso.com;sam@contoso.com` or `12;34`.

### Lookup

- Must be a numeric lookup item id (not the lookup display text).
- Written as `FieldNameId`.

### File / read-only / calculated / counter / Id

- Not treated as writable by schema filtering.
- `EnsureTableFields` also skips **Hidden** fields.
- Do not add these as patch columns.

### Clearing complex fields with `null`

| Field type | Wire effect |
| --- | --- |
| Person / Lookup | `FieldNameId: null` |
| MultiPerson | `FieldNameId: { results: [] }` |
| MultiChoice | empty `Collection(Edm.String)` results |
| Date / text / other | JSON `null` |

## Call flow (`UpdateList`)

1. `stdSharepointAuthenticator.Create(SiteURL)`
2. `auth.protEnsureAuthenticated`
3. `stdSharepointList.CreateFromTitle(SiteURL, ListTitle, auth)`
4. `list.FieldsFetchSchema`
5. Verify field names — mismatch → fail all rows, no write
6. Validate each row; invalid rows are logged and skipped
7. `BatchItemsCreate` / `BatchItemsSet` in 500-item chunks for valid rows

## Results

`UpdateList` returns a 1-based array (dump with `DumpResults`):

| ID | SuccessFailure | FailureReason |
| --- | --- | --- |

- Input row order is preserved.
- Create Success rows get the new SharePoint Id in `ID`.
- Per-item HTTP failures include SharePoint `error.message.value` when present.
- On an update timeout (including an HTTP 408 or 504 for one item), Excel reads the affected IDs in groups of 25. A row shows **Success** if every submitted field is already at its requested value. This confirms the observed state, including values that already matched before the run. Omitted fields do not participate. If the row is missing, still differs, or cannot be read, its **FailureReason** records the timeout. SharePoint may still finish the write later.
- The updater never resends a timed-out write. If a create response is lost, Excel has no assigned ID to check. Verify whether the item exists before retrying. Other HTTP and validation errors remain failures.
- Every update changeset repeats its first PATCH at the end for SharePoint compatibility. Excel consumes one HTTP result per original item and ignores the additional workaround response when present. Creates send each item once.

## Test batched person resolution

Open a signed-in page on the target SharePoint site. In its browser developer console, run the contents of [tests/ensureuser-batch.js](tests/ensureuser-batch.js). Enter the site URL and two distinct, valid test accounts. Use accounts new to the site if you want to exercise first-time resolution. Like ordinary `ensureuser`, this can register those accounts with the site. The test does not change list items.

The script sends the two `ensureuser` POSTs together, prints the raw multipart response, then makes two individual calls for comparison. A pass requires two successful item responses with distinct IDs matching the individual controls. An outer HTTP 200 alone is insufficient. The uploader continues to resolve people individually.

## `EnsureTableFields` / `mainFetchFields`

- Ensures `Type` and `ID` exist (inserts at the front if missing).
- Appends any missing non-hidden writable fields.
- New headers use **Title**, falling back to **InternalName**.
- Existing columns and data are kept.
- Returns the number of columns added.

## Tips

- Prefer `mainFetchFields` before filling data so person/date/choice columns get the correct field-type transforms.
- For people, use work email or site user id — not display name.
- For lookups, use the related item’s numeric Id.
- Leave cells blank to skip; type `null` only when you intentionally want to clear.
