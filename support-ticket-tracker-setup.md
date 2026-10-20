# Support Ticket Tracker

## Deploy

1. Run the complete current `supabase-migration-support-ticket-tracker.sql` in the SQL Editor of the same Supabase project used by the portal. It is safe to rerun and upgrades existing import tables to the version 2 staged-import schema.
2. Deploy `support-ticket-tracker.js` and `tracker-multi-filters.js` alongside `index_revised.html` and reload the portal.
3. Open **Project & Task Tracker > Support Ticket Tracker**, immediately above Daily Tracker.

The migration adds independent ticket/import tables and functions. It does not change Project Overview data or tables. No Edge Function or service-role key is needed in the browser.

If the portal reports `Could not find the table public.support_ticket_imports in the schema cache`, the database setup has not been applied to that project or PostgREST has not reloaded its cache. The migration creates `public.support_ticket_imports` and `public.support_tickets` and ends with `NOTIFY pgrst, 'reload schema';` inside its transaction. After executing it, reopen the importer. If needed, run that NOTIFY statement again and reload the portal. The importer's setup error also provides a Download Database Migration action. The browser cannot create PostgreSQL tables using the public Supabase client.

To verify the tables in SQL Editor:

```sql
SELECT to_regclass('public.support_ticket_imports'),
	   to_regclass('public.support_tickets');
NOTIFY pgrst, 'reload schema';
```

## Permissions

- Active Admin and Level 1 users can import, create, edit, and retrieve original source workbooks.
- Active Level 2 and Read-Only users can view/search/filter tickets, export the tickets they can read, and download the blank template.
- Anonymous and inactive users have no access. Supabase RLS enforces these rules.

## Import And Storage

- `.xls` and `.xlsx` workbooks are accepted without template metadata or reformatting. Ticket sheets are detected by RequestID and Subject headers, including after report preambles. Header matching ignores case, spaces, and punctuation.
- RequestID and Subject are mandatory. RequestID is a trimmed, case-sensitive, immutable text identifier with a database primary key. Leading zeros are retained; existing tickets are never overwritten by import.
- The 23 listed columns and additional columns are stored without application length limits in `support_tickets.fields`. Original cell values, formats, formulas, header spelling, sheet names, and row numbers remain in source metadata.
- The complete parsed workbook, including summary sheets, and the original file bytes are retained in `support_ticket_imports`. The original workbook can be downloaded from an imported ticket by an authorized editor. Repeated labels across sheets remain unchanged.
- Before preview/confirmation, the importer checks the version 2 database setup and the actual PostgREST table/column endpoints using zero-row reads, then refreshes existing tickets. A missing table or stale schema cache leaves rows Not processed at 0%, rather than labeling every row failed. No import or ticket writes can start until these checks succeed. After applying the migration, use Check Database Again to retry with the same workbook. A workbook with missing mandatory values cannot be confirmed.
- Confirming validates every RequestID in database-backed staging batches of 200. The progress bar uses the same `import-progress-track` and `import-progress-bar` styling as Project Overview, with processed rows, total rows, percentage, and validation/save phase labels. Staging does not insert or update live ticket records.
- After every row is validated, one database transaction inserts all new tickets and records the final report. Existing RequestIDs are skipped, including concurrent duplicates discovered at commit. Existing ticket values, versions, and audit timestamps remain unchanged. The completion summary shows successful, updated, duplicate/skipped, and failed counts; updated is zero because this module deliberately skips duplicates rather than overwriting them.
- Any invalid row or database failure during final commit rolls back all new ticket rows in that import. Failed/cancelled staging is cleared; source-file/job metadata may remain solely as a failed-import audit record, not as partial ticket data. Validation reports list every row; the dialog displays the first 50 and provides a complete report download.
- If a commit response is lost, recovery checks the same import UUID and returns the stored result without inserting again or deleting completed tickets. If the server cannot be reached to confirm the outcome, the dialog explicitly reports it as unconfirmed and offers Recover Import; the database still guarantees all tickets or no tickets. Re-importing the workbook also remains duplicate-safe. Supabase's deployment-specific request and storage limits still apply.
- Ticket edits use a version check to avoid overwriting a concurrent edit. Imported source metadata remains immutable.
- Closing and reopening the importer retains its current detailed preview/completion report and progress. Complete reports are also stored with the database import job and can be downloaded from the dialog.

## Views And Exports

Filter dropdowns on Overview, Programs, Initiatives, Tracks, Milestones, Daily Tracker, and Support Ticket Tracker support searchable checkbox multi-selection. Selected values use OR within a dropdown and AND across dropdowns; All clears that dropdown. Reset clears the current page's selections. Selection state is isolated by page/account and retained during rendering/navigation, without database changes. Sort, pagination, summary-dimension controls, and create/edit form dropdowns remain single-select. Date ranges remain date controls.

Tabular View is the default. Cards and paginated summaries remain available, sharing search/filters. The leading grid/export columns are ID, Subject, Department, Group, Stage, Status, Subgroup, Created Date, Created Time, Requestor, Age in Days, and Ageing Bucket, followed by remaining fields. ID maps to RequestID, Requestor to Requester, and Subgroup to Sub Category where those imported names are used. Missing Department values remain blank rather than being inferred from another field.

Stage is Open when Closure Date is blank and Closed when it is present, independently of Status. A Closure Date input is available in create/edit dialogs. Age is the absolute difference between calendar dates: Created Date to today's local date for Open tickets, or Created Date to Closure Date for Closed tickets. Time-of-day, weekends, holidays, daylight-saving changes, and working-day calendars do not alter the day count. Excel date serials and their original formats are used when available; combined Created Time timestamps supply the creation date if no separate Created Date was supplied.

Ageing Bucket is 0-5, 6-10, 11-20, or More than 20 days. Age counts carry green, amber, orange, or red flag icons respectively. Missing/unparseable dates show a neutral Check dates flag and are excluded from numeric averages. Stage, age, bucket, and flag are computed from current ticket values rather than stored as stale database fields. They recalculate on rendering, filtering, sorting, export, save/reload, and immediately in the edit preview when dates are changed. Raw imported date/Aging values remain preserved in source fields and the original workbook; editing an unrelated field does not replace the creation timestamp with its derived display value.

Summaries cover Request Type, Status, Location, Zone, Category, Priority, Group, Engineer Name, and Ageing Bucket. Open/closed summary counts use derived Stage, and average age uses the derived calendar-day count. Stage and Ageing Bucket are available as filter dropdowns.

Export includes all filtered tickets, the grid's derived/remaining fields, and nine summary sheets, not just the visible page. The old raw Aging column is replaced in the grid/export by Age in Days and Ageing Bucket. Templates retain the original 23 import headers and add optional Department, Created Date, and Closure Date columns; old 23-column workbooks still import as-is. Excel template/export cells are text to retain identifiers and contact numbers. The unchanged original workbook is available separately. No database migration or backfill is required for these derived display fields.

## Verification

Open `tests/support-ticket-tracker.html` in a browser with CDN access. It uses mocked Supabase methods and in-memory PostgreSQL; no live data is changed.

The named `Novel_-_TSD_ ( Oct 19, 2026 09_41 AM ).xls` workbook was not available in the workspace or readable attachments during implementation. Tests use a generated legacy `.xls` with the supplied 23-column structure, extra/reordered columns, number formats, duplicates, missing fields, and a summary sheet. The actual workbook still needs an import-preview verification before production import.
