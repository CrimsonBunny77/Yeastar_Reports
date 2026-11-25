# Yeastar Transfer Report

Streamlit web app that combines Yeastar PBX CDR exports (`.xls` / `.xlsx`), identifies
inter-department call transfers, calculates statistics and produces a PDF report
(`Transfer_Report_YYYY_MM.pdf`).

## Folder structure

```
yeastar_transfer_report/
├── app.py                     # Streamlit UI
├── department_mapping.xlsx    # Extension -> Department (edit this, no code changes needed)
├── requirements.txt
├── README.md
├── make_mapping.py            # Regenerates the sample mapping file
├── extension_names.xlsx       # Extension -> name (used by Extension call stats)
├── make_extension_names.py    # Rebuilds extension_names.xlsx from a Yeastar user export
├── Que_Mapping_Table.xlsx     # Queue number -> queue name (Queue overflow report)
├── CDR_Example/               # CDR exports that can be selected in the Queue overflow report
├── tests/test_queue_overflow.py
└── transfer_report/
    ├── config.py              # Constants, column names, regexes
    ├── errors.py              # Friendly exception types
    ├── mapping.py             # Mapping loader / template writer
    ├── loader.py              # Reads one CDR file (dtype=str) and parses transfers
    ├── processor.py           # Combine, map, flag, aggregate, audit
    ├── charts.py              # Plotly charts
    ├── queue_mapping.py       # Queue mapping loader (Queue overflow report)
    ├── queue_overflow.py      # Queue visits, routing type, outcomes, exceptions
    ├── queue_report.py        # Queue overflow PDF + Excel
    └── pdf_report.py          # ReportLab PDF
```

## Install (Python 3.12+)

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Run

```bash
streamlit run app.py
```

Open the URL shown (default http://localhost:8501), drop your CDR files, click
**Generate Transfer Report**, then **Download PDF**.

## Department mapping

`department_mapping.xlsx` has a **Mapping** sheet with columns `Extension` and `Department`
(plus an optional **Departments** list sheet). **The extensions shipped in the file are
placeholders. Replace them with your real extensions.** Format the Extension column as
text so leading zeros (e.g. `031`) are kept. Add or rename departments freely; no code changes
are required. You can also upload a different mapping file from the sidebar.

### When an extension moves department part-way through a month

The mapping has an optional third column, **`Effective From`** (a date such as `2026-10-22`; blank means "from the
beginning"). To record a move, **keep the old row and add a new row** for the same extension with the new
department and the first day it applies. A row applies until the next row for that extension starts, so there is no
end-date column.

| Extension | Department    | Effective From |
|-----------|---------------|----------------|
| 078       | Customer Care |                |
| 078       | Complaints    | 2026-10-22     |

Every report then attributes a call to the department the extension was in **on the date of that call** (the change
takes effect at 00:00 on the effective date):

* **Transfer report** - each side of a transfer is mapped at the transfer time, so a transfer from ext 078 to
  Customer Care is a same-department transfer (excluded) before 22 Oct and a Complaints -> Customer Care transfer
  from 22 Oct.
* **Department call list / Extension call stats / Extension performance detail** - the extension appears under both
  departments, each with only the calls from the dates it belonged there. An extension is not listed for a
  department it left before the report month started.
* Each report shows a notice naming the department changes that fall inside the report month.
* Calls from before an extension's first *dated* row are treated as unmapped, so give an extension's first row a
  blank date unless it really did not exist before that date.
* Two rows for the same extension with the same date (or both blank) are reported as a warning and the first is used.
  A value in the column that is not a date stops the load with a message naming the row.
* A mapping with no `Effective From` values behaves exactly as before.

## Department call list (second report type)

Choose **Department call list** at the top of the page, upload the same CDR files, pick a department and click
**Generate Call List**. You get a PDF and an Excel (.xlsx) download.

* **Outbound:** every call made from one of the department's extensions (extensions come from the mapping file),
  with the number dialled and the call status.
* **Inbound:** every call from an outside number that was *answered* by one of the department's extensions
  (including calls transferred in), with the caller's number. Internal calls are excluded.
* Both outputs contain, per direction, a list of distinct numbers (calls, answered, last call, talk time) and the
  full call log (date/time, extension, agent, number, status, total and talk time, call ID).
  The PDF shows at most 10,000 rows per table; the Excel file always has every row.
* Talk time is the "Handling Duration" for outbound calls and the answered agent leg's "Pure Talk Duration" for
  inbound calls.
* The **Unmapped extensions** tab lists people's extensions seen on these calls that are not in the mapping file
  (IVRs and queues are left out), because their calls cannot be attributed to a department.
* **All departments:** choose "All departments" in the dropdown for one combined list covering every mapped
  extension. Call logs gain a **Department** column, each number shows how many times it was dialled / came in and
  which departments were involved, and a **By department** summary is added (PDF summary page, app tab and an
  Excel sheet).
* Columns are located by their header names (not fixed positions), so Yeastar exports with extra summary columns
  such as "Second Participant" / "Last Participant" are handled.
* This report is separate from the transfer report: it lives in `calls.py` and `calls_report.py` and does not
  change any transfer-report logic.

## Extension call stats (third report type)

Choose **Extension call stats**, upload the CDR files, pick a department and click **Generate Extension Stats**
(PDF + Excel). For each extension of the department, with its name: Total Calls, Answered, No Answer, Busy,
Voicemail, Abandoned, Transfers Not Completed and Answer % (Answered / Total Calls), plus a TOTAL row. The same table is also given for
**Inbound** and **Outbound** calls separately (and **Internal** if you tick *Include internal calls*).

* **Counting rule (same as the PBX's own "Extension Call Statistics" report):** every call leg the extension was on,
  as caller or as callee, counts once under that leg's own status, including the extension's own voicemail-box leg
  ("Voicemail <name><ext>"). Queue, IVR and ring-group legs are not extensions and are ignored. Checked against a PBX
  report: an extension with complete data matched on every column.
* **Inbound / Outbound / Internal** tables = the type of the call the leg belongs to. *Include internal calls* is on
  by default because the PBX's "All" includes them.
* **Only calls inside the selected report month are counted.** If some days of the month have no calls at all
  (an incomplete export), a warning names the missing days, because the totals will then be lower than a PBX report
  for the whole month.
* Every extension mapped to the department is listed, even with zero calls (tick *Hide extensions with no calls*
  to drop them).
* **Names** come from `extension_names.xlsx` (Extension, Name), built from your Yeastar user export; if an
  extension isn't in it, the name shown in the CDR is used. To refresh the names permanently run
  `python make_extension_names.py <export.csv>`, or upload an export in the sidebar for one session. Only
  extension numbers and names are read from the export. It also holds registration passwords and voicemail PINs, so
  keep the original file private and don't leave it on the server.

## Queue overflow report (fourth report type)

Choose **Queue Overflow Report**, then in the sidebar upload (or keep the default) queue mapping file
(`Que_Mapping_Table.xlsx`: a queue number column and a queue name column; headers are matched loosely). Upload CDR
files, or select exports stored in `CDR_Example/`, pick one or more queues and click **Generate Queue Overflow Report**
(PDF + Excel). Nothing about the queues is hard-coded.

* The figures are built to match the PBX's own **Queue Performance** report (checked against a PBX export for one
  queue over a full month: total, answered, missed, abandoned, ring/talk/hold times and agents all agreed row by row).
  **Summary by queue**: Total Calls (every queue visit), Answered / Missed / Abandoned (how each visit ended), the
  rates, Unique Calls (distinct Call IDs), calls with a recorded agent answer and calls overflowed out. A second
  table gives unique calls by final status. **Routing type** table: queue visits by how they were reached.
  **Call detail**: one row per queue visit in the PBX's layout (Call ID, Time, Caller, Agent, Status, Ring / Talk /
  Hold Duration) plus routing columns, filterable by date, queue, status and routing type, with the Call ID and source
  file on every row. **Exceptions**: mapping problems, blank statuses, duplicates, undetermined routing, unmatched
  answering agents, and visits that were not counted because they rang too briefly.
* **Definitions** (also shown in the app and in the files):
  * *Automatic overflow*: the call entered the queue within 2 seconds of ringing out of its previous queue visit, with
    no transfer in between.
  * *Manual transfer*: an agent transfer (blind or attended) event precedes the queue entry. Never counted as overflow.
  * *Direct*: reached the queue from an IVR/inbound route. *Undetermined*: the events do not allow a decision.
  * *Answered by queue agent*: Yes when the queue's own leg was answered by an extension that the CDR shows the queue
    ringing (the mapping lists queues only, so queue members come from the CDR). *Unknown* when it was answered by
    someone else, for example by call pickup.
* **How it counts (as the PBX does):** every entry into a queue is counted, so a call that rings the queue and
  returns counts each time (Unique Calls shows the distinct Call IDs); all call types that reach the queue are
  included (internal calls too, see the Call Type column); a visit's Time is when the call entered the queue and the
  report month uses that time; visits that rang for 1 second or less are not counted (change this in the sidebar to
  match the PBX's "exclude calls within X seconds" settings). Overlapping exports are counted once.
* Tests: `python -m pytest tests/test_queue_overflow.py`.

## Overlapping files

Exports often overlap (for example one file ends on the 27th and the next one starts on the 26th). The app finds
calls that appear in more than one uploaded file, counts each call once (the later export's copy is kept) and says
so: a notice in the app, and a line on the report's audit page listing which files overlap and by how many calls.
Switch **Count each call once when files overlap** off in the sidebar only if you really want the repeats counted.

## Logo

`logo.png` (next to `app.py`) is used on the cover page of **both** report types automatically: centred at the
top, scaled to fit. To change the permanent logo, replace that file and restart the service. In the sidebar you can
upload a different PNG/JPG for the current session, or untick **Show logo on the report cover**.

## Report name and month

Set the **Report name** and **Report month / Year** above the Generate button. The name appears on the
cover page and in the file name (`<Report_name>_YYYY_MM.pdf`); the month is printed on the cover exactly as
selected (it defaults to last month). If the dates in the files point to a different month, a warning is shown.

## Transfers received by department

Below the Executive Summary, the PDF lists the transfers each department *received*, broken down by the
department they came from, with Answered / No Answer / Busy / Voicemail / Abandoned / Transfers Not Completed counts and a
shaded 'Total' row (with Answer %) per receiving department.

Below the Detailed Transfer Matrix, **Transfers Sent by Department** shows the reverse view: for each
sending department, its transfers to every other department with the same outcome counts and a shaded
'Total' row (combined transfers to all other departments, with Answer %).

## Choosing departments

The sidebar lists every department found in the mapping file. Untick any you want left out.
When some are unticked, choose whether a transfer is included when **both** departments are
selected (default) or when **either** is selected. The cover page of the PDF lists the departments included.

## Unmapped extensions

Transfers whose From or To extension is not in the mapping are excluded. After each run the
**Unmapped extensions** tab lists every such extension with counts, and an Excel download
contains the summary, every excluded record, and a "Mapping additions" sheet to fill in and
paste into `department_mapping.xlsx`.  The PDF does not list them (only a count appears on its audit page).

## Processing rules (as implemented)

The Yeastar export is layered: one summary row per call, then one row per call leg, then one row
per event, each block with its own repeated header row. The app reads it by column position and
handles this layout (a `.xls` file that really contains `.xlsx` data is also accepted).

1. Every cell is read as text (`dtype=str`), so `ID` is never converted to a number.
2. Transfers are `operate` **events** whose JSON contains `blind_to` (Blind) or `attended_to` (Attended).
   Attended transfers are also logged on the receiving leg without a `"from"` key; those duplicates are ignored.
3. `TransferFrom` comes from `"from"`, `TransferTo` from `"blind_to"` / `"attended_to"`, and the transfer
   time from the event's start time.
4. **Outcome:** operate events have no `Last Status`, so the outcome is taken from the call legs created for the
   transfer target: same call `ID`, leg "Call To" extension = `TransferTo`, starting at (or within 2 s before) the
   transfer time. If there are several such legs the **final outcome** is used: ANSWERED if the target answered at
   any point (a later leg never downgrades an answered transfer), otherwise the status of the last leg, so "no
   answer / busy, then fell through to the target's voicemail" is counted as Voicemail. Transfers with no such leg
   (the target never rang) are counted under **Transfers Not Completed** and shown as `NO MATCHING CALL LEG` in the Status values tab.
5. Extensions are mapped to departments; rows with an unmapped side are removed.
6. Same-department transfers are removed (`From Department != To Department`).
7. Statuses become flags (Answered, No Answer, Busy, Voicemail, Abandoned, Transfers Not Completed), matched
   case-insensitively and ignoring spaces/underscores.
8. Per department pair: Transfers, Answered, No Answer, Busy, Voicemail, Abandoned, Transfers Not Completed,
   `Answer % = Answered / Transfers * 100`.
9. The report month is the one you select; a warning appears if the dates in the files point elsewhere or if any
   day of that month has no calls at all (an incomplete export).

## Performance notes

Each file is read with only the needed columns and is filtered to transfer rows immediately,
so memory use stays low even for 1,000,000+ total rows. Reading large `.xlsx` files is the slowest
step (Excel parsing is inherently slow); expect a few minutes per million rows.

## PDF charts

Charts are rendered with Plotly via `kaleido==0.2.1`. If image export is unavailable the PDF
automatically falls back to native ReportLab bar charts, so a report is always produced.

## Error handling

Friendly messages are shown for: no files uploaded, unreadable or corrupted files, empty files,
missing columns, a missing or invalid mapping file, and runs that yield no transfers.
Problem files are skipped and listed as long as at least one file is valid.
