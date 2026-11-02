# DataGuard — Intelligent Cross-System Data Integrity Auditor

A UiPath RPA project (college PEP project) that compares employee data between two
Excel sources — **Master.xlsx** and **HR_Data.xlsx** — matches records by `EmployeeID`,
detects **Exact Matches**, **Value Mismatches**, **Missing Records**, and **Duplicate
Records**, assigns a **severity** (LOW / MEDIUM / HIGH), and writes an **Audit_Report.xlsx**
with a results sheet and a summary sheet.

No Flask, no Python, no external API, no external database — everything runs inside
UiPath Studio using modern, non-deprecated Excel/System activities, and is designed to be
called from a UiPath Apps dashboard.

---

## 1. What's in this project

```
DataGuard/
├── Main.xaml                     ← entry point / orchestrator
├── project.json                  ← project manifest (dependencies, entry-point args)
├── Workflows/
│   ├── ReadMaster.xaml           ← reads Master.xlsx into a DataTable
│   ├── ReadHR.xaml               ← reads HR_Data.xlsx into a DataTable
│   ├── CompareData.xaml          ← the comparison engine (core logic)
│   ├── CalculateSummary.xaml     ← counts totals for the dashboard
│   └── GenerateReport.xaml       ← writes Output/Audit_Report.xlsx
├── Input/
│   ├── Master.xlsx                ← sample data (matches the spec you gave)
│   └── HR_Data.xlsx               ← sample data with 2 intentional mismatches
└── Output/                        ← Audit_Report.xlsx is written here at runtime
```

`Main.xaml` arguments (these are what UiPath Apps binds to):

| Direction | Name | Type | Purpose |
|---|---|---|---|
| In | `in_MasterFilePath` | String | full path to Master Excel file |
| In | `in_HRFilePath` | String | full path to HR Excel file |
| Out | `out_ResultDataTable` | DataTable | row-per-difference audit table |
| Out | `out_TotalRecords` | Int32 | total employees audited |
| Out | `out_ExactMatches` | Int32 | count of fully matching employees |
| Out | `out_Mismatches` | Int32 | count of mismatched fields |
| Out | `out_MissingRecords` | Int32 | count of employees missing from HR |
| Out | `out_Duplicates` | Int32 | count of duplicate EmployeeIDs found |
| Out | `out_HighRisk` / `out_MediumRisk` / `out_LowRisk` | Int32 | severity counts |
| Out | `out_StatusMessage` | String | human-readable result message |
| Out | `out_ReportPath` | String | full path to the generated report |

---

## A. Opening the project in UiPath Studio

1. Unzip `DataGuard.zip` to a folder on your machine, e.g. `C:\Projects\DataGuard`.
2. Open **UiPath Studio (Community 2026.0.x)**.
3. On the Studio start screen, click **Open** → **Open Project** and select the
   `project.json` file inside the unzipped `DataGuard` folder (do **not** select the
   folder itself — select `project.json`).
4. Studio will open the project and show `Main.xaml` in the **Project** panel along with
   the `Workflows`, `Input`, and `Output` folders.
5. Studio will automatically try to restore the NuGet dependencies listed in
   `project.json` (`UiPath.Excel.Activities`, `UiPath.System.Activities`). If it doesn't
   restore automatically, go to **Manage Packages** (top ribbon) → **All Packages** →
   click **Restore All Packages**. This requires an internet connection the first time.
6. Double-click `Main.xaml` to open it in the designer. You should see the full sequence:
   file checks → Try/Catch → Read Master → Read HR → Compare Data → Calculate Summary →
   Generate Report → set status message.

**If Studio flags a validation warning on first open** (common with any hand-authored
XAML, since Studio re-serializes files in its own format once you save): open the
workflow that's flagged, right-click the highlighted activity, and Studio will usually
offer to auto-fix it, or you can delete and re-drag the same activity from the Activities
panel using the property values documented in Section F below. This project was built and
schema-checked against UiPath's standard modern-activity XAML format, but Studio's
internal activity versions can shift slightly between builds — see **Troubleshooting**
at the end of this file for the most common fixes.

---

## B. Connecting the process to UiPath Apps

1. In Studio, first **publish** the process: ribbon → **Publish** → **Publish Process**
   → choose your feed (Orchestrator Tenant Processes feed, or Orchestrator Personal
   Workspace feed if you're testing solo) → **Publish**.
2. Go to **Orchestrator** (cloud.uipath.com or your tenant) → make sure the DataGuard
   process package appears under **Automations → Processes** (or in your Personal
   Workspace).
3. Open **UiPath Apps** (**apps.uipath.com**, or the **Apps** module inside Orchestrator).
4. Click **Create new app** → **Blank App**.
5. In the App's left panel, click the **Automations** tab → **Add Automation** → select
   the **DataGuard** process you just published.
6. Studio Apps will read `Main.xaml`'s arguments automatically (because `project.json`
   declares them as the entry point's `input`/`output`) and list them as bindable
   fields — this is what Section C and D below use.

---

## C. Configuring App inputs (Master / HR file upload controls)

1. Drag a **File Upload** control onto the canvas, label it **Master Data**.
2. Drag a second **File Upload** control, label it **HR Data**.
3. Drag a **Button** control, label it **Run Audit**.
4. Select the **Run Audit** button → in its **Events** panel, add an action:
   **Run Automation** → choose the **DataGuard** automation.
5. In the automation's input mapping, set:
   - `in_MasterFilePath` = the **Master Data** file-upload control's file path/URL
     output (UiPath Apps exposes this as `<ControlName>.file.path` or similar — pick it
     from the field picker, don't type it by hand).
   - `in_HRFilePath` = the **HR Data** file-upload control's equivalent field.
6. (Optional but recommended) Add a **Rule** on the Run Audit button so it's disabled
   until both files are selected — e.g. `IsEmpty(MasterUpload.file) OR IsEmpty(HRUpload.file)`.

---

## D. Configuring App outputs (dashboard, table, status)

After the **Run Automation** action for the Run Audit button:

1. **Summary cards** — add 8 **Text**/**Card** controls and bind their text to:
   `DataGuard.out_TotalRecords`, `DataGuard.out_ExactMatches`, `DataGuard.out_Mismatches`,
   `DataGuard.out_MissingRecords`, `DataGuard.out_Duplicates`, `DataGuard.out_HighRisk`,
   `DataGuard.out_MediumRisk`, `DataGuard.out_LowRisk`.
   (`DataGuard` here is whatever name you gave the automation in step B.5 — the field
   picker will show these automatically once the automation has been run once in
   Preview mode.)
2. **Results table** — add a **Table** control, set its data source to
   `DataGuard.out_ResultDataTable`, and map columns: `EmployeeID`, `Field`,
   `MasterValue`, `HRValue`, `Status`, `Severity`. Optionally add a conditional style
   rule on the `Severity` column (HIGH = red badge, MEDIUM = amber, LOW = green).
3. **Status message** — add a **Text** control bound to `DataGuard.out_StatusMessage`.
4. **Generate Audit Report button** — add a second button labeled **Generate Audit
   Report**. Since the report is already written to disk by the same automation run
   (`out_ReportPath`), the simplest option is:
   - Bind a **Text**/**Link** control to `DataGuard.out_ReportPath` so the user can see
     where the file was saved, **or**
   - Add a **Download File** action on the button pointing at `DataGuard.out_ReportPath`
     if your UiPath Apps runtime instance exposes file-download actions for Automations
     with local/attended execution.
5. Click **Preview** (top-right) to test the whole flow end-to-end before publishing the
   app.

---

## E. Running the project

**From Studio (fastest way to test the automation itself):**
1. Open `Main.xaml`.
2. Click **Run File** (F6) in the ribbon.
3. Studio will prompt for `in_MasterFilePath` and `in_HRFilePath` — browse to
   `Input\Master.xlsx` and `Input\HR_Data.xlsx` (or type the full paths).
4. After it runs, check the **Locals**/**Output** panel for the returned arguments, and
   open `Output\Audit_Report.xlsx` to see the generated report.

**From UiPath Apps (end-to-end, as your professor will see it):**
1. Open the published DataGuard app.
2. Upload a Master file and an HR file.
3. Click **Run Audit**.
4. Watch the summary cards and results table populate, and read the status message.
5. Click **Generate Audit Report** (or open the link) to access `Audit_Report.xlsx`.

---

## F. How the comparison works

For every row in **Master**:

1. **Duplicate check** — the workflow counts how many rows in Master share the same
   `EmployeeID`. If more than one, a single `Duplicate` row (Severity `HIGH`) is added
   (added only once per duplicate group, not once per repeated row).
2. **Missing check** — it looks for a row with the same `EmployeeID` in HR. If none is
   found, a `Missing` row (Severity `HIGH`) is added and no field-by-field comparison is
   attempted for that employee.
3. **Field comparison** — if a match exists in HR, `Name`, `Salary`, and `Phone` are
   each compared (case-insensitive, trimmed):
   - `Name` differs → `Mismatch`, Severity `MEDIUM`
   - `Salary` differs → `Mismatch`, Severity `HIGH`
   - `Phone` differs → `Mismatch`, Severity `MEDIUM`

   Each differing field becomes its **own row** in the result table — an employee with
   two mismatched fields produces two rows (this matches your spec's E002/E003 example).
4. **Exact match** — if none of the three fields differ and the employee is not a
   duplicate, a single `Exact Match` row (Severity `LOW`) is added.

`CalculateSummary.xaml` then counts rows in the result table by `Status` and `Severity`
to populate the dashboard cards, and `GenerateReport.xaml` writes the result table plus
a Metric/Value summary table to `Output\Audit_Report.xlsx` (two sheets: `AuditResults`
and `Summary`).

**Design choice:** duplicates are checked within Master (the primary/reference source),
per your spec's step ordering ("For every EmployeeID in Master … Check duplicate
EmployeeIDs"). If you also want HR-side duplicates flagged, that's a small addition to
`CompareData.xaml` — ask and it can be added as a second loop over `in_HR`.

---

## Error handling (implemented in Main.xaml)

| Condition | Message |
|---|---|
| No Master file selected | "Please select Master Data file." |
| No HR file selected | "Please select HR Data file." |
| Excel can't be read / any processing exception | "Unable to read Excel file. Details: …" |
| Audit runs clean | "No data inconsistencies found." |
| Audit finds issues | "Audit completed. N data inconsistencies detected." |

---

## Troubleshooting (first open in Studio)

These XAML files were hand-built to match UiPath's standard modern-activity schema
(`UiPath.Excel.Activities` for `WorkbookReadRange`/`WorkbookWriteRange`/`AddDataRow`/
`ForEachRow`, plus stock Workflow Foundation activities for `Sequence`/`If`/`Assign`/
`TryCatch`/`InvokeMethod`). If your exact Studio build's package version renders
anything as an "unknown activity":

1. **Manage Packages → Restore All Packages** first — most issues are just a missing
   NuGet restore.
2. If one specific activity (e.g. `Read Range Workbook`) shows as broken: open the
   Activities panel, search for **"Read Range Workbook"** (Excel category), drag a
   fresh copy next to the broken one, set its **WorkbookPath** to `in_FilePath`, its
   **DataTable** output to `out_DataTable`, delete the broken one, and reconnect the
   sequence.
3. If `Add Data Row` shows as broken: search **"Add Data Row"** in Activities, set
   **DataTable** to `dtResult`, and set **ArrayRow** to the same `New Object() {...}`
   expression shown in that step's original `DisplayName` in the file (each is
   self-descriptive, e.g. "Add Row: Salary Mismatch").
4. Any workflow can always be opened, inspected, and rebuilt activity-by-activity using
   Section F above as the logic spec — nothing here depends on external services, so a
   rebuilt activity behaves identically to the original.

## Requirements recap (already followed)

- UiPath Studio Community 2026.0.x, Windows 11 ✅
- Modern activities only — `WorkbookReadRange`/`WorkbookWriteRange` (no `Excel
  Application Scope`), `ForEachRow`, `If`, `Assign`, `AddDataRow`, `InvokeMethod`,
  `TryCatch` ✅
- Excel files as the only data source ✅
- No Flask, no Python, no external API/database, no LLM calls ✅
