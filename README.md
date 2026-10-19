# Project 001 — Contract Management Intelligence System

**Power BI Service | Power Query (M) | Data Quality | Interactive Executive Reporting**

![Completed dashboard](screenshots/dashboard-final.png)

## Project overview

This portfolio case study documents the design and development of a **Contract Management Intelligence Dashboard** from 25 fictional contract records. The goal was to turn inconsistent contract data into a traceable, interactive management report supporting contract oversight, financial visibility, and portfolio exploration. The project deliberately does **not** assign contract risk scores.

The work was completed primarily in **Power BI Service through a web browser on macOS**, without Power BI Desktop. It documents both the capabilities of that environment and the practical friction encountered along the way.

## Business task

A contract-management team needs a consolidated view of contract volume, known contract value, status, expiration timing, departmental distribution, vendor concentration, and manager assignments. Before reporting, the source requires normalization and validation so that management does not confuse missing amounts with zero, count duplicate identifiers as unique contracts, or treat naming variants as separate entities.

## Dataset and tools

- **Dataset:** 25 fictional contracts, 12 original source columns; no real organizations' contract records or sensitive data.
- **Source:** `Power_BI_Contract_Data_Cleaning_Lab.xlsx` (included in this project package).
- **Platform:** Power BI Service (browser), browser-based Power Query, semantic model and report editor.
- **Environment:** MacBook Air M2, macOS; no native Power BI Desktop.
- **Language:** Power Query M for repeatable cleaning and validation.
- **Report:** Three executive KPI cards, four charts, and four interactive list slicers.

## Methodology

1. **Assess the raw data.** Inspect contract identifiers, financial text, dates, status categories, departments, vendor names, and manager names.
2. **Preserve raw fields.** Create cleaned and validated columns rather than overwriting original source evidence.
3. **Normalize and validate.** Parse numeric values, handle `TBD` separately from missing amounts, standardize text, convert mixed-format dates, and flag invalid date relationships.
4. **Reconcile portfolio counts.** A KPI exposed 25 rows but only 24 distinct contract IDs. Investigation found two legitimate contracts assigned `CTR-0023`; a targeted fictional-data correction assigned `CTR-0024` to *Consulting Agreement 24* while preserving the original ID.
5. **Build and test visuals.** Replace raw identifiers and labels with cleaned reporting fields, validate aggregates, and test slicer combinations.
6. **Refine presentation.** Apply a consistent executive layout, aligned cards, restrained colors, borders/shadows, and short display labels for narrow slicers.

## Dashboard and validated results

| Indicator | Baseline result | Definition |
|---|---:|---|
| Total contracts | 25 | Distinct `Clean Contract ID` |
| Known contract value | ~$4.15M | Sum of valid `Clean Contract Value`; excludes nulls |
| Active contracts | 19 | Distinct cleaned IDs filtered to Active |
| Pending contracts | 5 | Status category count |
| Expired contracts | 1 | Status category count |
| Expirations in 2025 / 2027 / 2029 | 7 / 10 / 7 | 24 dated expirations; one missing expiration date |

**Interpretation:** Known contract value represents recorded contract amounts, **not actual spending or payments**. Two contract amounts are unavailable and are excluded from the financial sum. Expiration charts omit the one record with no expiration date. These are baseline values; slicer selections change the filter context.

### Interactive report

Four list slicers filter the KPI cards and charts: **Department**, **Contract Status**, **Contract Type**, and **Contract Manager**. The Department slicer uses presentation labels `HR` and `IT` while the department-value chart retains full names. Combined Department + Status filtering was tested successfully.

![Dashboard before slicers](screenshots/dashboard-before-slicers.png)

## Data quality and validation findings

| Finding | Treatment | Why it matters |
|---|---|---|
| Blank and `TBD` contract amounts | Separate validation statuses; numeric nulls | Avoids presenting unknown value as $0 |
| Mixed numeric strings | `Number.FromText` with `try` | Prevents conversion errors from corrupting sums |
| Mixed date strings and Excel serial dates | Culture-aware parsing and Excel date origin | Enables accurate expiration-year reporting |
| One missing expiration date | Explicit date validation flag | Keeps incomplete record visible to QA |
| `ACTIVE` / `active` / `Active` | Capitalize Each Word | Avoids fragmented status categories |
| Department synonyms | Explicit IT mapping | Consolidates departmental financial totals |
| Vendor naming variants | Explicit vendor mappings | Avoids duplicate vendor categories |
| Two different contracts shared `CTR-0023` | Targeted `Clean Contract ID` correction | Restores 25 distinct reporting IDs |
| `J. Lee` and `Jordan Lee` | Consolidated in fictional dataset | Avoids splitting manager portfolio |

**Important distinction:** Row-level checks cannot establish dataset-level uniqueness. The distinct-count KPI exposed an identifier collision that was not visible from the earlier status charts. In a production system, changing a contract ID or merging person identities would require confirmation against an authoritative register.

## Power Query formulas

The following formulas were entered as **Custom Column expressions**, not full standalone queries. Original fields were retained for traceability. `Clean Effective Date` and `Clean Expiration Date` were converted to Date; `Clean Contract Value` was set to fixed decimal/currency. `Contract Status` was standardized with **Transform → Format → Capitalize Each Word**.

### Contract Value Validation

```powerquery
let
    value = [Contract Value],
    cleaned = if value = null then "" else Text.Trim(value),
    numeric = try Number.FromText(cleaned, "en-US")
in
    if cleaned = "" then "Missing"
    else if Text.Upper(cleaned) = "TBD" then "Pending Determination"
    else if numeric[HasError] then "Review Required"
    else "Available"
```

### Clean Contract Value

```powerquery
let
    status = [Contract Value Validation],
    result = try Number.FromText(
        Text.Trim([Contract Value]),
        "en-US"
    )
in
    if status = "Available" and not result[HasError]
    then result[Value]
    else null
```

### Clean Effective Date

```powerquery
let
    raw = [Effective Date],
    cleaned = if raw = null then "" else Text.Trim(Text.From(raw)),
    serial = try Number.FromText(cleaned, "en-US"),
    textDate = try Date.FromText(cleaned, [Culture = "en-US"])
in
    if cleaned = "" then null
    else if not serial[HasError] then
        try Date.AddDays(#date(1899, 12, 30), serial[Value]) otherwise null
    else if not textDate[HasError] then
        textDate[Value]
    else null
```

### Clean Expiration Date

```powerquery
let
    raw = [Expiration Date],
    cleaned = if raw = null then "" else Text.Trim(Text.From(raw)),
    serial = try Number.FromText(cleaned, "en-US"),
    textDate = try Date.FromText(cleaned, [Culture = "en-US"])
in
    if cleaned = "" then null
    else if not serial[HasError] then
        try Date.AddDays(#date(1899, 12, 30), serial[Value]) otherwise null
    else if not textDate[HasError] then
        textDate[Value]
    else null
```

### Contract Date Validation

```powerquery
if [Clean Effective Date] = null then "Missing Effective Date"
else if [Clean Expiration Date] = null then "Missing Expiration Date"
else if [Clean Expiration Date] < [Clean Effective Date] then "Review Required"
else "Valid"
```

### Clean Vendor Name

```powerquery
let
    original = [Vendor Name],
    cleaned = if original = null then "" else Text.Trim(Text.From(original)),
    normalized = Text.Upper(cleaned)
in
    if cleaned = "" then null
    else if List.Contains(
        {"APEX TECH LLC", "APEX TECHNOLOGY LLC"},
        normalized
    ) then "Apex Technology LLC"
    else if List.Contains(
        {"COASTAL OFFICE SUPPLY", "COASTAL OFFICE SUPPLIES"},
        normalized
    ) then "Coastal Office Supplies"
    else cleaned
```

### Clean Contract ID

```powerquery
let
    id = Text.Trim(Text.From([Contract ID])),
    contractName = Text.Trim(Text.From([Contract Name]))
in
    if id = "CTR-0023"
       and contractName = "Consulting Agreement 24"
    then "CTR-0024"
    else id
```

### Clean Department

```powerquery
let
    original = [Department],
    cleaned = if original = null then "" else Text.Trim(Text.From(original)),
    normalized = Text.Upper(cleaned)
in
    if cleaned = "" then null
    else if List.Contains(
        {"INFORMATION TECHNOLOGY", "INFO TECHNOLOGY", "IT"},
        normalized
    ) then "Information Technology"
    else cleaned
```

### Department Display

```powerquery
if [Clean Department] = "Information Technology" then "IT"
else if [Clean Department] = "Human Resources" then "HR"
else [Clean Department]
```

### Clean Contract Manager

```powerquery
let
    original = [Contract Manager],
    cleaned = if original = null then "" else Text.Trim(Text.From(original)),
    normalized = Text.Upper(cleaned)
in
    if cleaned = "" then null
    else if List.Contains(
        {"J. LEE", "JORDAN LEE"},
        normalized
    ) then "Jordan Lee"
    else cleaned
```

## Dashboard design decisions

- **Layout:** Three KPIs across the top, financial charts in the middle, operational charts below, four slicers in a compact right rail.
- **Palette:** Navy `#172B4D`, teal `#008C95`, blue `#2878C8`, slate `#64748B`, and a muted blue-gray canvas. Status categories use teal (Active), amber (Pending), and red (Expired).
- **Visual hierarchy:** White card surfaces, subtle borders/shadows, centered KPI values, concise chart titles, and consistent spacing.
- **Formatting:** The new Card visual's automatic **Category label** (e.g., “Count of Clean Contract ID”) was hidden using **Format visual → Visual → Search `label` → Category label Off**. A custom numeric format retained the `$` prefix while displaying millions with two decimals.
- **Workflow:** Grouping helped move visual elements uniformly, but did not apply formatting to all grouped elements. Duplicating a fully formatted visual is often faster than styling each copy from scratch.

## Browser-based Power BI: lessons and limitations

This project demonstrates that meaningful Power Query transformations, semantic-model reporting, KPI cards, charts, slicers, and interactive filtering can be completed in **Power BI Service** on a Mac. It also exposed practical limitations:

- **No native Power BI Desktop on macOS.** The browser environment was the available path without introducing Windows virtualization.
- **UI discoverability.** Card formatting options varied from older tutorials; searching the formatting pane for `label` proved more reliable than navigating nested menus.
- **Text box friction.** Editing, resizing, recoloring, and selecting a report-title text box were troublesome; the visible page header was deferred rather than destabilizing the report.
- **Visual grouping is not style synchronization.** Grouping moves items together but does not propagate all aesthetic settings.
- **Data/model updates require deliberate refresh or apply steps.** Newly created Power Query fields had to be made available to the semantic model before use in report visuals.
- **Feature differences are version-dependent.** This documents the observed browser workflow, not a claim that every Power BI Service tenant lacks a given feature.

## What I learned

1. **Data quality precedes design.** A polished chart can still be misleading if its categories and identifiers are inconsistent.
2. **Validation must be positive and explicit.** A catch-all “Available” status is not a substitute for successful numeric parsing.
3. **Distinct counts are diagnostic tools.** The difference between 25 rows and 24 unique IDs exposed a real modeling issue.
4. **Keep raw and cleaned fields separate.** This makes transformations inspectable and repeatable.
5. **Missing is not zero.** Financial and date nulls require transparent reporting caveats.
6. **A slicer changes filter context.** Multiple slicers intersect, and visual-level filters (such as Active-only KPIs) continue to apply.
7. **Presentation labels can differ from analytical labels.** `Department Display` permits compact HR/IT slicers while preserving full chart labels.
8. **Interface friction is part of the engineering process.** Documenting workarounds makes subsequent builds faster.

## Limitations and future improvements

- All 25 records are **fictional**, so findings illustrate workflow rather than real contract performance.
- No risk scoring, forecasting, spending analysis, or claims about contract compliance outcomes are included.
- One expiration date and two financial values are unavailable; known-value totals are incomplete by design.
- The report-title text box was not finalized in the browser interface.
- A reproducible exported Power BI report or live public report URL is **not included** in this package; screenshots document the report state.
- Future improvements: confirm model-level data types and formatting, re-run final slicer-reset QA, consider a dedicated date dimension and additional measures, and evaluate Power BI Desktop if a Windows environment becomes available.

## Acknowledgment

Special thanks to my alma mater, **the University of La Verne**. Access to my alumni organizational account provided an opportunity to explore Power BI Service and complete this hands-on learning exercise. This is a personal, independently developed educational project and is **not affiliated with or endorsed by the university**.

## Repository contents

```text
contract_management_project_001/
├── README.md
├── Power_BI_Contract_Data_Cleaning_Lab.xlsx
├── Contract_Management_Intelligence_Project_Report.docx
└── screenshots/
    ├── dashboard-final.png
    ├── dashboard-before-slicers.png
    ├── slicer-selection.png
    └── card-formatting.png
```

**Reproducibility note:** Screenshots and formula documentation are included. The complete Power BI semantic model/report has not been exported into this repository package. A fresh build requires recreating the transformations and visuals in an accessible Power BI environment.