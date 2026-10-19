# Power Query (M) transformations

These formulas document the transformations implemented in Power BI Service. See the [project README](../README.md) for the full case study.

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
