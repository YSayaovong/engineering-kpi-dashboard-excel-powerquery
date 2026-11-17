# Engineering KPI Workbook

**Engineering project tracking | Spreadsheet analysis | Configuration visibility**

I developed this workbook example to explore how structured engineering project information can support clearer status reviews and follow-up decisions. My background in mechanical design shapes how I think about project records, revisions and the need for consistent reporting.

## Business question

**How can engineering work be organized so a reviewer can understand project status and work in progress without relying on scattered updates?**

The primary artifact is [Project Tracking.xlsx](Project%20Tracking.xlsx), an Excel workbook example based on a mock engineering log.

## What I want this project to demonstrate

- Organizing engineering project records into a reviewable format
- Thinking critically about status definitions and consistent data entry
- Using spreadsheet reporting to support visibility and follow-up
- Recognizing the difference between a documented dashboard concept and a verified automated reporting pipeline

## Review the workbook

1. Open a copy of [Project Tracking.xlsx](Project%20Tracking.xlsx) in Microsoft Excel.
2. Review the actual sheets, tables, formulas, charts and any pivot tables.
3. Inspect external data connections before refreshing or enabling them.
4. Compare reported values with their source records before drawing conclusions.

I have included [validate_workbook.py](validate_workbook.py), which checks that the XLSX file is structurally readable and contains worksheet XML:

```bash
python validate_workbook.py
```

The [GitHub Actions workflow](.github/workflows/checks.yml) runs this structural check. **Passing this check does not verify spreadsheet formulas, pivot tables, Power Query refreshes or the accuracy of reported metrics.**

## Relevance to engineering analyst roles

I am interested in roles that connect engineering information, configuration records and business processes. This workbook illustrates that direction, while keeping its demonstrated capabilities separate from features that still need verification.

## Current limitations

The repository does not establish that a working Power Query pipeline is present, despite its repository name. I do not claim a live dashboard, a measured 20% reduction in review time, verified cost savings or completed automation. Any future enhancement should be documented only after its workbook behavior is tested.

This is a portfolio example, not a representation of confidential employer project data.
