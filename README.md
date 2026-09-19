# Excel Data Quality Sample

This repository contains a synthetic workbook that demonstrates a repeatable
CSV/XLSX cleanup workflow. No client data is included.

## What the sample shows

- Duplicate `Record_ID` values are quarantined rather than silently deleted.
- Missing emails, invalid dates, unknown statuses, and non-positive amounts are
  listed on a traceable `Exceptions` sheet.
- Clean records are normalized to consistent names, email casing, dates,
  statuses, and currency formatting.
- `VLOOKUP` checks owner ownership against a controlled lookup sheet.
- Conditional formatting makes failed quality checks visible.
- `SUMIFS`/`COUNTIFS` calculate a weekly owner summary.
- A chart turns the summary into a quick visual report.

## Reproduce

```bash
python build_sample.py
```

The command writes `cleanup_sample.xlsx` and runs structural assertions against
the generated workbook.

## Delivery boundary

Real work starts with a small sample and an agreed rule sheet. Ambiguous values
are flagged for the client instead of being guessed. The final handoff can
include the cleaned workbook, exception report, reusable script, and a short
change log.
