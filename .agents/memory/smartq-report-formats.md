---
name: SmartQ report formats
description: The source reports may arrive as detailed OrderLog rows or aggregated point summaries.
---

SmartQ source files are not always the same shape. Detailed exports contain User, User Type, Date, and Credits; aggregated exports can contain Email ID, User Type, Sum of Points, and Sum of Employee Paid, with no Date column.

**Why:** Treating every missing Date as a malformed detailed report caused valid summary exports to be rejected.

**How to apply:** Inspect the headers before rejecting a report. If the aggregated summary shape is present, make the field mapping and any report-date assumption visible to the user instead of silently treating the summary as detailed order data.