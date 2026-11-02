# MS-Teams-Calls-Reporting

## Reporting script

Use `TeamsReportingSolution.ps1` for the half-month captures and monthly report.
`TaskScheduler-TeamsReportingSolution.ps1` calls that filename; configure its
connection, report-folder, and email settings for the environment.

Run separately for each UPN, using the same UPN, report folder, and export base
name in both tasks.

## Scheduling

| Task | Suggested run time | Reporting period |
| --- | --- | --- |
| `-FirstHalfOfMonth` | 17th at 6:00 AM US Central | Current month, days 1-15 |
| `-SecondHalfOfMonth` | 2nd at 6:00 AM US Central | Previous month, days 16-end, merged with the saved first half |

Schedule after the Power BI dataset refresh. The 16th/1st also work if source
data is available by the run time. Omit `MonthNumber` and `Year` in scheduled
runs so the script selects the reporting month automatically.

Query boundaries are UTC, not Central calendar days. The first half ends at
00:00 UTC on the 16th; the second half ends at 00:00 UTC on the next month's
1st. Manual end dates are also exclusive.

Keep the first-half CSV until the monthly report is complete. Capturing it
midmonth preserves UPN data before CQD's 28-day identifiable-data limit.
Previously generated CSVs are not backfilled automatically.

Call durations retain fractional minutes. CQD queries are not limited to 502
rows, but Power BI/CQD service limits still apply. Retrieval errors stop the
report rather than publishing a known partial result.
Duplicate removal compares all report fields, so different calls with the same
start time remain in the report.

## License
Copyright 2022 Jake Gwynn

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
