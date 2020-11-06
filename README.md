# HCF Eyecare Python Tools

## Reports
* report-cltrial.py
...List of unfollowed up list of trials from two weeks ago
* report-cx_clrx_ex.py
...Glasses customers vs Exam report for BOM purposes with 4 month window
* report-cxex.py
...Glasses customers vs Exam report without script window.
* report-cxunique-weekly.py
...Count of unique members with accounts created this financial year designed to run on a weekly basis.
* report-cxunique.py
...Count of unique members with accounts created this financial year.
* report-cxutilisation.py
...Measure of accounts created, and deliveries done for a given timespan
* report-dispensers.py
...Report of dispenser's stats
* report-masterdata.py
...Report containing several columns of base data
* report-optometrists.py
...Report of optometrist's stats
* report-optom-fte.py
...Report of forward looking FTE for optometrist, generates XSLX
* report-optom-fte.py
...Report of forward looking FTE for optometrist, generates XSLX

All of the reports refer to reportconfig.py

There are also manual running versions of some reports designated by "-manual" at the end of the filename

### Tools
* ./tools/tools-le-checker.py
...Retired tool to check limit enquiry output and diff results
* ./tools/tools-le-checker-re.py
...Regular expression based limit enquiry checker (also needs config.py).
* ./tools-cleareligibility.py
...clears eligibility for members with appointments > 3 days into the future.

### Libraries
* report_email.py
...Refactored email sending code

### Config 
* report_config.py.sample 
...Configration for DB and Email calls
* ./emails/emails_config.py.sample
...Configuration for DB and Email calls for Cx Mailer

## Emails
