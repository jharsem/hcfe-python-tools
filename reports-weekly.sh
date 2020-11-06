#!/bin/sh
#
# Weekly Reports for HCF Eyecare
#
/usr/bin/python3 /opt/reports/report-cxunique-weekly.py
sleep 5
/usr/bin/python3 /opt/reports/report-cxutilisation-weekly.py
sleep 5
/usr/bin/python3 /opt/reports/report-masterdata-weekly.py
