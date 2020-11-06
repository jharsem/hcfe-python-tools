#!/bin/sh
#
# Monthly Reports for HCF Eyecare
#
/usr/bin/python3 /opt/reports/report-cxutilisation.py
sleep 5
/usr/bin/python3 /opt/reports/report-cxunique.py
sleep 5
/usr/bin/python3 /opt/reports/report-cltrial.py
sleep 5
/usr/bin/python3 /opt/reports/report-masterdata.py
sleep 5
/usr/bin/python3 /opt/reports/report-dispensers.py
sleep 5
/usr/bin/python3 /opt/reports/report-optometrists.py