#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import untangle
import dotenv
import logging
import excel_helper

from datetime import date as sysdate


logging.basicConfig(
    format='%(asctime)s [%(levelname)s] - %(message)s', level=logging.DEBUG)

dotenv.load()
timestamps_file = os.environ.get("TIMESTAMPS_FILE")
target_excel = os.environ.get("TARGET_EXCEL")

logging.debug(
    "Using timestamps file '%s' and target excel '%s'.", timestamps_file, target_excel)

timestamps = untangle.parse(timestamps_file)
excel = excel_helper.ExcelHelper(target_excel)

all = False
if "--all" in sys.argv:
    logging.debug("Processing all days.")
    all = True
else:
    logging.debug("Processing only today (%s)", str(sysdate.today()))

found = False
for day in timestamps.TimeList.Date:
    date = day["value"]
    if not all and date == str(sysdate.today()):
        found = True
        if not all:
            logging.debug("Found stamps for today.")
        if len(day.Time) > 4:
            logging.warn("Too many stamps for day %s.", date)
        else:
            if len(day.Time) < 4:
                logging.warn("Not enough stamps for day %s.", date)
            i = 0
            for stamp in day.Time:
                time = stamp.cdata
                excel.write_entry(date, time, i)
                i += 1

if not found and not all:
    logging.warn("No stamps found today (%s).", str(sysdate.today()))
