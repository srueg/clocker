#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import untangle
import dotenv
import logging
import excel_helper
import datetime

from datetime import date as sysdate


logging.basicConfig(
    format='%(asctime)s [%(levelname)s] - %(message)s', level=logging.DEBUG)

if any(x in sys.argv for x in ["--help", "-h"]):
    print     '''
    Usage:  --all [-a]  process all enries in timestamps file
            --help [-h] show this help
    '''
    exit()

dotenv.load()
timestamps_file = os.environ.get("TIMESTAMPS_FILE")
target_excel = os.environ.get("TARGET_EXCEL")

logging.debug(
    "Using timestamps file '%s' and target excel '%s'.", timestamps_file, target_excel)

timestamps = untangle.parse(timestamps_file)
excel = excel_helper.ExcelHelper(target_excel)

all = False
if any(x in sys.argv for x in ["--all", "-a"]):
    logging.debug("Processing all days.")
    all = True
else:
    logging.debug("Processing only today (%s)", str(sysdate.today()))

found = False
for day in timestamps.TimeList.Date:
    date = day["value"]
    if all or date == str(sysdate.today()):
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
                d = datetime.datetime.strptime(date, '%Y-%m-%d').date()
                excel.write_entry(d, time, i)
                i += 1

if found:
    excel.save(target_excel)

if not found and not all:
    logging.warn("No stamps found today (%s).", str(sysdate.today()))
