#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import untangle
import dotenv

from enum import Enum
from openpyxl import load_workbook


class EntryType(Enum):
    MORNING_IN = 0
    MORNING_OUT = 1
    AFTERNOON_IN = 2
    AFTERNOON_OUT = 3


def write_entry(date, time, type):
    if type == EntryType.MORNING_IN:
        print "Morning: " + time
    elif type == EntryType.MORNING_OUT:
        print "Noon: " + time
    elif type == EntryType.AFTERNOON_IN:
        print "Afternoon: " + time
    elif type == EntryType.AFTERNOON_OUT:
        print "Bye: " + time
    else:
        print "UUPs"

dotenv.load()
timestamps_file = os.environ.get("TIMESTAMPS_FILE")
target_excel = os.environ.get("TARGET_EXCEL")

print "Using timestamps file '{0}' and target excel '{1}'".format(timestamps_file, target_excel)

timestamps = untangle.parse(timestamps_file)
wb = load_workbook(filename=target_excel)
sheets = wb.get_sheet_names()

for day in timestamps.TimeList.Date:
    date = day["value"]
    print "New day: " + date
    i = 0
    for stamp in day.Time:
        time = stamp.cdata
        write_entry(date, time, i)
        i += 1
        if i > 3:
            break
    print
