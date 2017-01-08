# -*- coding: utf-8 -*-

import logging

from enum import Enum
from openpyxl import load_workbook

logger = logging.getLogger(__name__)


class ExcelHelper:

    def __init__(self, excel_file):
        self.wb = load_workbook(filename=excel_file)
        sheets = self.wb.get_sheet_names()

    def write_entry(self, date, time, type):
        if type == EntryType.MORNING_IN:
            print "Morning: " + time
        elif type == EntryType.MORNING_OUT:
            print "Noon: " + time
        elif type == EntryType.AFTERNOON_IN:
            print "Afternoon: " + time
        elif type == EntryType.AFTERNOON_OUT:
            print "Bye: " + time
        else:
            logger.error("Unknown entry type on %s!", date)


class EntryType(Enum):
    MORNING_IN = 0
    MORNING_OUT = 1
    AFTERNOON_IN = 2
    AFTERNOON_OUT = 3
