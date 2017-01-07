#!/usr/bin/env python
# -*- coding: utf-8 -*-

import untangle
import os

timestamps_file = os.environ['TIMESTAMPS_FILE']
timestamps = untangle.parse(timestamps_file)

for stamp in timestamps.TimeList.Date:
    print stamp["value"]
    for time in stamp.Time:
        print time.cdata
   
