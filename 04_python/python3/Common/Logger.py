#!/usr/bin/python
# -*- coding: utf-8 -*-

################################################################################
# Write messeages on the log file
# Argument    : log file name including the path (requied paramter)
# Outpur      : log file
# Return Code : 0 successful termination
#             : 1 unsuccessful termination
################################################################################

import datetime
import getpass
import sys
import traceback

class Logger:
    # Set constants
    LEVEL_INFO = "INFO"
    LEVEL_DEBUG = "DEBUG"
    LEVEL_ERROR = "ERROR"

    # constructor
    def __init__(self, filepath):
        try:
            self.__file = open(filepath, 'a')
        except Exception as e:
            traceback.print_exc()
            sys.exit(1)

    # destructor
    def __del__(self):
        self.__file.close()
        del self.__file

    # write messeages on the log file 
    def write_log(self, loglevel, msg):
        date = datetime.datetime.now()
        date = date.strftime('%Y-%m-%d %H:%M:%S')
        username = getpass.getuser()
        log = date + " - " + username + " - " + loglevel + " - " + msg

        self.__file.write(log + "\n")

