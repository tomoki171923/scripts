#!/usr/bin/env python
# -*- coding: utf-8 -*-

import subprocess
import datetime
from io import StringIO

class ExecCmd:

    # constructor
    def __init__(self, filepath):
        try:
            f = open(filepath)
            self.__cmds = f.readlines()
            f.close()
        except Exception as e:
            traceback.print_exc()
            sys.exit(1)

    # destructor
    def __del__(self):
        del self.__cmds

    def ecec_cmd(self):
        #Set linefeed code
        ln = "/n"
        for cmd in self.__cmds:
            cmd = cmd.replace(ln,'')
            print("------------- start -------------")
            print("commond : " + cmd)
            print(subprocess.check_output(cmd,shell=True))
            print("------------- end -------------")

