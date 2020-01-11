#!/usr/bin/env python
# -*- coding: utf-8 -*-

import subprocess
import datetime
from io import StringIO

class ExecCmd:
    def ecec_cmd(self,filepath):
        f = open(filepath)
        cmds = f.readlines()
        f.close()

        #Set linefeed code
        ln = "/n"
        for cmd in cmds:
            cmd = cmd.replace(ln,'')
            print("------------- start -------------")
            print("commond : " + cmd)
            print(subprocess.check_output(cmd,shell=True))
            print("------------- end -------------")

