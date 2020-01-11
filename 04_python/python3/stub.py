#!/usr/bin/env python
# -*- coding: utf-8 -*-

from Cmd import ExecCmd
from Common import Logger

if __name__ == "__main__":

    # Stub of ExecCmd
    filepath = "./cmdlist.txt"
    cmd = ExecCmd.ExecCmd()
    cmd.ecec_cmd(filepath)
    del cmd

    # Stub of Logger
    l = Logger.Logger("./hogehoge.log")
    l.write_log(l.LEVEL_INFO, "hogehoge start")
    l.write_log(l.LEVEL_INFO, "hogehoge stop")
    l.write_log(l.LEVEL_ERROR, "hogehoge error")
    del l

