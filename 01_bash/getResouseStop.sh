#!/bin/sh

LANG=C

STOP_COMMANDS=("sar" "vmstat" "nvidia-smi" "start.sh")

for i in ${STOP_COMMANDS[@]}; do
  ps aux | grep $i | grep -v grep | awk '{ print "kill -9", $2 }' | sh
done

