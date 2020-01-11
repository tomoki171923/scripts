#!/bin/sh

LANG=C

#--------
# Initialization
#--------
INTERVAL=1
FILE_NAME="./`date +%Y%m%d_%H%M%S`_`hostname`_$1"

#--------
# CPU
#--------
echo "time,CPU,%user,%nice,%system,%iowait,%steal,%idle" >> ${FILE_NAME}_cpu.csv
sar -P ALL ${INTERVAL} | awk '$0' | awk '$2!="CPU"{print}' | awk 'NR >= 3' | tr -s ' ' ',' >> ${FILE_NAME}_cpu.csv &

#--------
# MEMORY & SWAP
#--------
echo "time,r,b,swpd,free,buff,cache,si,so,bi,bo,in,cs,us,sy,id,wa,st" >> ${FILE_NAME}_memory.csv
vmstat -n ${INTERVAL} | awk 'NR>2{print(strftime("%H:%M:%S"),$0);fflush();}' | tr -s ' ' ',' >> ${FILE_NAME}_memory.csv &

#--------
# PROCESS
#--------
echo "time,Rank,RSS,COMMAND/CMD,PID">> ${FILE_NAME}_process.csv
HEAD=20
while true :
do
  ps alx  | awk '{printf ("%d,%s,%d\n", $8,$13,$3)}' | sort -nr | head -${HEAD} | awk '{print NR, $0}' |awk '{print(strftime("%H:%M:%S"),$0);fflush();}' | tr -s ' ' ',' >> ${FILE_NAME}_process.csv &
  sleep ${INTERVAL}
done &

#--------
# GPU
#--------
nvidia-smi --query-gpu=index,timestamp,name,uuid,serial,utilization.gpu,utilization.memory,memory.total,memory.free,memory.used,clocks.sm,clocks.mem,clocks.gr --format=csv >> ${FILE_NAME}_gpu.csv -l 1 &

