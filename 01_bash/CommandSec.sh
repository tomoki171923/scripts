#bin/sh
#############################################
# Argument : Time (Sec)
# User: root
# Return Code : 0 successful termination
#               1 unsuccessful termination
#               255 Initialization ERROR
#############################################
#-----------------------------------------------------------------------#
# Check Data Validation of Argument
#-----------------------------------------------------------------------#
# Check that Argument is an integer .
L_VAL=$1
expr "${L_VAL}" + 1 > /dev/null 2>&1
if [ ! $? -eq 0 ];then
    echo "Argument is Failed"
    exit 255
fi

#-----------------------------------------------------------------------#
# Set variables
#-----------------------------------------------------------------------#
L_LOG_FILE=./CommandSec_`hostname`_`date +%Y%m%d%H%M%S`.log
L_CMD="hostname"

while true;do
    echo "`date +%Y%m%d%H%M%S`,`${L_CMD}`">>${L_LOG_FILE}
    sleep 1
done
