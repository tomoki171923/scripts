#bin/sh
#############################################
# Argument : List File Name
# User: root
# Return Code : 0 successful termination
#               1 unsuccessful termination
#               255 Initialization ERROR
#############################################
#-----------------------------------------------------------------------#
# Initialization
#-----------------------------------------------------------------------#
# Set a number of arguments
L_ARGC=1
L_CONF_LIST=lsCheck.lst
L_LOG_FILE=./lsCheck_`hostname`_`date +%Y%m%d%H%M%S`.log

# Check data validation of argument
if [ ! "$#" -eq 0 ]; then
    if [ "$#" -ne "${L_ARGC}" ]; then
        echo "Argument is Failed"
        exit 255
    else
        L_CONF_LIST="${1}"
    fi
fi

# Check that a file can be loaded
if [ ! -r "${L_CONF_LIST}" ]; then
    if [ -f "${L_CONF_LIST}" ]; then
        echo "${L_CONF_LIST} is not loaded"
        exit 255
    else
        echo "${L_CONF_LIST} is not found"
        exit 255
    fi
fi

#-----------------------------------------------------------------------#
# the finction to output log messeage
#-----------------------------------------------------------------------#
funcWriteLog(){
    L_MSG_TYPE=${1}
    L_MSG=${2}
    case "${L_MSG_TYPE}" in
        "ERROR")
            echo "`date +%Y/%m/%d:%H:%M:%S` | Failed           | ${L_MSG}" | tee -a ${L_LOG_FILE}
            ;;
        "SUCCESS")
            echo "`date +%Y/%m/%d:%H:%M:%S` | Successful    | ${L_MSG}" | tee -a ${L_LOG_FILE}
            ;;
    esac
}

#-----------------------------------------------------------------------#
# Main Processing
#-----------------------------------------------------------------------#
while read L_LINE
do
    # Skip a comment row of list file 
    if [[ $L_LINE = *#* ]]; then
        continue
    fi
    
    # Set variables form the `ls` data in list file .リストファイルのls情報を変数に格納
    L_PERMISSION=`echo ${L_LINE} | cut -d ',' -f 1`
    L_OWNER=`echo ${L_LINE} | cut -d ',' -f 2`
    L_GROUP=`echo ${L_LINE} | cut -d ',' -f 3`
    L_PATH=`echo ${L_LINE} | cut -d ',' -f 4`

    if [ -e ${L_PATH} ]; then
        # Set variables form the `ls` data in the server
        L_SV_PERMISSION=`stat -c%a ${L_PATH}`
        L_SV_OWNER=`stat -c%U ${L_PATH}`
        L_SV_GROUP=`stat -c%G ${L_PATH}`
    
        # Compare permisson
        if [ ${L_PERMISSION} = ${L_SV_PERMISSION} ]; then
            L_MSG="${L_PATH} 's PERMISSION (${L_SV_PERMISSION}) is OK."
            funcWriteLog "SUCCESS" "${L_MSG}"
        else
            L_MSG="${L_PATH} 's PERMISSION (${L_SV_PERMISSION}) is Faild. Change ${L_PERMISSION}."
            funcWriteLog "ERROR" "${L_MSG}"
        fi

        # Compare owner
        if [ ${L_OWNER} = ${L_SV_OWNER} ]; then
            L_MSG="${L_PATH} 's OWNER (${L_SV_OWNER}) is OK."
            funcWriteLog "SUCCESS" "${L_MSG}"
        else
            L_MSG="${L_PATH} 's OWNER (${L_SV_OWNER}) is Faild. Change ${L_OWNER}."
            funcWriteLog "ERROR" "${L_MSG}"
        fi

        # Compare group
        if [ ${L_GROUP} = ${L_SV_GROUP} ]; then
            L_MSG="${L_PATH} 's GROUP (${L_SV_GROUP}) is OK."
            funcWriteLog "SUCCESS" "${L_MSG}"
        else
            L_MSG="${L_PATH} 's GROUP (${L_SV_GROUP}) is Faild. Change ${L_GROUP}."
            funcWriteLog "ERROR" "${L_MSG}"
        fi     
        
    else
        # output error if PATH is not found 
        L_MSG="${L_PATH} is not found."
        funcWriteLog "ERROR" "${L_MSG}"
    fi
done < ${L_CONF_LIST}

exit 0
