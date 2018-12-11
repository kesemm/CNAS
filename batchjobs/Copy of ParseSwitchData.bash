#!/bin/bash

#	****************************************************************************************
#	*           VERSION CONTROL INFORMATION
#	****************************************************************************************
#	* CVS File:      $RCSfile: ParseSwitchData.bash,v $
#	* Commit Date:   $Date: 2007/07/20 19:52:16 $ (UTC)
#	* Committed by:  $Author: browng $
#	* CVS Revision:  $Revision: 1.7 $
#	* Checkout Tag:  $Name:  $ (Version/Build)
#	****************************************************************************************

# ------------------------------------------------------------------------------------

#####################
# DECLARE VARIABLES
#####################

# Directory Variables

BASEDIR="/cygdrive/d"
WORKDIR=${BASEDIR}/CNA/BIRRDS
DOSWORKDIR="D:\CNA\BIRRDS"
FileType="Switch"

#####################
# DECLARE FUNCTIONS
#####################

funcERRORNOTIFY ( ) {

echo "An Error was encountered while processing the automated import of the" > ${WORKDIR}/blatbody.txt
echo "LERG 7 (Switch) data file. The Error occured in the following module:" >> ${WORKDIR}/blatbody.txt
echo  >> ${WORKDIR}/blatbody.txt
echo  >> ${WORKDIR}/blatbody.txt

case $1 in

dts)
	echo "  DTS Job" >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt

 	blat "${DOSWORKDIR}\blatbody.txt" -subject "Import LERG 7 (Switch) Data - DTS Module" -to walshkel,browng -server 192.168.10.151 -from database@saiccanada.com

	;;

unzip)
	echo "  Unzip Process" >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt

 	blat "${DOSWORKDIR}\blatbody.txt" -subject "Import LERG 7 (Switch) Data - Unzip Module" -to walshkel,browng -server 192.168.10.151 -from database@saiccanada.com

	;;

wrongfile)
	echo "  Wrong File" >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt

 	blat "${DOSWORKDIR}\blatbody.txt" -subject "Import LERG 7 (Switch) Data - Wrong File" -to walshkel,browng -server 192.168.10.151 -from database@saiccanada.com

	;;
nozip)
	echo "  No Zip File" >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt

 	blat "${DOSWORKDIR}\blatbody.txt" -subject "Import LERG 7 (Switch) Data - No Zip File" -to walshkel,browng -server 192.168.10.151 -from database@saiccanada.com

	;;

esac

rm ${WORKDIR}/blatbody.txt

}


#####################
# START MAIN SCRIPT
#####################

# Rename the zip file
if [ -f ${WORKDIR}/Query.zip ]
then
mv ${WORKDIR}/Query.zip ${WORKDIR}/${FileType}.zip
else
	funcERRORNOTIFY "nozip"
	exit
fi

# Check/Create TEMPDIR & DATDIR
if [ -n ${FileType} ]		# if string is not blank
then
	if [ -d ${WORKDIR}/${FileType}TEMP ]
		then
		rm -R ${WORKDIR}/${FileType}TEMP
		mkdir -p ${WORKDIR}/${FileType}TEMP
	else
		mkdir -p ${WORKDIR}/${FileType}TEMP
	fi

	if [ -d ${WORKDIR}/${FileType}DAT ]
	then
		rm -R ${WORKDIR}/${FileType}DAT
		mkdir -p ${WORKDIR}/${FileType}DAT
	else
		mkdir -p ${WORKDIR}/${FileType}DAT
	fi
fi
# Upzip the file
unzip -o -p ${WORKDIR}/${FileType}.zip >${WORKDIR}/${FileType}DAT/${FileType}.dat
if [ $? -ne "0" ]
then
	funcERRORNOTIFY "unzip"
	exit
fi

#Check to make sure we have the correct file
Switch=$( cat ${WORKDIR}/${FileType}DAT/${FileType}.dat | grep -c 'REPORT OF SWITCHING ENTITY DB BY SEARCH CRITERIA' )

if [ ${Switch} -eq 0 ]
	then
		funcERRORNOTIFY "wrongfile"
		exit
fi

LineNumber=$( cat ${WORKDIR}/${FileType}DAT/${FileType}.dat | grep -m 1 -n 'FILE: SRD'|cut -f1 -d:)
cat ${WORKDIR}/${FileType}DAT/${FileType}.dat | sed 1,${LineNumber}d > ${WORKDIR}/${FileType}TEMP/Rows
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c1-11  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/SWITCH_ID			
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c15-16  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/CREATION_DATE_YEAR
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c17-19  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/CREATION_DATE_DAY
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c21-24  | tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/AOCN
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c69-70  | tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/EFF_DATE_YEAR
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c71-73  | tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/EFF_DATE_DAY
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c75-75  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/STATUS
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c77-80  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/OCN
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c88-92  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/EQPT_TYPE
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c94-98  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/MAJOR_VC
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c100-104  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/MAJOR_HC
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c106-106  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/IDDD
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c108-167  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/STREET
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c168-197  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/CITY
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c198-199  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/STATE
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c200-208  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/ZIP
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c256-266  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/ORIG_FG_D
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c300-310  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/ORIG_FG_D_INT
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c311-321  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/ORIG_LOCAL
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c377-387  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/TERM_FG_D
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c421-431  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/TERM_FG_D_INT
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c432-442  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/TERM_LOCAL
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c443-453  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/TERM_INTRA
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c487-497  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/STP_1
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c498-508  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/STP_2
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c542-552  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/ACTUAL_ID
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c553-563  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/CALL_AGENT
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c564-574  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/TRUNK_GATEWAY
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c597-597  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/END_OFF_IND
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c598-598  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/HOST_IND
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c601-601  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/CLASS_4_5_IND
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c602-602  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/WIRELESS_IND
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c612-612  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/INTERMED_OFF_IND
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c616-616  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/LOCAL_TDM_IND
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c634-634  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/LNP_CAPABLE_IND
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c645-645  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/CALL_AGENT_IND
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c816-817  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/LAST_CHANGE_YEAR
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c818-820  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/LAST_CHANGE_DAY

paste -d\| ${WORKDIR}/${FileType}TEMP/SWITCH_ID \
${WORKDIR}/${FileType}TEMP/CREATION_DATE_YEAR \
${WORKDIR}/${FileType}TEMP/CREATION_DATE_DAY \
${WORKDIR}/${FileType}TEMP/AOCN \
${WORKDIR}/${FileType}TEMP/EFF_DATE_YEAR \
${WORKDIR}/${FileType}TEMP/EFF_DATE_DAY \
${WORKDIR}/${FileType}TEMP/STATUS \
${WORKDIR}/${FileType}TEMP/OCN \
${WORKDIR}/${FileType}TEMP/EQPT_TYPE \
${WORKDIR}/${FileType}TEMP/MAJOR_VC \
${WORKDIR}/${FileType}TEMP/MAJOR_HC \
${WORKDIR}/${FileType}TEMP/IDDD \
${WORKDIR}/${FileType}TEMP/STREET \
${WORKDIR}/${FileType}TEMP/CITY \
${WORKDIR}/${FileType}TEMP/STATE \
${WORKDIR}/${FileType}TEMP/ZIP \
${WORKDIR}/${FileType}TEMP/ORIG_FG_D \
${WORKDIR}/${FileType}TEMP/ORIG_FG_D_INT \
${WORKDIR}/${FileType}TEMP/ORIG_LOCAL \
${WORKDIR}/${FileType}TEMP/TERM_FG_D \
${WORKDIR}/${FileType}TEMP/TERM_FG_D_INT \
${WORKDIR}/${FileType}TEMP/TERM_LOCAL \
${WORKDIR}/${FileType}TEMP/TERM_INTRA \
${WORKDIR}/${FileType}TEMP/STP_1 \
${WORKDIR}/${FileType}TEMP/STP_2 \
${WORKDIR}/${FileType}TEMP/ACTUAL_ID \
${WORKDIR}/${FileType}TEMP/CALL_AGENT \
${WORKDIR}/${FileType}TEMP/TRUNK_GATEWAY \
${WORKDIR}/${FileType}TEMP/END_OFF_IND \
${WORKDIR}/${FileType}TEMP/HOST_IND \
${WORKDIR}/${FileType}TEMP/CLASS_4_5_IND \
${WORKDIR}/${FileType}TEMP/WIRELESS_IND \
${WORKDIR}/${FileType}TEMP/INTERMED_OFF_IND \
${WORKDIR}/${FileType}TEMP/LOCAL_TDM_IND \
${WORKDIR}/${FileType}TEMP/LNP_CAPABLE_IND \
${WORKDIR}/${FileType}TEMP/CALL_AGENT_IND \
${WORKDIR}/${FileType}TEMP/LAST_CHANGE_YEAR \
${WORKDIR}/${FileType}TEMP/LAST_CHANGE_DAY >${WORKDIR}/${FileType}TEMP/LERG7

Echo "SWITCH_ID| \
CREATION_DATE_YEAR| \
CREATION_DATE_DAY| \
AOCN| \
EFF_DATE_YEAR| \
EFF_DATE_DAY| \
STATUS| \
OCN| \
EQPT_TYPE| \
MAJOR_VC| \
MAJOR_HC| \
IDDD| \
STREET| \
CITY| \
STATE| \
ZIP| \
ORIG_FG_D| \
ORIG_FG_D_INT| \
ORIG_LOCAL| \
TERM_FG_D| \
TERM_FG_D_INT| \
TERM_LOCAL| \
TERM_INTRA| \
STP_1| \
STP_2| \
ACTUAL_ID| \
CALL_AGENT| \
TRUNK_GATEWAY| \
END_OFF_IND| \
HOST_IND| \
CLASS_4_5_IND| \
WIRELESS_IND| \
INTERMED_OFF_IND| \
LOCAL_TDM_IND| \
LNP_CAPABLE_IND| \
CALL_AGENT_IND| \
LAST_CHANGE_YEAR| \
LAST_CHANGE_DAY" >${WORKDIR}/${FileType}TEMP/LERG7.txt

cat ${WORKDIR}/${FileType}TEMP/LERG7 >> ${WORKDIR}/${FileType}TEMP/LERG7.txt

 call DTS
dtsrun /Sinternet-svcs /E /NImport_LERG_7_Data

if [ $? -ne "0" ]
then
	funcERRORNOTIFY "dts"
	exit
fi


# House Keeping


if [ -d ${WORKDIR}/${FileType}TEMP ]
then
	rm -R ${WORKDIR}/${FileType}TEMP
fi

if [ -d ${WORKDIR}/${FileType}DAT ]
then
	rm -R ${WORKDIR}/${FileType}DAT
fi

if [ -f ${WORKDIR}/${FileType}.zip ]
then
rm ${WORKDIR}/${FileType}.zip
fi

# FOLLOWING LINES USED TO MAIL ADMINS FOR CONFIRMATION OF JOB

echo "  LERG 7 (Switch) Data Load Complete" >> ${WORKDIR}/blatbody.txt
echo >> ${WORKDIR}/blatbody.txt
echo >> ${WORKDIR}/blatbody.txt

blat "${DOSWORKDIR}\blatbody.txt" -subject "Import LERG 7 (Switch) Data Completed" -to browng -server 192.168.10.151 -from database@saiccanada.com

rm ${WORKDIR}/blatbody.txt

