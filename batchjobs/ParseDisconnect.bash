#!/bin/bash

#	****************************************************************************************
#	*           VERSION CONTROL INFORMATION
#	****************************************************************************************
#	* CVS File:      $RCSfile: ParseDisconnect.bash,v $
#	* Commit Date:   $Date: 2015/03/09 11:19:44 $ (UTC)
#	* Committed by:  $Author: walshkel $
#	* CVS Revision:  $Revision: 1.3 $
#	* Checkout Tag:  $Name:  $ (Version/Build)
#	****************************************************************************************

# ------------------------------------------------------------------------------------

#####################
# DECLARE VARIABLES
#####################

# Directory Variables

BASEDIR="/d"
WORKDIR=${BASEDIR}/CNA/BIRRDS
DOSWORKDIR="D:\CNA\BIRRDS"
FileType="Disconnect"

#####################
# DECLARE FUNCTIONS
#####################

funcERRORNOTIFY ( ) {

echo "An Error was encountered while processing the automated import of the" > ${WORKDIR}/blatbody.txt
echo "LERG 6 (Disconnect) data file. The Error occured in the following module:" >> ${WORKDIR}/blatbody.txt
echo  >> ${WORKDIR}/blatbody.txt
echo  >> ${WORKDIR}/blatbody.txt

case $1 in

dts)
	echo "  DTS Job" >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt

 	blat "${DOSWORKDIR}\blatbody.txt" -subject "Import LERG 6 (Disconnect) Data - DTS Module" -to walshkel,browng -server 192.168.10.51 -from database@leidos.ca

	;;

unzip)
	echo "  Unzip Process" >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt

 	blat "${DOSWORKDIR}\blatbody.txt" -subject "Import LERG 6 (Disconnect) Data - Unzip Module" -to walshkel,browng -server 192.168.10.51 -from database@leidos.ca

	;;

wrongfile)
	echo "  Wrong File" >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt

 	blat "${DOSWORKDIR}\blatbody.txt" -subject "Import LERG 6 (Disconnect) Data - Wrong File" -to walshkel,browng -server 192.168.10.51 -from database@leidos.ca

	;;
nozip)
	echo "  No Zip File" >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt

 	blat "${DOSWORKDIR}\blatbody.txt" -subject "Import LERG 6 (Disconnect) Data - No Zip File" -to walshkel,browng -server 192.168.10.51 -from database@leidos.ca

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
NXX=$( cat ${WORKDIR}/${FileType}DAT/${FileType}.dat | grep -c 'NXD' )
if [ ${NXX} -eq 0 ]
	then
		funcERRORNOTIFY "wrongfile"
		exit
fi

grep ^[0-9][0-9][0-9][0-9][0-9][0-9]A ${WORKDIR}/${FileType}DAT/${FileType}.dat > ${WORKDIR}/${FileType}TEMP/Rows

cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c1-3  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/NPA			
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c4-6  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/NXX			
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c7-7  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/BLOCK_ID			
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c12-12  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/PRODUCT			
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c13-16  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/AOCN_RESP			
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c22-22  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/STATUS
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c23-24  | tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/EFF_DATE_YEAR
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c25-27  | tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/EFF_DATE_DAY
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c29-30  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/LAST_CHANGE_YEAR
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c31-33  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/LAST_CHANGE_DAY
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c57-58  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/CREATION_DATE_YEAR
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c59-61  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/CREATION_DATE_DAY
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c75-76  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/LO_STATE
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c77-86  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/LO_NAME
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c89-91  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/COC_TYPE
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c92-95  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/SSC
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c96-97  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/EO
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c98-99  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/AT
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c101-104  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/OCN
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c105-115  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/SWITCH
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c116-117  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/SHA_IND
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c123-126 |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/TEST_LINE
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c259-260  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/RC_STATE
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c261-270  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/RC_NAME10
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c274-274  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/IDDD
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c275-275  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/DAYLIGHT
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c276-276  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/TIMEZONE
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c278-278  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/DIALABLE
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c279-279  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/POOL_IND
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c280-280  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/PORTABLE
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c284-285  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/NXX_TYPE
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c286-288  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/BILL_RAO
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c290-290  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/CO_TYPE
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c299-301  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/BO
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c305-306  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/OLR_STEP
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c308-317  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/PLNAME10
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c318-367  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/PLACE_NAME
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c368-369  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/PLACE_ST
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c370-419  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/LOC_NAME
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c420-421  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/LOC_ST
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c422-471  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/RC_NAME
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c472-473  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/RC_ST
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c476-480  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/MAJOR_VERT
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c482-486  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/MAJOR_HORZ

paste -d\| ${WORKDIR}/${FileType}TEMP/NPA \
${WORKDIR}/${FileType}TEMP/NXX \
${WORKDIR}/${FileType}TEMP/BLOCK_ID \
${WORKDIR}/${FileType}TEMP/PRODUCT \
${WORKDIR}/${FileType}TEMP/AOCN_RESP \
${WORKDIR}/${FileType}TEMP/STATUS \
${WORKDIR}/${FileType}TEMP/EFF_DATE_YEAR \
${WORKDIR}/${FileType}TEMP/EFF_DATE_DAY \
${WORKDIR}/${FileType}TEMP/LAST_CHANGE_YEAR \
${WORKDIR}/${FileType}TEMP/LAST_CHANGE_DAY \
${WORKDIR}/${FileType}TEMP/CREATION_DATE_YEAR \
${WORKDIR}/${FileType}TEMP/CREATION_DATE_DAY \
${WORKDIR}/${FileType}TEMP/LO_STATE \
${WORKDIR}/${FileType}TEMP/LO_NAME \
${WORKDIR}/${FileType}TEMP/COC_TYPE \
${WORKDIR}/${FileType}TEMP/SSC \
${WORKDIR}/${FileType}TEMP/EO \
${WORKDIR}/${FileType}TEMP/AT \
${WORKDIR}/${FileType}TEMP/OCN \
${WORKDIR}/${FileType}TEMP/SWITCH \
${WORKDIR}/${FileType}TEMP/SHA_IND \
${WORKDIR}/${FileType}TEMP/TEST_LINE \
${WORKDIR}/${FileType}TEMP/RC_STATE \
${WORKDIR}/${FileType}TEMP/RC_NAME10 \
${WORKDIR}/${FileType}TEMP/IDDD \
${WORKDIR}/${FileType}TEMP/DAYLIGHT \
${WORKDIR}/${FileType}TEMP/TIMEZONE \
${WORKDIR}/${FileType}TEMP/DIALABLE \
${WORKDIR}/${FileType}TEMP/POOL_IND \
${WORKDIR}/${FileType}TEMP/PORTABLE \
${WORKDIR}/${FileType}TEMP/NXX_TYPE \
${WORKDIR}/${FileType}TEMP/BILL_RAO \
${WORKDIR}/${FileType}TEMP/CO_TYPE \
${WORKDIR}/${FileType}TEMP/BO \
${WORKDIR}/${FileType}TEMP/OLR_STEP \
${WORKDIR}/${FileType}TEMP/PLNAME10 \
${WORKDIR}/${FileType}TEMP/PLACE_NAME \
${WORKDIR}/${FileType}TEMP/PLACE_ST \
${WORKDIR}/${FileType}TEMP/LOC_NAME \
${WORKDIR}/${FileType}TEMP/LOC_ST \
${WORKDIR}/${FileType}TEMP/RC_NAME \
${WORKDIR}/${FileType}TEMP/RC_ST \
${WORKDIR}/${FileType}TEMP/MAJOR_VERT \
${WORKDIR}/${FileType}TEMP/MAJOR_HORZ >${WORKDIR}/${FileType}TEMP/Disconnect

Echo "NPA|\
NXX|\
BLOCK_ID|\
PRODUCT|\
AOCN_RESP|\
STATUS|\
EFF_DATE_YEAR|\
EFF_DATE_DAY|\
LAST_CHANGE_YEAR|\
LAST_CHANGE_DAY|\
CREATION_DATE_YEAR|\
CREATION_DATE_DAY|\
LO_STATE|\
LO_NAME|\
COC_TYPE|\
SSC|\
EO|\
AT|\
OCN|\
SWITCH|\
SHA_IND|\
TEST_LINE|\
RC_STATE|\
RC_NAME10|\
IDDD|\
DAYLIGHT|\
TIMEZONE|\
DIALABLE|\
POOL_IND|\
PORTABLE|\
NXX_TYPE|\
BILL_RAO|\
CO_TYPE|\
BO|\
OLR_STEP|\
PLNAME10|\
PLACE_NAME|\
PLACE_ST|\
LOC_NAME|\
LOC_ST|\
RC_NAME|\
RC_ST|\
MAJOR_VERT|\
MAJOR_HORZ" >${WORKDIR}/${FileType}TEMP/Disconnect.txt

cat ${WORKDIR}/${FileType}TEMP/Disconnect >> ${WORKDIR}/${FileType}TEMP/Disconnect.txt


# call DTSX
CMD /c "DTExec /FILE D:\SQLServer\DTSX\Import_LERG6_Disconnect_Data.dtsx"


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

echo "  LERG 6 (Disconnect) Data Load Complete" >> ${WORKDIR}/blatbody.txt
echo >> ${WORKDIR}/blatbody.txt
echo >> ${WORKDIR}/blatbody.txt

blat "${DOSWORKDIR}\blatbody.txt" -subject "Import LERG 6 (Disconnect) Data Completed" -to browng,walshkel -server 192.168.10.51 -from database@leidos.ca

rm ${WORKDIR}/blatbody.txt

