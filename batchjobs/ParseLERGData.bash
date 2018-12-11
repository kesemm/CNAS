#!/bin/bash

# -----------------------------------------------------
                                            ###########
# ------------------------------------------------------------------------------------

#	****************************************************************************************
#	*           VERSION CONTROL INFORMATION
#	****************************************************************************************
#	* CVS File:      $RCSfile: ParseLERGData.bash,v $
#	* Commit Date:   $Date: 2014/01/28 16:27:57 $ (UTC)
#	* Committed by:  $Author: walshkel $
#	* CVS Revision:  $Revision: 1.2 $
#	* Checkout Tag:  $Name:  $ (Version/Build)
#	****************************************************************************************

# ------------------------------------------------------------------------------------

#####################
# DECLARE VARIABLES
#####################

# Directory Variables

BASEDIR="/cygdrive/d"
WORKDIR=${BASEDIR}/CNA/BIRRDS
DOSWORKDIR="d:\CNA\BIRRDS"

#####################
# DECLARE FUNCTIONS
#####################

funcERRORNOTIFY ( ) {

echo "An Error was encountered while processing the automated import of the" > ${WORKDIR}/blatbody.txt
echo "LERG 6 data. The Error occured in the following module:" >> ${WORKDIR}/blatbody.txt
echo  >> ${WORKDIR}/blatbody.txt
echo  >> ${WORKDIR}/blatbody.txt

case $1 in

dts)
	echo "  DTS Job" >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt

 	blat "${DOSWORKDIR}\blatbody.txt" -subject "Import BIRRDS Data - DTS Module" -to walshkel,browng -server 192.168.10.151 -from database@leidos.ca

	;;

unzip)
	echo "  Unzip Process" >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt

 	blat "${DOSWORKDIR}\blatbody.txt" -subject "Import BIRRDS Data - Unzip Module" -to walshkel,browng -server 192.168.10.151 -from database@leidos.ca

	;;

esac

rm ${WORKDIR}/blatbody.txt

}


#####################
# START MAIN SCRIPT
#####################

# Check/Create TEMPDIR & DATDIR

if [ -d ${WORKDIR}/TEMP ]
then
	rm -R ${WORKDIR}/TEMP
	mkdir -p ${WORKDIR}/TEMP
else
	mkdir -p ${WORKDIR}/TEMP
fi

if [ -d ${WORKDIR}/DAT ]
then
	rm -R ${WORKDIR}/DAT
	mkdir -p ${WORKDIR}/DAT
else
	mkdir -p ${WORKDIR}/DAT
fi

# Unzip LERG6 File

unzip -o -p ${WORKDIR}/LERG6.zip >${WORKDIR}/DAT/LERG6.dat

if [ $? -ne "0" ]
then
	funcERRORNOTIFY "unzip"
fi

grep ^[0-9][0-9][0-9][0-9][0-9][0-9]A ${WORKDIR}/DAT/LERG6.dat > ${WORKDIR}/TEMP/Rows

cat ${WORKDIR}/TEMP/Rows | cut -c1-3  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/NPA			
cat ${WORKDIR}/TEMP/Rows | cut -c4-6  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/NXX			
cat ${WORKDIR}/TEMP/Rows | cut -c7-7  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/BLOCK_ID			
cat ${WORKDIR}/TEMP/Rows | cut -c12-12  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/PRODUCT			
cat ${WORKDIR}/TEMP/Rows | cut -c13-16  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/AOCN_RESP			
cat ${WORKDIR}/TEMP/Rows | cut -c22-22  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/STATUS
cat ${WORKDIR}/TEMP/Rows | cut -c23-24  | tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/EFF_DATE_YEAR
cat ${WORKDIR}/TEMP/Rows | cut -c25-27  | tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/EFF_DATE_DAY
cat ${WORKDIR}/TEMP/Rows | cut -c29-30  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/LAST_CHANGE_YEAR
cat ${WORKDIR}/TEMP/Rows | cut -c31-33  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/LAST_CHANGE_DAY
cat ${WORKDIR}/TEMP/Rows | cut -c57-58  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/CREATION_DATE_YEAR
cat ${WORKDIR}/TEMP/Rows | cut -c59-61  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/CREATION_DATE_DAY
cat ${WORKDIR}/TEMP/Rows | cut -c75-76  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/LO_STATE
cat ${WORKDIR}/TEMP/Rows | cut -c77-86  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/LO_NAME
cat ${WORKDIR}/TEMP/Rows | cut -c89-91  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/COC_TYPE
cat ${WORKDIR}/TEMP/Rows | cut -c92-95  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/SSC
cat ${WORKDIR}/TEMP/Rows | cut -c96-97  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/EO
cat ${WORKDIR}/TEMP/Rows | cut -c98-99  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/AT
cat ${WORKDIR}/TEMP/Rows | cut -c101-104  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/OCN
cat ${WORKDIR}/TEMP/Rows | cut -c105-115  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/SWITCH
cat ${WORKDIR}/TEMP/Rows | cut -c116-117  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/SHA_IND
cat ${WORKDIR}/TEMP/Rows | cut -c123-126 |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/TEST_LINE
cat ${WORKDIR}/TEMP/Rows | cut -c259-260  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/RC_STATE
cat ${WORKDIR}/TEMP/Rows | cut -c261-270  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/RC_NAME10
cat ${WORKDIR}/TEMP/Rows | cut -c274-274  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/IDDD
cat ${WORKDIR}/TEMP/Rows | cut -c275-275  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/DAYLIGHT
cat ${WORKDIR}/TEMP/Rows | cut -c276-276  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/TIMEZONE
cat ${WORKDIR}/TEMP/Rows | cut -c278-278  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/DIALABLE
cat ${WORKDIR}/TEMP/Rows | cut -c279-279  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/POOL_IND
cat ${WORKDIR}/TEMP/Rows | cut -c280-280  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/PORTABLE
cat ${WORKDIR}/TEMP/Rows | cut -c284-285  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/NXX_TYPE
cat ${WORKDIR}/TEMP/Rows | cut -c286-288  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/BILL_RAO
cat ${WORKDIR}/TEMP/Rows | cut -c290-290  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/CO_TYPE
cat ${WORKDIR}/TEMP/Rows | cut -c299-301  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/BO
cat ${WORKDIR}/TEMP/Rows | cut -c305-306  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/OLR_STEP
cat ${WORKDIR}/TEMP/Rows | cut -c308-317  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/PLNAME10
cat ${WORKDIR}/TEMP/Rows | cut -c318-367  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/PLACE_NAME
cat ${WORKDIR}/TEMP/Rows | cut -c368-369  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/PLACE_ST
cat ${WORKDIR}/TEMP/Rows | cut -c370-419  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/LOC_NAME
cat ${WORKDIR}/TEMP/Rows | cut -c420-421  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/LOC_ST
cat ${WORKDIR}/TEMP/Rows | cut -c422-471  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/RC_NAME
cat ${WORKDIR}/TEMP/Rows | cut -c472-473  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/RC_ST
cat ${WORKDIR}/TEMP/Rows | cut -c476-480  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/MAJOR_VERT
cat ${WORKDIR}/TEMP/Rows | cut -c482-486  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/TEMP/MAJOR_HORZ

paste -d, ${WORKDIR}/TEMP/NPA \
${WORKDIR}/TEMP/NXX \
${WORKDIR}/TEMP/BLOCK_ID \
${WORKDIR}/TEMP/PRODUCT \
${WORKDIR}/TEMP/AOCN_RESP \
${WORKDIR}/TEMP/STATUS \
${WORKDIR}/TEMP/EFF_DATE_YEAR \
${WORKDIR}/TEMP/EFF_DATE_DAY \
${WORKDIR}/TEMP/LAST_CHANGE_YEAR \
${WORKDIR}/TEMP/LAST_CHANGE_DAY \
${WORKDIR}/TEMP/CREATION_DATE_YEAR \
${WORKDIR}/TEMP/CREATION_DATE_DAY \
${WORKDIR}/TEMP/LO_STATE \
${WORKDIR}/TEMP/LO_NAME \
${WORKDIR}/TEMP/COC_TYPE \
${WORKDIR}/TEMP/SSC \
${WORKDIR}/TEMP/EO \
${WORKDIR}/TEMP/AT \
${WORKDIR}/TEMP/OCN \
${WORKDIR}/TEMP/SWITCH \
${WORKDIR}/TEMP/SHA_IND \
${WORKDIR}/TEMP/TEST_LINE \
${WORKDIR}/TEMP/RC_STATE \
${WORKDIR}/TEMP/RC_NAME10 \
${WORKDIR}/TEMP/IDDD \
${WORKDIR}/TEMP/DAYLIGHT \
${WORKDIR}/TEMP/TIMEZONE \
${WORKDIR}/TEMP/DIALABLE \
${WORKDIR}/TEMP/POOL_IND \
${WORKDIR}/TEMP/PORTABLE \
${WORKDIR}/TEMP/NXX_TYPE \
${WORKDIR}/TEMP/BILL_RAO \
${WORKDIR}/TEMP/CO_TYPE \
${WORKDIR}/TEMP/BO \
${WORKDIR}/TEMP/OLR_STEP \
${WORKDIR}/TEMP/PLNAME10 \
${WORKDIR}/TEMP/PLACE_NAME \
${WORKDIR}/TEMP/PLACE_ST \
${WORKDIR}/TEMP/LOC_NAME \
${WORKDIR}/TEMP/LOC_ST \
${WORKDIR}/TEMP/RC_NAME \
${WORKDIR}/TEMP/RC_ST \
${WORKDIR}/TEMP/MAJOR_VERT \
${WORKDIR}/TEMP/MAJOR_HORZ >${WORKDIR}/TEMP/LERG6

Echo "NPA, \
NXX, \
BLOCK_ID, \
PRODUCT, \
AOCN_RESP, \
STATUS, \
EFF_DATE_YEAR, \
EFF_DATE_DAY, \
LAST_CHANGE_YEAR, \
LAST_CHANGE_DAY, \
CREATION_DATE_YEAR, \
CREATION_DATE_DAY, \
LO_STATE, \
LO_NAME, \
COC_TYPE, \
SSC, \
EO, \
AT, \
OCN, \
SWITCH, \
SHA_IND, \
TEST_LINE, \
RC_STATE, \
RC_NAME10, \
IDDD, \
DAYLIGHT, \
TIMEZONE, \
DIALABLE, \
POOL_IND, \
PORTABLE, \
NXX_TYPE, \
BILL_RAO, \
CO_TYPE, \
BO, \
OLR_STEP, \
PLNAME10, \
PLACE_NAME, \
PLACE_ST, \
LOC_NAME, \
LOC_ST, \
RC_NAME, \
RC_ST, \
MAJOR_VERT, \
MAJOR_HORZ" >${WORKDIR}/TEMP/LERG6.txt

cat ${WORKDIR}/TEMP/LERG6 >> ${WORKDIR}/TEMP/LERG6.txt

# call DTS
dtsrun /Sinternet-svcs /E /NImport_Daily_LERG_6_Data

if [ $? -ne "0" ]
then
	funcERRORNOTIFY "dts"
fi


# House Keeping


if [ -d ${WORKDIR}/TEMP ]
then
	rm -R ${WORKDIR}/TEMP
fi

if [ -d ${WORKDIR}/DAT ]
then
	rm -R ${WORKDIR}/DAT
fi

if [ -f ${WORKDIR}/LERG6.zip ]
then
	rm ${WORKDIR}/LERG6.zip
fi

# FOLLOWING LINES USED TO MAIL ADMINS FOR CONFIRMATION OF JOB

#echo "  BIRRDS Import" >> ${WORKDIR}/blatbody.txt
#echo >> ${WORKDIR}/blatbody.txt
#echo >> ${WORKDIR}/blatbody.txt

#blat "${DOSWORKDIR}\blatbody.txt" -subject "Import BIRRDS Data Completed" -to walshkel,browng -server 192.168.10.151 -from database@leidos.ca

#rm ${WORKDIR}/blatbody.txt

