#!/bin/bash

#	****************************************************************************************
#	*           VERSION CONTROL INFORMATION
#	****************************************************************************************
#	* CVS File:      $RCSfile: ParseLRNData.bash,v $
#	* Commit Date:   $Date: 2015/03/09 11:19:44 $ (UTC)
#	* Committed by:  $Author: walshkel $
#	* CVS Revision:  $Revision: 1.4 $
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
FileType="LRN"

#####################
# DECLARE FUNCTIONS
#####################

funcERRORNOTIFY ( ) {

echo "An Error was encountered while processing the automated import of the" > ${WORKDIR}/blatbody.txt
echo "LERG 12 (LRN) data file. The Error occured in the following module:" >> ${WORKDIR}/blatbody.txt
echo  >> ${WORKDIR}/blatbody.txt
echo  >> ${WORKDIR}/blatbody.txt

case $1 in

dts)
	echo "  DTS Job" >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt

 	blat "${DOSWORKDIR}\blatbody.txt" -subject "Import LERG 12 (LRN) Data - DTS Module" -to walshkel,browng -server 192.168.10.51 -from database@leidos.ca

	;;

unzip)
	echo "  Unzip Process" >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt

 	blat "${DOSWORKDIR}\blatbody.txt" -subject "Import LERG 12 (LRN) Data - Unzip Module" -to walshkel,browng -server 192.168.10.51 -from database@leidos.ca

	;;

wrongfile)
	echo "  Wrong File" >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt

 	blat "${DOSWORKDIR}\blatbody.txt" -subject "Import LERG 12 (LRN) Data - Wrong File" -to walshkel,browng -server 192.168.10.51 -from database@leidos.ca

	;;
nozip)
	echo "  No Zip File" >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt

 	blat "${DOSWORKDIR}\blatbody.txt" -subject "Import LERG 12 (LRN) Data - No Zip File" -to walshkel,browng -server 192.168.10.51 -from database@leidos.ca

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
LRN=$( cat ${WORKDIR}/${FileType}DAT/${FileType}.dat | grep -c 'REPORT OF LOCATION ROUTING' )
if [ ${LRN} -eq 0 ]
	then
		funcERRORNOTIFY "wrongfile"
		exit
fi

grep ^[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9] ${WORKDIR}/${FileType}DAT/${FileType}.dat | grep -v 12345678901234567890 > ${WORKDIR}/${FileType}TEMP/Rows

cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c1-3  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/NPA			
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c4-6  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/NXX			
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c7-10  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/LRN
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c12-15  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/AOCN
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c23-24  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/CREATION_DATE_YEAR
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c25-27  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/CREATION_DATE_DAY
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c31-32  | tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/EFF_DATE_YEAR
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c33-35  | tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/EFF_DATE_DAY
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c37-37  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/STATUS
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c39-39  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/LRN_TYPE
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c41-51  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/SWITCH
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c59-62  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/OCN
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c64-73  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/RC_NAME10
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c76-77  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/RC_STATE
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c214-215  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/LAST_CHANGE_YEAR
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c216-218  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/LAST_CHANGE_DAY
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c249-250  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/SHA_IND


paste -d\| ${WORKDIR}/${FileType}TEMP/NPA \
${WORKDIR}/${FileType}TEMP/NXX \
${WORKDIR}/${FileType}TEMP/LRN \
${WORKDIR}/${FileType}TEMP/AOCN \
${WORKDIR}/${FileType}TEMP/CREATION_DATE_YEAR \
${WORKDIR}/${FileType}TEMP/CREATION_DATE_DAY \
${WORKDIR}/${FileType}TEMP/EFF_DATE_YEAR \
${WORKDIR}/${FileType}TEMP/EFF_DATE_DAY \
${WORKDIR}/${FileType}TEMP/STATUS \
${WORKDIR}/${FileType}TEMP/LRN_TYPE \
${WORKDIR}/${FileType}TEMP/SWITCH \
${WORKDIR}/${FileType}TEMP/OCN \
${WORKDIR}/${FileType}TEMP/RC_NAME10 \
${WORKDIR}/${FileType}TEMP/RC_STATE \
${WORKDIR}/${FileType}TEMP/LAST_CHANGE_YEAR \
${WORKDIR}/${FileType}TEMP/LAST_CHANGE_DAY \
${WORKDIR}/${FileType}TEMP/SHA_IND  >${WORKDIR}/${FileType}TEMP/LERG12

Echo "NPA|\
NXX|\
LRN|\
AOCN|\
CREATION_DATE_YEAR|\
CREATION_DATE_DAY|\
EFF_DATE_YEAR|\
EFF_DATE_DAY|\
STATUS|\
LRN_TYPE|\
SWITCH|\
OCN|\
RC_NAME10|\
RC_STATE|\
LAST_CHANGE_YEAR|\
LAST_CHANGE_DAY|\
SHA_IND" >${WORKDIR}/${FileType}TEMP/LRN.txt

cat ${WORKDIR}/${FileType}TEMP/LERG12 >> ${WORKDIR}/${FileType}TEMP/LRN.txt


# call DTSX
CMD /c "DTExec /FILE D:\SQLServer\DTSX\Import_LERG12_Data.dtsx"

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

echo "  LERG 12 (LRN) Data Load Complete" >> ${WORKDIR}/blatbody.txt
echo >> ${WORKDIR}/blatbody.txt
echo >> ${WORKDIR}/blatbody.txt

blat "${DOSWORKDIR}\blatbody.txt" -subject "Import LERG 12 (LRN) Data Completed" -to browng,walshkel -server 192.168.10.51 -from database@leidos.ca

rm ${WORKDIR}/blatbody.txt