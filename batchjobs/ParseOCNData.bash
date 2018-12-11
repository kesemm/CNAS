#!/bin/bash

#	****************************************************************************************
#	*           VERSION CONTROL INFORMATION
#	****************************************************************************************
#	* CVS File:      $RCSfile: ParseOCNData.bash,v $
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
FileType="OCN"

#####################
# DECLARE FUNCTIONS
#####################

funcERRORNOTIFY ( ) {

echo "An Error was encountered while processing the automated import of the" > ${WORKDIR}/blatbody.txt
echo "LERG 1 (OCN) data file. The Error occured in the following module:" >> ${WORKDIR}/blatbody.txt
echo  >> ${WORKDIR}/blatbody.txt
echo  >> ${WORKDIR}/blatbody.txt

case $1 in

dts)
	echo "  DTS Job" >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt

 	blat "${DOSWORKDIR}\blatbody.txt" -subject "Import LERG 1 (OCN) Data - DTS Module" -to walshkel,browng -server 192.168.10.51 -from database@leidos.ca

	;;

unzip)
	echo "  Unzip Process" >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt

 	blat "${DOSWORKDIR}\blatbody.txt" -subject "Import LERG 1 (OCN) Data - Unzip Module" -to walshkel,browng -server 192.168.10.51 -from database@leidos.ca

	;;

wrongfile)
	echo "  Wrong File" >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt

 	blat "${DOSWORKDIR}\blatbody.txt" -subject "Import LERG 1 (OCN) Data - Wrong File" -to walshkel,browng -server 192.168.10.51 -from database@leidos.ca

	;;
nozip)
	echo "  No Zip File" >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt
	echo >> ${WORKDIR}/blatbody.txt

 	blat "${DOSWORKDIR}\blatbody.txt" -subject "Import LERG 1 (OCN) Data - No Zip File" -to walshkel,browng -server 192.168.10.51 -from database@leidos.ca

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
OCN=$( cat ${WORKDIR}/${FileType}DAT/${FileType}.dat | grep -c 'CRD' )
if [ ${OCN} -eq 0 ]
	then
		funcERRORNOTIFY "wrongfile"
		exit
fi

LineNumber=$( cat ${WORKDIR}/${FileType}DAT/${FileType}.dat | grep -m 1 -n 'FILE: CRD'|cut -f1 -d:)

cat ${WORKDIR}/${FileType}DAT/${FileType}.dat | sed 1,${LineNumber}d > ${WORKDIR}/${FileType}TEMP/Rows
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c1-4  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/OCN			
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c6-55  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/OCN_NAME			
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c78-79  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/OCN_ST
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c81-90  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/OCN_CODE
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c99-118  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/LAST_NAME
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c120-129  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/FIRST_NAME
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c133-182  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/COMPANY
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c184-213  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/TITLE
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c215-244  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/ADDRESS
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c277-296  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/CITY
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c298-299  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/STATE
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c301-309  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/ZIP
cat ${WORKDIR}/${FileType}TEMP/Rows | cut -c311-322  |tr -s "[ ]" "[ ]" |sed -e "s/ $//g" >${WORKDIR}/${FileType}TEMP/PHONE

paste -d\| ${WORKDIR}/${FileType}TEMP/OCN \
${WORKDIR}/${FileType}TEMP/OCN_NAME \
${WORKDIR}/${FileType}TEMP/OCN_ST \
${WORKDIR}/${FileType}TEMP/OCN_CODE \
${WORKDIR}/${FileType}TEMP/LAST_NAME \
${WORKDIR}/${FileType}TEMP/FIRST_NAME \
${WORKDIR}/${FileType}TEMP/COMPANY \
${WORKDIR}/${FileType}TEMP/TITLE \
${WORKDIR}/${FileType}TEMP/ADDRESS \
${WORKDIR}/${FileType}TEMP/CITY \
${WORKDIR}/${FileType}TEMP/STATE \
${WORKDIR}/${FileType}TEMP/ZIP \
${WORKDIR}/${FileType}TEMP/PHONE>${WORKDIR}/${FileType}TEMP/LERG1

Echo "OCN|\
OCN_NAME|\
OCN_ST|\
OCN_CODE|\
LAST_NAME|\
FIRST_NAME|\
COMPANY|\
TITLE|\
ADDRESS|\
CITY|\
STATE|\
ZIP|\
PHONE" >${WORKDIR}/${FileType}TEMP/LERG1.txt

cat ${WORKDIR}/${FileType}TEMP/LERG1 >> ${WORKDIR}/${FileType}TEMP/LERG1.txt



# call DTSX
CMD /c "DTExec /FILE D:\SQLServer\DTSX\Import_LERG1_Data.dtsx"


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

echo "  LERG 1 (OCN) Data Load Complete" >> ${WORKDIR}/blatbody.txt
echo >> ${WORKDIR}/blatbody.txt
echo >> ${WORKDIR}/blatbody.txt

blat "${DOSWORKDIR}\blatbody.txt" -subject "Import LERG 1 (OCN) Data Completed" -to browng,walshkel -server 192.168.10.51 -from database@leidos.ca

rm ${WORKDIR}/blatbody.txt