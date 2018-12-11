#!/bin/ksh

# -----------------------------------------------------
                                            ###########
# CheckCNAFTP.ksh                         ###############
                                        ###################
 ########################################                 ##
########################################  Kelly T. Walsh   ##
 #######################################    SAIC Canada    ##
      ####      ####       ####        ##                 ##
       ##        ##         ##          ###################
                                          ###############
                                            ###########
# LastModified: 2003-06-23 10:36

# What this script will do:

# 1) Obtain listing of ac*.htm and esrd*.htm and NAPANXX.zip files from
#    AT&T and Magma CNA Websites
# 2) Checks:
#             - That there are 47 files;
#             - The file date are all the same day; and,
#             - That the day is current day - 1 (i.e. yesterday)
# 3) Blats the results (positive too) to walshkel and browng

# What this script requires:

# 1) CygWin installed with ncFTPPut.

# ------------------------------------------------------------------------------------

############################
# EMAIL NOTIFICATION SWITCH
############################
#NOTIFY=OFF
NOTIFY=ON


####################
# DEBUG MODE SWITCH
####################
DEBUG=OFF
#DEBUG=ON


######################
# DIRECTORY VARIABLES
######################
WORKDIR=d:/CNA/External/
DEBUGDIR=d:/CNA/External/


########################
# SET VARIABLE DEFAULTS
########################
ATT=BAD
MAGMA=BAD


integer TODAY=$(date +%d)


## SORT THAT WHOLE 'NEW MONTH' THING OUT ##

if [ $TODAY = 1 ]; then

	case $(date +%m) in

	02|04|06|08|09|11)
		integer POSTDAY=31
		;;

	01|05|07|10|12)
		integer POSTDAY=30
		;;

	03)
		integer POSTDAY=28  ## NO LEAP YEAR CONSIDERATION - NEXT WORRY 2008 ##
		;;

	esac

else

	integer POSTDAY=$(expr $TODAY - 1)

fi


## DEBUG OUTPUT ##

if [ $DEBUG = ON ]; then
      echo $(date) > ${DEBUGDIR}debug.log   ## NOTE THE STOMP ON DEBUG.LOG ##
      echo ${TODAY} >> ${DEBUGDIR}debug.log
      echo ${POSTDAY}  >> ${DEBUGDIR}debug.log
fi


# AQUIRE AT&T LISTING

ncftpls -l -u cnac8702 -p 2LeTmEiN ftp://cnac.ca/htdocs/ | egrep ac.*htm\|esrd.*htm\|NPANXX.zip | tr -s [" "] > ${WORKDIR}~ATTlisting

if [ $(cat ${WORKDIR}~ATTlisting | wc -l) = 47 ]; then

	cat ${WORKDIR}~ATTlisting | cut -f7 -d ' ' | sort | uniq > ${WORKDIR}~ATTlisting2	# FIELD 7 FOR DAY ON AT&T

	if [ $(cat ${WORKDIR}~ATTlisting2 | wc -l  ) = 1 ]; then

		integer ATTDAY=$(cat ${WORKDIR}~ATTlisting2)	# CONVERT DAY TO INTEGER FOR COMPARE TO WORK

		if [ $ATTDAY = $POSTDAY ]; then

			ATT=GOOD

		fi

	fi

fi


## DEBUG OUTPUT ##

if [ $DEBUG = ON ]; then
      cat ${WORKDIR}~ATTlisting  >> ${DEBUGDIR}debug.log
      cat ${WORKDIR}~ATTlisting2  >> ${DEBUGDIR}debug.log
      echo ${ATT}  >> ${DEBUGDIR}debug.log
fi


# AQUIRE MAGMA LISTING

ncftpls -l -u saiccan -p Qu3en6 ftp://64.26.166.158/public_html/ | egrep ac.*htm\|esrd.*htm\|NPANXX.zip > ${WORKDIR}~MAGMAlisting

if [ $(cat ${WORKDIR}~MAGMAlisting | wc -l) = 47 ]; then

	cat ${WORKDIR}~MAGMAlisting | tr -s [" "]  | cut -f6 -d ' ' | sort | uniq > ${WORKDIR}~MAGMAlisting2	# FIELD 6 FOR DAY ON MAGMA

	if [ $(cat ${WORKDIR}~MAGMAlisting2 | wc -l  ) = 1 ]; then

		integer MAGMADAY=$(cat ${WORKDIR}~MAGMAlisting2)	# CONVERT DAY TO INTEGER FOR COMPARE TO WORK

		if [ $MAGMADAY = $POSTDAY ]; then

			MAGMA=GOOD

		fi

	fi

fi


## DEBUG OUTPUT ##
if [ $DEBUG = ON ]; then
      cat ${WORKDIR}~MAGMAlisting  >> ${DEBUGDIR}debug.log
      cat ${WORKDIR}~MAGMAlisting2  >> ${DEBUGDIR}debug.log
      echo ${MAGMA}  >> ${DEBUGDIR}debug.log
fi
	

## DRAFT AND BLAT APPROPRIATE RESULTS IF NOTIFY IS ON ##

if [ $NOTIFY = ON ]; then

      echo 'Here are the resulting listings from both sites for you to look at.' > ${WORKDIR}~blatbody
      echo >> ${WORKDIR}~blatbody
      echo 'AT&T Website:' >> ${WORKDIR}~blatbody
      echo >> ${WORKDIR}~blatbody
      cat ${WORKDIR}~ATTlisting >> ${WORKDIR}~blatbody
      echo >> ${WORKDIR}~blatbody
      echo >> ${WORKDIR}~blatbody
      echo 'MAGMA Website:' >> ${WORKDIR}~blatbody
      echo >> ${WORKDIR}~blatbody
      cat ${WORKDIR}~MAGMAlisting >> ${WORKDIR}~blatbody
      

      if [ $ATT != GOOD ] && [ $MAGMA != GOOD ]; then

      	blat ${WORKDIR}~blatbody -subject "CNA FTP Listing Error at AT&T and Magma" -to walshkel,browng -from "database@saiccanada.com" -server "192.168.10.50"

      elif [ $ATT = GOOD ] && [ $MAGMA != GOOD ]; then

      	blat ${WORKDIR}~blatbody -subject "CNA FTP Listing Error at Magma" -to walshkel,browng -from "database@saiccanada.com" -server "192.168.10.50"

      elif [ $ATT != GOOD ] && [ $MAGMA = GOOD ]; then

      	blat ${WORKDIR}~blatbody -subject "CNA FTP Listing Error at AT&T" -to walshkel,browng -from "database@saiccanada.com" -server "192.168.10.50"

      elif [ $ATT = GOOD ] && [  $MAGMA = GOOD ]; then

      	blat ${WORKDIR}~blatbody -subject "CNA FTP Listing A-OK" -to walshkel,browng -from "database@saiccanada.com" -server "192.168.10.50"

      fi

	rm ${WORKDIR}~blatbody

fi


rm ${WORKDIR}~*listing*


#############
# END SCRIPT
#############
