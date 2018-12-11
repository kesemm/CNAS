#!/bin/ksh

# -----------------------------------------------------
                                            ###########
# ------------------------------------------------------------------------------------

#	****************************************************************************************
#	*           VERSION CONTROL INFORMATION
#	****************************************************************************************
#	* CVS File:      $RCSfile:$
#	* Commit Date:   $Date:$ (UTC)
#	* Committed by:  $Author:$
#	* CVS Revision:  $Revision: $
#	* Checkout Tag:  $Name:  $ (Version/Build)
#	****************************************************************************************

# ------------------------------------------------------------------------------------

#####################
# DECLARE VARIABLES
#####################

# Directory Variables

BASEDIR="/cygdrive/d/"
DOSBASEDIR="d:\\"

YESTERDAY=$(date --date "1 day ago" +%Y%m%d)
FileName="http-${YESTERDAY}.log"

#####################
# START MAIN SCRIPT
#####################


# Get yesterday's log file from AT&T

ncftpget -d ${BASEDIR}ftpdebug.log -u cnac8702 -p 2LeTmEiN cnac.ca ${BASEDIR} /logs/${FileName}

#####################
# END SCRIPT
#####################
