

:: CO CODE STATUS PAGES OUT
:: ------------------------

:: EXEC THE DTS TO OUTPUT CO CODE STATUS FILES
DTExec /FILE D:\SQLServer\DTSX\COCodeStatus_ALL.dtsx
:: ZIP THE LARGE FILE
7za a -tzip D:\CNA\PublicCodeData\COCodeStatus_ALL.zip D:\CNA\PublicCodeData\COCodeStatus_ALL.csv
:: DELETE THE ORIGINAL LARGE FILE
DEL /F /Q D:\CNA\PublicCodeData\COCodeStatus_ALL.csv


:: ESRD CODE STATUS PAGES OUT
:: --------------------------

:: EXEC THE DTS TO OUTPUT ESRD CODE STATUS FILES
DTExec /FILE D:\SQLServer\DTSX\ESRDCodeStatus_ALL.dtsx
:: ZIP THE LARGE FILE
7za a -tzip D:\CNA\PublicCodeData\ESRDCodeStatus_ALL.zip D:\CNA\PublicCodeData\ESRDCodeStatus_ALL.csv
:: DELETE THE ORIGINAL LARGE FILE
DEL /F /Q D:\CNA\PublicCodeData\ESRDCodeStatus_ALL.csv


:: NONGEO CODE STATUS PAGES OUT
:: ----------------------------

:: EXEC THE DTS TO OUTPUT CO CODE STATUS FILES
DTExec /FILE D:\SQLServer\DTSX\NonGeoCodeStatus_ALL.dtsx
:: ZIP THE LARGE FILE
7za a -tzip D:\CNA\PublicCodeData\NonGeoCodeStatus_ALL.zip D:\CNA\PublicCodeData\NonGeoCodeStatus_ALL.csv
:: DELETE THE ORIGINAL LARGE FILE
DEL /F /Q D:\CNA\PublicCodeData\NonGeoCodeStatus_ALL.csv


:: MBI CODE STATUS PAGES OUT
:: -------------------------
:: CALL STORED PROCEDURE TO CREATE THE MBI OUTPUT DATA TABLES

sqlcmd -S "." -d XCA_DB -Q "Execute xca_db.dbo.MBIPublicData_CreateTables"


:: EXEC THE DTS TO OUTPUT MBI CODE STATUS FILES
DTExec /FILE D:\SQLServer\DTSX\MBICodeStatus_ALL.dtsx

:: ZIP THE LARGE FILE
7za a -tzip D:\CNA\MBIData\MBICodeStatus_ALL.zip D:\CNA\MBIData\MBICodeStatus_ALL.csv
:: DELETE THE ORIGINAL LARGE FILE
DEL /F /Q D:\CNA\MBIData\MBICodeStatus_ALL.csv

:: CALL STORED PROCEDURE TO DROP THE MBI OUTPUT DATA TABLES
sqlcmd -S "." -d XCA_DB -Q "Execute xca_db.dbo.MBIPublicData_DropTables"



:: FTP SYNC CO AND ESRD CODE PAGES TO PUBLIC WEBSITE
:: -------------------------------------------------

:: CALL FTP SYNC COMMAND FILE
CALL D:\CNA\PublicCode_ftpsync\PublicCode_PostDataFiles.cmd

:: FTP SYNC MBI CODE PAGES TO PUBLIC WEBSITE
:: -------------------------------------------------

:: CALL FTP SYNC COMMAND FILE
CALL D:\CNA\MBI_ftpsync\MBI_PostDataFiles.cmd


:: ADD WEBSITE STATS TO TABLE
::  37 NPAs plus 1 with all X 3
:: ----------------------------
sqlcmd -S "." -d XCA_DB -Q "INSERT INTO [XCA_DB].[dbo].[WebsiteUpdateStats] ([Count],[StatTypeID],[AutoStat]) VALUES (114,1,1)"

