

:: THIS REVISED CMD SCRIPT ASSUMES THAT WE'VE ALREADY RUN THE STATUS PAGE EXPORTER UTIL (D:\bin\ExportStatusPages.exe)



:: FTP SYNC CO AND ESRD CODE PAGES TO PUBLIC WEBSITE
:: -------------------------------------------------

:: CALL FTP SYNC COMMAND FILE
CALL D:\CNA\PublicCode_ftpsync\PublicCode_PostDataFiles.cmd

:: FTP SYNC MBI CODE PAGES TO PUBLIC WEBSITE
:: -------------------------------------------------

:: CALL FTP SYNC COMMAND FILE
CALL D:\CNA\MBI_ftpsync\MBI_PostDataFiles.cmd


:: ADD WEBSITE STATS TO TABLE
::  37 NPAs htm and csv plus 1 with all X 3  AND Nongeo is 3 total
:: ----------------------------
sqlcmd -S "." -d XCA_DB -Q "INSERT INTO [XCA_DB].[dbo].[WebsiteUpdateStats] ([Count],[StatTypeID],[AutoStat]) VALUES (228,1,1)"

