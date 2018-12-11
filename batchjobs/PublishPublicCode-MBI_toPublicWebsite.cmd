

:: THIS REVISED CMD SCRIPT ASSUMES THAT WE'VE ALREADY RUN THE STATUS PAGE EXPORTER UTIL (D:\bin\ExportStatusPages.exe)



:: SCP ALL THE DATA FILES TO BOTH WEBSITES USING WINSCP SCRIPT
:: -----------------------------------------------------------

:: CALL WinSCP command file
CALL D:\CNA\WinSCP\CALL_WinSCP_DataFiles_script.cmd


:: ADD WEBSITE STATS TO TABLE
::  37 NPAs htm and csv plus 1 with all X 3  AND Nongeo is 3 total
:: ----------------------------
sqlcmd -S "." -d XCA_DB -Q "INSERT INTO [XCA_DB].[dbo].[WebsiteUpdateStats] ([Count],[StatTypeID],[AutoStat]) VALUES (228,1,1)"

