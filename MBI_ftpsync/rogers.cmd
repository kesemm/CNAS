::
:: FTPSync MBI Data Files to Rogers Server
::

:: FTPSYNC FILES TO ROGERS
:repeat
D:\CNA\MBI_ftpsync_rogers\FTPSync.exe rogers
if errorlevel 4 goto inierror
if errorlevel 3 goto repeat
if errorlevel 2 goto repeat
if errorlevel 1 goto repeat
goto ftpcomplete

:inierror
:: IN CASE AN INI FILE ERROR
ECHO "******* FTPSync INI file error" >> D:\CNA\MBI_ftpsync_rogers\error.log

:ftpcomplete
:: CONCATENATE ALL THE LOG FILES
COPY D:\CNA\MBI_ftpsync_rogers\*.log D:\CNA\MBI_ftpsync_rogers\blatbody.txt

:: EMAIL LOG FILES TO ADMINS
blat "D:\CNA\MBI_ftpsync_rogers\blatbody.txt" -subject "FTPSync of MBI files to rogers completed" -to browng,walshkel -server 192.168.10.51 -from database@leidos.ca

:: CLEAN UP ALL LOG FILES
DEL /Q /F D:\CNA\MBI_ftpsync_rogers\*.log D:\CNA\MBI_ftpsync_rogers\blatbody.txt