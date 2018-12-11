::
:: FTPSync MBI Data Files to Allstream Server
::

:: FTPSYNC FILES TO ALLSTREAM
:repeat
D:\CNA\MBI_ftpsync_allstream\FTPSync.exe allstream
if errorlevel 4 goto inierror
if errorlevel 3 goto repeat
if errorlevel 2 goto repeat
if errorlevel 1 goto repeat
goto ftpcomplete

:inierror
:: IN CASE AN INI FILE ERROR
ECHO "******* FTPSync INI file error" >> D:\CNA\MBI_ftpsync_allstream\error.log

:ftpcomplete
:: CONCATENATE ALL THE LOG FILES
COPY D:\CNA\MBI_ftpsync_allstream\*.log D:\CNA\MBI_ftpsync_allstream\blatbody.txt

:: EMAIL LOG FILES TO ADMINS
blat "D:\CNA\MBI_ftpsync_allstream\blatbody.txt" -subject "FTPSync of MBI files to Allstream completed" -to browng,walshkel -server 192.168.10.151 -from database@saiccanada.com

:: CLEAN UP ALL LOG FILES
DEL /Q /F D:\CNA\MBI_ftpsync_allstream\*.log D:\CNA\MBI_ftpsync_allstream\blatbody.txt