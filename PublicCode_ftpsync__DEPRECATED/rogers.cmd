::
:: FTPSync COCode & ESRD Data Files to Rogers Server
::

:: FTPSYNC FILES TO ROGERS
:repeat
D:\CNA\PublicCode_ftpsync_magma\FTPSync.exe rogers
if errorlevel 4 goto inierror
if errorlevel 3 goto repeat
if errorlevel 2 goto repeat
if errorlevel 1 goto repeat
goto ftpcomplete

:inierror
:: IN CASE AN INI FILE ERROR
ECHO "******* FTPSync INI file error" >> D:\CNA\PublicCode_ftpsync_magma\error.log

:ftpcomplete
:: CONCATENATE ALL THE LOG FILES
COPY D:\CNA\PublicCode_ftpsync_magma\*.log D:\CNA\PublicCode_ftpsync_magma\blatbody.txt

:: EMAIL LOG FILES TO ADMINS
blat "D:\CNA\PublicCode_ftpsync_magma\blatbody.txt" -subject "FTPSync of COCode & ESRD files to Rogers completed" -to browng,walshkel -server 192.168.10.51 -from database@leidos.ca

:: CLEAN UP ALL LOG FILES
DEL /Q /F D:\CNA\PublicCode_ftpsync_magma\*.log D:\CNA\PublicCode_ftpsync_magma\blatbody.txt