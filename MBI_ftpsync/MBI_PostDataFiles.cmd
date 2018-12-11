@ECHO OFF
::     This Command file is used to post MBI Data files to the cnac.ca primary and backup websites
::
::     This command file is intended to be called as the last step in a scheduled database job
::
::     REQUIRED IN PATH:
::                           UnxUtils  (specifically the Windows Command compatible "grep" utility


:: CLEAR ALL FTPSYNC LOG FILES
IF EXIST D:\CNA\MBI_ftpsync\*.log DEL /F /Q D:\CNA\MBI_ftpsync\*.log


:: CLEAR ALL WORKING FILES
IF EXIST D:\CNA\MBI_ftpsync\ftpsync-logs.txt DEL /F /Q D:\CNA\MBI_ftpsync\ftpsync-logs.txt
IF EXIST D:\CNA\MBI_ftpsync\BlatBody.txt DEL /F /Q D:\CNA\MBI_ftpsync\BlatBody.txt
IF EXIST D:\CNA\MBI_ftpsync\ftpsync-allstream-logs.txt DEL /F /Q D:\CNA\MBI_ftpsync\ftpsync-allstream-logs.txt
IF EXIST D:\CNA\MBI_ftpsync\ftpsync-rogers-logs.txt DEL /F /Q D:\CNA\MBI_ftpsync\ftpsync-rogers-logs.txt


:: FTPSYNC FILES TO ALLSTREAM - Error checking added to repeat FTPSync
:repeatallstream
D:\CNA\MBI_ftpsync\FTPSync.exe allstream
if errorlevel 4 goto ftpcompleteallstream
if errorlevel 3 goto repeatallstream
if errorlevel 2 goto repeatallstream
if errorlevel 1 goto repeatallstream
goto ftpcompleteallstream

:ftpcompleteallstream


:: FTPSYNC FILES TO ROGERS - Error checking added to repeat FTPSync
:repeatrogers
D:\CNA\MBI_ftpsync\FTPSync.exe rogers
if errorlevel 4 goto ftpcompleterogers
if errorlevel 3 goto repeatrogers
if errorlevel 2 goto repeatrogers
if errorlevel 1 goto repeatrogers
goto ftpcompleterogers

:ftpcompleterogers


:: CONCATENATE THE FTPSync LOG FILES BASED ON SITE LOCATION TO PARSE FOR EXIT CODES
COPY /A /Y D:\CNA\MBI_ftpsync\ftpsync-allstream-*.log D:\CNA\MBI_ftpsync\ftpsync-allstream-logs.txt
COPY /A /Y D:\CNA\MBI_ftpsync\ftpsync-rogers-*.log D:\CNA\MBI_ftpsync\ftpsync-rogers-logs.txt


:: CONCATENATE THE TWO LOCATION-BASED LOG FILES TO USE AS NOTIFICATION EMAIL ATTACHMENT
COPY /A /Y D:\CNA\MBI_ftpsync\ftpsync-allstream-logs.txt+D:\CNA\MBI_ftpsync\ftpsync-rogers-logs.txt D:\CNA\MBI_ftpsync\ftpsync-logs.txt


:: BUILD THE NOTIFICATION TEXT EMAIL BODY
ECHO The scheduled job to post MBI Data files to the cnac.ca website has been completed. >> D:\CNA\MBI_ftpsync\BlatBody.txt
ECHO. >> D:\CNA\MBI_ftpsync\BlatBody.txt
ECHO. >> D:\CNA\MBI_ftpsync\BlatBody.txt
ECHO The following are exit codes from FTPSync to Allstream: >> D:\CNA\MBI_ftpsync\BlatBody.txt
ECHO. >> D:\CNA\MBI_ftpsync\BlatBody.txt
D:\ubin\grep "Exit code:" D:\CNA\MBI_ftpsync\ftpsync-allstream-logs.txt >> D:\CNA\MBI_ftpsync\BlatBody.txt
ECHO. >> D:\CNA\MBI_ftpsync\BlatBody.txt
ECHO. >> D:\CNA\MBI_ftpsync\BlatBody.txt
ECHO The following are exit codes from FTPSync to Rogers: >> D:\CNA\MBI_ftpsync\BlatBody.txt
ECHO. >> D:\CNA\MBI_ftpsync\BlatBody.txt
D:\ubin\grep "Exit code:" D:\CNA\MBI_ftpsync\ftpsync-rogers-logs.txt >> D:\CNA\MBI_ftpsync\BlatBody.txt
ECHO. >> D:\CNA\MBI_ftpsync\BlatBody.txt
ECHO. >> D:\CNA\MBI_ftpsync\BlatBody.txt
ECHO. >> D:\CNA\MBI_ftpsync\BlatBody.txt
ECHO. >> D:\CNA\MBI_ftpsync\BlatBody.txt
ECHO FTPSync logfiles have been attached for verification purposes as well. >> D:\CNA\MBI_ftpsync\BlatBody.txt


:: EMAIL NOTIFICATION
blat D:\CNA\MBI_ftpsync\BlatBody.txt -attach D:\CNA\MBI_ftpsync\ftpsync-logs.txt -subject "MBI Data posting job complete" -to walshkel@leidos.ca,browng@leidos.ca -server 192.168.10.51 -from database@leidos.ca


:: CLEAR ALL WORKING FILES
IF EXIST D:\CNA\MBI_ftpsync\ftpsync-logs.txt DEL /F /Q D:\CNA\MBI_ftpsync\ftpsync-logs.txt
IF EXIST D:\CNA\MBI_ftpsync\BlatBody.txt DEL /F /Q D:\CNA\MBI_ftpsync\BlatBody.txt
IF EXIST D:\CNA\MBI_ftpsync\ftpsync-allstream-logs.txt DEL /F /Q D:\CNA\MBI_ftpsync\ftpsync-allstream-logs.txt
IF EXIST D:\CNA\MBI_ftpsync\ftpsync-rogers-logs.txt DEL /F /Q D:\CNA\MBI_ftpsync\ftpsync-rogers-logs.txt
IF EXIST D:\CNA\MBI_ftpsync\Times.txt DEL /F /Q D:\CNA\MBI_ftpsync\Times.txt
