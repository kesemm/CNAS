ECHO OFF

:: CLEAR ALL WORKING FILES
IF EXIST D:\CNA\WinSCP\log\DataFiles.log DEL /F /Q D:\CNA\WinSCP\log\DataFiles.log
IF EXIST D:\CNA\WinSCP\Blat\BlatBody.txt DEL /F /Q D:\CNA\WinSCP\Blat\BlatBody.txt


:: CALL WinSCP AND THE SCRIPT TO PUBLISH ALL DATA FILES TO BOTH SITES
"D:\CNA\WinSCP\WinSCP.exe" /log="D:\CNA\WinSCP\log\DataFiles.log" /ini="D:\CNA\WinSCP\WinSCP.ini" /script="D:\CNA\WinSCP\WinSCP_DataFiles_script.txt"


:: BUILD THE NOTIFICATION TEXT EMAIL BODY
ECHO The scheduled job to SECURELY post the CNA Data files (COCode, ESRD, NonGeo and MBI) to the cnac.ca and cnac2.ca website has been completed. >> D:\CNA\WinSCP\Blat\BlatBody.txt
ECHO. >> D:\CNA\WinSCP\Blat\BlatBody.txt
ECHO. >> D:\CNA\WinSCP\Blat\BlatBody.txt
ECHO The following is the exit code from WinSCP: %ERRORLEVEL% >> D:\CNA\WinSCP\Blat\BlatBody.txt
ECHO. >> D:\CNA\WinSCP\Blat\BlatBody.txt
ECHO. >> D:\CNA\WinSCP\Blat\BlatBody.txt
ECHO The WinSCP logfile has been attached for verification purposes as well. >> D:\CNA\WinSCP\Blat\BlatBody.txt


:: EMAIL NOTIFICATION
blat D:\CNA\WinSCP\Blat\BlatBody.txt -attach D:\CNA\WinSCP\log\DataFiles.log -subject "Secure Data job complete" -to walshkel@leidos.ca,browng@leidos.ca -server 192.168.10.51 -from database@leidos.ca


:: CLEAR ALL WORKING FILES
IF EXIST D:\CNA\WinSCP\log\DataFiles.log DEL /F /Q D:\CNA\WinSCP\log\DataFiles.log
IF EXIST D:\CNA\WinSCP\Blat\BlatBody.txt DEL /F /Q D:\CNA\WinSCP\Blat\BlatBody.txt
