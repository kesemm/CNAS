@echo off
REM *** This batch file is part of an
REM *** SQL Server job to import LERG
REM *** Data. This is the first part of
REM *** the job that copies the mdb to
REM *** the HD to get a file lock.

echo Making Temporary Directory on Hard drive
MD D:\~LERGCD

echo.
echo.
echo Copying LERG Database to Temporary location
COPY E:\LERGDATA.MDB D:\~LERGCD


exit

