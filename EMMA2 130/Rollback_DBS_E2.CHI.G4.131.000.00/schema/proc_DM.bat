rem Loading Scripts Generated (IBM FARS) 9/22/2017 10:58:36
echo off 
echo .................................................................. 
echo usage proc.bat username password servername databasename 
echo .................................................................. 

date /T >PROCS_%4.OUT 
time /T >>PROCS_%4.OUT 

if "%1"=="" goto USAGE 
if "%2"=="" goto USAGE 
if "%3"=="" goto USAGE 
if "%4"=="" goto USAGE 

echo Running BUILD_RELEASE_NUMBER scripts... 

echo   1 of  1 START LOAD: ITSR013528_Revoke_DM.SQL  >> PROCS_"%4".OUT
echo   1 of  1 START LOAD: ITSR013528_Revoke_DM.SQL  
echo Creating Temp Script... started 
if NOT "%4" == "" echo USE %4  > TempSQL.SQL 
if NOT "%4" == "" echo GO     >> TempSQL.SQL 
type "ITSR013528_Revoke_DM.SQL"   >> TempSQL.SQL 
echo Running Script... ITSR013528_Revoke_DM.SQL		
isql -U%1 -P%2 -S%3 -Jcp850 -iTempSQL.SQL >> PROCS_"%4".OUT
echo   1 of  1 END LOAD: ITSR013528_Revoke_DM.SQL  
echo   1 of  1 END LOAD: ITSR013528_Revoke_DM.SQL  >> PROCS_"%4".OUT

del TempSQL.SQL 

echo Finished. 

date /T >>PROCS_%4.OUT 
time /T >>PROCS_%4.OUT 

EXIT /B 

:USAGE 
echo ERROR ENCOUNTERED ! 
echo PLEASE FOLLOW CORRECT SYNTAX AS SHOWN BELOW : 
echo %0 username password servername databasename 

