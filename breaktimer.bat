@echo off
setlocal

:start
set endH= %TIME:~0,2%
set /A endM=((%TIME:~3,2% + 20) %% 60)
if %endM% lss %TIME:~3,2% set /A endH +=1


title Timer
mode con cols=16 lines=2

:synchronize
for /F "tokens=1,2 delims=:" %%a in ("%time%") do set /A "minutes=(endH*60+endM)-(%%a*60+1%%b-100)-1, seconds=159"




:wait
timeout /T 1 /NOBREAK > NUL
echo Timer:  %minutes%:%seconds:~-2%
set /A seconds-=1
if %seconds% geq 100 goto wait
set /A minutes-=1, seconds=159, minMOD5=minutes %% 5
if %minutes% lss 0 goto :buzz
if %minMOD5% equ 0 goto synchronize
goto wait

:buzz
start
goto start