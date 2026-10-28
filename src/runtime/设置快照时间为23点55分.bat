@echo off
net session >nul 2>&1
if not "%errorlevel%"=="0" goto ELEV
goto RUN

:ELEV
echo Requesting administrator privileges...
powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
exit /b

:RUN
echo Changing HZDashBoardDailySave trigger to 23:55 ...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$t=Get-ScheduledTask -TaskName 'HZDashBoardDailySave'; $old=([datetime]$t.Triggers[0].StartBoundary).ToString('HH:mm'); $tr=$t.Triggers[0]; $tr.StartBoundary=(Get-Date).ToString('yyyy-MM-dd')+'T23:55:00'; Set-ScheduledTask -TaskName 'HZDashBoardDailySave' -Trigger $tr | Out-Null; $n=Get-ScheduledTask -TaskName 'HZDashBoardDailySave'; $i=$n | Get-ScheduledTaskInfo; Write-Host ('  time : '+$old+' -> '+([datetime]$n.Triggers[0].StartBoundary).ToString('HH:mm')); Write-Host ('  next : '+$i.NextRunTime); Write-Host ('  state: '+$n.State)"
echo.
pause
