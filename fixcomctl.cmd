rem @echo off
reg query "HKEY_CLASSES_ROOT\typelib\{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}\2.0"
if %errorlevel%==0 GOTO DELREGKEY
if %errorlevel%==1 GOTO REGISTEROCX

:DELREGKEY
reg delete hkcr\typelib\{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}\2.0 /f

:REGISTEROCX
if exist %systemroot%\SysWOW64\cscript.exe goto 64 
%systemroot%\system32\regsvr32 /u mscomctl.ocx /s
%systemroot%\system32\regsvr32 mscomctl.ocx /s
pause
exit

:64 
%systemroot%\sysWOW64\regsvr32 /u mscomctl.ocx /s
%systemroot%\sysWOW64\regsvr32 mscomctl.ocx /s
pause
exit
