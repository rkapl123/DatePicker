Set /P answr=deploy (r)elease (empty for debug)? 
@echo off
set source=bin\Debug
If "%answr%"=="r" (
	set source=bin\Release
)
if exist "C:\Program Files\Microsoft Office\root\" (
	echo 64bit office
	copy /Y %source%\DatePicker-AddIn64-packed.xll "%appdata%\Microsoft\AddIns\DatePicker.xll"
	copy /Y %source%\DatePicker.pdb "%appdata%\Microsoft\AddIns"
) else (
	echo 32bit office
	copy /Y %source%\DatePicker-AddIn-packed.xll "%appdata%\Microsoft\AddIns\DatePicker.xll"
	copy /Y %source%\DatePicker.pdb "%appdata%\Microsoft\AddIns"
)
if "%source%"=="bin\Release" (
	echo copying Distribution
	copy /Y %source%\DatePicker-AddIn64-packed.xll Distribution\DatePicker64.xll
	copy /Y %source%\DatePicker-AddIn-packed.xll Distribution\DatePicker32.xll
	copy /Y TestVBA.xlsm Distribution
)
pause
