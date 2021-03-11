rem copy Addin and settings...
@echo off
if exist "C:\Program Files\Microsoft Office\root\" (
	echo 64bit office
	copy /Y DatePicker64.xll "%appdata%\Microsoft\AddIns\DatePicker.xll"
) else (
	echo 32bit office
	copy /Y DatePicker32.xll "%appdata%\Microsoft\AddIns\DatePicker.xll"
)
enableAddin.vbs
pause
