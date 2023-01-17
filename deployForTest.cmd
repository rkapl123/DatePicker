@echo off
Set ans=
Set /P answr=deploy [r]elease [empty for debug]?
Set /P ans=deploy [u]npacked xll components [empty for packed in one xll]?
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
	If "%ans%"=="u" (
		echo unpacked components
		copy /Y %source%\DatePicker-AddIn.xll "%appdata%\Microsoft\AddIns\DatePicker.xll"
		copy /Y %source%\DatePicker-AddIn.dna "%appdata%\Microsoft\AddIns\DatePicker.dna"
		copy /Y %source%\DatePicker.dll "%appdata%\Microsoft\AddIns\"
	)
	If "%ans%"=="" (
		echo packed in one xll
		copy /Y %source%\DatePicker-AddIn-packed.xll "%appdata%\Microsoft\AddIns\DatePicker.xll"
	)
	echo pdb for debug
	copy /Y %source%\DatePicker.pdb "%appdata%\Microsoft\AddIns"
)
if "%source%"=="bin\Release" (
	echo copying Distribution
	copy /Y %source%\DatePicker-AddIn.xll Distribution\
	copy /Y %source%\DatePicker-AddIn.dna Distribution\
	copy /Y %source%\DatePicker.dll Distribution\
	copy /Y %source%\DatePicker-AddIn64-packed.xll Distribution\DatePicker64.xll
	copy /Y %source%\DatePicker-AddIn-packed.xll Distribution\DatePicker32.xll
	copy /Y TestVBA.xlsm Distribution
)
pause
