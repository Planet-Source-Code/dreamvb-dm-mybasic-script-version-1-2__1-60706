@echo off
CLS

echo This will compress engine.exe using UPX this process
echo will only need to preformed once.
echo.
pause


if exist "upx.exe" goto CompressExe

goto ende


:CompressExe
upx --best --crp-ms=100000 engine.exe 2>log.txt
goto Finsihed

:Finsihed
CLS
more log.txt
echo.
pause
del log.txt
goto Close

:ende
CLS
echo.
echo. UPX was not found on your system.
echo. Please make sure the application was not deleted by mistake.
echo.
pause

:Close
