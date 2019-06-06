@echo off

Del /s Client-Package.zip >nul 2>&1

copy ..\Mmos.exe Support\. > nul
copy ..\..\Loader\Loader.exe Support\. > nul

echo .
echo Now run MMOS.bat in the support folder
echo .

pause

"c:\Program Files\7-Zip\7z.exe" a Client-Package.zip -xr!*.bat -xr!.gitignore -xr!*.txt -x!Support\


echo .
pause