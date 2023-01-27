start/w setversion.exe

pyinstaller -w -F -i "logo.ico" ExcelToWorld.py

xcopy %CD%\*.xltx %CD%\dist /H /Y /C /R
xcopy %CD%\*.dotx %CD%\dist /H /Y /C /R
xcopy %CD%\*.ico %CD%\dist /H /Y /C /R
xcopy %CD%\*.ini %CD%\dist /H /Y /C /R

xcopy C:\vxvproj\tnnc-ExcelToWorld\tnnc-ExcelToWorld\dist C:\vxvproj\tnnc-ExcelToWorld\ConsoleApp\ /H /Y /C /R
