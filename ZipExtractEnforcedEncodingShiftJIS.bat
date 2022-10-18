@echo off

rem Preparation
rem - Place this batch script (or its shortcut) to the SendTo folder.
rem - 7zip is installed at C:\Program Files
rem - 7zip can be downloaded from https://www.7-zip.org/download.html
rem
rem How to use:
rem - Right click a zip archive.
rem - On file explorer, select this batch script from SendTo right-click
rem - Drag and drop the output folder
rem - Press enter

title Shift-JIS Zip Extraction

rem set code page of command prompt to 932 (Shift-JIS)
rem if other encoding is desired, change this number.
chcp 932

rem print the passed file path for debugging purpose.
echo Extracting %1 in Shift-JIS Encoding...

rem get output folder.
echo ------------------------------------------------------------
echo Drag and drop the output folder where you want to extract to.
echo In case of current directory, simply type in dot(.)
set /p "DirOut=Output Empty Folder: "
echo ------------------------------------------------------------

rem extract zip file with code page 932 at current directory.
rem command = "C:\Program Files\7-Zip\7z.exe" 7z.exe references 7z.dll.
rem arg[0]  = x: extract including the folder path
rem arg[1]  = %1: passed path from Sendto
rem arg[2]  = -o%DirOut%: output folder
rem arg[3]  = -mpc=932: use code page 932 for decoding file/folder name
"C:\Program Files\7-Zip\7z.exe" x %1 -o%DirOut% -mcp=932

rem prevent the command prompt from closing
pause
