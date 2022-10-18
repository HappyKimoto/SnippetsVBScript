@echo off

rem How to use
rem - Place this batch script at SendTo folder.

title Shift-JIS Zip Extraction

rem set code page of command prompt to 932 (Shift-JIS)
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
