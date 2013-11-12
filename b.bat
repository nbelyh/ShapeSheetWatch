@echo off

MKDIR x86\Release
MKDIR x86\Debug
MKDIR x64\Release
MKDIR x64\Debug

COPY AddinSetup\lib\x86\htmlayout.dll x86\Debug\htmlayout.dll
COPY AddinSetup\lib\x86\htmlayout.dll x86\Release\htmlayout.dll
COPY AddinSetup\lib\x64\htmlayout.dll x64\Debug\htmlayout.dll
COPY AddinSetup\lib\x64\htmlayout.dll x64\Release\htmlayout.dll

%WINDIR%\Microsoft.NET\Framework\v3.5\MSBuild.exe %* /p:Platform=x86 /p:Configuration=Release
%WINDIR%\Microsoft.NET\Framework\v3.5\MSBuild.exe %* /p:Platform=x64 /p:Configuration=Release

RMDIR /S /Q ship
MKDIR ship
COPY x64\Release\*.msi ship
COPY x86\Release\*.msi ship
