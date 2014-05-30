@echo off
SET DISABLEOUTOFPROCTASKHOST=1

SET HTMLayoutSDK=C:\Projects\Libs\HTMLayoutSDK\

"%ProgramFiles(x86)%\MSBuild\12.0\Bin\MSBuild.exe" %* /p:Platform=x86 /p:Configuration=Release
"%ProgramFiles(x86)%\MSBuild\12.0\Bin\MSBuild.exe" %* /p:Platform=x64 /p:Configuration=Release
"%ProgramFiles(x86)%\MSBuild\12.0\Bin\MSBuild.exe" %* /p:Platform=x86 /p:Configuration=Debug
"%ProgramFiles(x86)%\MSBuild\12.0\Bin\MSBuild.exe" %* /p:Platform=x64 /p:Configuration=Debug
