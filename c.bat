@echo off
IF EXIST ship RMDIR /S /Q ship
IF EXIST x86 RMDIR /S /Q x86
IF EXIST x64 RMDIR /S /Q x64
IF EXIST AddinSetup\bin RMDIR /S /Q AddinSetup\bin
IF EXIST AddinSetup\obj RMDIR /S /Q AddinSetup\obj
IF EXIST Addin\x64 RMDIR /S /Q Addin\x64
IF EXIST Addin\x86 RMDIR /S /Q Addin\x86

IF EXIST Addin\*_i.h DEL /S /Q Addin\*_i.h
IF EXIST Addin\*_i.c DEL /S /Q Addin\*_i.c
IF EXIST Addin\*.user DEL /S /Q Addin\*.user
IF EXIST *.ncb DEL /S /Q *.ncb
IF EXIST *.cache DEL /S /Q *.cache
IF EXIST *.aps DEL /S /Q *.aps
