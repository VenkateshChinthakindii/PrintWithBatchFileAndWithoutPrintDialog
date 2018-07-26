@echo off
RUNDLL32 PRINTUI.DLL,PrintUIEntry /y /n %1
pushd %2
for %%a in (*.pdf *.doc) do (
   call %3 "%%~fa"
)