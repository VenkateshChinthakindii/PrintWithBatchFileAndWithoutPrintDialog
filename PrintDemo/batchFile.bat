@echo off
pushd "E:\Practice\PrintFilesDi"
for %%a in (*.pdf *.docx *.txt *.xlsx *.xls) do (
   call E:\Practice\PrintDemo\PrintDemo\printjs.bat "%%~fa"  
)