@echo off

echo.
echo WANRING, DATA LOSS IF YOU CONTINUE
echo.
SET /P input=Delete Partitions? (y/N)?
IF /I "%input%" NEQ "Y" GOTO end

(echo Select Disk 0
echo clean
)  | diskpart

:end