@echo off

echo.
echo WANRING, DATA LOSS IF YOU CONTINUE
echo.
choice /C yn /N /M "Delete Partitions? (y/N)"
if /I "%errorlevel%" neq "1" goto end

(echo Select Disk 0
echo clean
)  | diskpart

:end