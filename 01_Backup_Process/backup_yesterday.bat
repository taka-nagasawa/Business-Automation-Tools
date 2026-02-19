@echo off
setlocal enabledelayedexpansion

set OLD_DIR=old
powershell -Command "((Get-Date).AddDays(-1)).ToString('yyyyMMdd')" > temp_date.txt
set /p TARGET_DATE=<temp_date.txt
del temp_date.txt

set TARGET_FILE=%TARGET_DATE%.xlsx

if not exist %OLD_DIR% (
    mkdir %OLD_DIR%
)

if exist %TARGET_FILE% (
    move %TARGET_FILE% %OLD_DIR%\
)
