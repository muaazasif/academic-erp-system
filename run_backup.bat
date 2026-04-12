@echo off
echo ============================================================
echo    MyERP Database Backup Tool
echo ============================================================
echo.

cd /d "%~dp0"

echo [1/3] Creating database backup...
python backup_database.py

echo.
echo [2/3] Checking backup files...
if exist "database_backups\*.db" (
    echo    OK - Database backups found
    dir /b /o-d database_backups\*.db | findstr /n "^" | findstr "^1:" 
) else (
    echo    WARNING - No database backups found!
)

if exist "data_exports\*.json" (
    echo    OK - JSON exports found
) else (
    echo    WARNING - No JSON exports found!
)

if exist "data_exports\*.xlsx" (
    echo    OK - Excel exports found
) else (
    echo    WARNING - No Excel exports found!
)

echo.
echo [3/3] Backup Summary
echo ============================================================
echo Backups stored in: %CD%\database_backups
echo Exports stored in: %CD%\data_exports
echo ============================================================
echo.
echo Press any key to exit...
pause >nul
