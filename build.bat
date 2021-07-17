@echo on
cd /d %cd%

set mode=production
call webpack

pause