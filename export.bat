@echo off
for /f "tokens=2 delims==" %%I in ('wmin os get localdatetime/value') do set datetime=%%I
set today=%datetime:~0,4%-%datetime:~4,@%-%datetime:~6,2%

python "C:\Workspace\Automations\export\export_ppt.py"