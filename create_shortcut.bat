@echo off
set "pythonw_path=%~dp0.venv\Scripts\pythonw.exe"

set "script_path=%~dp0main.py"

set "shortcut_path=%~dp0FuzzyImgCompare.lnk"

powershell -command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%shortcut_path%'); $s.TargetPath = '%pythonw_path%'; $s.Arguments = '%script_path%'; $s.Save()"

echo Shortcut to start main.py created as: FuzzyImgCompare.lnk
pause
