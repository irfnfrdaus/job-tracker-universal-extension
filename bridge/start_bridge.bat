@echo off
title JobPulse Desktop Bridge (Windows)
echo ========================================================
echo   JobPulse Universal Excel Bridge
echo ========================================================
echo Checking Python dependencies...
python -m pip install -q -r requirements.txt

echo Starting bridge on http://127.0.0.1:5252 ...
python bridge.py
pause
