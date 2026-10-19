#!/usr/bin/env bash
echo "========================================================"
echo "   JobPulse Universal Excel Bridge (macOS / Linux)"
echo "========================================================"
echo "Checking Python dependencies..."
python3 -m pip install -q -r requirements.txt 2>/dev/null || pip install -q -r requirements.txt

echo "Starting bridge on http://127.0.0.1:5252 ..."
python3 bridge.py || python bridge.py
