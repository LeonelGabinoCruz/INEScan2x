@echo off
cd /d "%~dp0"
title INEScan 2x
python escaneo_ine.py
if errorlevel 1 pause
