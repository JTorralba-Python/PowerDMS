@echo off

PowerShell -ExecutionPolicy ByPass -file "WebDriver.ps1"

python Fetch.py

@echo on
