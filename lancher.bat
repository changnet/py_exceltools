@echo off

python loader.py --input ./ --srv server/ --clt client/ --timeout -1 --suffix .xlsx -w lua

pause
