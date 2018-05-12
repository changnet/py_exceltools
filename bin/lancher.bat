@echo off

win10_x64\reader.exe --input ./ --srv server/ --clt client/ --timeout -1 --suffix .xlsx --swriter lua --cwriter json

pause
