@echo off

python reader.py --input ./ --srv server/ --clt client/ --timeout -1 --suffix .xlsx --swriter lua --cwriter json

pause
