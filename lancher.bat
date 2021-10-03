@echo off

python exceltools.py --input ./ --srv server/ --clt client/ --timeout -1 --suffix .xlsx --swriter lua --cwriter json

python exceltools.py --input ./ --srv server/ --clt client/ --timeout -1 --suffix .xlsx --cwriter lua
python exceltools.py --input ./ --srv server/ --clt client/ --timeout -1 --suffix .xlsx --cwriter xml

pause
