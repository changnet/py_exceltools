@echo off

python exceltools.py --input ./ --srv server/ --clt client/ --timeout -1 --swriter lua --cwriter json

python exceltools.py --input ./ --srv server/ --clt client/ --timeout -1 --cwriter lua
python exceltools.py --input ./ --srv server/ --clt client/ --timeout -1 --cwriter xml

pause
