@echo off

python reader.py --input ./ --srv server/ --clt client/ --timeout -1 --suffix .xlsx --swriter lua --cwriter json

python reader.py --input ./ --srv server/ --clt client/ --timeout -1 --suffix .xlsx --cwriter lua
python reader.py --input ./ --srv server/ --clt client/ --timeout -1 --suffix .xlsx --cwriter xml

pause
