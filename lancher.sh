#!/bin/bash

python3 exceltools.py --input ./ --srv server/ --clt client/ --timeout -1 --suffix .xlsx --swriter lua --cwriter json
python3 exceltools.py --input ./ --srv server/ --clt client/ --timeout -1 --suffix .xlsx --cwriter xml
python3 reaexceltoolsder.py --input ./ --srv server/ --clt client/ --timeout -1 --suffix .xlsx --cwriter lua

