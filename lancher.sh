#!/bin/bash

python3 exceltools.py --input ./ --srv server/ --clt client/ --timeout -1 --swriter lua --cwriter json
python3 exceltools.py --input ./ --srv server/ --clt client/ --timeout -1 --cwriter xml
python3 exceltools.py --input ./ --srv server/ --clt client/ --timeout -1 --cwriter lua

