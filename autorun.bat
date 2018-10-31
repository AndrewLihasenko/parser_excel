@echo on

:bootstart
rem pskill python

python parser_excel.py

timeout /T 3600

goto bootstart
