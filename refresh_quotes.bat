@echo off
rem Call stock quotes API, refresh stock quotes
rem The Windows CMD can not support UNC path 
rem So you have to map the network share path to local path before running

python qq_quotes.py

rem Keep the execution window, otherwise it will be closed automatically
pause