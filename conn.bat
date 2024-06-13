@echo off
"C:/Program Files/Firebird/Firebird_2_5/bin/isql.exe" -user sysdba -password Q5QIST "C:\robo\CONTROLE.FDB" -i "c:\robo\script.sql" -o "c:\robo\conteiner.txt"