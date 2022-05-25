@echo off

echo %time%
set query="%~f1.sql"
set out="%~f2.sql"

chcp 1252 > $null

REM SERVIDOR GCTI
::sqlcmd -S154.12.245.234 -Usa -PJuana123 -dmaster -I -w8100 -i%query% -o%out% -b

REM SERVIDOR MINAS
sqlcmd -S208.118.63.167 -Udb_a7c315_gesum_admin -PScott2021 -ddb_a7c315_gesum -I -w8100 -i%query% -o%out% -b -y8000

:: sqlcmd -S. -Usa -Pjuana -dSCP -I -w8100 -i%query% -o%out% -b
:: sqlcmd -S.\SQLEXPRESS -Usa -Pjuana -I -w8100 -i%query% -o%out% -b -y8000
:: -y8000



echo %time%
if errorlevel = 1 (type %out%)
