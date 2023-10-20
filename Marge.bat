@echo off

REM File marge

type *.dat > Marge.log

mkdir dat

move /-Y *.dat .\dat\

exit

