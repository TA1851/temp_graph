# ファイルをマージしてマージファイルを作成し、新規作成したフォルダに元データを移動させる
@echo off

REM File marge

type *.dat > Marge.log

mkdir dat

move /-Y *.dat .\dat\

exit

