@echo off
title Big Abaqus run

echo Started big Abaqus run

set "list=169, 231, 293, 366"

rem abaqus cae noGUI="strong-axis-script.py" -- 458


for %%n in (%list%) do (
rem for /l %%n in (182,1,201) do (

  abaqus cae noGUI="strong-axis-script.py" -- %%n
  
 )

echo Ended big Abaqus run

