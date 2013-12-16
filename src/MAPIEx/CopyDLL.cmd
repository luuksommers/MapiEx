@echo on
md ..\%1
copy %1\MAPIEx.dll ..\%1
md ..\TestNetMAPI\bin\%1
copy %1\MAPIEx.dll ..\TestNetMAPI\bin\%1
@echo off
