@echo off
cd /d %~dp0
if exist "C:\Program Files (x86)\AutoIt3\Aut2Exe\Aut2exe.exe" set a2e="C:\Program Files (x86)\AutoIt3\Aut2Exe\Aut2exe.exe" 

if exist "C:\Program Files\AutoIt3\Aut2Exe\Aut2exe.exe" set a2e="C:\Program Files\AutoIt3\Aut2Exe\Aut2exe.exe" 


%a2e% /In wincdpcmd.au3 /out wincdpcmd.exe /icon cisco.ico /comp 3 /console /x86
%a2e% /In wincdp.au3 /out wincdp.exe /icon cisco.ico /comp 3  /x86
dir