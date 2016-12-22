@echo off
cd /d %~dp0
"C:\Program Files (x86)\AutoIt3\Aut2Exe\Aut2exe.exe" /In wincdp2.au3 /out wincdpcmd.exe /icon cisco.ico /comp 3 /console /x86
