::** BAT file for FrapsJoin
::
::usage: 
::  - right click any AVI file
::  - "Open With...", browse here, Enter
::  - this BAT file will now appear in the "Open With" list for AVI files
::
start wscript "%~dp0\FrapsJoin.vbs" %1
exit 
