
"%CommonProgramFiles%\ASCOM\Telescope\EQMOD.EXE" /unregserver

copy "EQMOD.EXE" "%CommonProgramFiles%\ASCOM\Telescope\"
copy "eqcontrl.dll" "%CommonProgramFiles%\ASCOM\Telescope\"

cd "C:\Program Files\Common Files\ASCOM\Telescope"

"EQMOD.EXE" /regserver

pause