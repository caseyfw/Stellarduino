
copy "EQMOD.EXE" "%CommonProgramFiles%\ASCOM\Telescope"
copy "eqcontrl.dll" "%CommonProgramFiles%\ASCOM\Telescope"

cd "%CommonProgramFiles%\ASCOM\Telescope"

"EQMOD.EXE" /regserver

pause