
copy "EQMOD.EXE" "%CommonProgramFiles%\ASCOM\Telescope"
copy "EQMOD_SIM.EXE" "%CommonProgramFiles%\ASCOM\Telescope"
copy "eqcontrl.dll" "%CommonProgramFiles%\ASCOM\Telescope"
copy "eqmod??.dll" "%CommonProgramFiles%\ASCOM\Telescope"

cd "%CommonProgramFiles%\ASCOM\Telescope"

"EQMOD.EXE" /regserver
"EQMOD_SIM.EXE" /regserver

pause