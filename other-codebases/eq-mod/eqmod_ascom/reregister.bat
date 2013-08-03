"EQMOD.EXE" /unregserver
"EQMOD_SIM.EXE" /unregserver
"%CommonProgramFiles%\ASCOM\Telescope\EQMOD.EXE" /unregserver
"%CommonProgramFiles%\ASCOM\Telescope\EQMOD_SIM.EXE" /unregserver

register.bat
rem copy "EQMOD.EXE" "%CommonProgramFiles%\ASCOM\Telescope\"
rem copy "eqcontrl.dll" "%CommonProgramFiles%\ASCOM\Telescope\"

rem cd "%CommonProgramFiles%\ASCOM\Telescope\"

rem "EQMOD.EXE" /regserver

pause
