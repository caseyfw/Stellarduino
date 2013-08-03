'set scope = CreateObject("EQMOD_2.Telescope")
set scope = CreateObject("EQMOD.Telescope")
scope.Connected = true
msgbox("Motor Controller Firmware Version =" & scope.commandstring(":MOUNTVER#"))
