'set scope = CreateObject("EQMOD_sim.Telescope")
'set scope = CreateObject("EQMOD_2.Telescope")
set scope = CreateObject("EQMOD.Telescope")
scope.Connected = true
'msgbox("ra=" & scope.ramotor & " Dec=" & scope.decmotor )
msgbox("ra=" & scope.commandstring(":RA_ENC#") & " Dec=" & scope.commandstring(":DEC_ENC#") )
