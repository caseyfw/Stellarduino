dim ra 
dim dec

' connect to EQMOD telescope simulator (included for debug)
'set scope = CreateObject("EQMOD_sim.Telescope")

' connect to EQMOD telescope server Instance 2
'set scope = CreateObject("EQMOD_2.Telescope")


' connect to EQMOD telescope server
set scope = CreateObject("EQMOD.Telescope")

' connect to driver
scope.Connected = true

If scope.connected = true Then
	txt1=InputBox("RA Hours")
	txt2=InputBox("RA Mins")
	txt3=InputBox("RA Secs")

	ra = cdbl(txt1) + cdbl(txt2)/60 + cdbl(txt3)/3600

	txt1=InputBox("DEC Degrees")
	txt2=InputBox("DEC Mins")
	txt3=InputBox("DEC Secs")

	dec = cdbl(txt1) + cdbl(txt2)/60 + cdbl(txt3)/3600
	
	If scope.tracking Then

		' issue sync
		scope.SyncToCoordinates ra, dec
	Else
		msgbox("Sync failed - mount isn't tracking")
	End If
Else
	msgbox("Can't connect to driver")
End If

