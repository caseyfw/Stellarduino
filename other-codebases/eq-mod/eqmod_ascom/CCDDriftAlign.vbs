 Dim endtime

 ' connect to EQASCOM
' set scope = CreateObject("EQMOD_Sim.Telescope")
'set scope = CreateObject("EQMOD_2.Telescope")
 set scope = CreateObject("EQMOD.Telescope")
 scope.Connected = true

 ' Move mount west at sidereal x 2
 ' moveaxis(a,b) where a=0 for RA axis, a=1, for DEC axis, b is rate in degrees/sec
 scope.moveaxis 0,0.00832

 'wait 60 secs
 endtime = timer + 60
 do while timer < endtime
 loop

 ' Stop tracking and drift back
 scope.moveaxis 0,0
 scope.tracking = false

 'wait 60 secs
 endtime = timer + 60
 do while timer < endtime
 loop

 ' Start tracking again
 scope.tracking = true

