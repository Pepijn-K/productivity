#MaxThreadsPerHotkey 2
!#m::
coordmode,mouse,screen
one := A_Screenwidth*.46
two := A_Screenwidth*.54
vert := A_Screenheight/2

#MaxThreadsPerHotkey 1
	if keep_wiggling = y
		{
		keep_wiggling = n
		return
		}
	keep_wiggling = y
	loop
		{
		if keep_wiggling = n
			return
		MouseMove,%one%,%vert%
		sleep,1000
		MouseMove,%two%,%vert%
		sleep,1000
		}
return