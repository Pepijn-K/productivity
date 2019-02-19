#MaxThreadsPerHotkey 2
!#m::
one := A_Screenwidth*.2
two := A_Screenwidth*.3
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
		sleep,120000
		MouseMove,%two%,%vert%
		sleep,400
		}
return