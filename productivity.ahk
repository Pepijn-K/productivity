; ############ AUTO-EXECUTE ############

#noenv
setworkingdir,%a_scriptdir%
sendmode,input
coordmode,mouse,client
coordmode,pixel,client

; #### Common URLs ####

msgbox % sum(4,32),Working?,Run SAP related programs?,30
ifmsgbox,yes
	gosub,runsap

; ############ HOTKEYS ############

F10::reload
F11::exitapp
lshift::lctrl
capslock::lshift
NumpadMult::^w
NumpadDiv::+^Tab
NumpadSub::^Tab
NumpadAdd::^t

; #### Numpad keys ####
^Numpad1::pid1 := do("outlook",pid1)
^Numpad2::pid2 := do("onenote",pid2)
^Numpad3::pid3 := do("logon",pid3)
^Numpad4::pid4 := do("ssf",pid4)
^Numpad5::pid5 := do("ssf",pid5)
^Numpad6::pid6 := do("crm",pid6)
^Numpad7::pid7 := do("store",pid7)

; #### Mouse keys ####
+rbutton::send,^c
^rbutton::send,^v
+^rbutton::send,^v{enter}

; #### Winkey Tasks ####
#c::									; call multibox
	inputbox,usin,Multibox,,,240,120
	if errorlevel
		return
	else
		{
		iniread,val,lib/db.ini,mbsearch,%usin%,0
		if(val)
			run,%val%
		else
			search(usin)
		}
return
+#c::									; create or edit multibox alias
	inputbox,dest,Program/path to execute,,,240,120
	if errorlevel
		return
	inputbox,sc,Alias to use,,,240,120
	if errorlevel
		return
	iniwrite,%dest%,lib/db.ini,mbsearch,%sc%
	msgbox, % "Enter " sc " to run:`n" dest
return

#s::									; search selected
	clipboard := ""
	send,^c
	clipwait
	search(clipboard)
return

#d::									; create Outlook task from anywhere
	settitlematchmode,2
	if(winexist("ahk_pid " pid1))
		winactivate
	else
		return
	send,+^k
return

; #### Hotstrings ####
::mytel::{+}353 (0) 91 433532
:*:@ct::send,sap_cloud_terminations@sap.com{tab 2}^a{backspace}
:*:@jana::jana.kerschl.sudekova@sap.com
::@msol::marisol.torres@sap.com
::tod::
	date := gen_dt()
	send % date
return


; ############ SUBROUTINES ############

runsap:
	do("outlook",pid1)
	do("onenote",pid2)
	do("logon",pid3)
	do("ssf",pid4)
	do("ssf",pid5)
	do("crm",pid6)
return


; ############ FUNCTIONS ############

gen_dt(format := "yyyMMdd") {
	if(format = "date")
		formattime,res,,dd-MM
	else if(format = "time")
		formattime,res,,HH:mm
	else
		formattime,res,,%format%
	return res
}

search(query) {
	if(winexist("ahk_exe chrome.exe"))
		{
		winactivate
		send,^t
		}
	else
		{
		run,chrome.exe
		winwaitactive
		}
	sleep,200
	send %query% {enter}
}

do(p,id) {
	if(!id)
		{
		iniread,src,lib/db.ini,running,%p%
		run % src
		}
	else
		winactivate,ahk_pid %id%	
}

sum(x*) {
	tot := 0
	for i, y in x
		tot := y + tot
	return tot
}