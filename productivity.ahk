; ############ AUTO-EXECUTE ############

#noenv
setworkingdir,%a_scriptdir%
sendmode,input
coordmode,mouse,client
coordmode,pixel,client

; #### Common URLs ####
sfp := "https://sfp.wdf.sap.corp/sap(bD1lbiZjPTAwMSZkPW1pbg==)/bc/bsp/sap/crm_ui_start/default.htm"
crm := "https://icp.wdf.sap.corp/sap(bD1lbiZjPTAwMSZkPW1pbg==)/bc/bsp/sap/crm_ui_start/default.htm"

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
^Numpad1::
	if(winexist("ahk_pid " outlook_pid))
		winactivate
	else
		outlook_pid := run("outlook")
return

^Numpad2::
	if(winexist("ahk_pid " onenote_pid))
		winactivate
	else
		onenote_pid := run("onenote")
return

^Numpad3::
	if(winexist("ahk_pid " logon_pid))
		winactivate
	else
		logon_pid := run("saplogon")
return

^Numpad4::
	if(winexist("ahk_pid " ssf1_pid))
		winactivate
	else
		ssf1_pid := open(sfp)
return

^Numpad5::
	if(winexist("ahk_pid " ssf2_pid))
		winactivate
	else
		ssf2_pid := open(sfp)
return

^Numpad6::
	if(winexist("ahk_pid " crm_pid))
		winactivate
	else
		crm_pid := open(crm)
return

^Numpad7::
	if(winexist("ahk_pid " sto_pid))
		winactivate
	else
		sto_pid := run("firefox")
return


/*
loop,9
	{
	keyname := "Numpad" a_index
	hotkey

*/

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
	send,^c
	search(a_clipboard)
return

#d::									; create Outlook task from anywhere
	settitlematchmode,2
	if(winexist("ahk_pid " outlook_pid))
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
::dt::
	date := gen_dt()
	send % date
return


; ############ SUBROUTINES ############

runsap:
	run,outlook.exe,,,outlook_pid
	run,onenote.exe,,,onenote_pid
	run,saplogon.exe,,,logon_pid
	run,%ssf%,,,ssf1_pid
	run,%ssf%,,,ssf2_pid
	run,%crm%,,,crm_pid
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
	sleep,400
	send,%query%{enter}
}

run(exe) {
	if(exe = "firefox")
		exe .= ".exe -private-window https://www.sapstore.com/"
	else
		exe .= ".exe"
	run,%exe%,,,pid
	return pid
}

open(url) {
	run,%url%,,,pid
	return pid
}

sum(x*) {
	tot := 0
	for i, y in x
		tot := y + tot
	return tot
}