; ############ AUTO-EXECUTE ############

#noenv
#singleinstance,force
sendmode,input
setworkingdir,%a_scriptdir%
coordmode,mouse,client
coordmode,pixel,client


; ############ HOTKEYS ############

F10::reload
F11::exitapp
lshift::lctrl
capslock::lshift
;1::^w
;2::+^Tab
;3::^Tab
;4::^t
;5::^r

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


; #### Hotstrings ####
::mytel::{+}353 (0) 91 433532
::@ct::send,sap_cloud_terminations@sap.com{tab 2}^a{backspace}
::@jana::jana.kerschl.sudekova@sap.com
::dt::
	date := gen_dt()
	send % date
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
