; ############ AUTO-EXECUTE ############

#noenv
setworkingdir,%a_scriptdir%
sendmode,input
coordmode,mouse,client
coordmode,pixel,client

; #### VARs ####
global login := "C:\Users\I349302\OneDrive - SAP SE\Documents\sec.ini"

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
^Numpad7::pid7 := do("bo",pid7)
^Numpad8::pid8 := do("store",pid8)
^Numpad9::pid9 := do("bydstore",pid9)

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
			run % val
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
	if(winactive("- Task"))
		{
		sleep,200
		send,{lalt}1{alt}fc
		}
	else if(winexist("ahk_pid " pid1))
		{
		winactivate
		send,+^k
		}
	else
		return
return

#w::									; send working hours email
	inputbox,locnhrs,Location and hours per day,,,180,120
	if(errorlevel)
		return
	loop,parse,locnhrs,csv
		{
		if(a_loopfield = 8)
			lh := "8:00 - 16:00"
		else if(a_loopfield = 9)
			lh := "9:00 - 17:00"
		else if(a_loopfield = "h")
			lh := "home"
		else if(a_loopfield = "o")
			lh := "office"
		else
			lh := "NULL"
		lh%a_index% := lh
		}
	settitlematchmode,1
	if(winexist("ahk_pid " pid1))
		winactivate
	else
		{
		pid1 := do("outlook",pid1)
		winwaitactive
		}
	send,^n
	sleep,600
	send,ann{enter}car{enter}joh{enter}{tab}wain{enter}{tab 2}Working hours this week{tab}
(
Hi guys,

Monday: %lh1%, %lh2%
Tuesday: %lh3%, %lh4%
Wednesday: %lh5%, %lh6%
Thursday: %lh7%, %lh8%
Friday: %lh9%, %lh10%
)
return

!#l::									; login common websites
	sleep,300
	settitlematchmode,1
	if(winactive("Logon - Internet Explorer"))						; KCF Frontend
		{
		iniread,u,%login%,access,usr_gbl
		iniread,p,%login%,access,pw_s10
		Send,%u%{tab}%p%{enter}
		}
	else if WinActive("SAP Store | Buy SAP Software")				; ASM
		{
		iniread,u,%login%,access,usr_gbl
		iniread,p,%login%,access,pw_str
		Send,%u%{tab}%p%{enter}
		}
return


; #### Hotstrings ####
::mytel::{+}353 (0) 91 433532
:*:@ct::send,sap_cloud_terminations@sap.com{tab 2}^a{backspace}
:*:@jana::jana.kerschl.sudekova@sap.com
:*:@msol::marisol.torres@sap.com
::tfgit::Thanks for getting in touch.
::tyfyi::Thank you for your interest!
::tyffu::Thank you for following up.
::itaeichw::If there's anything else I can help with please don't hesitate to let me know.
::itdnsti::If this does not solve the issue please let me know in a reply and I will investigate further.
::sig::{lalt}eas{enter}
::tod::
	date := gen_dt()
	send % date
return


; ############ SUBROUTINES ############

runsap:
	pid1 := do("outlook",pid1)
	pid2 := do("onenote",pid2)
	pid3 := do("logon",pid3)
	pid4 := do("ssf",pid4)
	pid5 := do("ssf",pid5)
	pid6 := do("crm",pid6)
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
	if(winexist("ahk_pid " id))
		winactivate
	else
		{
		iniread,p_res,lib/db.ini,running,%p%
		run % p_res,,,id
		}
	return id
}

sum(x*) {
	tot := 0
	for i, y in x
		tot := y + tot
	return tot
}
