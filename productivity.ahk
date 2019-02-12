; ############ AUTO-EXECUTE ############

#noenv
setworkingdir,%a_scriptdir%
sendmode,input
coordmode,mouse,client
coordmode,pixel,client
settitlematchmode,2

; #### GUIs ####
#include lib/wait.ahk

; #### VARs ####
login := "C:\Users\I349302\OneDrive - SAP SE\Documents\sec.ini"

msgbox % sum(4,32),Working?,Run SAP related programs?,30
ifmsgbox,yes
	gosub,runsap

; ############ HOTKEYS ############

F10::reload
lshift::lctrl
capslock::lshift
NumpadMult::^w
NumpadDiv::+^Tab
NumpadSub::^Tab
NumpadAdd::^t

; #### Numpad keys ####
^Numpad1::winactivate,ahk_id %one%		; window: x: 592	y: 0	w: 1748	h: 1450
^Numpad2::winactivate,ahk_id %two%		; window: x: 990	y: 0	w: 1287	h: 1450
^Numpad3::winactivate,ahk_id %three%	; window ahk_class LyncTabFrameHostWindowClass: x: 722	y: 199	w: 1396	h: 1075
^Numpad4::winactivate,ahk_id %four%		; window: x: 321	y: 0	w: 2088	h: 1459
^Numpad5::winactivate,ahk_id %five% 	; window: x: 406	y: 0	w: 2003	h: 1459
^Numpad6::winactivate,ahk_id %six%		; window: x: 494	y: 0	w: 1915	h: 1459
/*
^Numpad7::seven_id := do(seven_id,"seven","seven_name") ; window: irr (FOLLOW CHROME)
^Numpad8::eight_id := do(eight_id,"eight","eight_name")	; window: x: 449	y: 70	w: 1794	h: 1320
^Numpad9::nine_id := do(nine_id,"nine","nine_name")	; window: x: 449	y: 70	w: 1794	h: 1320
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
	msgbox % "Enter " sc " to run:`n" dest
return

#s::									; search selected
	search_holder := clipboardall
	clipboard := ""
	send,^c
	clipwait
	search(clipboard)
	clipboard := search_holder
	search_holder := ""
return

#d::									; create Outlook task from anywhere
	settitlematchmode,2
	if(winactive("- Task"))
		send,{alt}hav
	else if(winexist("ahk_id " one))
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
	if(winexist("ahk_id " one))
		winactivate
	else
		{
		run,notepad.exe
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

#o::									; ask Silvia for processor of O2I ticket
	if(winexist("ahk_id " one))
		{
		winactivate
		send,^1{lalt}hn1
		}
	else
		{
		run,notepad.exe
		winwaitactive
		}
	send,silvia.peves@sap.com{tab 4}Ticket %clipboard%{tab}Hi Silvia,{enter 2}Could you please tell me who owns the above ticket?{enter 2}Many thanks!
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

#y::
	inputbox,patopa,Path to be converted,,,200,120
	if errorlevel
		return
	clipboard := strreplace(patopa,"\","/")
return

#include lib/wiggle.ahk

; #### Hotstrings ####
::mytel::{+}353 (0) 91 433532
:*:@ct::sap_cloud_terminations@sap.com{tab 2}^a{backspace}
:*:@jana::jana.kerschl.sudekova@sap.com
:*:@msol::marisol.torres@sap.com
:*:@jette:jette.bork-wagenblast@sap.common
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
	gui,wait:show,x0 y0 w%a_screenwidth% h%a_screenheight%
	run,outlook.exe			; window: x: 592	y: 0	w: 1748	h: 1450
	run,onenote.exe			; window: x: 990	y: 0	w: 1287	h: 1450
	run,lync.exe			; window ahk_class LyncTabFrameHostWindowClass: x: 722	y: 199	w: 1396	h: 1075
	one := winexist("ahk_exe OUTLOOK.EXE")
	two := winexist("ahk_exe ONENOTE.EXE")
	three := winexist("ahk_exe lync.exe")
	run,iexplore.exe https://sfp.wdf.sap.corp/sap(bD1lbiZjPTAwMSZkPW1pbg==)/bc/bsp/sap/crm_ui_start/default.htm		; window: x: 321	y: 0	w: 2088	h: 1459
	loop,25
		{
		four := winexist("SAP - [Select a business role: ]")
		if(four)
			break
		else
			{
			sleep,200
			continue
			}
		}
	controlsend,,{tab 6}{enter},ahk_id %four%
	winwait,SAP Store Operations - Level2 - Internet Explorer
	run,iexplore.exe https://sfp.wdf.sap.corp/sap(bD1lbiZjPTAwMSZkPW1pbg==)/bc/bsp/sap/crm_ui_start/default.htm 		; window: x: 406	y: 0	w: 2003	h: 1459
	loop,25
		{
		five := winexist("SAP - [Select a business role: ]")
		if(five)
			break
		else
			{
			sleep,200
			continue
			}
		}
	controlsend,,{tab 6}{enter},ahk_id %five%
	sleep,4000
	controlsend,,{tab 2}{down}{tab}{enter},ahk_id %five%
	run,iexplore.exe https://icp.wdf.sap.corp/sap(bD1lbiZjPTAwMSZkPW1pbg==)/bc/bsp/sap/crm_ui_start/default.htm			; window: x: 494	y: 0	w: 1915	h: 1459
	loop,25
		{
		six := winexist("Select a business role: - [SAP] - Internet Explorer")
		if(six)
			break
		else
			{
			sleep,200
			continue
			}
		}
	controlsend,,{tab 4}{enter},ahk_id %six%
	gui,wait:cancel
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
/*
do(id,pr,n) {
	settitlematchmode,2
	if(id)
		winactivate,ahk_id %id%
	else
		{
		iniread,app,lib/db.ini,running,%pr%
		iniread,appname,lib/db.ini,running,%n%
		run % app
		winwait,%appname%
		h := winexist(appname)
		}
	return h
}
*/
sum(x*) {
	tot := 0
	for i, y in x
		tot := y + tot
	return tot
}