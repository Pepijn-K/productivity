; ############ AUTO-EXECUTE ############

#noenv
#include C:/Users/I349302/Documents/AHK/func_lib/funcs.ahk
#notrayicon
#singleinstance,force
setworkingdir,%a_scriptdir%
sendmode,input
/*
msgbox % a_screenwidth

msgbox,4,,Continue?
ifmsgbox,no
	return
*/
; #### VARs ####
if(a_screenwidth = 1920)
	{
	scaling := 125
	dim := 100 / scaling
	screenheight := a_screenheight * dim
	screenwidth := a_screenwidth * dim
	}
else if(a_screenwidth = 1600)
	{
	scaling := 100
	screenheight := a_screenheight
	screenwidth := a_screenwidth
	}
	

; #### GUIs ####
#include lib/wait.ahk


runsap:
	settitlematchmode,1
	gui,wait:show,x0 y0 w%screenwidth% h%screenheight%,Waiting for stuff to load
	winset,transparent,216,Waiting for stuff to load
;	run,C:\Users\I349302\Documents\AHK\SSF_assistant\SSF_assistant.ahk
	run,outlook.exe			; window: x: 592	y: 0	w: 1748	h: 1450
	run,onenote.exe			; window: x: 990	y: 0	w: 1287	h: 1450
	run,lync.exe			; window ahk_class LyncTabFrameHostWindowClass: x: 722	y: 199	w: 1396	h: 1075
	winwait,Inbox - pepijn.krijnsen@sap.com,,30
	if(errorlevel)
		msgbox,Outlook did not start as expected
;	msgbox % num1
	num1 := assign("outlook")
;	msgbox % num1
	num2 := assign("onenote")
;	num3 := assign("lync")
	run,iexplore.exe https://sfp.wdf.sap.corp/sap(bD1lbiZjPTAwMSZkPW1pbg==)/bc/bsp/sap/crm_ui_start/default.htm,,,num4		; window: x: 321	y: 0	w: 2088	h: 1459
;	run,iexplore.exe https://sfp.wdf.sap.corp/sap(bD1lbiZjPTAwMSZkPW1pbg==)/bc/bsp/sap/crm_ui_start/default.htm,,,num5
;	run,iexplore.exe https://icp.wdf.sap.corp/sap(bD1lbiZjPTAwMSZkPW1pbg==)/bc/bsp/sap/crm_ui_start/default.htm,,,num6		; window: x: 494	y: 0	w: 1915	h: 1459
	winmove,ahk_id %num1%,,473,0,1398,1160
;	winmove,ahk_id %num2%,,792,0,1030,1160
	gui,wait:cancel
	msgbox % "Let's check some IDs`n`nOutlook:`t" num1 "`nOnenote:`t" num2 "`nLync:`t" num3 "`nSSF:`t" num4