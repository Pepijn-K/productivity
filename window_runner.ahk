; ############ AUTO-EXECUTE ############

#noenv
#notrayicon
#singleinstance,force
setworkingdir,%a_scriptdir%
sendmode,input
coordmode,mouse,client
coordmode,pixel,client

msgbox % a_screenwidth

; #### VARs ####
; if(a_screenwidth = 
if(!scaling)
	{
	inputbox,scaling,Scaling percentage,Please go to "Change display settings" option in Windows to check the current display scaling setting for your main monitor. Enter the number without any other characters like the percentage sign.
	iniwrite,%scaling%,lib/db.ini,settings,scaling
	}
dim := 100 / scaling
screenheight := a_screenheight * dim
screenwidth := a_screenwidth * dim

msgbox % sum(4,32),Working?,Run SAP related programs?,300
ifmsgbox,yes
	gosub,runsap

runsap:
	gui,wait:show,x0 y0 w%screenwidth% h%screenheight%,Waiting for stuff to load
	winset,transparent,240,Waiting for stuff to load
	run,C:\Users\I349302\Documents\AHK\SSF_assistant\SSF_assistant.ahk
	run,outlook.exe			; window: x: 592	y: 0	w: 1748	h: 1450
	run,onenote.exe			; window: x: 990	y: 0	w: 1287	h: 1450
	run,lync.exe			; window ahk_class LyncTabFrameHostWindowClass: x: 722	y: 199	w: 1396	h: 1075
	num1 := assign("outlook")
	num2 := assign("onenote")
	num3 := assign("lync")
	run,iexplore.exe https://sfp.wdf.sap.corp/sap(bD1lbiZjPTAwMSZkPW1pbg==)/bc/bsp/sap/crm_ui_start/default.htm,,,num4		; window: x: 321	y: 0	w: 2088	h: 1459
	run,iexplore.exe https://sfp.wdf.sap.corp/sap(bD1lbiZjPTAwMSZkPW1pbg==)/bc/bsp/sap/crm_ui_start/default.htm,,,num5
	run,iexplore.exe https://icp.wdf.sap.corp/sap(bD1lbiZjPTAwMSZkPW1pbg==)/bc/bsp/sap/crm_ui_start/default.htm,,,num6		; window: x: 494	y: 0	w: 1915	h: 1459
	winmove,ahk_id %num1%,,473,0,1398,1160
	winmove,ahk_id %num2%,,792,0,1030,1160
	gui,wait:cancel
return

assign(app)
	{
	iniread,apptitle,lib/db.ini,running,%app%
	loop
		{
		if(a_index = 25)
			{
			msgbox % "Could not assign " app " to key."
			break
			}
		numpadkey := winexist(apptitle)
		if(numpadkey)
			{
			return numpadkey
			break
			}
		else
			{
			sleep,500
			continue
			}
		}
	}

assign_r(pid)
	{
	loop
		{
		if(a_index = 25)
			{
			msgbox % "Could not assign process ID " pid " to key."
			break
			}
		numpadkey := winexist("ahk_pid " pid)
;		msgbox % "Round " a_index "`nI see PID: " pid " and HWND " numpadkey
		if(numpadkey)
			{
			return numpadkey
			break
			}
		else
			{
			sleep,500
			continue
			}
		}
	}
