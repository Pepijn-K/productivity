textx := (screenwidth / 2) - 250
texty := screenheight * 0.4

gui,wait:new
gui,wait:-caption +alwaysontop

gui,font,s30 bold
gui,add,text,x%textx% y%texty% w500 +center,RUNNING YOUR STUFF
gui,add,text,x%textx% y+24 w500 +center,please wait