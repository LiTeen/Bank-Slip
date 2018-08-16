SetTitleMatchMode 2  
#SingleInstance, Force
;=====================GET DAY IN A MONTH====================
;InputBox, monthinput, Enter month number, Please enter the month to print in number., , 380, 150

;press / to auto enter day, press*to print

y= %A_YYYY%
m= %A_MM% 
day:=0

;MsgBox %y% y %m% m %day% d

k:=days_in_month(y,m)

days_in_month(year,month)
{
  month++	
  early=%year%%month%
  month++
  if month = 13
  {
    month = 1
    year++
  }
  later=%year%%month% 
  envsub, later, %early%, days
  return later
}	



NumpadDiv::

day++

if ( day <= k)
{

WinActivate, deposit_format.xlsx
Send, ^g	;chg day
Send, Z12
Send {enter}

Sendinput %day%
Send {enter}

return
}
else {	
	MsgBox, 4,, Printing done. Start again?
IfMsgBox Yes
    	day := 0
	return
IfMsgBox No   
	return
}


NumpadMult::
Send, ^p	;print
Send {Enter}
Sleep, 2000
Send {Up}
return

;NumpadSub::


;Date := a_now
;FormatTime, ndate, %date%, dd/MM
;Send %ndate%




;^q::

ifWinExist, deposit_format.xlsx
{
	WinActivate	
	return
}
else
{
	Run \\IT_TEAM-PC\IT_Team\Dennise\deposit slip\deposit_format.xlsx
	WinActivate, deposit_format.xlsx
	return
}