SetTitleMatchMode 2  	

day:=0

^q::

ifWinExist, deposit_format.xlsx
{
	WinActivate	
	return
}
else
{
	Run GET YOUR PATH\deposit_format.xlsx
	WinActivate, deposit_format.xlsx
	return
}


NumpadDiv::

day++

WinActivate, deposit_format.xlsx
Send, ^g	;chg day
Send, Z12
Send {enter}

Sendinput %day%
Send {enter}

return


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

cnt = 0
FormatTime next_month, % A_YYYY+(A_MM+cnt)//12 . Mod(A_MM+cnt,12)+1, MM

Send, % LDOM(220418)

LDOM(TimeStr="") {
  If TimeStr=
     TimeStr = %next_month%
  StringLeft Date,TimeStr,6 ; YearMonth
  Date1 = %Date%
  Date1+= 31,D              ; A day in next month
  StringLeft Date1,Date1,6  ; YearNextmonth
  Date1-= %Date%,D          ; Difference in days
  Return Date1
}

return
