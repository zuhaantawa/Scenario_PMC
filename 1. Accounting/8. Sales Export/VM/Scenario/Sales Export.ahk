; This script was created using Pulover's Macro Creator
; www.macrocreator.com

#NoEnv
SetWorkingDir %A_ScriptDir%
CoordMode, Mouse, Window
SendMode Input
#SingleInstance Force
SetTitleMatchMode 2
#WinActivateForce
SetControlDelay 1
SetWinDelay 0
SetKeyDelay -1
SetMouseDelay -1
SetBatchLines -1


Macro1:
FormatTime, timeNow, , yyyy-MM-dd HH:mm:ss
Goto, Macro2
selesai:
FormatTime, timeEnd, , yyyy-MM-dd HH:mm:ss
FileAppend, ACCOUNTING`, Sales Export`, %timeNow%`, %timeEnd%`n, D:\PMC\Scenarios\1. Accounting\8. Sales Export\Log\sales.export.txt
Sleep, 500
MsgBox, 0, , Sales Export Selesai !
Return

Macro2:
Run, D:\PMC\Scenarios\1. Accounting\8. Sales Export\New folder\Support Files\1. WEEKLY REPORT .xls, D:\PMC\Scenarios\1. Accounting\8. Sales Export\New folder\Support Files
Sleep, 2000
IfWinActive, Microsoft Office Activation Wizard
{
    Sleep, 500
    Send, {Alt Down}{F4}{Alt Up}
    Sleep, 1
}
Sleep, 2000
WinMaximize, 1. WEEKLY REPORT .xls  [Protected View] - Excel
Sleep, 333
Sleep, 1000
Run, D:\PMC\Scenarios\1. Accounting\8. Sales Export\New folder\Support Files\2. Rekap Sales Export.xls, D:\PMC\Scenarios\1. Accounting\8. Sales Export\New folder\Support Files
WinMaximize, 2. Rekap Sales Export.xls  [Compatibility Mode] - Excel
Sleep, 333
Sleep, 1000
Run, D:\PMC\Scenarios\1. Accounting\8. Sales Export\New folder\Support Files\3. Unit Price.xlsx, D:\PMC\Scenarios\1. Accounting\8. Sales Export\New folder\Support Files
WinMaximize, 3. Unit Price.xlsx - Excel
Sleep, 333
Sleep, 1000
Run, D:\PMC\Scenarios\1. Accounting\8. Sales Export\New folder\Support Files\4. SALES REPORT - 1.xlsx, D:\PMC\Scenarios\1. Accounting\8. Sales Export\New folder\Support Files
WinMaximize, 4. SALES REPORT - 1.xlsx - Excel
Sleep, 333
Sleep, 1000
InputBox, inputDate, Date File, Input Date File
Sleep, 1
Sleep, 1000
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{PgDn}{Control Up}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 1000
Run, D:\PMC\Scenarios\1. Accounting\8. Sales Export\New folder\Support Files\5. SALES REPORT.xlsx, D:\PMC\Scenarios\1. Accounting\8. Sales Export\New folder\Support Files
WinMaximize, 5. SALES REPORT.xlsx - Excel
Sleep, 333
Sleep, 2000
Click, 66, 699 Right, 1
Sleep, 100
Sleep, 3000
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Home}{Enter}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 3000
Click, 1169, 324, 0
Sleep, 100
Sleep, 2000
Send, {Control Down}{PgUp 6}{Control Up}
Sleep, 100
Sleep, 300
Goto, week1
Return

week1:
WinActivate, 1. WEEKLY REPORT .xls  [Compatibility Mode] - Excel  ; input week 1 
Sleep, 333
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {B}{5}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}
Sleep, 100
Sleep, 300  ; input week 1 
Send, {Control Down}{c}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
dateETD := Clipboard  ; input week 1 
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {C}{5}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}
Sleep, 100
Sleep, 300  ; input week 1 
Send, {Control Down}{c}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
sONO := Clipboard  ; input week 1 
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {E}{5}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}
Sleep, 100
Sleep, 300  ; input week 1 
Send, {Control Down}{c}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
invoiceno := Clipboard  ; input week 1 
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {F}{5}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}
Sleep, 100
Sleep, 300  ; input week 1 
Send, {Control Down}{c}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
model := Clipboard  ; input week 1 
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {G}{5}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}
Sleep, 100
Sleep, 300  ; input week 1 
Send, {Control Down}{c}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
qty := Clipboard  ; input week 1 
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {H}{5}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}
Sleep, 100
Sleep, 300  ; input week 1 
Send, {Control Down}{c}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
uprice := Clipboard  ; input week 1 
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {I}{5}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}
Sleep, 100
Sleep, 300  ; input week 1 
Send, {Control Down}{c}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
amount := Clipboard  ; input week 1 
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {J}{5}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}
Sleep, 100
Sleep, 300  ; input week 1 
Send, {Control Down}{c}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
destination := Clipboard  ; input week 1 
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {K}{5}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}
Sleep, 100
Sleep, 300  ; input week 1 
Send, {Control Down}{c}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
accountee := Clipboard  ; input week 1 
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {L}{5}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}
Sleep, 100
Sleep, 300  ; input week 1 
Send, {Control Down}{c}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
pebno := Clipboard  ; input week 1 
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {M}{5}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}
Sleep, 100
Sleep, 300  ; input week 1 
Send, {Control Down}{c}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
pebdate := Clipboard  ; input week 1 
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {N}{5}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}
Sleep, 100
Sleep, 300  ; input week 1 
Send, {Control Down}{c}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
noaju := Clipboard  ; input week 1 
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {O}{5}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}
Sleep, 100
Sleep, 300  ; input week 1 
Send, {Control Down}{c}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
npe := Clipboard  ; input week 1 
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {A}{3}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
/*
Sleep, 300
*/
Send, {Control Down}{c}{Control Up}
Sleep, 100
Sleep, 300
week := Clipboard
Sleep, 300
WinActivate, 2. Rekap Sales Export.xls  [Compatibility Mode] - Excel  ; input week 1 
Sleep, 333
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {B}{6}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Clipboard := dateETD  ; input week 1 
Sleep, 300  ; input week 1 
Send, {Control Down}{v}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {C}{6}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Clipboard := model  ; input week 1 
Sleep, 300  ; input week 1 
Send, {Control Down}{v}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {D}{6}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Clipboard := sONO  ; input week 1 
Sleep, 300  ; input week 1 
Send, {Control Down}{v}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {E}{6}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Clipboard := invoiceno  ; input week 1 
Sleep, 300  ; input week 1 
Send, {Control Down}{v}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {F}{6}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Clipboard := qty  ; input week 1 
Sleep, 300  ; input week 1 
Send, {Control Down}{v}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {G}{6}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Clipboard := uprice  ; input week 1 
Sleep, 300  ; input week 1 
Send, {Control Down}{v}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {H}{6}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Clipboard := amount  ; input week 1 
Sleep, 300  ; input week 1 
Send, {Control Down}{v}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {L}{6}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Clipboard := destination  ; input week 1 
Sleep, 300  ; input week 1 
Send, {Control Down}{v}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {O}{6}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Clipboard := accountee  ; input week 1 
Sleep, 300  ; input week 1 
Send, {Control Down}{v}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {P}{6}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Clipboard := pebno  ; input week 1 
Sleep, 300  ; input week 1 
Send, {Control Down}{v}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Q}{6}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Clipboard := pebdate  ; input week 1 
Sleep, 300  ; input week 1 
Send, {Control Down}{v}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {R}{6}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Clipboard := noaju  ; input week 1 
Sleep, 300  ; input week 1 
Send, {Control Down}{v}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {S}{6}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
Clipboard := npe  ; input week 1 
Sleep, 300  ; input week 1 
Send, {Control Down}{v}{Control Up}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
Send, {Control Down}{Down}{Control Up}
Sleep, 100
Sleep, 300
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Left}{Left}{Left}{Left}{Left}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300
Clipboard := week
Sleep, 300
Send, {Control Down}{v}{Control Up}
Sleep, 100
Sleep, 300
Send, {Control Down}{c}{Control Up}
Sleep, 100
Sleep, 300
Send, {Shift Down}{Control Down}{Up}{Control Up}{Shift Up}
Sleep, 100
Sleep, 300
Send, {Shift Down}{Down}{Shift Up}
Sleep, 100
Sleep, 300
Send, {Control Down}{v}{Control Up}
Sleep, 100
Sleep, 300
Goto, week2
Return

week2:
WinActivate, 1. WEEKLY REPORT .xls  [Compatibility Mode] - Excel  ; input week 2
Sleep, 333
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {B}{5}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Control Down}{c}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 1 
dateETD := Clipboard  ; input week 2
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {C}{5}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Control Down}{c}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 1 
sONO := Clipboard  ; input week 2
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {E}{5}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Control Down}{c}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 1 
invoiceno := Clipboard  ; input week 2
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {F}{5}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Control Down}{c}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 1 
model := Clipboard  ; input week 2
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {G}{5}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Control Down}{c}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 1 
qty := Clipboard  ; input week 2
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {H}{5}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Control Down}{c}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 1 
uprice := Clipboard  ; input week 2
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {I}{5}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Control Down}{c}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 1 
amount := Clipboard  ; input week 2
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {J}{5}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Control Down}{c}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 1 
destination := Clipboard  ; input week 2
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {K}{5}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Control Down}{c}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 1 
accountee := Clipboard  ; input week 2
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {L}{5}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Control Down}{c}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 1 
pebno := Clipboard  ; input week 2
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {M}{5}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Control Down}{c}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 1 
pebdate := Clipboard  ; input week 2
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {N}{5}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Control Down}{c}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 1 
noaju := Clipboard  ; input week 2
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {O}{5}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Shift Down}{Down}{Shift Up}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Control Down}{c}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 1 
npe := Clipboard  ; input week 2
Sleep, 300  ; input week 2
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {A}{3}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Control Up}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300
Send, {Control Down}{c}{Control Up}
Sleep, 100
Sleep, 300
week := Clipboard
Sleep, 300
WinActivate, 2. Rekap Sales Export.xls  [Compatibility Mode] - Excel  ; input week 2
Sleep, 333
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {B}{6}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Clipboard := dateETD  ; input week 2
Sleep, 300  ; input week 2
Send, {Control Down}{v}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {C}{6}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Clipboard := model  ; input week 2
Sleep, 300  ; input week 2
Send, {Control Down}{v}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {D}{6}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Clipboard := sONO  ; input week 2
Sleep, 300  ; input week 2
Send, {Control Down}{v}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {E}{6}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Clipboard := invoiceno  ; input week 2
Sleep, 300  ; input week 2
Send, {Control Down}{v}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {F}{6}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Clipboard := qty  ; input week 2
Sleep, 300  ; input week 2
Send, {Control Down}{v}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {G}{6}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Clipboard := uprice  ; input week 2
Sleep, 300  ; input week 2
Send, {Control Down}{v}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {H}{6}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Clipboard := amount  ; input week 2
Sleep, 300  ; input week 2
Send, {Control Down}{v}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {L}{6}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Clipboard := destination  ; input week 2
Sleep, 300  ; input week 2
Send, {Control Down}{v}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {O}{6}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Clipboard := accountee  ; input week 2
Sleep, 300  ; input week 2
Send, {Control Down}{v}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {P}{6}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Clipboard := pebno  ; input week 2
Sleep, 300  ; input week 2
Send, {Control Down}{v}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Q}{6}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Clipboard := pebdate  ; input week 2
Sleep, 300  ; input week 2
Send, {Control Down}{v}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {R}{6}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Clipboard := noaju  ; input week 2
Sleep, 300  ; input week 2
Send, {Control Down}{v}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {F5}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {S}{6}{Enter}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 2
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 2
Send, {Down}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Clipboard := npe  ; input week 2
Sleep, 300  ; input week 2
Send, {Control Down}{v}{Control Up}  ; input week 2
Sleep, 100
Sleep, 300  ; input week 2
Send, {Control Down}{Down}{Control Up}
Sleep, 100
Sleep, 300
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
Loop, 5
{
    SendEvent, {Left}
}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300
Clipboard := week
Sleep, 300
Send, {Control Down}{v}{Control Up}
Sleep, 100
Sleep, 300
Send, {Control Down}{c}{Control Up}
Sleep, 100
Sleep, 300
Send, {Shift Down}{Control Down}{Up}{Control Up}{Shift Up}
Sleep, 100
Sleep, 300
Send, {Shift Down}{Down}{Shift Up}
Sleep, 100
Sleep, 300
Send, {Control Down}{v}{Control Up}
Sleep, 100
Sleep, 300
Goto, week3
Return

week3:
WinActivate, 1. WEEKLY REPORT .xls  [Compatibility Mode] - Excel  ; input week 3
Sleep, 333
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {B}{5}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Control Down}{c}{Control Up}  ; input week 3
Sleep, 300
dateETD := Clipboard  ; input week 3
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {C}{5}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Control Down}{c}{Control Up}  ; input week 3
Sleep, 300
sONO := Clipboard  ; input week 3
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {E}{5}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Control Down}{c}{Control Up}  ; input week 3
Sleep, 300
invoiceno := Clipboard  ; input week 3
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {F}{5}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Control Down}{c}{Control Up}  ; input week 3
Sleep, 300
model := Clipboard  ; input week 3
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {G}{5}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Control Down}{c}{Control Up}  ; input week 3
Sleep, 300
qty := Clipboard  ; input week 3
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {H}{5}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Control Down}{c}{Control Up}  ; input week 3
Sleep, 300
uprice := Clipboard  ; input week 3
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {I}{5}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Control Down}{c}{Control Up}  ; input week 3
Sleep, 300
amount := Clipboard  ; input week 3
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {J}{5}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Control Down}{c}{Control Up}  ; input week 3
Sleep, 300
destination := Clipboard  ; input week 3
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {K}{5}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Control Down}{c}{Control Up}  ; input week 3
Sleep, 300
accountee := Clipboard  ; input week 3
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {L}{5}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Control Down}{c}{Control Up}  ; input week 3
Sleep, 300
pebno := Clipboard  ; input week 3
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {M}{5}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Control Down}{c}{Control Up}  ; input week 3
Sleep, 300
Sleep, 200
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {N}{5}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Control Down}{c}{Control Up}  ; input week 3
Sleep, 300
noaju := Clipboard  ; input week 3
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {O}{5}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Shift Down}{Down}{Shift Up}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Control Down}{c}{Control Up}  ; input week 3
Sleep, 300
Sleep, 200
npe := Clipboard  ; input week 3
Sleep, 200  ; input week 3
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 200  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {A}{3}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Control Up}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200
Send, {Control Down}{c}{Control Up}
Sleep, 300
Sleep, 200
week := Clipboard
Sleep, 200
WinActivate, 2. Rekap Sales Export.xls  [Compatibility Mode] - Excel  ; input week 3
Sleep, 333
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {B}{6}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Clipboard := dateETD  ; input week 3
Sleep, 200  ; input week 3
Send, {Control Down}{v}{Control Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {C}{6}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Clipboard := model  ; input week 3
Sleep, 200  ; input week 3
Send, {Control Down}{v}{Control Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {D}{6}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Clipboard := sONO  ; input week 3
Sleep, 200  ; input week 3
Send, {Control Down}{v}{Control Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {E}{6}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Clipboard := invoiceno  ; input week 3
Sleep, 200  ; input week 3
Send, {Control Down}{v}{Control Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {F}{6}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Clipboard := qty  ; input week 3
Sleep, 200  ; input week 3
Send, {Control Down}{v}{Control Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {G}{6}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Clipboard := uprice  ; input week 3
Sleep, 200  ; input week 3
Send, {Control Down}{v}{Control Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {H}{6}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Clipboard := amount  ; input week 3
Sleep, 200  ; input week 3
Send, {Control Down}{v}{Control Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {L}{6}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Clipboard := destination  ; input week 3
Sleep, 200  ; input week 3
Send, {Control Down}{v}{Control Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {O}{6}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Clipboard := accountee  ; input week 3
Sleep, 200  ; input week 3
Send, {Control Down}{v}{Control Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {P}{6}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Clipboard := pebno  ; input week 3
Sleep, 200  ; input week 3
Send, {Control Down}{v}{Control Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Q}{6}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Clipboard := pebdate  ; input week 3
Sleep, 200  ; input week 3
Send, {Control Down}{v}{Control Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {R}{6}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Clipboard := noaju  ; input week 3
Sleep, 200  ; input week 3
Send, {Control Down}{v}{Control Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {F5}  ; input week 3
Sleep, 100
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {S}{6}{Enter}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Control Down}{Down}{Control Up}  ; input week 3
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200  ; input week 3
Send, {Down}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Clipboard := npe  ; input week 3
Sleep, 200  ; input week 3
Send, {Control Down}{v}{Control Up}  ; input week 3
Sleep, 300
Sleep, 200  ; input week 3
Send, {Control Down}{Down}{Control Up}
Sleep, 300
Sleep, 200
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Left}{Left}{Left}{Left}{Left}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 200
Clipboard := week
Sleep, 200
Send, {Control Down}{v}{Control Up}
Sleep, 300
Sleep, 200
Send, {Control Down}{c}{Control Up}
Sleep, 300
Sleep, 200
Send, {Shift Down}{Control Down}{Up}{Control Up}{Shift Up}
Sleep, 300
Sleep, 200
Send, {Shift Down}{Down}{Shift Up}
Sleep, 300
Sleep, 200
Send, {Control Down}{v}{Control Up}
Sleep, 300
Sleep, 200
Goto, Macro6
Return

Macro6:
WinActivate, 1. WEEKLY REPORT .xls  [Compatibility Mode] - Excel  ; Input week 4 
Sleep, 333
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {B}{5}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Control Down}{c}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
dateETD := Clipboard  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {C}{5}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Control Down}{c}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
sONO := Clipboard  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {E}{5}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Control Down}{c}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
invoiceno := Clipboard  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {F}{5}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Control Down}{c}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
model := Clipboard  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {G}{5}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Control Down}{c}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
qty := Clipboard  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {H}{5}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Control Down}{c}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
uprice := Clipboard  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {I}{5}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Control Down}{c}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
amount := Clipboard  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {J}{5}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Control Down}{c}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
destination := Clipboard  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {K}{5}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Control Down}{c}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
accountee := Clipboard  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {L}{5}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Control Down}{c}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
pebno := Clipboard  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {M}{5}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Control Down}{c}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
pebdate := Clipboard  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {N}{5}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Control Down}{c}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
noaju := Clipboard  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {O}{5}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Shift Down}{Down}{Shift Up}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Control Down}{c}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
npe := Clipboard  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {A}{3}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300
Send, {Control Down}{c}{Control Up}
Sleep, 100
Sleep, 300
week := Clipboard
Sleep, 300
WinActivate, 2. Rekap Sales Export.xls  [Compatibility Mode] - Excel  ; Input week 4 
Sleep, 333
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {B}{6}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Clipboard := dateETD  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {Control Down}{v}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {C}{6}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Clipboard := model  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {Control Down}{v}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {D}{6}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Clipboard := sONO  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {Control Down}{v}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {E}{6}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Clipboard := invoiceno  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {Control Down}{v}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {F}{6}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Clipboard := qty  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {Control Down}{v}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {G}{6}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Clipboard := uprice  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {Control Down}{v}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {H}{6}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Clipboard := amount  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {Control Down}{v}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {L}{6}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Clipboard := destination  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {Control Down}{v}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {O}{6}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Clipboard := accountee  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {Control Down}{v}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {P}{6}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Clipboard := pebno  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {Control Down}{v}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Q}{6}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Clipboard := pebdate  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {Control Down}{v}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {R}{6}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Clipboard := noaju  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {Control Down}{v}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {F5}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {S}{6}{Enter}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; Input week 4 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input week 4 
Send, {Down}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Clipboard := npe  ; Input week 4 
Sleep, 300  ; Input week 4 
Send, {Control Down}{v}{Control Up}  ; Input week 4 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {Control Down}{Down}{Control Up}
Sleep, 100
Sleep, 300
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Left}{Left}{Left}{Left}{Left}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300
Clipboard := week
Sleep, 300
Send, {Control Down}{v}{Control Up}
Sleep, 100
Sleep, 300
Send, {Control Down}{c}{Control Up}
Sleep, 100
Sleep, 300
Send, {Shift Down}{Control Down}{Up}{Control Up}{Shift Up}
Sleep, 100
Sleep, 300
Send, {Shift Down}{Down}{Shift Up}
Sleep, 100
Sleep, 300
Send, {Control Down}{v}{Control Up}
Sleep, 100
Sleep, 300
Goto, Macro7
Return

Macro7:
WinActivate, 1. WEEKLY REPORT .xls  [Compatibility Mode] - Excel  ; input week 5 
Sleep, 333
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {B}{5}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Control Down}{c}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; Input week 4 
dateETD := Clipboard  ; input week 5 
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {C}{5}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Control Down}{c}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; Input week 4 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
sONO := Clipboard  ; input week 5 
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {E}{5}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Control Down}{c}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; Input week 4 
invoiceno := Clipboard  ; input week 5 
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {F}{5}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Control Down}{c}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; Input week 4 
model := Clipboard  ; input week 5 
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {G}{5}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Control Down}{c}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; Input week 4 
qty := Clipboard  ; input week 5 
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {H}{5}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Control Down}{c}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; Input week 4 
uprice := Clipboard  ; input week 5 
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {I}{5}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Control Down}{c}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; Input week 4 
amount := Clipboard  ; input week 5 
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {J}{5}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Control Down}{c}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; Input week 4 
destination := Clipboard  ; input week 5 
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {K}{5}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Control Down}{c}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; Input week 4 
accountee := Clipboard  ; input week 5 
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {L}{5}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Control Down}{c}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; Input week 4 
pebno := Clipboard  ; input week 5 
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {M}{5}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Control Down}{c}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; Input week 4 
pebdate := Clipboard  ; input week 5 
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {N}{5}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Control Down}{c}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; Input week 4 
noaju := Clipboard  ; input week 5 
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {O}{5}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Shift Down}{Down}{Shift Up}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Control Down}{c}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; Input week 4 
npe := Clipboard  ; input week 5 
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 1 
Sleep, 100
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {A}{3}{Enter}  ; input week 1 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 1 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300
Send, {Control Down}{c}{Control Up}
Sleep, 100
Sleep, 300
week := Clipboard
Sleep, 300
WinActivate, 2. Rekap Sales Export.xls  [Compatibility Mode] - Excel  ; input week 5 
Sleep, 333
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {B}{6}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Clipboard := dateETD  ; input week 5 
Sleep, 300  ; input week 5 
Send, {Control Down}{v}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {C}{6}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Clipboard := model  ; input week 5 
Sleep, 300  ; input week 5 
Send, {Control Down}{v}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {D}{6}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Clipboard := sONO  ; input week 5 
Sleep, 300  ; input week 5 
Send, {Control Down}{v}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {E}{6}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Clipboard := invoiceno  ; input week 5 
Sleep, 300  ; input week 5 
Send, {Control Down}{v}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {F}{6}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Clipboard := qty  ; input week 5 
Sleep, 300  ; input week 5 
Send, {Control Down}{v}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {G}{6}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Clipboard := uprice  ; input week 5 
Sleep, 300  ; input week 5 
Send, {Control Down}{v}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {H}{6}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Clipboard := amount  ; input week 5 
Sleep, 300  ; input week 5 
Send, {Control Down}{v}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {L}{6}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Clipboard := destination  ; input week 5 
Sleep, 300  ; input week 5 
Send, {Control Down}{v}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {O}{6}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Clipboard := accountee  ; input week 5 
Sleep, 300  ; input week 5 
Send, {Control Down}{v}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {P}{6}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Clipboard := pebno  ; input week 5 
Sleep, 300  ; input week 5 
Send, {Control Down}{v}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Q}{6}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Clipboard := pebdate  ; input week 5 
Sleep, 300  ; input week 5 
Send, {Control Down}{v}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {R}{6}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Clipboard := noaju  ; input week 5 
Sleep, 300  ; input week 5 
Send, {Control Down}{v}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {F5}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {S}{6}{Enter}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Control Up}  ; input week 5 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; input week 5 
Send, {Down}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Clipboard := npe  ; input week 5 
Sleep, 300  ; input week 5 
Send, {Control Down}{v}{Control Up}  ; input week 5 
Sleep, 100
Sleep, 300  ; input week 5 
Send, {Control Down}{Down}{Control Up}
Sleep, 100
Sleep, 300
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Left}{Left}{Left}{Left}{Left}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300
Clipboard := week
Sleep, 300
Send, {Control Down}{v}{Control Up}
Sleep, 100
Sleep, 300
Send, {Control Down}{c}{Control Up}
Sleep, 100
Sleep, 300
Send, {Shift Down}{Control Down}{Up}{Control Up}{Shift Up}
Sleep, 100
Sleep, 300
Send, {Shift Down}{Down}{Shift Up}
Sleep, 100
Sleep, 300
Send, {Control Down}{v}{Control Up}
Sleep, 100
Sleep, 300
Goto, Macro8
Return

Macro8:
WinActivate, 1. WEEKLY REPORT .xls  [Compatibility Mode] - Excel  ; Copy paste jumlah qty dan amount
Sleep, 333
Sleep, 300  ; Copy paste jumlah qty dan amount
Send, {F5}  ; Copy paste jumlah qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste jumlah qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {G}{5}{Enter}  ; Copy paste jumlah qty dan amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Copy paste jumlah qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; Copy paste jumlah qty dan amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Copy paste jumlah qty dan amount
Send, {Control Down}{c}{Control Up}  ; Copy paste jumlah qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste jumlah qty dan amount
totalqty := Clipboard  ; Copy paste jumlah qty dan amount
Sleep, 300  ; Input week 4 
Send, {F5}  ; Copy paste jumlah qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste jumlah qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {I}{5}{Enter}  ; Copy paste jumlah qty dan amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Copy paste jumlah qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Down}{Control Up}  ; Copy paste jumlah qty dan amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Copy paste jumlah qty dan amount
Send, {Control Down}{c}{Control Up}  ; Copy paste jumlah qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste jumlah qty dan amount
totalamount := Clipboard  ; Copy paste jumlah qty dan amount
WinActivate, 2. Rekap Sales Export.xls  [Compatibility Mode] - Excel  ; Copy paste jumlah qty dan amount
Sleep, 333
Sleep, 300  ; Copy paste jumlah qty dan amount
Send, {F5}  ; Copy paste jumlah qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste jumlah qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {F}{6}{Enter}  ; Copy paste jumlah qty dan amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Copy paste jumlah qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Control Up}  ; Copy paste jumlah qty dan amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Copy paste jumlah qty dan amount
Send, {Down}  ; Copy paste jumlah qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste jumlah qty dan amount
Clipboard := totalqty  ; Copy paste jumlah qty dan amount
Sleep, 300  ; Input week 4 
Send, {Control Down}{v}{Control Up}  ; Copy paste jumlah qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste jumlah qty dan amount
Send, {F5}  ; Copy paste jumlah qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste jumlah qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {H}{6}{Enter}  ; Copy paste jumlah qty dan amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Copy paste jumlah qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Down}{Down}{Control Up}  ; Copy paste jumlah qty dan amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Copy paste jumlah qty dan amount
Send, {Down}  ; Copy paste jumlah qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste jumlah qty dan amount
Clipboard := totalamount  ; Copy paste jumlah qty dan amount
Sleep, 300  ; Input week 4 
Send, {Control Down}{v}{Control Up}  ; Copy paste jumlah qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste jumlah qty dan amount
Send, {F5}  ; make pivot tabel
Sleep, 100
Sleep, 300  ; make pivot tabel
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {H}{5}{Enter}  ; make pivot tabel
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; make pivot tabel
Send, {Control Down}{Down}{Control Up}
Sleep, 100
Sleep, 300  ; make pivot tabel
Send, {Down}
Sleep, 100
Sleep, 300  ; make pivot tabel
Send, {Right}
Sleep, 100
Sleep, 300  ; make pivot tabel
Send, {Shift Down}{Control Down}{Right}{Control Up}{Shift Up}
Sleep, 100
Sleep, 300  ; make pivot tabel
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}
Sleep, 100
Sleep, 300  ; make pivot tabel
Send, {Delete}
Sleep, 100
Sleep, 300  ; make pivot tabel
Send, {F5}  ; make pivot tabel
Sleep, 100
Sleep, 300  ; make pivot tabel
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {B}{6}{Enter}  ; make pivot tabel
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; make pivot tabel
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Shift Down}{Down}{Right}{Right}{Right}{Right}{Shift Up}{Control Up}  ; make pivot tabel
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; make pivot tabel
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Alt}{N}{V}{Enter}  ; make pivot tabel
SetKeyDelay, %CurrentKeyDelay%
Sleep, 500  ; make pivot tabel
Send, {F6}
Sleep, 100
Sleep, 300
Send, {Down}  ; make pivot tabel
Sleep, 100
Sleep, 300  ; make pivot tabel
Send, {Space}  ; make pivot tabel
Sleep, 100
Sleep, 300  ; make pivot tabel
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down}{Down}{Down}  ; make pivot tabel
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; make pivot tabel
Send, {AppsKey}  ; make pivot tabel
Sleep, 100
Sleep, 300  ; make pivot tabel
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down}{Down}{Enter}  ; make pivot tabel
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; make pivot tabel
Send, {Down}  ; make pivot tabel
Sleep, 100
Sleep, 300  ; make pivot tabel
Send, {AppsKey}  ; make pivot tabel
Sleep, 100
Sleep, 300  ; make pivot tabel
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down}{Down}{Down}{Down}{Enter}  ; make pivot tabel
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; make pivot tabel
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down}{Down}{Down}{Down}  ; make pivot tabel
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; make pivot tabel
Send, {AppsKey}  ; make pivot tabel
Sleep, 100
Sleep, 300  ; make pivot tabel
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down}{Down}{Down}{Down}{Enter}  ; make pivot tabel
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; make pivot tabel
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down}{Down}{Down}{Down}  ; make pivot tabel
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; make pivot tabel
Send, {AppsKey}  ; make pivot tabel
Sleep, 100
Sleep, 300  ; make pivot tabel
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down}{Enter}  ; make pivot tabel
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; make pivot tabel
Send, {Down}  ; make pivot tabel
Sleep, 100
Sleep, 300  ; make pivot tabel
Send, {AppsKey}  ; make pivot tabel
Sleep, 100
Sleep, 300  ; make pivot tabel
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down}{Enter}  ; make pivot tabel
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; make pivot tabel
Send, {Control Down}{f}{Control Up}
Sleep, 100
Sleep, 300  ; make pivot tabel
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {A}{A}{M}{O}{U}{N}{T}{Enter}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; make pivot tabel
Send, {Escape}
Sleep, 100
Sleep, 300
Send, {F5}  ; make pivot tabel
Sleep, 100
Sleep, 300  ; make pivot tabel
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {B}{2}{Enter}  ; make pivot tabel
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; make pivot tabel
Send, {Alt Down}{Down}{Alt Up}  ; make pivot tabel
Sleep, 100
Sleep, 300  ; make pivot tabel
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down}{Down}{Down}{Enter}  ; make pivot tabel
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; make pivot tabel
Click, 662, 422 Left, 1
Sleep, 10
Sleep, 300
Send, {F5}  ; copy SKE Kuning
Sleep, 100
Sleep, 300  ; copy SKE Kuning
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {A}{5}{Enter}  ; copy SKE Kuning
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; copy SKE Kuning
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Alt}{J}{T}{G}  ; copy SKE Kuning
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; copy SKE Kuning
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Alt Down}{E}{Alt Up}{Enter}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Right}{Down}  ; copy SKE Kuning
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; copy SKE Kuning
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Shift Down}{Right}{Down}{Shift Up}{Control Up}  ; copy SKE Kuning
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; copy SKE Kuning
Send, {Control Down}{c}{Control Up}  ; copy SKE Kuning
Sleep, 100
Sleep, 300  ; copy SKE Kuning
pivotTabel := Clipboard  ; copy SKE Kuning
Sleep, 300  ; copy SKE Kuning
Goto, Macro9
Return

Macro9:
WinActivate, 4. SALES REPORT - 1.xlsx - Excel  ; Input tabel kuning 
Sleep, 333
Sleep, 300  ; Input tabel kuning 
Send, {Control Down}{PgUp 3}{Control Up}
Sleep, 100
Sleep, 300  ; Input tabel kuning 
Send, {Control Down}{PgDn}{Control Up}
Sleep, 100
Sleep, 300  ; Input tabel kuning 
Send, {F5}  ; Input tabel kuning 
Sleep, 100
Sleep, 300  ; Input tabel kuning 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {C}{1}{5}{Enter}  ; Input tabel kuning 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input tabel kuning 
Clipboard := pivotTabel  ; Input tabel kuning 
Sleep, 300  ; Input tabel kuning 
Send, {Control Down}{v}{Control Up}  ; Input tabel kuning 
Sleep, 100
Sleep, 300
Send, {F5}  ; Input tabel kuning 
Sleep, 100
Sleep, 300  ; Input tabel kuning 
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {E}{4}{Enter}  ; Input tabel kuning 
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input tabel kuning 
Clipboard := inputDate  ; Input tabel kuning 
Sleep, 300  ; Input tabel kuning 
Send, {Control Down}{v}{Control Up}  ; Input tabel kuning 
Sleep, 100
Sleep, 300  ; Input tabel kuning 
WinActivate, 2. Rekap Sales Export.xls  [Compatibility Mode] - Excel  ; Input tabel JKC biru
Sleep, 333
Sleep, 300  ; Input tabel JKC biru
Send, {F5}  ; Input tabel JKC biru
Sleep, 100
Sleep, 300  ; Input tabel JKC biru
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {B}{2}{Enter}  ; Input tabel JKC biru
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input tabel JKC biru
Send, {Alt Down}{Down}{Alt Up}  ; Input tabel JKC biru
Sleep, 100
Sleep, 300  ; Input tabel JKC biru
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down}{Up}{Enter}  ; Input tabel JKC biru
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input tabel JKC biru
Goto, mulai  ; Input tabel JKC biru
Sleep, 300  ; Input tabel JKC biru
loopingMulai:  ; Input tabel JKC biru
Sleep, 300  ; Input tabel JKC biru
jumlahPerulangan := 0  ; Input tabel JKC biru
mulai:  ; Input tabel JKC biru
Sleep, 300  ; Input tabel JKC biru
If (jumlahPerulangan = 5)  ; Input tabel JKC biru
{
    Goto, loopingSelesai  ; Input tabel JKC biru
}
jumlahPerulangan += 1  ; Input tabel JKC biru
WinActivate, 2. Rekap Sales Export.xls  [Compatibility Mode] - Excel  ; Input tabel JKC biru
Sleep, 333
Sleep, 300  ; Input tabel JKC biru
Send, {F5}  ; Input tabel JKC biru
Sleep, 100
Sleep, 300  ; Input tabel JKC biru
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {B}{1}{Enter}  ; Input tabel JKC biru
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input tabel JKC biru
Send, {Alt Down}{Down}{Alt Up}  ; Input tabel JKC biru
Sleep, 100
Sleep, 300  ; Input tabel JKC biru
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down}{Down}{Enter}  ; Input tabel JKC biru
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input tabel JKC biru
Send, {F5}  ; Input tabel JKC biru
Sleep, 100
Sleep, 300  ; Input tabel JKC biru
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {A}{5}{Enter}  ; Input tabel JKC biru
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input tabel JKC biru
Send, {Down}  ; Input tabel JKC biru
Sleep, 100
Sleep, 300
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {F2}{Shift Down}{Home}{Shift Up}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300
Send, {Control Down}{c}{Control Up}
Sleep, 100
Sleep, 300
Send, {Escape}
Sleep, 100
Sleep, 300
hastag := Clipboard  ; Input tabel kuning 
Sleep, 300
/*
MsgBox, 0, , hastag = %hastag%
*/
If hastag contains Grand Total
{
    Goto, monthlyReport
}
Send, {Right 3}
Sleep, 100
Sleep, 300  ; Input tabel JKC biru
Send, {Shift Down}{Control Down}{Down}{Control Up}{Shift Up}
Sleep, 100
Sleep, 300  ; Input tabel JKC biru
Send, {Shift Down}{Left 2}{Shift Up}
Sleep, 100
Sleep, 300  ; Input tabel JKC biru
Send, {Shift Down}{Up}{Shift Up}
Sleep, 100
Sleep, 300
Send, {Control Down}{c}{Control Up}  ; Input tabel JKC biru
Sleep, 100
Sleep, 300  ; Input tabel JKC biru
pivot1 := Clipboard  ; Input tabel JKC biru
Sleep, 300  ; Input tabel JKC biru
WinActivate, 5. SALES REPORT.xlsx - Excel  ; Input tabel JKC biru
Sleep, 333
Sleep, 300  ; Input tabel JKC biru
Send, {F5}  ; Input tabel JKC biru
Sleep, 100
Sleep, 300  ; Input tabel JKC biru
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {D}{1}{3}{Enter}  ; Input tabel JKC biru
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Input tabel JKC biru
Clipboard := pivot1  ; Input tabel JKC biru
Sleep, 300  ; Input tabel JKC biru
Send, {Control Down}{v}{Control Up}  ; Input tabel JKC biru
Sleep, 100
Sleep, 300  ; Input tabel JKC biru
Send, {Control Down}{PgDn}{Control Up}
Sleep, 100
Sleep, 300
Goto, loopingMulai  ; Input tabel JKC biru
loopingSelesai:  ; Input tabel JKC biru
Sleep, 300
monthlyReport:
Sleep, 300
WinActivate, 5. SALES REPORT.xlsx - Excel
Sleep, 333
Sleep, 300
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{PgDn}{PgDn}{PgDn}{PgDn}{PgDn}{Control Up}  ; Copy paste total qty dan amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{PgUp}{PgUp}{PgUp}{PgUp}{PgUp}{Control Up}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Copy paste total qty dan amount
Send, {F5}  ; Copy paste total qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {E}{1}{3}{Enter}  ; Copy paste total qty dan amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Copy paste total qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {F2}{Shift Down}{Home}{Shift Up}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300
Send, {Control Down}{c}{Control Up}
Sleep, 100
Sleep, 300
Send, {Escape}
Sleep, 100
Sleep, 300
mobileQty := Clipboard  ; Input tabel kuning 
Sleep, 300
/*
MsgBox, 0, , Mobile Quantity = %mobileQty%
*/
If (mobileQty =)
{
    Goto, nextPage
}
Sleep, 300  ; Copy paste total qty dan amount
Send, {Control Down}{f}{Control Up}
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {T}{O}{T}{A}{L}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Copy paste total qty dan amount
Send, {Enter}
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
Send, {Escape}
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
Send, {Right}
Sleep, 100
Sleep, 300
Send, {Control Down}{c}{Control Up}  ; Copy paste total qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
qty1 := Clipboard  ; Copy paste total qty dan amount
Sleep, 300  ; Copy paste total qty dan amount
Send, {Right}
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
Send, {Control Down}{c}{Control Up}  ; Copy paste total qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
amount1 := Clipboard  ; Copy paste total qty dan amount
Sleep, 300  ; Copy paste total qty dan amount
Send, {Control Down}{PgDn}{Control Up}  ; Copy paste total qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
Send, {F5}  ; Copy paste total qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {E}{1}{3}{Enter}  ; Copy paste total qty dan amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Copy paste total qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {F2}{Shift Down}{Home}{Shift Up}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300
Send, {Control Down}{c}{Control Up}
Sleep, 100
Sleep, 300
Send, {Escape}
Sleep, 100
Sleep, 300
mobileQty := Clipboard  ; Input tabel kuning 
Sleep, 300
/*
MsgBox, 0, , Mobile Quantity = %mobileQty%
*/
If (mobileQty =)
{
    Goto, nextPage
}
Sleep, 300  ; Copy paste total qty dan amount
Send, {Control Down}{f}{Control Up}
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {T}{O}{T}{A}{L}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Copy paste total qty dan amount
Send, {Enter}
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
Send, {Escape}
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
Send, {Right}
Sleep, 100
Sleep, 300
Send, {Control Down}{c}{Control Up}  ; Copy paste total qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
qty2 := Clipboard  ; Copy paste total qty dan amount
Sleep, 300  ; Copy paste total qty dan amount
Send, {Right}
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
Send, {Control Down}{c}{Control Up}  ; Copy paste total qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
amount2 := Clipboard  ; Copy paste total qty dan amount
Sleep, 300  ; Copy paste total qty dan amount
Send, {Control Down}{PgDn}{Control Up}  ; Copy paste total qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
Send, {F5}  ; Copy paste total qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {E}{1}{3}{Enter}  ; Copy paste total qty dan amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Copy paste total qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {F2}{Shift Down}{Home}{Shift Up}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300
Send, {Control Down}{c}{Control Up}
Sleep, 100
Sleep, 300
Send, {Escape}
Sleep, 100
Sleep, 300
mobileQty := Clipboard  ; Input tabel kuning 
Sleep, 300
/*
MsgBox, 0, , Mobile Quantity = %mobileQty%
*/
If (mobileQty =)
{
    Goto, nextPage
}
Sleep, 300  ; Copy paste total qty dan amount
Send, {Control Down}{f}{Control Up}
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {T}{O}{T}{A}{L}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Copy paste total qty dan amount
Send, {Enter}
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
Send, {Escape}
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
Send, {Right}
Sleep, 100
Sleep, 300
Send, {Control Down}{c}{Control Up}  ; Copy paste total qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
qty3 := Clipboard  ; Copy paste total qty dan amount
Sleep, 300  ; Copy paste total qty dan amount
Send, {Right}
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
Send, {Control Down}{c}{Control Up}  ; Copy paste total qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
amount3 := Clipboard  ; Copy paste total qty dan amount
Sleep, 300  ; Copy paste total qty dan amount
Send, {Control Down}{PgDn}{Control Up}  ; Copy paste total qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
Send, {F5}  ; Copy paste total qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {E}{1}{3}{Enter}  ; Copy paste total qty dan amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Copy paste total qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {F2}{Shift Down}{Home}{Shift Up}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300
Send, {Control Down}{c}{Control Up}
Sleep, 100
Sleep, 300
Send, {Escape}
Sleep, 100
Sleep, 300
mobileQty := Clipboard  ; Input tabel kuning 
Sleep, 300
/*
MsgBox, 0, , Mobile Quantity = %mobileQty%
*/
If (mobileQty =)
{
    Goto, nextPage
}
Sleep, 300  ; Copy paste total qty dan amount
Send, {Control Down}{f}{Control Up}
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {T}{O}{T}{A}{L}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Copy paste total qty dan amount
Send, {Enter}
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
Send, {Escape}
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
Send, {Right}
Sleep, 100
Sleep, 300
Send, {Control Down}{c}{Control Up}  ; Copy paste total qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
qty4 := Clipboard  ; Copy paste total qty dan amount
Sleep, 300  ; Copy paste total qty dan amount
Send, {Right}
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
amount4 := Clipboard  ; Copy paste total qty dan amount
Sleep, 300  ; Copy paste total qty dan amount
Send, {Control Down}{PgDn}{Control Up}  ; Copy paste total qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
Send, {F5}  ; Copy paste total qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {E}{1}{3}{Enter}  ; Copy paste total qty dan amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Copy paste total qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {F2}{Shift Down}{Home}{Shift Up}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300
Send, {Control Down}{c}{Control Up}
Sleep, 100
Sleep, 300
Send, {Escape}
Sleep, 100
Sleep, 300
mobileQty := Clipboard  ; Input tabel kuning 
Sleep, 300
/*
MsgBox, 0, , Mobile Quantity = %mobileQty%
*/
If (mobileQty =)
{
    Goto, nextPage
}
Sleep, 300
Send, {Control Down}{f}{Control Up}
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {T}{O}{T}{A}{L}
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Copy paste total qty dan amount
Send, {Enter}
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
Send, {Escape}
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
Send, {Right}
Sleep, 100
Sleep, 300
Sleep, 300  ; Copy paste total qty dan amount
Send, {Control Down}{c}{Control Up}  ; Copy paste total qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
qty5 := Clipboard  ; Copy paste total qty dan amount
Sleep, 300  ; Copy paste total qty dan amount
Send, {Right}
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
Send, {Control Down}{c}{Control Up}  ; Copy paste total qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
amount5 := Clipboard  ; Copy paste total qty dan amount
Sleep, 300
nextPage:
Sleep, 300
Send, {Control Down}{PgDn}{Control Up}
Sleep, 100
Sleep, 300
Send, {F5}  ; Copy paste total qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {C}{2}{3}{Enter}  ; Copy paste total qty dan amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Copy paste total qty dan amount
Clipboard := qty1  ; Copy paste total qty dan amount
Sleep, 300
Send, {Control Down}{v}{Control Up}
Sleep, 100
Sleep, 300
Send, {Enter}
Sleep, 100
Sleep, 300
Clipboard := qty2  ; Copy paste total qty dan amount
Sleep, 300
Send, {Control Down}{v}{Control Up}
Sleep, 100
Sleep, 300
Send, {Enter}
Sleep, 100
Sleep, 300
Clipboard := qty3  ; Copy paste total qty dan amount
Sleep, 300
Send, {Control Down}{v}{Control Up}
Sleep, 100
Sleep, 300
Send, {Enter}
Sleep, 100
Sleep, 300
Clipboard := qty4  ; Copy paste total qty dan amount
Sleep, 300
Send, {Control Down}{v}{Control Up}
Sleep, 100
Sleep, 300
Send, {Enter}
Sleep, 100
Sleep, 300
Clipboard := qty5  ; Copy paste total qty dan amount
Sleep, 300
Send, {Control Down}{v}{Control Up}
Sleep, 100
Sleep, 300
Send, {F5}  ; Copy paste total qty dan amount
Sleep, 100
Sleep, 300  ; Copy paste total qty dan amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {D}{2}{3}{Enter}  ; Copy paste total qty dan amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Copy paste total qty dan amount
Clipboard := amount1  ; Copy paste total qty dan amount
Sleep, 300
Send, {Control Down}{v}{Control Up}
Sleep, 100
Sleep, 300
Send, {Enter}
Sleep, 100
Sleep, 300
Clipboard := amount2  ; Copy paste total qty dan amount
Sleep, 300
Send, {Control Down}{v}{Control Up}
Sleep, 100
Sleep, 300
Send, {Enter}
Sleep, 100
Sleep, 300
Clipboard := amount3  ; Copy paste total qty dan amount
Sleep, 300
Send, {Control Down}{v}{Control Up}
Sleep, 100
Sleep, 300
Send, {Enter}
Sleep, 100
Sleep, 300
Clipboard := amount4  ; Copy paste total qty dan amount
Sleep, 300
Send, {Control Down}{v}{Control Up}
Sleep, 100
Sleep, 300
Send, {Enter}
Sleep, 100
Sleep, 300
Clipboard := amount5  ; Copy paste total qty dan amount
Sleep, 300
Send, {Control Down}{v}{Control Up}
Sleep, 100
Sleep, 300
Goto, selesai
Return

