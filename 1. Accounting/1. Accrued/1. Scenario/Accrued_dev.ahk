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
FormatTime, timeStart, , yyyy-MM-dd HH:mm:ss
Sleep, 100
/*
IfExist, D:\PMC\Log\Accrued.txt  ; Write Log
{
    FileAppend, Start`, %timeStart%
}
FileAppend, `nS %timeStart%, D:\PMC\Log\Accrued.txt  ; Write Log
Sleep, 100
*/
OpenFiles()
IsiAccruedEmpty()
FillLedgerImport()
FillAccrued()
FillExcRate()
/*
IfExist, D:\PMC\Log\Accrued.txt  ; Write Log
{
    */
    FileAppend, Accounting`,Accrued`,%timeStart%`,%timeEnd%`n, D:\PMC\Scenario\1. Accounting\1. Accrued\4. Log\Accrued_Log.txt
    /*
}
*/
MsgBox, 0, , Proses Selesai :)
Return

OpenFiles()
{
    Run, D:\PMC\Scenario\1. Accounting\1. Accrued\2. Support Files\1. Cashbill.xlsx  ; Cashbill Open
    Sleep, 1000
    Sleep, 3000  ; Cashbill Open
    WinWaitActive, Microsoft Office Activation Wizard, , 20
    Sleep, 333
    Sleep, 300
    WinClose, Microsoft Office Activation Wizard
    Sleep, 333
    Sleep, 300
    WinMaximize, 1. Cashbill.xlsx - Excel  ; Cashbill Open
    Sleep, 333
    Sleep, 1000  ; Cashbill Open
    Run, D:\PMC\Scenario\1. Accounting\1. Accrued\2. Support Files\2. List Accrued EMPTY.xls  ; Accrued Empty Open
    Sleep, 1000
    Sleep, 3000  ; Accrued Empty Open
    WinMaximize, 2. List Accrued EMPTY.xls  [Compatibility Mode] - Excel  ; Accrued Empty Open
    Sleep, 333
    Sleep, 1000  ; Accrued Empty Open
    Run, D:\PMC\Scenario\1. Accounting\1. Accrued\2. Support Files\3. LEDGER IMPORT.xls  ; Ledger Import Open
    Sleep, 1000
    Sleep, 3000  ; Ledger Import Open
    WinMaximize, 3. LEDGER IMPORT.xls  [Compatibility Mode] - Excel  ; Ledger Import Open
    Sleep, 333
    Sleep, 1000  ; Ledger Import Open
    Run, D:\PMC\Scenario\1. Accounting\1. Accrued\2. Support Files\4. ACCRUED.xlsx  ; Accrued Open
    Sleep, 1000
    Sleep, 3000  ; Accrued Open
    WinMaximize, 4. ACCRUED.xlsx - Excel  ; Accrued Open
    Sleep, 333
    Sleep, 1000  ; Accrued Open
    Run, D:\PMC\Scenario\1. Accounting\1. Accrued\2. Support Files\5. EXC.RATE.xlsx  ; Exc Rate Open
    Sleep, 1000
    Sleep, 3000  ; Exc Rate Open
    WinMaximize, 5. EXC.RATE.xlsx - Excel  ; Exc Rate Open
    Sleep, 333
    Sleep, 1000  ; Exc Rate Open
    /*
    WinActivate, 1. Cashbill.xlsx - Excel  ; SGD_Invoice Date - Supp Name
    Sleep, 333
    Sleep, 300  ; SGD_Invoice Date - Supp Name
    Send, {Control Down}{Home}{Control Up}  ; SGD_Invoice Date - Supp Name
    Sleep, 100
    Sleep, 300  ; SGD_Invoice Date - Supp Name
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Right 2}{Control Down}{Down}{Control Up}{Shift Up}  ; SGD_Invoice Date - Supp Name
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Invoice Date - Supp Name
    Send, {Control Down}{c}{Control Up}  ; SGD_Invoice Date - Supp Name
    Sleep, 300
    Sleep, 300  ; SGD_Invoice Date - Supp Name
    WinActivate, 5. EXC.RATE.xlsx - Excel  ; SGD_Invoice Date - Supp Name
    Sleep, 333
    Sleep, 1000  ; SGD_Invoice Date - Supp Name
    Send, {F5}  ; SGD_Invoice Date - Supp Name
    Sleep, 100
    Sleep, 100  ; SGD_Invoice Date - Supp Name
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, B8{Enter}  ; SGD_Invoice Date - Supp Name
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Invoice Date - Supp Name
    Send, {Control Down}{Down}{Control Up}
    Sleep, 300  ; SGD_Invoice Date - Supp Name
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; SGD_Invoice Date - Supp Name
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 100  ; SGD_Invoice Date - Supp Name
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; USD_CC - Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    */
    MsgBox, 0, , end of open files, 1
    return return
}

IsiAccruedEmpty()
{
    WinActivate, 2. List Accrued EMPTY.xls  [Compatibility Mode] - Excel  ; Accrued Empty Open
    Sleep, 333
    Sleep, 1000  ; Accrued Empty Open
    InputBox, periode, Periode Accrued, Input Periode Accrued  ; Isi Period
    Sleep, 100
    /*
    MsgBox, 0, , Period : %periode%, 100  ; Isi Period
    */
    Sleep, 300
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {F5}  ; Isi Period
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {c}{7}{Enter}  ; Isi Period
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300
    Clipboard := periode  ; Isi Period
    Sleep, 300
    /*
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{v}{Control Up}  ; Isi Period
    SetKeyDelay, %CurrentKeyDelay%
    */
    WinActivate, 1. Cashbill.xlsx - Excel  ; Accrued Empty Open
    Sleep, 333
    Sleep, 300
    Send, {Control Down}{Home}{Control Up}  ; Get Invoice Date
    Sleep, 100
    Sleep, 300
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Control Down}{Shift Down}{Down}{Shift Up}{Control Up}  ; Get Invoice Date
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300
    Send, {Control Down}{c}{Control Up}  ; Get Invoice Date
    Sleep, 100
    Sleep, 300
    invoiceDate := Clipboard  ; Get Invoice Date
    Sleep, 300
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Up}{Right}{Down}  ; Get Invoice Date
    SetKeyDelay, %CurrentKeyDelay%
    /*
    MsgBox, 0, , invoices = %invoiceDate%, 100  ; Get Invoice Date
    */
    Sleep, 300
    Send, {Control Down}{Home}{Control Up}  ; Get Invoice Number
    Sleep, 100
    Sleep, 300
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Right}{Down}{Control Down}{Shift Down}{Down}{Shift Up}{Control Up}  ; Get Invoice Number
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300
    Send, {Control Down}{c}{Control Up}  ; Get Invoice Number
    Sleep, 100
    Sleep, 300
    invoiceNumber := Clipboard  ; Get Invoice Number
    Sleep, 300
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Up}{Right}{Down}  ; Get Invoice Number
    SetKeyDelay, %CurrentKeyDelay%
    /*
    MsgBox, 0, , invoices = %invoiceNumber%, 100  ; Get Invoice Number
    */
    Sleep, 300
    Send, {Control Down}{Home}{Control Up}  ; Get Slip Number
    Sleep, 100
    Sleep, 300
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Right 2}{Down}{Control Down}{Shift Down}{Down}{Shift Up}{Control Up}  ; Get Slip Number
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300
    Send, {Control Down}{c}{Control Up}  ; Get Slip Number
    Sleep, 100
    Sleep, 300
    slipNumber := Clipboard  ; Get Slip Number
    Sleep, 300
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Up}{Right}{Down}  ; Get Slip Number
    SetKeyDelay, %CurrentKeyDelay%
    /*
    MsgBox, 0, , slip = %slipNumber%, 100  ; Get Slip Number
    */
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Shift Down}{Down}{Shift Up}{Control Up}  ; Get Supplier Name
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300
    Send, {Control Down}{c}{Control Up}  ; Get Supplier Name
    Sleep, 100
    Sleep, 300  ; Get Supplier Name
    supplierName := Clipboard  ; Get Supplier Name
    Sleep, 300  ; Get Supplier Name
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Up}{Right}{Down}  ; Get Supplier Name
    SetKeyDelay, %CurrentKeyDelay%
    /*
    MsgBox, 0, , supplierName = %supplierName%, 100  ; Get Supplier Name
    */
    Sleep, 300  ; Get Description
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Shift Down}{Down}{Shift Up}{Control Up}  ; Get Description
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Get Description
    Sleep, 300  ; Get Description
    Send, {Control Down}{c}{Control Up}  ; Get Description
    Sleep, 100
    Sleep, 300  ; Get Description
    description := Clipboard  ; Get Description
    Sleep, 300  ; Get Description
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Up}{Right}{Down}  ; Get Description
    SetKeyDelay, %CurrentKeyDelay%
    /*
    MsgBox, 0, , desc = %description%, 100  ; Get Description
    */
    Sleep, 300  ; Get Description
    Send, {F5}  ; Get CC
    Sleep, 100
    Sleep, 300
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, G2{Enter}  ; Get CC
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Shift Down}{Down}{Shift Up}{Control Up}  ; Get CC
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300
    Send, {Control Down}{c}{Control Up}  ; Get CC
    Sleep, 100
    Sleep, 300
    costCenter := Clipboard  ; Get CC
    Sleep, 300
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Up}{Right}{Down}  ; Get CC
    SetKeyDelay, %CurrentKeyDelay%
    /*
    MsgBox, 0, , cc = %costCenter%, 100  ; Get CC
    */
    Sleep, 300
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Shift Down}{Down}{Shift Up}{Control Up}  ; Get Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300
    Send, {Control Down}{c}{Control Up}  ; Get Debit Account
    Sleep, 100
    Sleep, 300
    debitAccount := Clipboard  ; Get Debit Account
    Sleep, 300
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Up}{Right}{Down}  ; Get Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    /*
    MsgBox, 0, , debitAccount = %debitAccount%, 100  ; Get Debit Account
    */
    Sleep, 300
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Shift Down}{Down}{Shift Up}{Control Up}  ; Get credit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300
    Send, {Control Down}{c}{Control Up}  ; Get credit Account
    Sleep, 100
    Sleep, 300
    creditAccount := Clipboard  ; Get credit Account
    Sleep, 300
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Up}{Right}{Down}  ; Get credit Account
    SetKeyDelay, %CurrentKeyDelay%
    /*
    MsgBox, 0, , creditAccount = %creditAccount%, 100  ; Get credit Account
    */
    Sleep, 300
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Shift Down}{Down}{Shift Up}{Control Up}  ; Get Currency
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300
    Send, {Control Down}{c}{Control Up}  ; Get Currency
    Sleep, 100
    Sleep, 300
    currency := Clipboard  ; Get Currency
    Sleep, 300
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Up}{Right}{Down}  ; Get Currency
    SetKeyDelay, %CurrentKeyDelay%
    /*
    MsgBox, 0, , currency = %currency%, 100  ; Get Currency
    */
    Sleep, 300
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Shift Down}{Down}{Shift Up}{Control Up}  ; Get Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300
    Send, {Control Down}{c}{Control Up}  ; Get Amount
    Sleep, 100
    Sleep, 300
    amount := Clipboard  ; Get Amount
    Sleep, 300
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Up}{Right}{Down}  ; Get Amount
    SetKeyDelay, %CurrentKeyDelay%
    /*
    MsgBox, 0, , amount = %amount%, 100  ; Get Amount
    */
    WinActivate, 2. List Accrued EMPTY.xls  [Compatibility Mode] - Excel  ; Accrued Empty Open
    Sleep, 333
    Clipboard := ""
    Send, {F5}  ; Set Invoice Date
    Sleep, 100
    Sleep, 300  ; Set Invoice Date
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, B10{Enter}  ; Set Invoice Date
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Set Invoice Date
    Clipboard := invoiceDate  ; Set Invoice Date
    Sleep, 300  ; Set Invoice Date
    Send, {Control Down}{v}{Control Up}  ; Set Invoice Date
    Sleep, 100
    Sleep, 300  ; Set Invoice Date
    Clipboard := ""
    Send, {F5}  ; Set Invoice Number
    Sleep, 100
    Sleep, 300  ; Set Invoice Number
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, C10{Enter}  ; Set Invoice Number
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Set Invoice Number
    Clipboard := invoiceNumber  ; Set Invoice Number
    Sleep, 300  ; Set Invoice Number
    Send, {Control Down}{v}{Control Up}  ; Set Invoice Number
    Sleep, 100
    Sleep, 300  ; Set Invoice Number
    Clipboard := ""
    Send, {F5}  ; Set Slip Number
    Sleep, 100
    Sleep, 300  ; Set Slip Number
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, D10{Enter}  ; Set Slip Number
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Set Slip Number
    Clipboard := slipNumber  ; Set Slip Number
    Sleep, 300  ; Set Slip Number
    Send, {Control Down}{v}{Control Up}  ; Set Slip Number
    Sleep, 100
    Sleep, 300  ; Set Slip Number
    Clipboard := ""
    Send, {F5}  ; Set Supplier Name
    Sleep, 100
    Sleep, 300  ; Set Supplier Name
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, E10{Enter}  ; Set Supplier Name
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Set Supplier Name
    Clipboard := supplierName  ; Set Supplier Name
    Sleep, 300  ; Set Supplier Name
    Send, {Control Down}{v}{Control Up}  ; Set Supplier Name
    Sleep, 100
    Sleep, 300  ; Set Supplier Name
    /*
    Clipboard := ""
    */
    Send, {F5}  ; Set Description
    Sleep, 100
    Sleep, 300  ; Set Description
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, F10{Enter}  ; Set Description
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Set Description
    Clipboard := description  ; Set Slip Number
    Sleep, 300  ; Set Description
    Send, {Control Down}{v}{Control Up}  ; Set Description
    Sleep, 100
    Sleep, 300  ; Set Description
    Clipboard := ""
    Send, {F5}  ; Set Cost Center
    Sleep, 100
    Sleep, 300  ; Set Cost Center
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, I10{Enter}  ; Set Cost Center
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Set Cost Center
    Clipboard := costCenter  ; Set Cost Center
    Sleep, 300  ; Set Cost Center
    Send, {Control Down}{v}{Control Up}  ; Set Cost Center
    Sleep, 100
    Sleep, 300  ; Set Cost Center
    Clipboard := ""
    Send, {F5}  ; Set Debit Account
    Sleep, 100
    Sleep, 300  ; Set Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, J10{Enter}  ; Set Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Set Debit Account
    Clipboard := debitAccount  ; Set Debit Account
    Sleep, 300  ; Set Debit Account
    Send, {Control Down}{v}{Control Up}  ; Set Debit Account
    Sleep, 100
    Sleep, 300  ; Set Debit Account
    Clipboard := ""
    Send, {F5}  ; Set Cost Center
    Sleep, 100
    Sleep, 300  ; Set Cost Center
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, L10{Enter}  ; Set Cost Center
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Set Cost Center
    Clipboard := costCenter  ; Set Cost Center
    Sleep, 300  ; Set Cost Center
    Send, {Control Down}{v}{Control Up}  ; Set Cost Center
    Sleep, 100
    Sleep, 300  ; Set Cost Center
    Clipboard := ""
    Send, {F5}  ; Set Credit Account
    Sleep, 100
    Sleep, 300  ; Set Credit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, M10{Enter}  ; Set Credit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Set Credit Account
    Clipboard := creditAccount  ; Set Credit Account
    Sleep, 300  ; Set Credit Account
    Send, {Control Down}{v}{Control Up}  ; Set Credit Account
    Sleep, 100
    Sleep, 300  ; Set Credit Account
    Clipboard := ""
    Send, {F5}  ; Set currency
    Sleep, 100
    Sleep, 300  ; Set currency
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, O10{Enter}  ; Set currency
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Set currency
    Clipboard := currency  ; Set currency
    Sleep, 300  ; Set currency
    Send, {Control Down}{v}{Control Up}  ; Set currency
    Sleep, 100
    Sleep, 300  ; Set currency
    Clipboard := ""
    Send, {F5}  ; Set Amount
    Sleep, 100
    Sleep, 300  ; Set Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, P10{Enter}  ; Set Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Set Amount
    Clipboard := amount  ; Set Amount
    Sleep, 300  ; Set Amount
    Send, {Control Down}{v}{Control Up}  ; Set Amount
    Sleep, 100
    Sleep, 300  ; Set Amount
    Send, {F5}  ; Replace Cost Center
    Sleep, 100
    Sleep, 300  ; Replace Cost Center
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, I9{Enter}  ; Replace Cost Center
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Replace Cost Center
    Send, {Alt Down}{Down}{Alt Up}
    Sleep, 100
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 6}{Right}{Enter}
    SetKeyDelay, %CurrentKeyDelay%
    WinWaitActive, Custom AutoFilter, , 10
    Sleep, 333
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {0}{Enter}
    SetKeyDelay, %CurrentKeyDelay%
    Send, {F5}  ; Replace Cost Center
    Sleep, 100
    Sleep, 300  ; Replace Cost Center
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, I9{Enter}  ; Replace Cost Center
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Replace Cost Center
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Enter}{'}{0}{0}{Enter}
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Replace Cost Center
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Up}{Control Down}{c}{Control Up}{Control Down}{Shift Down}{Down}{Shift Up}{Control Up}{Enter}
    SetKeyDelay, %CurrentKeyDelay%
    Send, {F5}  ; Replace Cost Center
    Sleep, 100
    Sleep, 300  ; Replace Cost Center
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, L9{Enter}  ; Replace Cost Center
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Replace Cost Center
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Enter}{'}{0}{0}{Enter}
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Replace Cost Center
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Up}{Control Down}{c}{Control Up}{Control Down}{Shift Down}{Down}{Shift Up}{Control Up}{Enter}
    SetKeyDelay, %CurrentKeyDelay%
    MsgBox, 0, , 
    (LTrim
    Accrued Empty Finish
    
    Next >> Ledger Import
    )
    return return
}

FillLedgerImport()
{
    WinActivate, 2. List Accrued EMPTY.xls  [Compatibility Mode] - Excel  ; Get Ledger Data
    Sleep, 333
    Sleep, 300  ; Get Ledger Data
    Send, {Control Down}{PgDn}{Control Up}  ; Get Ledger Data
    Sleep, 100
    Sleep, 300  ; Get Ledger Data
    Send, {Control Down}{PgDn}{Control Up}  ; Get Ledger Data
    Sleep, 100
    Sleep, 300  ; Get Ledger Data
    Send, {Control Down}{PgDn}{Control Up}  ; Get Ledger Data
    Sleep, 100
    Sleep, 300  ; Get Ledger Data
    Send, {F5}  ; Get Ledger Data
    Sleep, 100
    Sleep, 300  ; Get Ledger Data
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, B8{Enter}  ; Get Ledger Data
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Get Ledger Data
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Shift Down}{Space}{Control Down}{Down}{Control Up}{Shift Up}  ; Get Ledger Data
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Get Ledger Data
    Send, {Control Down}{c}{Control Up}  ; Get Ledger Data
    Sleep, 100
    Sleep, 300  ; Get Ledger Data
    WinActivate, 3. LEDGER IMPORT.xls  [Compatibility Mode] - Excel  ; Paste to Ledger Import
    Sleep, 333
    Sleep, 500  ; Paste to Ledger Import
    Send, {F5}  ; Paste to Ledger Import
    Sleep, 100
    Sleep, 300  ; Paste to Ledger Import
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, A8{Enter}  ; Paste to Ledger Import
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Paste to Ledger Import
    Send, {Control Down}{v}{Control Up}  ; Paste to Ledger Import
    Sleep, 100
    Sleep, 3000  ; Paste to Ledger Import
    MsgBox, 0, , 
    (LTrim
    Ledger Import Finish 
    
    Next >> Accrued
    ), 1
    return return
}

FillAccrued()
{
    Sleep, 300
    WinActivate, 1. Cashbill.xlsx - Excel  ; Fill USD
    Sleep, 333
    Sleep, 300
    Send, {Control Down}{Home}{Control Up}
    Sleep, 300
    Send, {Alt}{a}{t}
    Sleep, 1000  ; Fill USD
    Send, {F5}  ; Fill USD
    Sleep, 100
    Sleep, 300  ; Fill USD
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, J1{Enter}  ; Fill USD
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Fill USD
    Send, {Alt Down}{Down}{Alt Up}  ; Fill USD
    Sleep, 100
    Sleep, 300  ; Fill USD
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 7}  ; Fill USD
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Fill USD
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, USD{Enter}  ; Fill USD
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Fill USD
    Send, {Control Down}{Home}{Control Up}  ; Delete C
    Sleep, 100
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Right 4}{Control Down}{Down}{Control Up}{Shift Up}  ; USD_Model - Desc
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; USD_Model - Desc
    Send, {Control Down}{c}{Control Up}  ; USD_Model - Desc
    Sleep, 300
    Sleep, 300  ; USD_Model - Desc
    WinActivate, 4. ACCRUED.xlsx - Excel  ; USD_Model - Desc
    Sleep, 333
    Sleep, 1000  ; USD_Model - Desc
    Send, {F5}  ; USD_Model - Desc
    Sleep, 100
    Sleep, 100  ; USD_Model - Desc
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, B8{Enter}  ; USD_Model - Desc
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; USD_Model - Desc
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; USD_Model - Desc
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 100  ; USD_Model - Desc
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; USD_CC - Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    WinActivate, 1. Cashbill.xlsx - Excel  ; Fill USD
    Sleep, 333
    Sleep, 300  ; USD_CC - Debit Account
    Send, {F5}  ; USD_CC - Debit Account
    Sleep, 100
    Sleep, 300  ; USD_CC - Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, G1{Enter}  ; USD_CC - Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; USD_CC - Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Right}{Control Down}{Down}{Control Up}{Shift Up}  ; USD_CC - Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; USD_CC - Debit Account
    Send, {Control Down}{c}{Control Up}  ; USD_CC - Debit Account
    Sleep, 300
    Sleep, 300  ; USD_CC - Debit Account
    WinActivate, 4. ACCRUED.xlsx - Excel  ; USD_CC - Debit Account
    Sleep, 333
    Sleep, 1000  ; USD_CC - Debit Account
    Send, {F5}  ; USD_CC - Debit Account
    Sleep, 100
    Sleep, 300  ; USD_CC - Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, G8{Enter}  ; USD_CC - Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; USD_CC - Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; USD_CC - Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; USD_CC - Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    WinActivate, 1. Cashbill.xlsx - Excel  ; Fill USD
    Sleep, 333
    Sleep, 300  ; USD_Amount
    Send, {F5}  ; USD_Amount
    Sleep, 100
    Sleep, 300  ; USD_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, K1{Enter}  ; USD_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; USD_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; USD_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; USD_Amount
    Send, {Control Down}{c}{Control Up}  ; USD_Amount
    Sleep, 300
    Sleep, 300  ; USD_Amount
    WinActivate, 4. ACCRUED.xlsx - Excel  ; USD_Amount
    Sleep, 333
    Sleep, 1000  ; USD_Amount
    Send, {F5}  ; USD_Amount
    Sleep, 100
    Sleep, 300  ; USD_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, N8{Enter}  ; USD_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; USD_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; USD_Amount
    SetKeyDelay, %CurrentKeyDelay%
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; USD_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; USD_Amount
    MsgBox, 0, , 
    (LTrim
    USD Selesai
    
    Next >> SGD
    )
    WinActivate, 1. Cashbill.xlsx - Excel  ; Fill SGD
    Sleep, 333
    Sleep, 1000  ; Fill SGD
    Send, {F5}  ; Fill SGD
    Sleep, 100
    Sleep, 300  ; Fill SGD
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, J1{Enter}  ; Fill SGD
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Fill SGD
    Send, {Alt Down}{Down}{Alt Up}  ; Fill SGD
    Sleep, 100
    Sleep, 300  ; Fill SGD
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 7}  ; Fill SGD
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Fill SGD
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, SGD{Enter}  ; Fill SGD
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Fill SGD
    WinActivate, 1. Cashbill.xlsx - Excel  ; Fill SGD
    Sleep, 333
    Sleep, 300  ; Fill SGD
    Send, {Control Down}{Home}{Control Up}  ; SGD_Model - Desc
    Sleep, 100
    Sleep, 300  ; Fill SGD
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Right 4}{Control Down}{Down}{Control Up}{Shift Up}  ; SGD_Model - Desc
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Model - Desc
    Send, {Control Down}{c}{Control Up}  ; SGD_Model - Desc
    Sleep, 300
    Sleep, 300  ; SGD_Model - Desc
    WinActivate, 4. ACCRUED.xlsx - Excel  ; SGD_Model - Desc
    Sleep, 333
    Sleep, 1000  ; SGD_Model - Desc
    Send, {F5}  ; SGD_Model - Desc
    Sleep, 100
    Sleep, 300  ; SGD_Model - Desc
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, B8{Enter}  ; SGD_Model - Desc
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Model - Desc
    Send, {Control Down}{Down}{Control Up}
    Sleep, 300  ; SGD_Model - Desc
    Send, {Down}
    Sleep, 100
    Sleep, 300  ; SGD_Model - Desc
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; SGD_Model - Desc
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Model - Desc
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; SGD_Model - Desc
    SetKeyDelay, %CurrentKeyDelay%
    WinActivate, 1. Cashbill.xlsx - Excel  ; SGD_CC - Debit Account
    Sleep, 333
    Sleep, 300  ; SGD_CC - Debit Account
    Send, {F5}  ; SGD_CC - Debit Account
    Sleep, 100
    Sleep, 300  ; SGD_CC - Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, G1{Enter}  ; SGD_CC - Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_CC - Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Right}{Control Down}{Down}{Control Up}{Shift Up}  ; SGD_CC - Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_CC - Debit Account
    Send, {Control Down}{c}{Control Up}  ; SGD_CC - Debit Account
    Sleep, 300
    Sleep, 300  ; SGD_CC - Debit Account
    WinActivate, 4. ACCRUED.xlsx - Excel  ; SGD_CC - Debit Account
    Sleep, 333
    Sleep, 1000  ; SGD_CC - Debit Account
    Send, {F5}  ; SGD_CC - Debit Account
    Sleep, 100
    Sleep, 300  ; SGD_CC - Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, G8{Enter}  ; SGD_CC - Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_CC - Debit Account
    Send, {Control Down}{Down}{Control Up}  ; SGD_CC - Debit Account
    Sleep, 300  ; SGD_CC - Debit Account
    Send, {Down}  ; SGD_CC - Debit Account
    Sleep, 100
    Sleep, 300  ; SGD_CC - Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; SGD_CC - Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; SGD_CC - Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    WinActivate, 1. Cashbill.xlsx - Excel  ; SGD_Amount
    Sleep, 333
    Sleep, 300  ; SGD_Amount
    Send, {F5}  ; SGD_Amount
    Sleep, 100
    Sleep, 300  ; SGD_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, K1{Enter}  ; SGD_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; SGD_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Amount
    Send, {Control Down}{c}{Control Up}  ; SGD_Amount
    Sleep, 300
    Sleep, 300  ; SGD_Amount
    WinActivate, 4. ACCRUED.xlsx - Excel  ; SGD_Amount
    Sleep, 333
    Sleep, 1000  ; SGD_Amount
    Send, {F5}  ; SGD_Amount
    Sleep, 100
    Sleep, 300  ; SGD_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, N8{Enter}  ; SGD_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Amount
    Send, {Control Down}{Down}{Control Up}  ; SGD_Amount
    Sleep, 100
    Sleep, 300  ; SGD_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Right 2}
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; SGD_Amount
    SetKeyDelay, %CurrentKeyDelay%
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; SGD_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Amount
    Send, {F5}  ; SGD Daily Rate
    Sleep, 100
    Sleep, 300  ; SGD Daily Rate
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, P6{Enter}  ; SGD Daily Rate
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD Daily Rate
    Send, {Control Down}{Down}{Control Up}  ; SGD_Amount
    Sleep, 100
    Sleep, 300  ; SGD Daily Rate
    Loop, 6
    {
        Send, {Left}  ; SGD Daily Rate
        Sleep, 100
    }
    Sleep, 300  ; SGD_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, =VLOOKUP({Home}{Right}`,{Control Down}{PgUp}{Home}{Shift Down}{Right}{Down}{Shift Up}{Control Up}`,3`,0{Enter}  ; SGD Daily Rate
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD Daily Rate
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Up}{Control Down}{c}{Control Up}{Left 2}{Control Down}{Down}{Control Up}{Right 2}  ; SGD Daily Rate
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD Daily Rate
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Shift Down}{Up}{Shift Up}{v}{Control Up}  ; SGD Daily Rate
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Rate B
    Send, {F5}  ; SGD_Rate B
    Sleep, 100
    Sleep, 300  ; SGD_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, J8{Enter}  ; SGD_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Rate B
    Send, {Control Down}{Down}{Control Up}  ; SGD_Rate B
    Sleep, 100
    Sleep, 300  ; SGD_Rate B
    Loop, 3
    {
        Send, {Right}  ; SGD_Rate B
        Sleep, 100
    }
    Sleep, 300  ; SGD_Rate B
    Send, {Control Down}{PgUp}{Control Up}  ; SGD_Rate B
    Sleep, 100
    Sleep, 300  ; SGD_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Home}{Down}{Control Up}{Right 2}  ; SGD_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Rate B
    Send, {Control Down}{c}{Control Up}  ; SGD_Rate B
    Sleep, 100
    Sleep, 300  ; SGD_Rate B
    Send, {Control Down}{PgDn}{Control Up}  ; SGD_Rate B
    Sleep, 100
    Sleep, 300  ; SGD_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; SGD_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; SGD_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{c}{Control Up}{Control Down}{Left}{Down}{Control Up}{Right 3}  ; SGD_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Shift Down}{Up}{Shift Up}{v}{Control Up}  ; SGD_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Rate B
    Send, {F5}  ; Rumus Round USD
    Sleep, 100
    Sleep, 300  ; Rumus Round USD
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, N8{Enter}  ; Rumus Round USD
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Rumus Round USD
    Loop, 2
    {
        Send, {Control Down}{Down}{Control Up}  ; SGD_Amount
        Sleep, 100
    }
    Sleep, 300  ; round Rumus SGD
    Send, {Control Down}{c}{Control Up}  ; round Rumus SGD
    Sleep, 100
    Sleep, 300  ; round Rumus SGD
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Up}{Control Up}{Down}  ; round Rumus SGD
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; round Rumus SGD
    Send, {Control Down}{v}{Control Up}  ; round Rumus SGD
    Sleep, 100
    Sleep, 300  ; round Rumus SGD
    Send, {Control Down}{c}{Control Up}  ; round Rumus SGD
    Sleep, 100
    Sleep, 300  ; round Rumus SGD
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Left}{Down}{Control Up}  ; round Rumus SGD
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; round Rumus SGD
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Right}{Control Down}{Shift Down}{Up}{Shift Up}{Control Up}  ; round Rumus SGD
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; SGD_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; SGD_Amount
    SetKeyDelay, %CurrentKeyDelay%
    MsgBox, 0, , SGD  ; Set Currency to JPY
    WinActivate, 1. Cashbill.xlsx - Excel  ; Set Currency to JPY
    Sleep, 333
    Sleep, 1000  ; Set Currency to JPY
    Send, {F5}  ; Set Currency to JPY
    Sleep, 100
    Sleep, 300  ; Set Currency to JPY
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, J1{Enter}  ; Set Currency to JPY
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Set Currency to JPY
    Send, {Alt Down}{Down}{Alt Up}  ; Set Currency to JPY
    Sleep, 100
    Sleep, 300  ; Set Currency to JPY
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 7}  ; Set Currency to JPY
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Set Currency to JPY
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, JPY{Enter}  ; Set Currency to JPY
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Set Currency to JPY
    WinActivate, 1. Cashbill.xlsx - Excel  ; JPY_Invoice Date - Description
    Sleep, 333
    Sleep, 300  ; JPY_Invoice Date - Description
    Send, {Control Down}{Home}{Control Up}  ; JPY_Invoice Date - Description
    Sleep, 100
    Sleep, 300  ; JPY_Invoice Date - Description
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Right 4}{Control Down}{Down}{Control Up}{Shift Up}  ; JPY_Invoice Date - Description
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Invoice Date - Description
    Send, {Control Down}{c}{Control Up}  ; JPY_Invoice Date - Description
    Sleep, 300
    Sleep, 300  ; JPY_Invoice Date - Description
    WinActivate, 4. ACCRUED.xlsx - Excel  ; JPY_Invoice Date - Description
    Sleep, 333
    Sleep, 1000  ; JPY_Invoice Date - Description
    Send, {F5}  ; JPY_Invoice Date - Description
    Sleep, 100
    Sleep, 300  ; JPY_Invoice Date - Description
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, B8{Enter}  ; JPY_Invoice Date - Description
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Invoice Date - Description
    Send, {Control Down}{Down}{Control Up}  ; JPY_Invoice Date - Description
    Sleep, 300  ; JPY_Invoice Date - Description
    Send, {Down}  ; JPY_Invoice Date - Description
    Sleep, 100
    Sleep, 300  ; JPY_Invoice Date - Description
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; JPY_Invoice Date - Description
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Invoice Date - Description
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; JPY_Invoice Date - Description
    SetKeyDelay, %CurrentKeyDelay%
    WinActivate, 1. Cashbill.xlsx - Excel  ; JPY_CC - Debit Account
    Sleep, 333
    Sleep, 300  ; JPY_CC - Debit Account
    Send, {F5}  ; JPY_CC - Debit Account
    Sleep, 100
    Sleep, 300  ; JPY_CC - Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, G1{Enter}  ; JPY_CC - Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_CC - Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Right}{Control Down}{Down}{Control Up}{Shift Up}  ; JPY_CC - Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_CC - Debit Account
    Send, {Control Down}{c}{Control Up}  ; JPY_CC - Debit Account
    Sleep, 300
    Sleep, 300  ; JPY_CC - Debit Account
    WinActivate, 4. ACCRUED.xlsx - Excel  ; JPY_CC - Debit Account
    Sleep, 333
    Sleep, 1000  ; JPY_CC - Debit Account
    Send, {F5}  ; JPY_CC - Debit Account
    Sleep, 100
    Sleep, 300  ; JPY_CC - Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, G8{Enter}  ; JPY_CC - Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_CC - Debit Account
    Send, {Control Down}{Down}{Control Up}  ; JPY_CC - Debit Account
    Sleep, 100
    Sleep, 300  ; JPY_CC - Debit Account
    Send, {Down}  ; JPY_CC - Debit Account
    Sleep, 100
    Sleep, 300  ; JPY_CC - Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; JPY_CC - Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; JPY_CC - Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    WinActivate, 1. Cashbill.xlsx - Excel  ; JPY_Amount
    Sleep, 333
    Sleep, 300  ; JPY_Amount
    Send, {F5}  ; JPY_Amount
    Sleep, 100
    Sleep, 300  ; JPY_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, K1{Enter}  ; JPY_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; JPY_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Amount
    Send, {Control Down}{c}{Control Up}  ; JPY_Amount
    Sleep, 300
    Sleep, 300  ; JPY_Amount
    WinActivate, 4. ACCRUED.xlsx - Excel  ; JPY_Amount
    Sleep, 333
    Sleep, 1000  ; JPY_Amount
    Send, {F5}  ; JPY_Amount
    Sleep, 100
    Sleep, 300  ; JPY_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, N8{Enter}  ; JPY_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Amount
    Send, {Control Down}{Down}{Control Up}  ; JPY_Amount
    Sleep, 100
    Sleep, 300  ; JPY_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Right 2}  ; JPY_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; JPY_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; JPY_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Amount
    Send, {F5}  ; JPY_Daily Rate Look Up
    Sleep, 100
    Sleep, 300  ; JPY_Daily Rate Look Up
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, P6{Enter}  ; JPY_Daily Rate Look Up
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Daily Rate Look Up
    Send, {Control Down}{Down}{Control Up}  ; JPY_Daily Rate Look Up
    Sleep, 100
    Sleep, 300  ; JPY_Daily Rate Look Up
    Loop, 6
    {
        Send, {Left}  ; JPY_Daily Rate Look Up
        Sleep, 100
    }
    Sleep, 300  ; JPY_Daily Rate Look Up
    Send, {Control Down}{Down}{Control Up}  ; JPY_Daily Rate Look Up
    Sleep, 100
    Sleep, 300  ; JPY_Daily Rate Look Up
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Right}
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Daily Rate Look Up
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, =VLOOKUP({Home}{Right}`,{Control Down}{PgUp}{Home}{Shift Down}{Right}{Down}{Shift Up}{Control Up}`,4`,0{Enter}  ; JPY_Daily Rate Look Up
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Daily Rate Look Up
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Up}{Control Down}{c}{Control Up}{Left 3}{Control Down}{Down}{Control Up}{Right 3}  ; JPY_Daily Rate Look Up
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Daily Rate Look Up
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Shift Down}{Up}{Shift Up}{v}{Control Up}  ; JPY_Daily Rate Look Up
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Daily Rate Look Up
    Sleep, 300  ; JPY_Rate B
    Send, {F5}  ; JPY_Rate B
    Sleep, 100
    Sleep, 300  ; JPY_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, J8{Enter}  ; JPY_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Rate B
    Loop, 2
    {
        Send, {Control Down}{Down}{Control Up}  ; JPY_Rate B
        Sleep, 100
    }
    Sleep, 300  ; JPY_Rate B
    Send, {Down}
    Sleep, 300  ; JPY_Rate B
    Loop, 3
    {
        Send, {Right}  ; JPY_Rate B
        Sleep, 100
    }
    Sleep, 300  ; JPY_Rate B
    Send, {Control Down}{PgUp}{Control Up}  ; JPY_Rate B
    Sleep, 100
    Sleep, 300  ; JPY_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Home}{Down}{Control Up}{Right 3}  ; JPY_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Rate B
    Send, {Control Down}{c}{Control Up}  ; JPY_Rate B
    Sleep, 100
    Sleep, 300  ; JPY_Rate B
    Send, {Control Down}{PgDn}{Control Up}  ; JPY_Rate B
    Sleep, 100
    Sleep, 300  ; JPY_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; JPY_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; JPY_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{c}{Control Up}{Control Down}{Left}{Down}{Control Up}{Right 2}  ; JPY_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Shift Down}{Up}{Shift Up}{v}{Control Up}  ; JPY_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Rate B
    Send, {F5}  ; JPY_Round
    Sleep, 100
    Sleep, 300  ; JPY_Round
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, N8{Enter}  ; JPY_Round
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Round
    Send, {Control Down}{Down}{Control Up}  ; JPY_Round
    Sleep, 100
    Sleep, 300  ; JPY_Round
    Send, {Down}  ; JPY_Round
    Sleep, 300  ; JPY_Round
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, =ROUND(({Right 2}/{Left 3})`,2){Enter}  ; JPY_Round
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Round
    Send, {Up}  ; JPY_Round
    Sleep, 100
    Sleep, 300  ; JPY_Round
    Send, {Control Down}{c}{Control Up}  ; JPY_Round
    Sleep, 100
    Sleep, 300  ; JPY_Round
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Left}{Down}{Control Up}  ; JPY_Round
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Round
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Right 1}{Control Down}{Shift Down}{Up}{Shift Up}{Control Up}  ; JPY_Round
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Round
    Send, {Control Down}{v}{Control Up}
    Sleep, 100
    MsgBox, 0, , JPY  ; IDR
    WinActivate, 1. Cashbill.xlsx - Excel  ; Set IDR to Cash Bill
    Sleep, 333
    Sleep, 1000  ; Set IDR to Cash Bill
    Send, {F5}  ; Set IDR to Cash Bill
    Sleep, 100
    Sleep, 300  ; Set IDR to Cash Bill
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, J1{Enter}  ; Set IDR to Cash Bill
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Set IDR to Cash Bill
    Send, {Alt Down}{Down}{Alt Up}  ; Set IDR to Cash Bill
    Sleep, 100
    Sleep, 300  ; Set IDR to Cash Bill
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 7}  ; Set IDR to Cash Bill
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Set IDR to Cash Bill
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, RP{Enter}  ; Set IDR to Cash Bill
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Set IDR to Cash Bill
    WinActivate, 1. Cashbill.xlsx - Excel  ; RP_Invoice Date - Desc
    Sleep, 333
    Sleep, 300  ; RP_Invoice Date - Desc
    Send, {Control Down}{Home}{Control Up}  ; RP_Invoice Date - Desc
    Sleep, 100
    Sleep, 300  ; RP_Invoice Date - Desc
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Right 4}{Control Down}{Down}{Control Up}{Shift Up}  ; RP_Invoice Date - Desc
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Invoice Date - Desc
    Send, {Control Down}{c}{Control Up}  ; RP_Invoice Date - Desc
    Sleep, 300
    Sleep, 300  ; RP_Invoice Date - Desc
    WinActivate, 4. ACCRUED.xlsx - Excel  ; RP_Invoice Date - Desc
    Sleep, 333
    Sleep, 1000  ; RP_Invoice Date - Desc
    Send, {F5}  ; RP_Invoice Date - Desc
    Sleep, 100
    Sleep, 300  ; RP_Invoice Date - Desc
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, B8{Enter}  ; RP_Invoice Date - Desc
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Invoice Date - Desc
    Loop, 2
    {
        Send, {Control Down}{Down}{Control Up}  ; RP_Invoice Date - Desc
        Sleep, 100
    }
    Sleep, 300  ; RP_Invoice Date - Desc
    Loop, 2
    {
        Send, {Down}  ; RP_Invoice Date - Desc
        Sleep, 100
    }
    Sleep, 300  ; RP_Invoice Date - Desc
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; RP_Invoice Date - Desc
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Invoice Date - Desc
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; RP_Invoice Date - Desc
    SetKeyDelay, %CurrentKeyDelay%
    WinActivate, 1. Cashbill.xlsx - Excel  ; RP_CC - Debit Account
    Sleep, 333
    Sleep, 300  ; RP_CC - Debit Account
    Send, {F5}  ; RP_CC - Debit Account
    Sleep, 100
    Sleep, 300  ; RP_CC - Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, G1{Enter}  ; RP_CC - Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_CC - Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Right}{Control Down}{Down}{Control Up}{Shift Up}  ; RP_CC - Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_CC - Debit Account
    Send, {Control Down}{c}{Control Up}  ; RP_CC - Debit Account
    Sleep, 300
    Sleep, 300  ; RP_CC - Debit Account
    WinActivate, 4. ACCRUED.xlsx - Excel  ; RP_CC - Debit Account
    Sleep, 333
    Sleep, 1000  ; RP_CC - Debit Account
    Send, {F5}  ; RP_CC - Debit Account
    Sleep, 100
    Sleep, 300  ; RP_CC - Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, G8{Enter}  ; RP_CC - Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_CC - Debit Account
    Loop, 2
    {
        Send, {Control Down}{Down}{Control Up}  ; RP_CC - Debit Account
        Sleep, 100
    }
    Sleep, 300  ; RP_CC - Debit Account
    Loop, 2
    {
        Send, {Down}  ; RP_CC - Debit Account
        Sleep, 100
    }
    Sleep, 300  ; RP_CC - Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; RP_CC - Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_CC - Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; RP_CC - Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    WinActivate, 1. Cashbill.xlsx - Excel  ; RP_Amount
    Sleep, 333
    Sleep, 300  ; RP_Amount
    Send, {F5}  ; RP_Amount
    Sleep, 100
    Sleep, 300  ; RP_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, K1{Enter}  ; RP_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; RP_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Amount
    Send, {Control Down}{c}{Control Up}  ; RP_Amount
    Sleep, 300
    Sleep, 300  ; RP_Amount
    WinActivate, 4. ACCRUED.xlsx - Excel  ; RP_Amount
    Sleep, 333
    Sleep, 1000  ; RP_Amount
    Send, {F5}  ; RP_Amount
    Sleep, 100
    Sleep, 300  ; RP_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, B8{Enter}  ; RP_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Amount
    Loop, 2
    {
        Send, {Control Down}{Down}{Control Up}  ; RP_Amount
        Sleep, 100
    }
    Sleep, 300  ; RP_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Control Down}{Right}{Control Up}{Right 7}  ; RP_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; RP_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; RP_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Amount
    Send, {F5}  ; RP_Daily Rate
    Sleep, 100
    Sleep, 300  ; RP_Daily Rate
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, B8{Enter}  ; RP_Daily Rate
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Daily Rate
    Loop, 2
    {
        Send, {Control Down}{Down}{Control Up}  ; RP_Daily Rate
        Sleep, 100
    }
    Sleep, 300  ; RP_Daily Rate
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Control Down}{Right}{Control Up}{Right 3}  ; RP_Daily Rate
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Daily Rate
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, =VLOOKUP({Home}{Right}`,{Control Down}{PgUp}{Home}{Shift Down}{Right}{Down}{Shift Up}{Control Up}{F4}`,2`,0{Enter}  ; RP_Daily Rate
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Daily Rate
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Up}{Control Down}{c}{Control Up}{Left 4}{Control Down}{Down}{Control Up}{Right 4}  ; RP_Daily Rate
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Daily Rate
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Shift Down}{Up}{Shift Up}{v}{Control Up}  ; RP_Daily Rate
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Daily Rate
    Send, {F5}  ; RP_Rumus Round
    Sleep, 100
    Sleep, 300  ; RP_Rumus Round
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, B8{Enter}  ; RP_Rumus Round
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Rumus Round
    Loop, 2
    {
        Send, {Control Down}{Down}{Control Up}  ; RP_Rumus Round
        Sleep, 100
    }
    Sleep, 300  ; RP_Rumus Round
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Control Down}{Right}{Control Up}{Right 5}  ; RP_Rumus Round
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Rumus Round
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, =ROUND(({Right 2}/{Left 2})`,2){Enter}  ; RP_Rumus Round
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Rumus Round
    Send, {Up}  ; RP_Rumus Round
    Sleep, 100
    Sleep, 300  ; RP_Rumus Round
    Send, {Control Down}{c}{Control Up}  ; RP_Rumus Round
    Sleep, 100
    Sleep, 300  ; RP_Rumus Round
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Left}{Down}{Control Up}  ; RP_Rumus Round
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Rumus Round
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Right 2}{Control Down}{Shift Down}{Up}{Shift Up}{Control Up}  ; RP_Rumus Round
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Rumus Round
    Send, {Control Down}{v}{Control Up}  ; RP_Rumus Round
    Sleep, 100
    Sleep, 300  ; RP_Rumus Round
    MsgBox, 0, , IDR Finish, 1  ; IDR
    return return
}

FillExcRate()
{
    WinActivate, 1. Cashbill.xlsx - Excel  ; Fill USD
    Sleep, 333
    Sleep, 1000  ; Fill USD
    Send, {F5}  ; Delete Slip Number
    Sleep, 100
    Sleep, 300  ; Delete Slip Number
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, C1{Enter}  ; Delete Slip Number
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Delete Slip Number
    Send, {Control Down}{Space}{Control Up}  ; Delete Slip Number
    Sleep, 300  ; Delete Slip Number
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {AppsKey}{d}  ; Delete Slip Number
    SetKeyDelay, %CurrentKeyDelay%
    Send, {F5}  ; Fill USD
    Sleep, 100
    Sleep, 300  ; Fill USD
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, I1{Enter}  ; Fill USD
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Fill USD
    Send, {Alt Down}{Down}{Alt Up}  ; Fill USD
    Sleep, 100
    Sleep, 300  ; Fill USD
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 7}  ; Fill USD
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Fill USD
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, USD{Enter}  ; Fill USD
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Fill USD
    Send, {Control Down}{Home}{Control Up}  ; USD MODEL
    Sleep, 100
    Sleep, 300  ; USD Invoice Date
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Right 2}{Control Down}{Down}{Control Up}{Shift Up}  ; USD Invoice Date
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; USD Invoice Date
    Send, {Control Down}{c}{Control Up}  ; USD Invoice Date
    Sleep, 300
    Sleep, 300  ; USD Invoice Date
    WinActivate, 5. EXC.RATE.xlsx - Excel  ; USD Invoice Date
    Sleep, 333
    Sleep, 1000  ; USD Invoice Date
    Send, {F5}  ; USD Invoice Date
    Sleep, 100
    Sleep, 100  ; USD Invoice Date
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, B8{Enter}  ; USD Invoice Date
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; USD Invoice Date
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; USD Invoice Date
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 100  ; USD Invoice Date
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; USD_CC - Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    WinActivate, 1. Cashbill.xlsx - Excel  ; USD_Debit Account
    Sleep, 333
    Sleep, 300  ; USD_Debit Account
    Send, {F5}  ; USD_Debit Account
    Sleep, 100
    Sleep, 300  ; USD_Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, G1{Enter}  ; USD_Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; USD_Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; USD_Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; USD_Debit Account
    Send, {Control Down}{c}{Control Up}  ; USD_Debit Account
    Sleep, 300
    Sleep, 300  ; USD_Debit Account
    WinActivate, 5. EXC.RATE.xlsx - Excel  ; USD_Debit Account
    Sleep, 333
    Sleep, 1000  ; USD_Debit Account
    Send, {F5}  ; USD_Debit Account
    Sleep, 100
    Sleep, 300  ; USD_Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, F8{Enter}  ; USD_Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; USD_Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; USD_Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; USD_Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; USD_Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    WinActivate, 1. Cashbill.xlsx - Excel  ; Fill USD
    Sleep, 333
    Sleep, 300  ; USD_Amount
    Send, {F5}  ; USD_Amount
    Sleep, 100
    Sleep, 300  ; USD_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, J1{Enter}  ; USD_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; USD_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; USD_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; USD_Amount
    Send, {Control Down}{c}{Control Up}  ; USD_Amount
    Sleep, 300
    Sleep, 300  ; USD_Amount
    WinActivate, 5. EXC.RATE.xlsx - Excel  ; USD_Amount
    Sleep, 333
    Sleep, 1000  ; USD_Amount
    Send, {F5}  ; USD_Amount
    Sleep, 100
    Sleep, 300  ; USD_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, K8{Enter}  ; USD_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; USD_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; USD_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; USD_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; USD_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; USD_Amount
    MsgBox, 0, , 
    (LTrim
    USD Selesai
    
    Next >> SGD
    )
    WinActivate, 1. Cashbill.xlsx - Excel  ; Fill SGD
    Sleep, 333
    Sleep, 1000  ; Fill SGD
    Send, {F5}  ; Fill SGD
    Sleep, 100
    Sleep, 300  ; Fill SGD
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, I1{Enter}  ; Fill SGD
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Fill SGD
    Send, {Alt Down}{Down}{Alt Up}  ; Fill SGD
    Sleep, 100
    Sleep, 300  ; Fill SGD
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 7}  ; Fill SGD
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Fill SGD
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, SGD{Enter}  ; Fill SGD
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Fill SGD
    WinActivate, 1. Cashbill.xlsx - Excel  ; SGD_Invoice Date - Supp Name
    Sleep, 333
    Sleep, 300  ; SGD_Invoice Date - Supp Name
    Send, {Control Down}{Home}{Control Up}  ; SGD_Invoice Date - Supp Name
    Sleep, 100
    Sleep, 300  ; SGD_Invoice Date - Supp Name
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Right 2}{Control Down}{Down}{Control Up}{Shift Up}  ; SGD_Invoice Date - Supp Name
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Invoice Date - Supp Name
    Send, {Control Down}{c}{Control Up}  ; SGD_Invoice Date - Supp Name
    Sleep, 300
    Sleep, 300  ; SGD_Invoice Date - Supp Name
    WinActivate, 5. EXC.RATE.xlsx - Excel  ; SGD_Invoice Date - Supp Name
    Sleep, 333
    Sleep, 1000  ; SGD_Invoice Date - Supp Name
    Send, {F5}  ; SGD_Invoice Date - Supp Name
    Sleep, 100
    Sleep, 100  ; SGD_Invoice Date - Supp Name
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, B8{Enter}  ; SGD_Invoice Date - Supp Name
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Invoice Date - Supp Name
    Send, {Control Down}{Down}{Control Up}  ; SGD_Invoice Date - Supp Name
    Sleep, 300  ; SGD_Model - Desc
    Send, {Down}
    Sleep, 100
    Sleep, 300  ; SGD_Invoice Date - Supp Name
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; SGD_Invoice Date - Supp Name
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 100  ; SGD_Invoice Date - Supp Name
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; SGD_Invoice Date - Supp Name
    SetKeyDelay, %CurrentKeyDelay%
    WinActivate, 1. Cashbill.xlsx - Excel  ; SGD_Debit Account
    Sleep, 333
    Sleep, 300  ; SGD_Debit Account
    Send, {F5}  ; SGD_Debit Account
    Sleep, 100
    Sleep, 300  ; SGD_Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, G1{Enter}  ; SGD_Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; SGD_Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Debit Account
    Send, {Control Down}{c}{Control Up}  ; SGD_Debit Account
    Sleep, 300
    Sleep, 300  ; SGD_Debit Account
    WinActivate, 5. EXC.RATE.xlsx - Excel  ; SGD_Debit Account
    Sleep, 333
    Sleep, 1000  ; SGD_Debit Account
    Send, {F5}  ; SGD_Debit Account
    Sleep, 100
    Sleep, 300  ; SGD_Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, F8{Enter}  ; SGD_Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Debit Account
    Send, {Control Down}{Down}{Control Up}  ; SGD_Debit Account
    Sleep, 300  ; SGD_Debit Account
    Send, {Down}  ; SGD_Debit Account
    Sleep, 100
    Sleep, 300  ; SGD_Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; SGD_Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; SGD_Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    WinActivate, 1. Cashbill.xlsx - Excel  ; SGD_Amount
    Sleep, 333
    Sleep, 300  ; SGD_Amount
    Send, {F5}  ; SGD_Amount
    Sleep, 100
    Sleep, 300  ; SGD_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, J1{Enter}  ; SGD_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; SGD_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Amount
    Send, {Control Down}{c}{Control Up}  ; SGD_Amount
    Sleep, 300
    Sleep, 300  ; SGD_Amount
    WinActivate, 5. EXC.RATE.xlsx - Excel  ; SGD_Amount
    Sleep, 333
    Sleep, 1000  ; SGD_Amount
    Send, {F5}  ; SGD_Amount
    Sleep, 100
    Sleep, 300  ; SGD_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, K8{Enter}  ; SGD_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Amount
    Send, {Control Down}{Down}{Control Up}  ; SGD_Amount
    Sleep, 300  ; SGD_Amount
    Send, {Down}  ; SGD_Amount
    Sleep, 100
    Sleep, 300  ; SGD_Amount
    Loop, 2
    {
        Send, {Right}  ; SGD_Amount
        Sleep, 100
    }
    Sleep, 300  ; SGD_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; SGD_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; SGD_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Amount
    WinActivate, 4. ACCRUED.xlsx - Excel  ; SGD_Daily Rate
    Sleep, 333
    Send, {Control Down}{PgUp}{Control Up}  ; SGD_Rate B
    Sleep, 100
    Sleep, 300  ; SGD_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Home}{Down}{Control Up}{Right 2}  ; SGD_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Rate B
    Send, {Control Down}{c}{Control Up}  ; SGD_Rate B
    Sleep, 100
    Sleep, 300  ; SGD_Rate B
    Send, {Control Down}{PgDn}{Control Up}  ; SGD_Rate B
    Sleep, 100
    Sleep, 300  ; SGD_Rate B
    WinActivate, 5. EXC.RATE.xlsx - Excel  ; SGD_Daily Rate
    Sleep, 333
    Sleep, 300  ; SGD_Rate B
    Send, {F5}  ; SGD_Rate B
    Sleep, 100
    Sleep, 300  ; SGD_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, K8{Enter}  ; SGD_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Rate B
    Send, {Control Down}{Down}{Control Up}  ; SGD_Rate B
    Sleep, 100
    Sleep, 300  ; SGD_Rate B
    Send, {Down}
    Sleep, 100
    Sleep, 300  ; SGD_Rate B
    Loop, 4
    {
        Send, {Left}
        Sleep, 100
    }
    Sleep, 300  ; SGD_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; SGD_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; SGD_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{c}{Control Up}{Control Down}{Left}{Down}{Control Up}{Right}  ; SGD_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Shift Down}{Up}{Shift Up}{v}{Control Up}  ; SGD_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Rate B
    Send, {F5}  ; SGD_Round Rumus
    Sleep, 100
    Sleep, 300  ; SGD_Round Rumus
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, K8{Enter}  ; SGD_Round Rumus
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Round Rumus
    Send, {Control Down}{Down}{Control Up}  ; SGD_Round Rumus
    Sleep, 100
    Sleep, 300  ; SGD_Round Rumus
    Send, {Down}  ; SGD_Round Rumus
    Sleep, 100
    Sleep, 300  ; SGD_Round Rumus
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, =ROUND(({Right 2}/{Left 4})`,2){Enter}  ; SGD_Round Rumus
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Round Rumus
    Send, {Up}  ; SGD_Round Rumus
    Sleep, 100
    Sleep, 300  ; SGD_Round Rumus
    Send, {Control Down}{c}{Control Up}  ; SGD_Round Rumus
    Sleep, 300  ; SGD_Round Rumus
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Left}{Down}{Control Up}{Right 4}{Shift Down}{Control Down}{Up}{Control Up}{Down}{Shift Up}  ; SGD_Round Rumus
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; SGD_Round Rumus
    Send, {Control Down}{v}{Control Up}  ; SGD_Round Rumus
    Sleep, 100
    Sleep, 300  ; SGD_Round Rumus
    WinActivate, 1. Cashbill.xlsx - Excel  ; Fill JPY
    Sleep, 333
    Sleep, 1000  ; Fill JPY
    Send, {F5}  ; Fill JPY
    Sleep, 100
    Sleep, 300  ; Fill JPY
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, I1{Enter}  ; Fill JPY
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Fill JPY
    Send, {Alt Down}{Down}{Alt Up}  ; Fill JPY
    Sleep, 100
    Sleep, 300  ; Fill JPY
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 7}  ; Fill JPY
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Fill JPY
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, JPY{Enter}  ; Fill JPY
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Fill JPY
    WinActivate, 1. Cashbill.xlsx - Excel  ; JPY_Inv date - Supp Name
    Sleep, 333
    Sleep, 300  ; JPY_Inv date - Supp Name
    Send, {Control Down}{Home}{Control Up}  ; JPY_Inv date - Supp Name
    Sleep, 100
    Sleep, 300  ; JPY_Inv date - Supp Name
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Right 2}{Control Down}{Down}{Control Up}{Shift Up}  ; JPY_Inv date - Supp Name
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Inv date - Supp Name
    Send, {Control Down}{c}{Control Up}  ; JPY_Inv date - Supp Name
    Sleep, 300
    Sleep, 300  ; JPY_Inv date - Supp Name
    WinActivate, 5. EXC.RATE.xlsx - Excel  ; JPY_Inv date - Supp Name
    Sleep, 333
    Sleep, 1000  ; JPY_Inv date - Supp Name
    Send, {F5}  ; JPY_Inv date - Supp Name
    Sleep, 100
    Sleep, 100  ; JPY_Inv date - Supp Name
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, B8{Enter}  ; JPY_Inv date - Supp Name
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Inv date - Supp Name
    Send, {Control Down}{Down}{Control Up}  ; JPY_Inv date - Supp Name
    Sleep, 100
    Sleep, 300  ; JPY_Inv date - Supp Name
    Send, {Down}  ; JPY_Inv date - Supp Name
    Sleep, 100
    Sleep, 300  ; JPY_Inv date - Supp Name
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; JPY_Inv date - Supp Name
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 100  ; JPY_Inv date - Supp Name
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; JPY_Inv date - Supp Name
    SetKeyDelay, %CurrentKeyDelay%
    WinActivate, 1. Cashbill.xlsx - Excel  ; JPY_Debit Account
    Sleep, 333
    Sleep, 300  ; JPY_Debit Account
    Send, {F5}  ; JPY_Debit Account
    Sleep, 100
    Sleep, 300  ; JPY_Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, G1{Enter}  ; JPY_Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; JPY_Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Debit Account
    Send, {Control Down}{c}{Control Up}  ; JPY_Debit Account
    Sleep, 300
    Sleep, 300  ; JPY_Debit Account
    WinActivate, 5. EXC.RATE.xlsx - Excel  ; JPY_Debit Account
    Sleep, 333
    Sleep, 1000  ; JPY_Debit Account
    Send, {F5}  ; JPY_Debit Account
    Sleep, 100
    Sleep, 300  ; JPY_Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, F8{Enter}  ; JPY_Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Debit Account
    Send, {Control Down}{Down}{Control Up}  ; JPY_Debit Account
    Sleep, 300  ; JPY_Debit Account
    Send, {Down}  ; JPY_Debit Account
    Sleep, 100
    Sleep, 300  ; JPY_Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; JPY_Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; JPY_Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    WinActivate, 1. Cashbill.xlsx - Excel  ; JPY_Amount
    Sleep, 333
    Sleep, 300  ; JPY_Amount
    Send, {F5}  ; JPY_Amount
    Sleep, 100
    Sleep, 300  ; JPY_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, J1{Enter}  ; JPY_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; JPY_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Amount
    Send, {Control Down}{c}{Control Up}  ; JPY_Amount
    Sleep, 300
    Sleep, 300  ; JPY_Amount
    WinActivate, 5. EXC.RATE.xlsx - Excel  ; JPY_Amount
    Sleep, 333
    Sleep, 1000  ; JPY_Amount
    Send, {F5}  ; JPY_Amount
    Sleep, 100
    Sleep, 300  ; JPY_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, K8{Enter}  ; JPY_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Amount
    Send, {Control Down}{Down}{Control Up}  ; JPY_Amount
    Sleep, 300  ; JPY_Amount
    Send, {Down}  ; JPY_Amount
    Sleep, 100
    Sleep, 300  ; JPY_Amount
    Loop, 2
    {
        Send, {Right}  ; JPY_Amount
        Sleep, 100
    }
    Sleep, 300  ; JPY_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; JPY_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; JPY_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Amount
    WinActivate, 4. ACCRUED.xlsx - Excel  ; JPY_Rate B
    Sleep, 333
    Send, {Control Down}{PgUp}{Control Up}  ; JPY_Rate B
    Sleep, 100
    Sleep, 300  ; JPY_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Home}{Down}{Control Up}{Right 3}  ; JPY_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Rate B
    Send, {Control Down}{c}{Control Up}  ; JPY_Rate B
    Sleep, 100
    Sleep, 300  ; JPY_Rate B
    Send, {Control Down}{PgDn}{Control Up}  ; JPY_Rate B
    Sleep, 100
    Sleep, 300  ; JPY_Rate B
    WinActivate, 5. EXC.RATE.xlsx - Excel  ; JPY_Rate B
    Sleep, 333
    Sleep, 300  ; JPY_Rate B
    Send, {F5}  ; JPY_Rate B
    Sleep, 100
    Sleep, 300  ; JPY_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, K8{Enter}  ; JPY_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Rate B
    Send, {Control Down}{Down}{Control Up}  ; JPY_Rate B
    Sleep, 100
    Sleep, 300  ; JPY_Rate B
    Send, {Down}  ; JPY_Rate B
    Sleep, 100
    Sleep, 300  ; JPY_Rate B
    Loop, 3
    {
        Send, {Left}  ; JPY_Rate B
        Sleep, 100
    }
    Sleep, 300  ; JPY_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; JPY_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; JPY_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{c}{Control Up}{Control Down}{Left}{Down}{Control Up}{Right 2}  ; JPY_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Shift Down}{Up}{Shift Up}{v}{Control Up}  ; JPY_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Rate B
    Send, {F5}  ; JPY_Round Rumus
    Sleep, 100
    Sleep, 300  ; JPY_Round Rumus
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, K8{Enter}  ; JPY_Round Rumus
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Round Rumus
    Loop, 2
    {
        Send, {Control Down}{Down}{Control Up}  ; JPY_Round Rumus
        Sleep, 100
    }
    Sleep, 300  ; JPY_Round Rumus
    Send, {Shift Down}{Right}{Shift Up}  ; JPY_Round Rumus
    Sleep, 100
    Sleep, 300  ; JPY_Round Rumus
    Send, {Control Down}{c}{Control Up}  ; JPY_Round Rumus
    Sleep, 100
    Sleep, 300  ; JPY_Round Rumus
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Up}{Control Up}{Down}  ; JPY_Round Rumus
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Round Rumus
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Left}{Down}{Control Up}{Right 3}{Shift Down}{Control Down}{Up}{Control Up}{Down}{Shift Up}  ; JPY_Round Rumus
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; JPY_Round Rumus
    Send, {Control Down}{v}{Control Up}  ; JPY_Round Rumus
    Sleep, 100
    Sleep, 300  ; JPY_Round Rumus
    MsgBox, 0, , JPY Finish  ; IDR
    WinActivate, 1. Cashbill.xlsx - Excel  ; Set IDR to Cash Bill
    Sleep, 333
    Sleep, 1000  ; Set IDR to Cash Bill
    Send, {F5}  ; Set IDR to Cash Bill
    Sleep, 100
    Sleep, 300  ; Set IDR to Cash Bill
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, I1{Enter}  ; Set IDR to Cash Bill
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Set IDR to Cash Bill
    Send, {Alt Down}{Down}{Alt Up}  ; Set IDR to Cash Bill
    Sleep, 100
    Sleep, 300  ; Set IDR to Cash Bill
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 7}  ; Set IDR to Cash Bill
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Set IDR to Cash Bill
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, RP{Enter}  ; Set IDR to Cash Bill
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; Set IDR to Cash Bill
    WinActivate, 1. Cashbill.xlsx - Excel  ; RP_Invoice Date - Desc
    Sleep, 333
    Sleep, 300  ; RP_Invoice Date - Desc
    Send, {Control Down}{Home}{Control Up}  ; RP_Invoice Date - Desc
    Sleep, 100
    Sleep, 300  ; RP_Invoice Date - Desc
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Right 2}{Control Down}{Down}{Control Up}{Shift Up}  ; RP_Invoice Date - Desc
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Invoice Date - Desc
    Send, {Control Down}{c}{Control Up}  ; RP_Invoice Date - Desc
    Sleep, 300
    Sleep, 300  ; RP_Invoice Date - Desc
    WinActivate, 5. EXC.RATE.xlsx - Excel  ; RP_Invoice Date - Desc
    Sleep, 333
    Sleep, 1000  ; RP_Invoice Date - Desc
    Send, {F5}  ; RP_Invoice Date - Desc
    Sleep, 100
    Sleep, 300  ; RP_Invoice Date - Desc
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, B8{Enter}  ; RP_Invoice Date - Desc
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Invoice Date - Desc
    Loop, 2
    {
        Send, {Control Down}{Down}{Control Up}  ; RP_Invoice Date - Desc
        Sleep, 100
    }
    Sleep, 300  ; RP_Invoice Date - Desc
    Loop, 2
    {
        Send, {Down}  ; RP_Invoice Date - Desc
        Sleep, 100
    }
    Sleep, 300  ; RP_Invoice Date - Desc
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; RP_Invoice Date - Desc
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Invoice Date - Desc
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; RP_Invoice Date - Desc
    SetKeyDelay, %CurrentKeyDelay%
    WinActivate, 1. Cashbill.xlsx - Excel  ; RP_Debit Account
    Sleep, 333
    Sleep, 300  ; RP_Debit Account
    Send, {F5}  ; RP_Debit Account
    Sleep, 100
    Sleep, 300  ; RP_Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, G1{Enter}  ; RP_Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; RP_Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Debit Account
    Send, {Control Down}{c}{Control Up}  ; RP_Debit Account
    Sleep, 300
    Sleep, 300  ; RP_Debit Account
    WinActivate, 5. EXC.RATE.xlsx - Excel  ; RP_Debit Account
    Sleep, 333
    Sleep, 1000  ; RP_Debit Account
    Send, {F5}  ; RP_Debit Account
    Sleep, 100
    Sleep, 300  ; RP_Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, F8{Enter}  ; RP_Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Debit Account
    Loop, 2
    {
        Send, {Control Down}{Down}{Control Up}  ; RP_Debit Account
        Sleep, 100
    }
    Sleep, 300  ; RP_Debit Account
    Loop, 2
    {
        Send, {Down}  ; RP_Debit Account
        Sleep, 100
    }
    Sleep, 300  ; RP_Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; RP_Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Debit Account
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; RP_Debit Account
    SetKeyDelay, %CurrentKeyDelay%
    WinActivate, 1. Cashbill.xlsx - Excel  ; RP_Amount
    Sleep, 333
    Sleep, 300  ; RP_Amount
    Send, {F5}  ; RP_Amount
    Sleep, 100
    Sleep, 300  ; RP_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, J1{Enter}  ; RP_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down}{Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; RP_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Amount
    Send, {Control Down}{c}{Control Up}  ; RP_Amount
    Sleep, 300
    Sleep, 300  ; RP_Amount
    WinActivate, 5. EXC.RATE.xlsx - Excel  ; RP_Amount
    Sleep, 333
    Sleep, 1000  ; RP_Amount
    Send, {F5}  ; RP_Amount
    Sleep, 100
    Sleep, 300  ; RP_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, B8{Enter}  ; RP_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Amount
    Loop, 2
    {
        Send, {Control Down}{Down}{Control Up}  ; RP_Amount
        Sleep, 100
    }
    Sleep, 300  ; RP_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Control Down}{Right}{Control Up}{Right 9}  ; RP_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; RP_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Amount
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; RP_Amount
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Amount
    WinActivate, 4. ACCRUED.xlsx - Excel  ; RP_End Month Rate IDR
    Sleep, 333
    Send, {Control Down}{PgUp}{Control Up}  ; RP_End Month Rate IDR
    Sleep, 100
    Sleep, 300  ; RP_End Month Rate IDR
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Home}{Down}{Control Up}{Right}  ; RP_End Month Rate IDR
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_End Month Rate IDR
    Send, {Control Down}{c}{Control Up}  ; RP_End Month Rate IDR
    Sleep, 100
    Sleep, 300  ; RP_End Month Rate IDR
    Send, {Control Down}{PgDn}{Control Up}  ; RP_End Month Rate IDR
    Sleep, 100
    Sleep, 300  ; RP_End Month Rate IDR
    WinActivate, 5. EXC.RATE.xlsx - Excel  ; RP_End Month Rate IDR
    Sleep, 333
    Sleep, 300  ; RP_End Month Rate IDR
    Send, {F5}  ; RP_End Month Rate IDR
    Sleep, 100
    Sleep, 300  ; RP_End Month Rate IDR
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, B8{Enter}  ; RP_End Month Rate IDR
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_End Month Rate IDR
    Loop, 3
    {
        Send, {Control Down}{Down}{Control Up}  ; RP_End Month Rate IDR
        Sleep, 100
    }
    Sleep, 300  ; RP_End Month Rate IDR
    Loop, 7
    {
        Send, {Right}  ; RP_End Month Rate IDR
        Sleep, 100
    }
    Sleep, 300  ; RP_End Month Rate IDR
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; RP_End Month Rate IDR
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_End Month Rate IDR
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; RP_End Month Rate IDR
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_End Month Rate IDR
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{c}{Control Up}{Control Down}{Left}{Down}{Control Up}{Right 3}  ; RP_End Month Rate IDR
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_End Month Rate IDR
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Shift Down}{Up}{Shift Up}{v}{Control Up}  ; RP_End Month Rate IDR
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_End Month Rate IDR
    WinActivate, 4. ACCRUED.xlsx - Excel  ; RP_Rate B
    Sleep, 333
    Send, {Control Down}{PgUp}{Control Up}  ; RP_Rate B
    Sleep, 100
    Sleep, 300  ; RP_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Home}{Down}{Control Up}{Right}  ; RP_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Rate B
    Send, {Control Down}{c}{Control Up}  ; RP_Rate B
    Sleep, 100
    Sleep, 300  ; RP_Rate B
    Send, {Control Down}{PgDn}{Control Up}  ; RP_Rate B
    Sleep, 100
    Sleep, 300  ; RP_Rate B
    WinActivate, 5. EXC.RATE.xlsx - Excel  ; RP_Rate B
    Sleep, 333
    Sleep, 300  ; RP_Rate B
    Send, {F5}  ; RP_Rate B
    Sleep, 100
    Sleep, 300  ; RP_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, B8{Enter}  ; RP_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Rate B
    Loop, 3
    {
        Send, {Control Down}{Down}{Control Up}  ; RP_Rate B
        Sleep, 100
    }
    Sleep, 300  ; RP_Rate B
    Loop, 6
    {
        Send, {Right}  ; RP_Rate B
        Sleep, 100
    }
    Sleep, 300  ; RP_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 300
    SendEvent, {Alt}{h}{v}  ; RP_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Enter}  ; RP_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{c}{Control Up}{Control Down}{Left}{Down}{Control Up}{Right 2}  ; RP_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Rate B
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Shift Down}{Up}{Shift Up}{v}{Control Up}  ; RP_Rate B
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Rate B
    /*
    Send, {F5}  ; RP_Rumus Round
    Sleep, 100
    Sleep, 300  ; RP_Rumus Round
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, B8{Enter}  ; RP_Rumus Round
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Rumus Round
    Loop, 2
    {
        Send, {Control Down}{Down}{Control Up}  ; RP_Rumus Round
        Sleep, 100
    }
    Sleep, 300  ; RP_Rumus Round
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Down 2}{Control Down}{Right}{Control Up}{Right 5}  ; RP_Rumus Round
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Rumus Round
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, =ROUND(({Right 2}/{Left 2})`,2){Enter}  ; RP_Rumus Round
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Rumus Round
    Send, {Up}  ; RP_Rumus Round
    Sleep, 100
    Sleep, 300  ; RP_Rumus Round
    Send, {Control Down}{c}{Control Up}  ; RP_Rumus Round
    Sleep, 100
    Sleep, 300  ; RP_Rumus Round
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Control Down}{Left}{Down}{Control Up}  ; RP_Rumus Round
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Rumus Round
    CurrentKeyDelay := A_KeyDelay
    SetKeyDelay, 100
    SendEvent, {Right 2}{Control Down}{Shift Down}{Up}{Shift Up}{Control Up}  ; RP_Rumus Round
    SetKeyDelay, %CurrentKeyDelay%
    Sleep, 300  ; RP_Rumus Round
    Send, {Control Down}{v}{Control Up}  ; RP_Rumus Round
    Sleep, 100
    */
    Sleep, 300  ; RP_Rumus Round
    MsgBox, 0, , IDR Finish, 1  ; IDR
    return return
}

F6::
TestMacro:
WinActivate, 1. Cashbill.xlsx - Excel  ; Fill USD
Sleep, 333
Sleep, 300
Send, {Control Down}{Home}{Control Up}
Sleep, 300
Send, {Alt}{a}{t}
Sleep, 1000  ; Fill USD
Send, {F5}  ; Delete Slip Number
Sleep, 100
Sleep, 300  ; Delete Slip Number
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, C1{Enter}  ; Delete Slip Number
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Delete Slip Number
Send, {Control Down}{Space}{Control Up}  ; Delete Slip Number
Sleep, 300  ; Delete Slip Number
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {AppsKey}{d}  ; Delete Slip Number
SetKeyDelay, %CurrentKeyDelay%
Send, {F5}  ; Fill USD
Sleep, 100
Sleep, 300  ; Fill USD
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, I1{Enter}  ; Fill USD
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Fill USD
Send, {Alt Down}{Down}{Alt Up}  ; Fill USD
Sleep, 100
Sleep, 300  ; Fill USD
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down 7}  ; Fill USD
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Fill USD
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, USD{Enter}  ; Fill USD
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Fill USD
Send, {Control Down}{Home}{Control Up}  ; USD MODEL
Sleep, 100
Sleep, 300  ; USD Invoice Date
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down}{Shift Down}{Right 2}{Control Down}{Down}{Control Up}{Shift Up}  ; USD Invoice Date
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; USD Invoice Date
Send, {Control Down}{c}{Control Up}  ; USD Invoice Date
Sleep, 300
Sleep, 300  ; USD Invoice Date
WinActivate, 5. EXC.RATE.xlsx - Excel  ; USD Invoice Date
Sleep, 333
Sleep, 1000  ; USD Invoice Date
Send, {F5}  ; USD Invoice Date
Sleep, 100
Sleep, 100  ; USD Invoice Date
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, B8{Enter}  ; USD Invoice Date
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; USD Invoice Date
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Alt}{h}{v}  ; USD Invoice Date
SetKeyDelay, %CurrentKeyDelay%
Sleep, 100  ; USD Invoice Date
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down 2}{Enter}  ; USD_CC - Debit Account
SetKeyDelay, %CurrentKeyDelay%
WinActivate, 1. Cashbill.xlsx - Excel  ; USD_Debit Account
Sleep, 333
Sleep, 300  ; USD_Debit Account
Send, {F5}  ; USD_Debit Account
Sleep, 100
Sleep, 300  ; USD_Debit Account
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, G1{Enter}  ; USD_Debit Account
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; USD_Debit Account
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down}{Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; USD_Debit Account
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; USD_Debit Account
Send, {Control Down}{c}{Control Up}  ; USD_Debit Account
Sleep, 300
Sleep, 300  ; USD_Debit Account
WinActivate, 5. EXC.RATE.xlsx - Excel  ; USD_Debit Account
Sleep, 333
Sleep, 1000  ; USD_Debit Account
Send, {F5}  ; USD_Debit Account
Sleep, 100
Sleep, 300  ; USD_Debit Account
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, F8{Enter}  ; USD_Debit Account
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; USD_Debit Account
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Alt}{h}{v}  ; USD_Debit Account
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; USD_Debit Account
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down 2}{Enter}  ; USD_Debit Account
SetKeyDelay, %CurrentKeyDelay%
WinActivate, 1. Cashbill.xlsx - Excel  ; Fill USD
Sleep, 333
Sleep, 300  ; USD_Amount
Send, {F5}  ; USD_Amount
Sleep, 100
Sleep, 300  ; USD_Amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, J1{Enter}  ; USD_Amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; USD_Amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down}{Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; USD_Amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; USD_Amount
Send, {Control Down}{c}{Control Up}  ; USD_Amount
Sleep, 300
Sleep, 300  ; USD_Amount
WinActivate, 5. EXC.RATE.xlsx - Excel  ; USD_Amount
Sleep, 333
Sleep, 1000  ; USD_Amount
Send, {F5}  ; USD_Amount
Sleep, 100
Sleep, 300  ; USD_Amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, K8{Enter}  ; USD_Amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; USD_Amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Alt}{h}{v}  ; USD_Amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; USD_Amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down 2}{Enter}  ; USD_Amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; USD_Amount
MsgBox, 0, , 
(LTrim
USD Selesai

Next >> SGD
)
WinActivate, 1. Cashbill.xlsx - Excel  ; Fill SGD
Sleep, 333
Sleep, 1000  ; Fill SGD
Send, {F5}  ; Fill SGD
Sleep, 100
Sleep, 300  ; Fill SGD
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, I1{Enter}  ; Fill SGD
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Fill SGD
Send, {Alt Down}{Down}{Alt Up}  ; Fill SGD
Sleep, 100
Sleep, 300  ; Fill SGD
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down 7}  ; Fill SGD
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Fill SGD
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, SGD{Enter}  ; Fill SGD
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Fill SGD
WinActivate, 1. Cashbill.xlsx - Excel  ; SGD_Invoice Date - Supp Name
Sleep, 333
Sleep, 300  ; SGD_Invoice Date - Supp Name
Send, {Control Down}{Home}{Control Up}  ; SGD_Invoice Date - Supp Name
Sleep, 100
Sleep, 300  ; SGD_Invoice Date - Supp Name
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down}{Shift Down}{Right 2}{Control Down}{Down}{Control Up}{Shift Up}  ; SGD_Invoice Date - Supp Name
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; SGD_Invoice Date - Supp Name
Send, {Control Down}{c}{Control Up}  ; SGD_Invoice Date - Supp Name
Sleep, 300
Sleep, 300  ; SGD_Invoice Date - Supp Name
WinActivate, 5. EXC.RATE.xlsx - Excel  ; SGD_Invoice Date - Supp Name
Sleep, 333
Sleep, 1000  ; SGD_Invoice Date - Supp Name
Send, {F5}  ; SGD_Invoice Date - Supp Name
Sleep, 100
Sleep, 100  ; SGD_Invoice Date - Supp Name
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, B8{Enter}  ; SGD_Invoice Date - Supp Name
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; SGD_Invoice Date - Supp Name
Send, {Control Down}{Down}{Control Up}  ; SGD_Invoice Date - Supp Name
Sleep, 300  ; SGD_Model - Desc
Send, {Down}
Sleep, 100
Sleep, 300  ; SGD_Invoice Date - Supp Name
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Alt}{h}{v}  ; SGD_Invoice Date - Supp Name
SetKeyDelay, %CurrentKeyDelay%
Sleep, 100  ; SGD_Invoice Date - Supp Name
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down 2}{Enter}  ; SGD_Invoice Date - Supp Name
SetKeyDelay, %CurrentKeyDelay%
WinActivate, 1. Cashbill.xlsx - Excel  ; SGD_Debit Account
Sleep, 333
Sleep, 300  ; SGD_Debit Account
Send, {F5}  ; SGD_Debit Account
Sleep, 100
Sleep, 300  ; SGD_Debit Account
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, G1{Enter}  ; SGD_Debit Account
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; SGD_Debit Account
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down}{Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; SGD_Debit Account
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; SGD_Debit Account
Send, {Control Down}{c}{Control Up}  ; SGD_Debit Account
Sleep, 300
Sleep, 300  ; SGD_Debit Account
WinActivate, 5. EXC.RATE.xlsx - Excel  ; SGD_Debit Account
Sleep, 333
Sleep, 1000  ; SGD_Debit Account
Send, {F5}  ; SGD_Debit Account
Sleep, 100
Sleep, 300  ; SGD_Debit Account
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, F8{Enter}  ; SGD_Debit Account
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; SGD_Debit Account
Send, {Control Down}{Down}{Control Up}  ; SGD_Debit Account
Sleep, 300  ; SGD_Debit Account
Send, {Down}  ; SGD_Debit Account
Sleep, 100
Sleep, 300  ; SGD_Debit Account
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Alt}{h}{v}  ; SGD_Debit Account
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; SGD_Debit Account
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down 2}{Enter}  ; SGD_Debit Account
SetKeyDelay, %CurrentKeyDelay%
WinActivate, 1. Cashbill.xlsx - Excel  ; SGD_Amount
Sleep, 333
Sleep, 300  ; SGD_Amount
Send, {F5}  ; SGD_Amount
Sleep, 100
Sleep, 300  ; SGD_Amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, J1{Enter}  ; SGD_Amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; SGD_Amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down}{Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; SGD_Amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; SGD_Amount
Send, {Control Down}{c}{Control Up}  ; SGD_Amount
Sleep, 300
Sleep, 300  ; SGD_Amount
WinActivate, 5. EXC.RATE.xlsx - Excel  ; SGD_Amount
Sleep, 333
Sleep, 1000  ; SGD_Amount
Send, {F5}  ; SGD_Amount
Sleep, 100
Sleep, 300  ; SGD_Amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, K8{Enter}  ; SGD_Amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; SGD_Amount
Send, {Control Down}{Down}{Control Up}  ; SGD_Amount
Sleep, 300  ; SGD_Amount
Send, {Down}  ; SGD_Amount
Sleep, 100
Sleep, 300  ; SGD_Amount
Loop, 2
{
    Send, {Right}  ; SGD_Amount
    Sleep, 100
}
Sleep, 300  ; SGD_Amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Alt}{h}{v}  ; SGD_Amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; SGD_Amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down 2}{Enter}  ; SGD_Amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; SGD_Amount
WinActivate, 4. ACCRUED.xlsx - Excel  ; SGD_Daily Rate
Sleep, 333
Send, {Control Down}{PgUp}{Control Up}  ; SGD_Rate B
Sleep, 100
Sleep, 300  ; SGD_Rate B
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Home}{Down}{Control Up}{Right 2}  ; SGD_Rate B
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; SGD_Rate B
Send, {Control Down}{c}{Control Up}  ; SGD_Rate B
Sleep, 100
Sleep, 300  ; SGD_Rate B
Send, {Control Down}{PgDn}{Control Up}  ; SGD_Rate B
Sleep, 100
Sleep, 300  ; SGD_Rate B
WinActivate, 5. EXC.RATE.xlsx - Excel  ; SGD_Daily Rate
Sleep, 333
Sleep, 300  ; SGD_Rate B
Send, {F5}  ; SGD_Rate B
Sleep, 100
Sleep, 300  ; SGD_Rate B
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, K8{Enter}  ; SGD_Rate B
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; SGD_Rate B
Send, {Control Down}{Down}{Control Up}  ; SGD_Rate B
Sleep, 100
Sleep, 300  ; SGD_Rate B
Send, {Down}
Sleep, 100
Sleep, 300  ; SGD_Rate B
Loop, 4
{
    Send, {Left}
    Sleep, 100
}
Sleep, 300  ; SGD_Rate B
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Alt}{h}{v}  ; SGD_Rate B
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; SGD_Rate B
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down 2}{Enter}  ; SGD_Rate B
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; SGD_Rate B
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{c}{Control Up}{Control Down}{Left}{Down}{Control Up}{Right}  ; SGD_Rate B
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; SGD_Rate B
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Shift Down}{Up}{Shift Up}{v}{Control Up}  ; SGD_Rate B
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; SGD_Rate B
Send, {F5}  ; SGD_Round Rumus
Sleep, 100
Sleep, 300  ; SGD_Round Rumus
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, K8{Enter}  ; SGD_Round Rumus
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; SGD_Round Rumus
Send, {Control Down}{Down}{Control Up}  ; SGD_Round Rumus
Sleep, 100
Sleep, 300  ; SGD_Round Rumus
Send, {Down}  ; SGD_Round Rumus
Sleep, 100
Sleep, 300  ; SGD_Round Rumus
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, =ROUND(({Right 2}/{Left 4})`,2){Enter}  ; SGD_Round Rumus
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; SGD_Round Rumus
Send, {Up}  ; SGD_Round Rumus
Sleep, 100
Sleep, 300  ; SGD_Round Rumus
Send, {Control Down}{c}{Control Up}  ; SGD_Round Rumus
Sleep, 300  ; SGD_Round Rumus
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Left}{Down}{Control Up}{Right 4}{Shift Down}{Control Down}{Up}{Control Up}{Down}{Shift Up}  ; SGD_Round Rumus
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; SGD_Round Rumus
Send, {Control Down}{v}{Control Up}  ; SGD_Round Rumus
Sleep, 100
Sleep, 300  ; SGD_Round Rumus
WinActivate, 1. Cashbill.xlsx - Excel  ; Fill JPY
Sleep, 333
Sleep, 1000  ; Fill JPY
Send, {F5}  ; Fill JPY
Sleep, 100
Sleep, 300  ; Fill JPY
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, I1{Enter}  ; Fill JPY
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Fill JPY
Send, {Alt Down}{Down}{Alt Up}  ; Fill JPY
Sleep, 100
Sleep, 300  ; Fill JPY
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down 7}  ; Fill JPY
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Fill JPY
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, JPY{Enter}  ; Fill JPY
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Fill JPY
WinActivate, 1. Cashbill.xlsx - Excel  ; JPY_Inv date - Supp Name
Sleep, 333
Sleep, 300  ; JPY_Inv date - Supp Name
Send, {Control Down}{Home}{Control Up}  ; JPY_Inv date - Supp Name
Sleep, 100
Sleep, 300  ; JPY_Inv date - Supp Name
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down}{Shift Down}{Right 2}{Control Down}{Down}{Control Up}{Shift Up}  ; JPY_Inv date - Supp Name
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; JPY_Inv date - Supp Name
Send, {Control Down}{c}{Control Up}  ; JPY_Inv date - Supp Name
Sleep, 300
Sleep, 300  ; JPY_Inv date - Supp Name
WinActivate, 5. EXC.RATE.xlsx - Excel  ; JPY_Inv date - Supp Name
Sleep, 333
Sleep, 1000  ; JPY_Inv date - Supp Name
Send, {F5}  ; JPY_Inv date - Supp Name
Sleep, 100
Sleep, 100  ; JPY_Inv date - Supp Name
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, B8{Enter}  ; JPY_Inv date - Supp Name
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; JPY_Inv date - Supp Name
Send, {Control Down}{Down}{Control Up}  ; JPY_Inv date - Supp Name
Sleep, 100
Sleep, 300  ; JPY_Inv date - Supp Name
Send, {Down}  ; JPY_Inv date - Supp Name
Sleep, 100
Sleep, 300  ; JPY_Inv date - Supp Name
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Alt}{h}{v}  ; JPY_Inv date - Supp Name
SetKeyDelay, %CurrentKeyDelay%
Sleep, 100  ; JPY_Inv date - Supp Name
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down 2}{Enter}  ; JPY_Inv date - Supp Name
SetKeyDelay, %CurrentKeyDelay%
WinActivate, 1. Cashbill.xlsx - Excel  ; JPY_Debit Account
Sleep, 333
Sleep, 300  ; JPY_Debit Account
Send, {F5}  ; JPY_Debit Account
Sleep, 100
Sleep, 300  ; JPY_Debit Account
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, G1{Enter}  ; JPY_Debit Account
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; JPY_Debit Account
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down}{Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; JPY_Debit Account
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; JPY_Debit Account
Send, {Control Down}{c}{Control Up}  ; JPY_Debit Account
Sleep, 300
Sleep, 300  ; JPY_Debit Account
WinActivate, 5. EXC.RATE.xlsx - Excel  ; JPY_Debit Account
Sleep, 333
Sleep, 1000  ; JPY_Debit Account
Send, {F5}  ; JPY_Debit Account
Sleep, 100
Sleep, 300  ; JPY_Debit Account
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, F8{Enter}  ; JPY_Debit Account
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; JPY_Debit Account
Send, {Control Down}{Down}{Control Up}  ; JPY_Debit Account
Sleep, 300  ; JPY_Debit Account
Send, {Down}  ; JPY_Debit Account
Sleep, 100
Sleep, 300  ; JPY_Debit Account
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Alt}{h}{v}  ; JPY_Debit Account
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; JPY_Debit Account
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down 2}{Enter}  ; JPY_Debit Account
SetKeyDelay, %CurrentKeyDelay%
WinActivate, 1. Cashbill.xlsx - Excel  ; JPY_Amount
Sleep, 333
Sleep, 300  ; JPY_Amount
Send, {F5}  ; JPY_Amount
Sleep, 100
Sleep, 300  ; JPY_Amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, J1{Enter}  ; JPY_Amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; JPY_Amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down}{Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; JPY_Amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; JPY_Amount
Send, {Control Down}{c}{Control Up}  ; JPY_Amount
Sleep, 300
Sleep, 300  ; JPY_Amount
WinActivate, 5. EXC.RATE.xlsx - Excel  ; JPY_Amount
Sleep, 333
Sleep, 1000  ; JPY_Amount
Send, {F5}  ; JPY_Amount
Sleep, 100
Sleep, 300  ; JPY_Amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, K8{Enter}  ; JPY_Amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; JPY_Amount
Send, {Control Down}{Down}{Control Up}  ; JPY_Amount
Sleep, 300  ; JPY_Amount
Send, {Down}  ; JPY_Amount
Sleep, 100
Sleep, 300  ; JPY_Amount
Loop, 2
{
    Send, {Right}  ; JPY_Amount
    Sleep, 100
}
Sleep, 300  ; JPY_Amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Alt}{h}{v}  ; JPY_Amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; JPY_Amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down 2}{Enter}  ; JPY_Amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; JPY_Amount
WinActivate, 4. ACCRUED.xlsx - Excel  ; JPY_Rate B
Sleep, 333
Send, {Control Down}{PgUp}{Control Up}  ; JPY_Rate B
Sleep, 100
Sleep, 300  ; JPY_Rate B
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Home}{Down}{Control Up}{Right 3}  ; JPY_Rate B
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; JPY_Rate B
Send, {Control Down}{c}{Control Up}  ; JPY_Rate B
Sleep, 100
Sleep, 300  ; JPY_Rate B
Send, {Control Down}{PgDn}{Control Up}  ; JPY_Rate B
Sleep, 100
Sleep, 300  ; JPY_Rate B
WinActivate, 5. EXC.RATE.xlsx - Excel  ; JPY_Rate B
Sleep, 333
Sleep, 300  ; JPY_Rate B
Send, {F5}  ; JPY_Rate B
Sleep, 100
Sleep, 300  ; JPY_Rate B
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, K8{Enter}  ; JPY_Rate B
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; JPY_Rate B
Send, {Control Down}{Down}{Control Up}  ; JPY_Rate B
Sleep, 100
Sleep, 300  ; JPY_Rate B
Send, {Down}  ; JPY_Rate B
Sleep, 100
Sleep, 300  ; JPY_Rate B
Loop, 3
{
    Send, {Left}  ; JPY_Rate B
    Sleep, 100
}
Sleep, 300  ; JPY_Rate B
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Alt}{h}{v}  ; JPY_Rate B
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; JPY_Rate B
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down 2}{Enter}  ; JPY_Rate B
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; JPY_Rate B
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{c}{Control Up}{Control Down}{Left}{Down}{Control Up}{Right 2}  ; JPY_Rate B
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; JPY_Rate B
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Shift Down}{Up}{Shift Up}{v}{Control Up}  ; JPY_Rate B
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; JPY_Rate B
Send, {F5}  ; JPY_Round Rumus
Sleep, 100
Sleep, 300  ; JPY_Round Rumus
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, K8{Enter}  ; JPY_Round Rumus
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; JPY_Round Rumus
Loop, 2
{
    Send, {Control Down}{Down}{Control Up}  ; JPY_Round Rumus
    Sleep, 100
}
Sleep, 300  ; JPY_Round Rumus
Send, {Shift Down}{Right}{Shift Up}  ; JPY_Round Rumus
Sleep, 100
Sleep, 300  ; JPY_Round Rumus
Send, {Control Down}{c}{Control Up}  ; JPY_Round Rumus
Sleep, 100
Sleep, 300  ; JPY_Round Rumus
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Up}{Control Up}{Down}  ; JPY_Round Rumus
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; JPY_Round Rumus
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Left}{Down}{Control Up}{Right 3}{Shift Down}{Control Down}{Up}{Control Up}{Down}{Shift Up}  ; JPY_Round Rumus
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; JPY_Round Rumus
Send, {Control Down}{v}{Control Up}  ; JPY_Round Rumus
Sleep, 100
Sleep, 300  ; JPY_Round Rumus
MsgBox, 0, , JPY Finish  ; IDR
WinActivate, 1. Cashbill.xlsx - Excel  ; Set IDR to Cash Bill
Sleep, 333
Sleep, 1000  ; Set IDR to Cash Bill
Send, {F5}  ; Set IDR to Cash Bill
Sleep, 100
Sleep, 300  ; Set IDR to Cash Bill
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, I1{Enter}  ; Set IDR to Cash Bill
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Set IDR to Cash Bill
Send, {Alt Down}{Down}{Alt Up}  ; Set IDR to Cash Bill
Sleep, 100
Sleep, 300  ; Set IDR to Cash Bill
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down 7}  ; Set IDR to Cash Bill
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Set IDR to Cash Bill
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, RP{Enter}  ; Set IDR to Cash Bill
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; Set IDR to Cash Bill
WinActivate, 1. Cashbill.xlsx - Excel  ; RP_Invoice Date - Desc
Sleep, 333
Sleep, 300  ; RP_Invoice Date - Desc
Send, {Control Down}{Home}{Control Up}  ; RP_Invoice Date - Desc
Sleep, 100
Sleep, 300  ; RP_Invoice Date - Desc
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down}{Shift Down}{Right 2}{Control Down}{Down}{Control Up}{Shift Up}  ; RP_Invoice Date - Desc
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_Invoice Date - Desc
Send, {Control Down}{c}{Control Up}  ; RP_Invoice Date - Desc
Sleep, 300
Sleep, 300  ; RP_Invoice Date - Desc
WinActivate, 5. EXC.RATE.xlsx - Excel  ; RP_Invoice Date - Desc
Sleep, 333
Sleep, 1000  ; RP_Invoice Date - Desc
Send, {F5}  ; RP_Invoice Date - Desc
Sleep, 100
Sleep, 300  ; RP_Invoice Date - Desc
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, B8{Enter}  ; RP_Invoice Date - Desc
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_Invoice Date - Desc
Loop, 2
{
    Send, {Control Down}{Down}{Control Up}  ; RP_Invoice Date - Desc
    Sleep, 100
}
Sleep, 300  ; RP_Invoice Date - Desc
Loop, 2
{
    Send, {Down}  ; RP_Invoice Date - Desc
    Sleep, 100
}
Sleep, 300  ; RP_Invoice Date - Desc
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Alt}{h}{v}  ; RP_Invoice Date - Desc
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_Invoice Date - Desc
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down 2}{Enter}  ; RP_Invoice Date - Desc
SetKeyDelay, %CurrentKeyDelay%
WinActivate, 1. Cashbill.xlsx - Excel  ; RP_Debit Account
Sleep, 333
Sleep, 300  ; RP_Debit Account
Send, {F5}  ; RP_Debit Account
Sleep, 100
Sleep, 300  ; RP_Debit Account
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, G1{Enter}  ; RP_Debit Account
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_Debit Account
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down}{Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; RP_Debit Account
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_Debit Account
Send, {Control Down}{c}{Control Up}  ; RP_Debit Account
Sleep, 300
Sleep, 300  ; RP_Debit Account
WinActivate, 5. EXC.RATE.xlsx - Excel  ; RP_Debit Account
Sleep, 333
Sleep, 1000  ; RP_Debit Account
Send, {F5}  ; RP_Debit Account
Sleep, 100
Sleep, 300  ; RP_Debit Account
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, F8{Enter}  ; RP_Debit Account
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_Debit Account
Loop, 2
{
    Send, {Control Down}{Down}{Control Up}  ; RP_Debit Account
    Sleep, 100
}
Sleep, 300  ; RP_Debit Account
Loop, 2
{
    Send, {Down}  ; RP_Debit Account
    Sleep, 100
}
Sleep, 300  ; RP_Debit Account
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Alt}{h}{v}  ; RP_Debit Account
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_Debit Account
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down 2}{Enter}  ; RP_Debit Account
SetKeyDelay, %CurrentKeyDelay%
WinActivate, 1. Cashbill.xlsx - Excel  ; RP_Amount
Sleep, 333
Sleep, 300  ; RP_Amount
Send, {F5}  ; RP_Amount
Sleep, 100
Sleep, 300  ; RP_Amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, J1{Enter}  ; RP_Amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_Amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down}{Shift Down}{Control Down}{Down}{Control Up}{Shift Up}  ; RP_Amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_Amount
Send, {Control Down}{c}{Control Up}  ; RP_Amount
Sleep, 300
Sleep, 300  ; RP_Amount
WinActivate, 5. EXC.RATE.xlsx - Excel  ; RP_Amount
Sleep, 333
Sleep, 1000  ; RP_Amount
Send, {F5}  ; RP_Amount
Sleep, 100
Sleep, 300  ; RP_Amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, B8{Enter}  ; RP_Amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_Amount
Loop, 2
{
    Send, {Control Down}{Down}{Control Up}  ; RP_Amount
    Sleep, 100
}
Sleep, 300  ; RP_Amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down 2}{Control Down}{Right}{Control Up}{Right 9}  ; RP_Amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_Amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Alt}{h}{v}  ; RP_Amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_Amount
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down 2}{Enter}  ; RP_Amount
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_Amount
WinActivate, 4. ACCRUED.xlsx - Excel  ; RP_End Month Rate IDR
Sleep, 333
Send, {Control Down}{PgUp}{Control Up}  ; RP_End Month Rate IDR
Sleep, 100
Sleep, 300  ; RP_End Month Rate IDR
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Home}{Down}{Control Up}{Right}  ; RP_End Month Rate IDR
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_End Month Rate IDR
Send, {Control Down}{c}{Control Up}  ; RP_End Month Rate IDR
Sleep, 100
Sleep, 300  ; RP_End Month Rate IDR
Send, {Control Down}{PgDn}{Control Up}  ; RP_End Month Rate IDR
Sleep, 100
Sleep, 300  ; RP_End Month Rate IDR
WinActivate, 5. EXC.RATE.xlsx - Excel  ; RP_End Month Rate IDR
Sleep, 333
Sleep, 300  ; RP_End Month Rate IDR
Send, {F5}  ; RP_End Month Rate IDR
Sleep, 100
Sleep, 300  ; RP_End Month Rate IDR
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, B8{Enter}  ; RP_End Month Rate IDR
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_End Month Rate IDR
Loop, 3
{
    Send, {Control Down}{Down}{Control Up}  ; RP_End Month Rate IDR
    Sleep, 100
}
Sleep, 300  ; RP_End Month Rate IDR
Loop, 7
{
    Send, {Right}  ; RP_End Month Rate IDR
    Sleep, 100
}
Sleep, 300  ; RP_End Month Rate IDR
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Alt}{h}{v}  ; RP_End Month Rate IDR
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_End Month Rate IDR
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down 2}{Enter}  ; RP_End Month Rate IDR
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_End Month Rate IDR
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{c}{Control Up}{Control Down}{Left}{Down}{Control Up}{Right 3}  ; RP_End Month Rate IDR
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_End Month Rate IDR
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Shift Down}{Up}{Shift Up}{v}{Control Up}  ; RP_End Month Rate IDR
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_End Month Rate IDR
WinActivate, 4. ACCRUED.xlsx - Excel  ; RP_Rate B
Sleep, 333
Send, {Control Down}{PgUp}{Control Up}  ; RP_Rate B
Sleep, 100
Sleep, 300  ; RP_Rate B
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Home}{Down}{Control Up}{Right}  ; RP_Rate B
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_Rate B
Send, {Control Down}{c}{Control Up}  ; RP_Rate B
Sleep, 100
Sleep, 300  ; RP_Rate B
Send, {Control Down}{PgDn}{Control Up}  ; RP_Rate B
Sleep, 100
Sleep, 300  ; RP_Rate B
WinActivate, 5. EXC.RATE.xlsx - Excel  ; RP_Rate B
Sleep, 333
Sleep, 300  ; RP_Rate B
Send, {F5}  ; RP_Rate B
Sleep, 100
Sleep, 300  ; RP_Rate B
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, B8{Enter}  ; RP_Rate B
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_Rate B
Loop, 3
{
    Send, {Control Down}{Down}{Control Up}  ; RP_Rate B
    Sleep, 100
}
Sleep, 300  ; RP_Rate B
Loop, 6
{
    Send, {Right}  ; RP_Rate B
    Sleep, 100
}
Sleep, 300  ; RP_Rate B
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 300
SendEvent, {Alt}{h}{v}  ; RP_Rate B
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_Rate B
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down 2}{Enter}  ; RP_Rate B
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_Rate B
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{c}{Control Up}{Control Down}{Left}{Down}{Control Up}{Right 2}  ; RP_Rate B
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_Rate B
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Shift Down}{Up}{Shift Up}{v}{Control Up}  ; RP_Rate B
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_Rate B
Send, {F5}  ; RP_Rumus Round
Sleep, 100
Sleep, 300  ; RP_Rumus Round
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, B8{Enter}  ; RP_Rumus Round
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_Rumus Round
Loop, 2
{
    Send, {Control Down}{Down}{Control Up}  ; RP_Rumus Round
    Sleep, 100
}
Sleep, 300  ; RP_Rumus Round
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Down 2}{Control Down}{Right}{Control Up}{Right 7}  ; RP_Rumus Round
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_Rumus Round
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, =ROUND(({Right 2}/{Left 2}){`,}2){Enter}  ; RP_Rumus Round
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_Rumus Round
Send, {Up}  ; RP_Rumus Round
Sleep, 100
Sleep, 300  ; RP_Rumus Round
Send, {Control Down}{c}{Control Up}  ; RP_Rumus Round
Sleep, 100
Sleep, 300  ; RP_Rumus Round
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Control Down}{Left}{Down}{Control Up}  ; RP_Rumus Round
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_Rumus Round
CurrentKeyDelay := A_KeyDelay
SetKeyDelay, 100
SendEvent, {Right 2}{Control Down}{Shift Down}{Up}{Shift Up}{Control Up}  ; RP_Rumus Round
SetKeyDelay, %CurrentKeyDelay%
Sleep, 300  ; RP_Rumus Round
Send, {Control Down}{v}{Control Up}  ; RP_Rumus Round
Sleep, 100
Sleep, 300  ; RP_Rumus Round
MsgBox, 0, , IDR Finish  ; IDR
Return

