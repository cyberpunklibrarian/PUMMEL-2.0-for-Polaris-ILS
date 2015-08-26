; PUMMEL - Processing Unclaimed Materials Manually is Essentially Lame
; Daniel Messer - Web Content Manager
; Maricopa County Library District

; --------------------------------------------------------------------
; SET MATCH MODE FOR LATER PROCESSING

SetTitleMatchMode, 1
SetTitleMatchMode, Slow

; --------------------------------------------------------------------

; --------------------------------------------------------------------
; SET UP THE GUI AND VARIABLES

; Speed- How fast does PUMMEL process the unclaim?
; Gui, font, s11 bold, Arial
; Gui, Add, Text,, Speed
; Gui, font, s11 norm, Arial
; Gui, Add, Radio, vSpeedRadio, Slow
; Gui, Add, Radio,, Medium
; Gui, Add, Radio,, Fast

; Select Branch Library
Gui, font, s11 bold, Arial
Gui, Add, Text, ym, Select Branch
Gui, Add, DropDownList, vBranch, AD|AG|EM|FA|FH|GB|GO|GU|HH|LP|NV|NW|PE|QC|RO|SE|SC|WT

; Check in after processing- Give the option
Gui, font, s11 bold, Arial
Gui, Add, Text, ym, Check In After Processing
Gui, font, s11 norm, Arial
Gui, Add, Radio, vCheckInRadio, Yes
Gui, Add, Radio, checked, No

Gui, Add, Button, default ym, Submit

; Process from file- Interface to read froom a file full of barcodes
Gui, font, s11 bold, Arial
Gui, Add, Text, ym, Process From File
Gui, font, s11 norm, Arial
Gui, Add, Button, gOpenFile, Browse



Gui, Show, W610 H175, PUMMEL 2.0 Development Build
return


ButtonSubmit:
Gui, Submit, NoHide

ProcessSpeed := 1200

if CheckInRadio = 1
	CheckIn := 1
else if CheckInRadio = 2
	CheckIn := 0
else
	CheckIn := 0

if CheckIn = 1
	WillCheck := "Yes"
else
	WillCheck := "No"

if Branch = AD
	BranchSelect := 0
if Branch = AG
	BranchSelect := 1
if Branch = RO
	BranchSelect := 4
if Branch = EM
	BranchSelect := 5
if Branch = FA
	BranchSelect := 6
if Branch = FH
	BranchSelect := 7
if Branch = GB
	BranchSelect := 8
if Branch = GO
	BranchSelect := 9
if Branch = GU
	BranchSelect := 10
if Branch = HH
	BranchSelect := 11
if Branch = LP
	BranchSelect := 12
if Branch = NV
	BranchSelect := 13
if Branch = NW
	BranchSelect := 14
if Branch = PE
	BranchSelect := 16
if Branch = QC
	BranchSelect := 17
if Branch = SE
	BranchSelect := 18
if Branch = SC
	BranchSelect := 19
if Branch = WT
	BranchSelect := 21

MsgBox Settings saved. `nBranch Selected: %Branch%`nCheck In: %WillCheck%`nClick okay to set up Polaris.


; --------------------------------------------------------------------
; RUN AND LOGIN TO POLARIS

run, "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Polaris\4.1R2\Polaris ILS" && pause
Sleep, 5000
send {Tab}{Tab}{Enter}
Sleep, 1000
send ^l
Sleep, 1000
send +{Tab}**USERNAME**{Tab} ;INSERT YOUR POLARIS USER NAME HERE
Sleep, 1000
send **PASSWORD**{Enter} ;INSERT YOUR POLARIS USER PASSWORD HERE
Sleep, 2000
send {Down %BranchSelect%}
Sleep, 2000
send {Enter}
Sleep, 2000
send !ci
; --------------------------------------------------------------------


; --------------------------------------------------------------------
; UNCLAIMED PROCESSING WORKFLOW

ProcessUnclaim:
^!u:: ;Triggered with CTRL ATL U manually, but also will be called by subroutine
send {Enter}
Sleep, ProcessSpeed
send ^+{End}
send ^c
Sleep, ProcessSpeed
send {Tab}{Tab}{Enter}
Sleep, ProcessSpeed
send !va^c2.00{Tab}
Sleep, ProcessSpeed
send r
Sleep, 250
send r
Sleep, 250
send r
Sleep, 250
send r
Sleep, 250
send r
Sleep, ProcessSpeed
send {Tab}^v{Tab}{Tab}
send Processed by PUMMEL v2.0 build 1.{Tab}{Enter}
Sleep, ProcessSpeed
send !fc
	IfWinExist, Item Record 
		{
			WinActivate
			Sleep, ProcessSpeed
			send !{F4}
			Sleep, ProcessSpeed
			send !n
			return
		}
	else
		{
			MsgBox, Item record window not found!
			ExitApp
		}
; --------------------------------------------------------------------
; PROCESS FROM FILE SUBROUTINE


OpenFile:
FileSelectFile, SelectedFile, 3, , Open a file, Text Documents (*.txt)
Goto ProcFile


ProcFile:
WinActivate, Item Records - Barcode Find Tool
	
Loop, read, %SelectedFile%
	{
	Clipboard = %A_LoopReadLine%
	GoSub, PasteBarcode
		if ErrorLevel = 1
		break
	}

if CheckIn = 1
	clipboard = %SelectedFile%
	MsgBox, 1, Ready to check in?
		IfMsgBox, Ok
			GoTo, ProcCheckIn
		IfMsgBox, Cancel
			return
if CheckIn = 0
	return

PasteBarcode:
IfWinExist, Item Records - Barcode Find Tool
	{
	WinActivate
	Sleep, ProcessSpeed
	send %clipboard%
	send {Enter}
	Sleep, ProcessSpeed
	GoSub, ProcessUnclaim
	return
	}
		else MsgBox, Item record window not found!
	return

; --------------------------------------------------------------------	
; PROCESS CHECK IN SUBROUTINE

ProcCheckIn:
IfWinExist, Polaris ILS
	{
	WinActivate
	Sleep, ProcessSpeed
	send !wa
	Sleep, 1000
	send {Enter}
	Sleep, 1000
	send {F2}
	Sleep, 2000
	send ^!i
	send C%clipboard%
	send {enter}	
	}
		else MsgBox, Error


GuiClose:
ExitApp