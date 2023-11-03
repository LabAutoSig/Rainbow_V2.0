;=#GENERAL INFORMATION # ====================================================================================================================
; Name...........: Rainbow - Timer
; Description ...: Rainbow Mainscript opens the timer with a specific countdown of minutes
; Author ........: Maria Fernanda Avila Vazquez modified from wakillon https://www.autoitscript.com/forum/topic/127667-how-to-create-a-countdown-timer-in-autoit/
;Co-Author.......: Nicole Rupp, Jeaninne Opara
;Note............: To the users: please follow the instruction of ReadMe.pdf. Adjustment to your page is necessary wherever you see a "!A". Please make the changes in increasing numbering order.


	AutoItSetOption("WinTitleMatchMode", -2)
	#include <MsgBoxConstants.au3>
	#include <AutoItConstants.au3>
	#include <Excel.au3>
	#include <WindowsConstants.au3>
	#include <StaticConstants.au3>
	#include <Array.au3>
	#include <Clipboard.au3>
	#include <File.au3>
	#include <Word.au3>
	#include <Process.au3>

Global $vText = ClipGet()

Sleep(1000)
Global $sRainbow = ClipPut("Please set your timer: say number of minutes and the word minutes.")
_VoiceOutput()
Sleep(500)
_GoogleListen2()
_SeparateInputSpeechToText()
Sleep(500)
Local $TimerArray = StringSplit($vText," ")
Local $iMinutes = ($TimerArray[1])
Sleep(500)
If IsNumber($iMinutes) = 0 Then
	If $iMinutes = "one" Then
		$iReplaceNum = StringReplace($TimerArray[1],"one","1")
		$iNewNumber = Number($iReplaceNum)
		$iMinutes = $iNewNumber
	EndIf
EndIf

Global $MinutesInMs = $iMinutes * 60000
Global $_Wait1Minute = Int($MinutesInMs)
Sleep(500)
Local $sRainbow = ClipPut("You set a timer for "&$iMinutes&" minutes")
_VoiceOutput()
Sleep(500)
$_GuiCountDown = GUICreate("RCountdown...", 700, 200, @DesktopWidth / 2 - 250, @DesktopHeight / 2 - 100, -1, $WS_EX_TOPMOST) ;Start GUI
GUISetBkColor(0xFFFFFF)
$TimeLabel = GUICtrlCreateLabel("", 0, 0, 700, 200, $SS_CENTER)
GUICtrlSetFont(-1, 125, 800)
GUISetState() ;End GUI
Global $TimeTicks = TimerInit()
While 1
	_Check()
	Sleep(1000)
Wend

Func _Check()
Local $Lapsed = ($_Wait1Minute - TimerDiff($TimeTicks)) / 1000 ; seconds lapsed
Local $_Hours = Floor($Lapsed / 3600)  ;Returns a number rounded down to the closest integer.
Local $_Minutes = Mod(Floor($Lapsed / 60), 60) ;Returns the remainder of a division after one number is divided by another - in form of a rounded down to the closes integer.
Local $_Seconds = Mod($Lapsed, 60) ;Returns the remainder of a division after one number is divided by another

If($_Hours + $_Minutes + $_Seconds) < 0 Then
	GUISetBkColor(0xFF0000, $_GuiCountDown)
	GUICtrlSetData($TimeLabel, "END")
	Sleep(1000)
	GUIDelete()
	Exit
Else
	GUICtrlSetData($TimeLabel, StringFormat("%02u:%02u:%02u", $_Hours, $_Minutes, $_Seconds))
	If $_Hours = 0 And $_Minutes = 0 And $_Seconds <= 10 Then
	Beep(1200, 100)
	GUISetBkColor(0xC0C0C0, $_GuiCountDown)
	EndIf
EndIf
EndFunc


Func _VoiceOutput()
Sleep(500)
WinMove("Google Übersetzer - Google Chrome","",1258,424,658,614) ;Move window 											;!A1. Please insert the same numbers as !3 mainscript
Sleep(500)
WinActivate("Google Übersetzer - Google Chrome")
Sleep(500)
MouseClick("left",1866,698) ;X-button																					;!A2. Please insert the same numbers as !6 mainscript
Sleep(250)
Send("^v")
Sleep(250)
MouseClick("left",1582,706) ;Field																						;!A3. Please insert the same numbers as !8 mainscript
Sleep(500)
Do
	$error = PixelGetColor(1290,777) ;revise if there is a correction in Google Translate								;!A4. Please insert the same numbers as !19 mainscript
	Sleep(250)
	$iA = 2
Until $iA = 2
If $error = 1733608 Then ;Check for blue pixel
	WinActivate("Google Übersetzer - Google Chrome")
	Sleep(500)
	Send("{TAB}")
	Sleep(500) ;save time with less ms?
	Send("{TAB}")
	Sleep(500)
	Send("{ENTER}")  ;Accept correction
	Sleep(500)
	MouseClick("left",1579,698) ;Field																					;!A5. Please insert the same numbers as !8 mainscript
EndIf

Sleep(200)  ;Normal output
Local $CheckBefore = PixelChecksum(1269,627,1885,943)																	;!A6. Please insert the same numbers as !21 mainscript
Sleep(500)
Send("{TAB}")
Sleep(200)
Send("{TAB}")
Sleep(200)
Send("{TAB}")
Sleep(200)
Send("{ENTER}")
Sleep(1000)
Send("{TAB}")
Sleep(300)
MouseClick("left",1582,706) ;Field - lab																				;!A7. Please insert the same numbers as !8 mainscript
Do         ;Pixel function PC lab:
$CheckAfter = PixelChecksum(1269,627,1885,943)  ;check pixel checksum 													;!A8. Please insert the same numbers as !21 mainscript
Sleep(500)
Until $CheckAfter = $CheckBefore ;pixel count vergleich
Sleep(250)
MouseClick("left",1866,698) ;X-button 																					;!A9. Please insert the same numbers as !6 mainscript
EndFunc

Func _GoogleListen2()
WinActivate("Google Übersetzer - Google Chrome")
Sleep(500)
MouseClick("left",1296,777) ;Input button - Lab																			;!A10. Please insert the same numbers as !19 mainscript
Beep(200,500)  ;signal for user
Sleep(7000)
MouseClick("left",1296,777) ;Input button - Lab																			;!A11. Please insert the same numbers as !19 mainscript
Sleep(250)
MouseClick("left",1582,706);Field - Lab																					;!A12. Please insert the same numbers as !8 mainscript
Send("^a") ;Select all
Sleep(500)
Send("^x")
Global $vText = ClipGet()
Sleep(250)
EndFunc

Func _SeparateInputSpeechToText() ;Parsing function: separate every word
Global $vText = ClipGet() ;the text stored in the ClipBoard assigned to variable $vText
Sleep(500)
If StringInStr($vText,"Microsoft") Then ;,0,2) Then
	Global $a = StringSplit($vText," ") ;$vText would be separated by empty spaces and saved in $a (an array)
	Global $vTask = ($a[1]) ;first product of StringSplit is an array with X elements - first element $a[1] is stored as $vTask
	Global $cMicrosoft = ($a[2]) ;second element of $a is stored as $vProgram
	Global $vProgram = ($a[3]) ;third element of $a is stored in $vProgram when the second element is "microsoft"
Else ;EndIf
Global $a = StringSplit($vText," ") ;$vText would be separated by empty spaces and saved in $a (an array)
Global $vTask = ($a[1]) ;first product of StringSplit is an array with X elements - first element $a[1] is stored as $vTask
Global $vProgram = ($a[2]) ;second element of $a is stored as $vProgram
Endif
EndFunc
