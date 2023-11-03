	AutoItSetOption("WinTitleMatchMode", -2)
	#include <MsgBoxConstants.au3>
	#include <AutoItConstants.au3>
	#include <Excel.au3>  ;include this library
	#include <WindowsConstants.au3>
	#include <StaticConstants.au3>
	#include <Array.au3>
	#include <Clipboard.au3>
	#include <File.au3>
	#include <Word.au3>
	#include <Process.au3>

;~ FileChangeDir("C:\Users\Public\Documents\SpektraMax")

_CheckExistingSMP()


Func _CheckExistingSMP() ;OK
If WinExists("SoftMax Pro 7.0.3") Then
	WinActivate("SoftMax Pro 7.0.3")
	Sleep(100)
	WinSetState("SoftMax Pro 7.0.3","",@SW_MAXIMIZE)
	Sleep(500)
	Global $iPixel = PixelGetColor(793,563)
	If $iPixel = 16777215 Then
		Send("{F10}")
		Sleep(500)
		Send("{1}")
	EndIf
	_SMPTaskMenu() ;Enter SMP
Else
Send("{LWIN}")
Sleep(1000)
Send("softm")    ;search for SoftMax
Sleep(1000)
Send("{ENTER}")
WinWaitActive("SoftMax Pro 7.0.3") ;Activate Software
Sleep(1000)
WinActivate("SoftMax Pro 7.0.3")
Sleep(100)
WinSetState("SoftMax Pro 7.0.3","",@SW_MAXIMIZE)
Sleep(500)
Global $iPixel = PixelGetColor(793,563)
If $iPixel = 16777215 Then
	Send("{F10}")
	Sleep(500)
	Send("{1}")
EndIf
_SMPTaskMenu()
EndIf
EndFunc

Func _SMPSeparateInputSpeechToText() ;OK
	Global $sSMPText = ClipGet() ;receive command from google translate
	Global $b = StringSplit($sSMPText," ")
	Global $sSMPTask = ($b[1])
	Global $sSecondWord = ($b[2])
EndFunc

Func _SMPTaskMenu()
	Global $sSMPRainbow = ClipPut("Please say command for SoftMax Pro.")
	_VoiceOutput()
	Sleep(500)
	_GoogleListen2()
	_SMPSeparateInputSpeechToText()
	Sleep(500)
	_SMPChooseTask()
EndFunc

Func _SMPFollowUp()
	_GoogleListen2()
	Global $sSMPText = Clipget()
	Sleep(500)
	If $sSMPText = "cancel" Then
		Local $sSMPRainbow = ClipPut("Closing SoftMax Pro Task Menu.")
		_VoiceOutput()
		Sleep(500)
;~ 		WinClose("SoftMax Pro 7.0.3")
;~ 		Sleep(500)
		Exit
	Else
	_SMPSeparateInputSpeechToText()
	_SMPChooseTask()
EndIf
EndFunc

Func _SMPChooseTask()
While 1
		Switch $sSMPTask
			Case "Connect" ;Device
				_connectDevice()
				Global $iA = 2
				If $iA = 2 Then
					_SMPFollowUp()
					ExitLoop ;or exit
				EndIf
			Case "Open" ;Protocol
				_openProtocol()
				Global $iA = 2
				If $iA = 2 Then
					_SMPFollowUp()
					ExitLoop ;or exit
				EndIf
			Case "New" ;Plate
				_NewPlate()
				Global $iA = 2
				If $iA = 2 Then
					_SMPFollowUp()
					ExitLoop
				EndIf
			Case "Safe" ;File
				_SaveAs()
				Global $iA = 2
				If $iA = 2 Then
					_SMPFollowUp()
					ExitLoop
				EndIf
			Case "Save" ;File
				_SaveAs()
				Global $iA = 2
				If $iA = 2 Then
					_SMPFollowUp()
					ExitLoop
				EndIf
			Case "Activate" ;Lid ;Close drawer; tray; lid; box
				_OpenCloseDrawer()
				Global $iA = 2
				If $iA = 2 Then
					_SMPFollowUp()
					ExitLoop
				EndIf
			Case "Actuate"
				_OpenCloseDrawer()
				Global $iA = 2
				If $iA = 2 Then
					_SMPFollowUp()
					ExitLoop
				EndIf
			Case "Read" ;plate
				_Read()
				Global $iA = 2
				If $iA = 2 Then
					_SMPFollowUp()
					ExitLoop
				EndIf
			Case "Reach"
				_Read()
				Global $iA = 2
				If $iA = 2 Then
					_SMPFollowUp()
					ExitLoop
				EndIf
			Case Else
				Local $sSMPRainbow = ClipPut("I didn't quite get that. Repeat command. Otherwise, say cancel.")
				_VoiceOutput()
				_SMPFollowUp()
		EndSwitch
WEnd
EndFunc

Func _connectDevice() ;Connect software to hardware.
	Sleep(200)
	WinActivate("SoftMax Pro 7.0.3") ;active the software window
	Sleep(500)
	Send("{F10}")
	Sleep(1000)
	Send("h")
	Sleep(1000)
	Send("s")
	Sleep(1000)
	WinWaitActive("Instrument Connection") ;window with software/hardware status opens
	Sleep(500)
	Global $sumConnect = PixelChecksum(811,446,1101,464)
	If $sumConnect = 1863126818 Then
		Do
		Global $sumConnect = PixelChecksum(811,446,1101,464)  ;check unti the pixel checksum is x to then click enter
		Sleep(500)
		Until $sumConnect = 1863126818 ;pixel count
		Sleep(500)
		Send("{ENTER}")
		Sleep(500)
		WinWaitActive("SoftMax Pro 7.0.3")
		Sleep(500)
		Local $sSMPRainbow = ClipPut("SpectraMax iD5 connected. Say next command. Otherwise, say cancel.")
		_VoiceOutput()
	Else
	Do
	Global $sumConnect = PixelChecksum(811,446,1101,464)  ;check unti the pixel checksum is x to then click enter
	Sleep(500)
	Until $sumConnect = 1863126818 ;pixel count
	Sleep(500)
	MouseClick("left",816,624) ;Simulate on
	Sleep(500)
	Send("{ENTER}")
	WinWaitActive("SoftMax Pro 7.0.3")
	Sleep(500)
	Local $sSMPRainbow = ClipPut("SpectraMax iD5 connected. Say next command. Otherwise, say cancel.")
	_VoiceOutput()
	EndIf
EndFunc

Func _openProtocol() ;Open existing protocol in SMP.
WinActivate("SoftMax Pro 7.0.3") ;wait till connection is established to open menu again
Send("{F10}")
Sleep(500)
Send("f")
Sleep(500)
Send("o")
WinWaitActive("Öffnen")
Sleep(500)
WinMove("Öffnen","Adresse: C:\Users\Public\Documents\SpektraMax",7,7)
Sleep(500)
Local $sSMPRainbow = ClipPut("Say protocol name to open.")
_VoiceOutput()
Sleep(500)
_GoogleListen2()
Global $sSMPText = ClipGet()
;~ MsgBox(0,"",$sSMPText)
Sleep(500)
WinActivate("Öffnen")
MouseClick("left",843,274)
Global $sSMPFileName = StringLeft($sSMPText,1) ;this may need some improvement, since only the first letter of file is assigned to variable
Sleep(1000)
Send($sSMPFileName);("20200305_Lissamin")
;~ MsgBox(0,"",$sSMPFileName)
Sleep(800)
Send("{ENTER}")
Sleep(500)
Local $sSMPRainbow = ClipPut("Protocol open. Say next command. Otherwise, say cancel.")
_VoiceOutput()
EndFunc

Func _NewPlate()
WinActivate("SoftMax Pro 7.0.3")
Sleep(500)
MouseClick("left",179,246)
Local $sSMPRainbow = ClipPut("New plate. Say next command. Otherwise, say cancel.")
_VoiceOutput()
EndFunc

Func _SaveAs() ;Save protocol as $FileName.
WinActivate("SoftMax Pro 7.0.3")
Send("{F10}")
Sleep(500)
Send("4")
Sleep(500)
WinWaitActive("Speichern unter")
Local $sRainbow = ClipPut("Say new file name.")
_VoiceOutput()
Sleep(500)
_GoogleListen2()
Sleep(500)
WinActivate("Speichern unter")
Sleep(500)
Send("{CTRLdown}")
Send("{v}")
Send("{CTRLup}")
Sleep(500)
Sleep(500)
Send("{ENTER}")
If WinExists("SoftMax Pro","The file already exists. Do you want to overwrite it?") Then
	Local $sRainbow = ("The file already exists. Do you want to overwrite it?")
	_VoiceOutput()
	Sleep(500)
	_GoogleListen2()
	Global $sSMPText = ClipGet()
	If $sSMPText = "yes" then
		WinActivate("SoftMax Pro","The file already exists. Do you want to overwrite it?")
		Send("{tab}")
		Sleep(500)
		Send("{enter}")
		Sleep(500)
		Local $sSMPRainbow = ClipPut("Protocol saved. Say next command. Otherwise, say cancel.")
		_VoiceOutput()
	Else
		WinActivate("SoftMax Pro","The file already exists. Do you want to overwrite it?")
		Sleep(500)
		Send("{enter}")
		Sleep(500)
		Local $sSMPRainbow = ClipPut("Protocol not saved. Say next command. Otherwise, say cancel.")
		_VoiceOutput()
	EndIf
Else
Local $sSMPRainbow = ClipPut("Protocol saved. Say next command. Otherwise, say cancel.")
_VoiceOutput()
EndIf
EndFunc

Func _OpenCloseDrawer() ;Open or close drawer
WinActivate("SoftMax Pro 7.0.3")
Sleep(1000)
Send("{F10}")
Sleep(1000)
Send("h")
Sleep(1000)
Send("o")
Local $sSMPRainbow = ClipPut("Active Drawer. Say next command. Otherwise, say cancel.")
_VoiceOutput()
EndFunc

 Func _Read() ;Reads plate
WinActivate("SoftMax Pro 7.0.3")
Sleep(500)
Send("{F10}")
Sleep(20)
Send("h")
Sleep(20)
Send("r1")
If WinExists("SoftMax Pro","The settings contain some illegal parameters.") Then
	WinActivate("SoftMax Pro","The settings contain some illegal parameters.")
	Sleep(500)
	Send("{ENTER}")
EndIf
Sleep(500)
WinWaitActive("Pre-Read Optimization Options")
Sleep(200)
Send("{ENTER}")
Sleep(200)
Local $sSMPRainbow = ClipPut("Plate read. Say next command. Otherwise, say cancel.")
_VoiceOutput()
 EndFunc

Func _VoiceOutput()    ;Output voice assistant $sRainbow
WinActivate("Google Übersetzer - Google Chrome")
Sleep(500)
WinMove("Google Übersetzer - Google Chrome","",1258,424,658,614) ;Move window - lab
Sleep(500)
WinActivate("Google Übersetzer - Google Chrome")
Sleep(500)
MouseClick("left",1866,698) ;X-button - lab
Sleep(250)
MouseClick("left",1582,706) ;Field - lab
Sleep(250)
Send("^v");&$sRainbow)
Sleep(250)
MouseClick("left",1582,706) ;Field - lab
Sleep(500)
Do
	$error = PixelGetColor(1290,777) ;revise if there is a correction in Google Translate
	Sleep(250)
	$iA = 2
Until $iA = 2
If $error = 1733608 Then ;Check for blue pixel
	WinActivate("Google Übersetzer - Google Chrome")
	Sleep(500)
	Send("{TAB}")
	Sleep(500) ;save time with less ms?
	Send("{TAB}")
	Sleep(600)
	Send("{ENTER}")  ;Accept correction
	Sleep(500)
	MouseClick("left",1579,698) ;Field - lab
EndIf
Sleep(500)  ;Normal output
Send("{TAB}")
Sleep(500)
Send("{TAB}")
Sleep(500)
Send("{TAB}")
Sleep(500)
Send("{ENTER}")
Sleep(500)
WinActivate("Google Übersetzer - Google Chrome")
Do         ;Pixel function PC lab:
$sum = PixelChecksum(1475,450,1488,462)  ;check unti the pixel checksum is x to then click enter
Sleep(500)
Until $sum = 802824189 ;pixel count
Sleep(250)
MouseClick("left",1582,706) ;Field - lab
Sleep(250)
EndFunc

Func _GoogleListen2() ;Voice Input with mouse clicks
WinActivate("Google Übersetzer - Google Chrome")
WinMove("Google Übersetzer - Google Chrome","",1258,424,658,614)  ;Lab PC
Sleep(500)
WinActivate("Google Übersetzer - Google Chrome")
Sleep(500)
MouseClick("left",1866,698) ;X-button - Lab
Sleep(250)
MouseClick("left",1582,706) ;Field - Lab
Sleep(250)
Send("{TAB}")
Sleep(500)
Send("{ENTER}")
Sleep(500)
Beep(200,500)  ;signal for user
Sleep(7000)
Send("{ENTER}")
Sleep(250)
MouseClick("left",1582,706);Field - Lab
Send("^a") ;Select all
Sleep(500)
Send("^c") ;Copy
Global $sSMPText = ClipGet()
Sleep(250)
MouseClick("left",1866,698) ;X-button - Lab
Sleep(250)
EndFunc


