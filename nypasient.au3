#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile=nypasient.exe
#AutoIt3Wrapper_Run_Before=updversion.exe
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****
;~ #Region ;**** Directives created by AutoIt3Wrapper_GUI ****


; ================================
; 13/04/19 - initial prototype
; 15/04/19 v1. ==========
; 17/04/19 - added change filename register by right click
; 23/04/19 - Fixed bugs
;	  - Sex now identifed correctly - by 9th digit (was by 7th)
; 	- When cancel file overwrite return to edit (was Exit from program)
; 	- Write to file in Unocode mode.Can handle norwegian chars (was in Ascii)
;	- Replaced StringProper to work correctly with norwegian chars
; 26/04/19
;	  - changed verification algorithms to RegEx
;	  - added pasient with only f.dato: ddmmyy(k|m), ddmmyyyy(k|m),
; 27/04/19 - added tooltip with examples
; 28/04/19 - added version number in title (dd.mm.yy.hhmm)
; 30/04/19 v2. ==========
; 	- redesigned with functions front/backend
; 	- added unit testing
;	- added support for ddmmyyyy(s)
; 	- added automatic centure for ddmmyy (if fnr.year < now.year -> 1900, otherwise 2000
; 08/05/19 v2.1
;	- fixed error "no auto_.xml"
; 	- fixed error in text "fГёdselsnummer"
;	- fiexed proper in names with numbers
;
; ================================

#include <Array.au3>
#include <string.au3>
#include <excel.au3>
#include <ie.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <TreeViewConstants.au3>
#include <StaticConstants.au3>


;OnAutoItExitRegister("MyExitFunc")
#include <WinAPI.au3>
#include <Date.au3>
#include <FontConstants.au3>
#include <GUIConstantsEx.au3>

#include "constants.au3"

Opt('MustDeclareVars', 1)

#AutoIt3Wrapper_Res_Field = Timestamp|%date%.%time%

Local $ver
If @Compiled Then

$ver = FileGetVersion(@ScriptFullPath, "Timestamp")
$ver = StringReplace( $ver, "/", "." )
$ver = StringReplace( $ver, ":", "" )
$ver = StringLeft( $ver, StringLen($ver)-2 )

Else

$ver = "Not compiled"

EndIf

; Create input
GUICreate("Create new pasient - " & $ver, 450, 48)

Global $ctrlFile = GUICtrlCreateLabel("Pasient", 8, 16)
;GUICtrlSetFont( $ctrlName, 10, 600 )

Local $contextmenu = GUICtrlCreateContextMenu($ctrlFile)

Local $properItem = GUICtrlCreateMenuItem("Navn Etternavn", $contextmenu)
Local $upperItem = GUICtrlCreateMenuItem("NAVN ETTERNAVN", $contextmenu)

Global $ctrlName = GUICtrlCreateInput("navn etternavn f.nr", 60, 8, 380, 30 )
GUICtrlSetTip(-1, "fnr/dnr" & @CRLF & "DDMMYY" & @CRLF & "DDMMYYm" & @CRLF & "DDMMYYYYk")
GUICtrlSetFont( $ctrlName, 14, 600 )

GUICtrlSetState($ctrlName, $GUI_FOCUS)

GUISetState() ; will display an empty dialog box

Local $msg
Global $typetext ; 0 - lowcase, 1- uppercase



Do
	$msg = GUIGetMsg()

	Switch $msg
		Case $ctrlName

			;ParseInput( GUICtrlRead($ctrlName) )
			ProcessInput( GUICtrlRead($ctrlName) )
			GUICtrlSetState($ctrlName, $GUI_FOCUS)

		Case $GUI_EVENT_SECONDARYDOWN

			; Run the GUI until the dialog is closed
			While 1

				Switch GUIGetMsg()

					Case $properItem
						$typetext = 0
						ExitLoop
					Case $upperItem
						$typetext = 1
						ExitLoop
				EndSwitch

			WEnd

			Switch $typetext
				Case 0
					GUICtrlSetData($ctrlFile, _StringProper1(GUICtrlRead($ctrlFile)))
				Case 1
					GUICtrlSetData($ctrlFile, StringUpper(GUICtrlRead($ctrlFile)))

			EndSwitch

		Case $GUI_EVENT_CLOSE
			Exit

	EndSwitch

Until $msg = -1 ;> 0 ;= $Button_Ok

GUIDelete()

Exit ;


; -----------------------------------------------------------------------------
; Function: Process input
; Input: String
;
; -----------------------------------------------------------------------------
Func ProcessInput( $input )
	Local $name, $surname, $fnr, $sexid, $fdato
	Local $err

	; Strip of white space and place into array
	$input = StringStripWS( $input, 7 )

	; Get name and surname
	$err = GetNames( $input, $name, $surname, $fnr )
	if $err > 0 then
		FlagError( $err )
		return
	EndIf

	; get all other elements from fnr
	$err = GetElements( $fnr, $fnr, $fdato, $sexid)
	if $err > 0 then
		FlagError( $err )
		return
	EndIf

	; Create pasients file
	$err = CreatePasientFile($name, $surname, $fnr, $fdato, $sexid )
	if $err > 0 then
		FlagError( $err )
		Return
	EndIf


								  

EndFunc

; ==================================
; Get names from input
; Return:
; 	- Navn , Etternavn, Fnr
; ==================================

Func GetNames( $input, byref $name, byref $surname, byref $fnr )


	; acceptable format:
	; 	name middle surname ddmmyyxxxxx	- fnr
	;	name middle surname ddmmyy(s) 	- short fdato
	;	name middle surname ddmmyyyy(s)	- long fdato

	; if fnr goes last
	If StringRegExp( $input, "^(\S+ ){2,}(\d{6})((\d\d)?([kKmM]?)|\d{5})$" ) then

		$name    = _StringProper1( StringRegExpReplace( $input, "^(\S+) (.*)$", "$1" ))
		$surname = _StringProper1( StringRegExpReplace( $input, "^(.*?) (.*) (.*)$", "$2" ))
		$fnr     = StringRegExpReplace( $input, "^(.*) (\d+.*)$", "$2" )

	; if fnr goes first
	elseif  StringRegExp( $input, "^(\d{6})((\d\d)?([kKmM]?)|\d{5})( \S+){2,}$" ) then

		$name    = _StringProper1( StringRegExpReplace( $input, "^(\d+.*?) (\S+) (.*)$", "$2" ))
		$surname = _StringProper1( StringRegExpReplace( $input, "^(\d+.*?) (\S+) (.*)$", "$3" ))
		$fnr     = StringRegExpReplace( $input, "^(\d+.*?) (.*)$", "$1" )

	; if wrong format
	Else
		Return $ERR_FORMAT

	EndIf

	;ConsoleWrite( $name & "," & $surname& "," & $fnr & @CRLF )
	return $ERR_OK

EndFunc

; ==================================
; Get Elements
; Return:
; 	- Fnr, Fdato, SexId
; ==================================
Func GetElements( $input, byref $fnr, byref $fdato, byref $sexid )

	; acceptable format:
	; 	ddmmyyxxxxx	- fnr
	;	ddmmyy(s) 	- short fdato
	;	ddmmyyyy(s)	- long fdato

	; if normal fnr/dnr - ddmmyyxxxxx
	If StringRegExp( $input, "^([04][1-9]|[1256][0-9]|[37][01])(0[1-9]|1[012])(\d){7}$") then
			$sexid = 2 - mod( StringMid( $input, 9, 1), 2) ; odd=2, even=1
			$fnr = $input
			$fdato = GetFdato( $input )
			return 0

	; if short fdato - ddmmyy(s)
	elseif StringRegExp( $input, "^(0[1-9]|[12][0-9]|3[01])(0[1-9]|1[012])[0-9][0-9][kKmM]?$") then
			$sexid = GetSexId(  StringRight($input,1) )
			$fnr = 0 ; StringLeft( $input, 6)
			$fdato = GetFdato( $input )
			return 0

	; if long fdato - ddmmyyyy(s)
	elseif StringRegExp( $input, "^(0[1-9]|[12][0-9]|3[01])(0[1-9]|1[012])\d{4}[kKmM]?$") then
			$sexid = GetSexId(  StringRight($input,1) )
			$fnr = 0 ;StringRegExpReplace( $input, "(\d\d\d\d)\d\d(\d\d).?", "$1$2")
			$fdato = GetFdato( $input )
			return 0

	EndIf

	; none fits - > wrong format
	Return $ERR_FORMAT

EndFunc

; -----------------------------------------------------------------------------
; Function: Create pasient file
; Input:
;	0 - name
;	1 - surname
; 	2 - fnr
; 	3 - fdato
;	4 - sexid (1-mann, 2-kvinne, 9-unknown)
; 	6 - namecase (0-lower, 1-upper)
; -----------------------------------------------------------------------------
Func CreatePasientFile( $name,$surname, $fnr, $fdato, $sexid )
	Local $id, $sex, $age
	Local $filetemplate
	Local $sString
	Local $fileoutput

	; Get GUID
	$id  = GetGUID()

	; Get Sex name
	$sex = GetSexName( $sexid )

	; Get age
	$age = GetAge( $fdato )
	if $age < 0 then
		return 2 ;ERR_UGYLDIG_FNR
	endif

	; Read file
	$filetemplate = @WorkingDir & "\auto_.xml"
	$sString = FileRead($filetemplate)
	If @error = 1 Then
		MsgBox(0, "Error", "Can't open file" & $filetemplate)
		Exit
	EndIf

	$sString = StringReplace($sString, "#name#", $name)
	$sString = StringReplace($sString, "#surname#", $surname)
	$sString = StringReplace($sString, "#birthdate#", $fdato)
	$sString = StringReplace($sString, "#fnr#", $fnr)
												
	$sString = StringReplace($sString, "#sex#", $sex)
	$sString = StringReplace($sString, "#sexid#", $sexid)
	$sString = StringReplace($sString, "#id#", $id)

	if $fnr = 0 then
		$sString = StringRegExpReplace( $sString, "(?s)(?i)<Ident>.*?FNR.*?</Ident>", "" )
;	Else
;		$sString = StringRegExpReplace( $sString, "(?s)(?i)<Kjonn.*?/>", "" )
	endif


	; change type
	if isDnr($fnr)  Then
		$sString = StringReplace($sString, '="FNR"', '="DNR"')
		$sString = StringReplace($sString, '="F'& ChrW(248) &'dselsnummer"', '="D-nummer"')

	Else
		$sString = StringReplace($sString, '="DNR"', '="FNR"')
		$sString = StringReplace($sString, '="D-nummer"', '="F'& ChrW(248) &'dselsnummer"')

	EndIf

	; Write file
	$fileoutput = StringReplace($filetemplate, "_", "_" & $name & " " & $surname, -1)
	ConsoleWrite( $typetext & @CRLF )
	If $typetext = 1 Then
		$fileoutput = StringReplace($filetemplate, "_", "_" & StringUpper($name & " " & $surname), -1)
	EndIf

	If FileExists($fileoutput) Then
		If MsgBox(1, "Error", "File " & $fileoutput & " alredy exists. Overwite? ") = 2 Then
			Return
		EndIf
		FileDelete($fileoutput)
	EndIf

	Local $file = FileOpen($fileoutput, 256 + 2)

	If $file = -1 Then
		MsgBox(0, "Error", "Can't open file " & $fileoutput)
		Return
	EndIf

	Local $sString
	If FileWrite($file, $sString) = 0 Then
		MsgBox(0, "Error", "Can't write the file")
		Return
	EndIf

	FileClose($file)

	$sString = ""
	$sString &= "Name : " & $name & " " & $surname & @CRLF
	$sString &= "Fnr  : " & $fnr & @CRLF
	$sString &= "Fdato: " & $fdato & "  (" & $sex & "-" & $age & ")" & @CRLF
	$sString &= "File : " & $fileoutput & @CRLF

	MsgBox(0, "Pasient successfully created", $sString)

EndFunc

; -----------------------------------------------------------------------------
; Function: Exit handler
; -----------------------------------------------------------------------------
Func _GenerateGUID()
	Local $oScriptlet = ObjCreate("Scriptlet.TypeLib")
	Return $oScriptlet.Guid
EndFunc   ;==>_GenerateGUID

; -----------------------------------------------------------------------------
; Function: _StringProper1
; -----------------------------------------------------------------------------
Func _StringProper1($s_String)
	Local $iX = 0
	Local $CapNext = 1
	Local $s_nStr = ""
	Local $s_CurChar
	For $iX = 1 To StringLen($s_String)
		$s_CurChar = StringMid($s_String, $iX, 1)
		Select
			Case $CapNext = 1
				If StringRegExp($s_CurChar, '[a-zA-ZА-я0-9' & ChrW(198) & ChrW(230) & ChrW(216) & ChrW(248) & ChrW(197) & ChrW(229) & ']') Then
					$s_CurChar = StringUpper($s_CurChar)
					$CapNext = 0
				EndIf
			Case Not StringRegExp($s_CurChar, '[a-zA-ZА-я0-9' & ChrW(198) & ChrW(230) & ChrW(216) & ChrW(248) & ChrW(197) & ChrW(229) & ']')
				$CapNext = 1
			Case Else
				$s_CurChar = StringLower($s_CurChar)
		EndSelect
		$s_nStr &= $s_CurChar
	Next
	Return $s_nStr
EndFunc ;==>_StringProper1

; ==================================
; Get century from fnr/dnr
; Return:
; 	0 - error - feil fnr
; 	1800,1900,2000 - otherwise
; Unit test:
; 	ut_GetCentury.au3
; ==================================
Func GetCentury( $fnr )
Local $pers = StringMid( $fnr, 7,3)
Local $yy = StringMid( $fnr, 5,2)
	; Check Alder
	; 000-499 - 1900
	; 500-749 - 1854-1899
	; 500-999 - 2000-2039
	; 900-999 - 1940-1999

	If $pers < 500 Then
		return 19
	ElseIf $yy < 40 Then
		return 20
	ElseIf $pers > 899 Then
		return 19
	ElseIf $pers < 750 and $yy > 53 Then
		return 18
	Else
		Return 0
	EndIf


EndFunc





; ==================================
; Get Sex name from sexid
; Return:
; 	sex name
;
; Unit test:
; 	ut_GetSexName.au3
; ==================================

Func GetSexName( $sexid )
	Local $sex

	switch $sexid
		case $SEX_MANN
			$sex = "Mann"
		case $SEX_KVINNE
			$sex = "Kvinne"
		case Else
			$sex = "Ukjent"
	EndSwitch

	Return $sex
EndFunc

; ==================================
; Get SexId from letter
; Return:
; 	sexid
;
; Unit test:
; 	ut_GetSexId.au3
; ==================================
Func GetSexId( $sSex )
	Local $id

	Switch StringUpper( $sSex  )
		case "M"
			$id = $SEX_MANN
		case "K"
			$id = $SEX_KVINNE
		case else
			$id = $SEX_UKJENT
	endswitch
	Return $id
EndFunc

; ==================================
; Get GUID
; Return:
; 	- uniq GUID
; ==================================
Func GetGUID()
	Local $id
	$id =_GenerateGUID()
	$id = StringMid($id, 2, StringLen($id) - 2)
	Return $id
EndFunc

; ==================================
; Check it is dnr
; Return: true=dnr/ false-fnr
; Unit test:
; 	ut_isDnr.au3
; ==================================
Func isDnr( $fnr)
	Return StringRegExp( $fnr, "^(4[1-9]|[56][0-9]|7[01])" )
EndFunc

; ==================================
; Get Fdato
; Input:
; 	ddmmyy
;	ddmmyyyy
;	ddmmyy00000

; Return:
; 	- fdato string yyyy-mm-dd
; ==================================
Func GetFdato( $fnr )

	Local $fdato

		; ddmmyy00000 - fnr&dnr
		if StringRegExp( $fnr, "^\d{11}$" ) then

			; check if dnr
			if isDnr( $fnr ) then
				$fnr = String(StringLeft( $fnr, 1)-4) & StringMid($fnr, 2, 10)
			Endif

			$fdato = StringRegExpReplace( $fnr, "^(\d\d)(\d\d)(\d\d)\d*", GetCentury($fnr) & "$3-$2-$1" )

		; ddmmyyyy - long fdato
		elseif StringRegExp( $fnr, "^\d{8}\D?$" ) then

			$fdato = StringRegExpReplace( $fnr, "(\d\d)(\d\d)(\d\d\d\d)\D?", "$3-$2-$1")

		; ddmmyy - short fdato
		elseif StringRegExp( $fnr, "^\d{6}\D?$" ) then

			$fdato = StringRegExpReplace( $fnr, "^(\d\d)(\d\d)(\d\d)\D?", "$3-$2-$1")

			; if fnr.year >= now.year then 1900, else 2000
			if StringMid( $fnr, 5,2) >= StringMid( _NowCalc(), 3,2) then
				$fdato = "19" & $fdato
			Else
				$fdato = "20" & $fdato
			EndIf

		else
			return -1
		EndIf

	Return $fdato

EndFunc

; ==================================
; Get Age
; Return:
; 	- Age
; ==================================
Func GetAge( $fdato )

	Local $age
	Local $yyyy, $mm, $dd

	$yyyy	= StringLeft($fdato, 4)
	$mm 	= StringMid( $fdato, 6, 2)
	$dd 	= StringMid( $fdato, 9, 2)

	$age = _DateDiff("Y", $yyyy & "/" & $mm & "/" & $dd, _NowCalc())
	if @error <> 0 Then
		Return -1
	EndIf

	Return $age

EndFunc


; ===================================
; Handle Error message
; ===================================

func FlagError( $err )
	Local $text
	switch $err
		case $ERR_OK
			return
		case $ERR_FORMAT
			$text = "name surname fnr"
		case $ERR_FNR
			$text = "ugyldig f" & ChrW(248) &  "dselsnummer"
		case 3
			$text = "3"
		case 4
			$text = "4"
		case Else
			$text = "Unknown error"

	EndSwitch

	MsgBox( 0, "Error" , $text )

EndFunc
