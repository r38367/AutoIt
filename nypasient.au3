#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile=nypasient.exe
#AutoIt3Wrapper_Run_Before=updversion.exe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;~ #Region ;**** Directives created by AutoIt3Wrapper_GUI ****


; ================================
; 13/04/19 - initial prototype
; 15/04/19 -
; 17/04/19 - added change filename register by right click
; 23/04/19 - Fixed bugs
;	- Sex now identifed correctly - by 9th digit (was by 7th)
; 	- When cancel file overwrite return to edit (was Exit from program)
; 	- Write to file in Unocode mode.Can handle norwegian chars (was in Ascii)
;	- Replaced StringProper to work correctly with norwegian chars
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
#Include <WinAPI.au3>
#include <Date.au3>
#include <FontConstants.au3>
#include <GUIConstantsEx.au3>

Global	$sString	; temp string
Global	$filetemplate	; filename for temaplte
Global	$fileoutput	; filename for output

Global	$name
Global	$surname
Global	$fnr
Global	$sex
Global	$sexid
Global	$fdato
Global	$id
Global	$age

Opt('MustDeclareVars', 1)


; Create input
GUICreate( "Create new pasient", 450, 48 )

Global $ctrlFile = GUICtrlCreateLabel("Pasient", 8, 16 )
;GUICtrlSetFont( $ctrlName, 10, 600 )

Local $contextmenu = GUICtrlCreateContextMenu($ctrlFile)

Local $properItem = GUICtrlCreateMenuItem("Navn Etternavn", $contextmenu)
Local $upperItem = GUICtrlCreateMenuItem("NAVN ETTERNAVN", $contextmenu)

Global $ctrlName = GUICtrlCreateInput("navn etternavn f.nr", 60, 8, 380, 30 )
GUICtrlSetFont( $ctrlName, 14, 600 )
GUICtrlSetState($ctrlName, $GUI_FOCUS)

GUISetState() ; will display an empty dialog box

Global $msg
Global $arrName
Global $arrFnr

Global $typetext = 1


Do
        $msg = GUIGetMsg()

		switch $msg
			case $ctrlName

				$arrName = StringSplit( StringStripWS (GUICtrlRead($ctrlName), 7), " " )
				ParseInput( $arrName )
				GUICtrlSetState($ctrlName, $GUI_FOCUS)

			case $GUI_EVENT_SECONDARYDOWN

				; Run the GUI until the dialog is closed
				While 1

					Switch GUIGetMsg()

						case $properItem
							$typetext = 1
							ExitLoop
						case $upperItem
							$typetext = 2
							ExitLoop
					EndSwitch

				WEnd

				switch $typetext
					case 1
						GUICtrlSetData( $ctrlFile, _StringProper1(GUICtrlRead($ctrlFile)))
					case 2
						GUICtrlSetData( $ctrlFile, StringUpper(GUICtrlRead($ctrlFile)))

				EndSwitch

			case $GUI_EVENT_CLOSE
				Exit

		EndSwitch

Until $msg = -1  ;> 0 ;= $Button_Ok

GUIDelete ()

Exit

; -----------------------------------------------------------------------------
; Function: Parse Input
; -----------------------------------------------------------------------------
Func ParseInput( $arrName )

	; if less than 3
	if $arrName[0] < 3 then
		MsgBox( 0, "Error", "Navn Etternavn fnr")
		Return
	endif

	; Navn og Fnr
	if StringIsDigit($arrName[1]) and StringLen($arrName[1])=11 then
		$fnr = $arrName[1]
		$name = $arrName[2]
		$surname = $arrName[3]
		for $i = 4 to $arrName[0]
			$surname = $surname & " " & $arrName[$i]
		Next
	Elseif StringIsDigit($arrName[$arrName[0]]) and StringLen($arrName[$arrName[0]])=11 then
		$name = $arrName[1]
		$surname = $arrName[2]
		for $i = 3 to $arrName[0]-1
			$surname = $surname & " " & $arrName[$i]
		Next
		$fnr = $arrName[$arrName[0]]
	Else
		MsgBox( 0, "Error", "Ugyldig Fnr")
		Return
	EndIf

	; split fnr
	$arrFnr = StringRegExp( $fnr, "(\d\d)(\d\d)(\d\d)(\d\d\d)(\d\d)", 1)
	global $dd = $arrFnr[0]
	if $dd > 40 then
		$dd -= 40
		if $dd < 10 then $dd = "0"&$dd
	endif

	global $mm	= $arrFnr[1]
	global $yy	= $arrFnr[2]
	global $pers = $arrFnr[3]
	global $cc	= $arrFnr[4]

	; check fdato

	if $dd = 0 or $dd > 31 then
		MsgBox( 0, "Error", "Ugyldig Fnr" )
		Return
	ElseIf $mm = 0 or $mm > 12 then
		MsgBox( 0, "Error", "Ugyldig Fnr" )
		Return
	endif

	; Get Sex
	; xxxxxx xx0xx - 0 - kvinne, 1 - mann
	if mod( StringRight( $pers,1), 2 ) = 0 Then
		$sex = "Kvinne"
		$sexid = 2
	Else
		$sex = "Mann"
		$sexid = 1
	EndIf

	; Check Alder
	; 000-499 - 1900
	; 500-749 - 1854-1899
	; 500-999 - 2000-2039
	; 900-999 - 1940-1999

	if $pers < 500 Then
		$yy = 1900 + $yy
	elseif $yy < 40 Then
		$yy = 2000 + $yy
	Elseif $pers >899 Then
		$yy = 1900 + $yy
	ElseIf $pers < 750 AND $yy > 53 Then
		$yy = 1800 + $yy
	Else
		MsgBox( 0, "Error", "Ugyldig Fnr" )
		Return
	endif

	; Get fdato
	$fdato = $yy & "-" & $mm & "-" & $dd

	; Get Age
	$age = _DateDiff("Y", $yy&"/"&$mm&"/"&$dd,_NowCalc())

	; Set GUID
	$id = _GenerateGUID()
	$id = StringMid($id,2,StringLen($id)-2)

	; Read file
	$filetemplate = @WorkingDir & "\auto_.xml"
	$sString = FileRead( $filetemplate )
	If @error = -1 Then
		MsgBox(0, "Error", "Can't open file" & $filetemplate )
		Exit
	endif
	$sString = StringReplace( $sString, "#name#", _StringProper1 ($name) )
	$sString = StringReplace( $sString, "#surname#", _StringProper1 ($surname) )
	$sString = StringReplace( $sString, "#birthdate#", $fdato )
	$sString = StringReplace( $sString, "#fnr#", $fnr )
	$sString = StringReplace( $sString, "#id#", $id )
	$sString = StringReplace( $sString, "#sex#", $sex )
	$sString = StringReplace( $sString, "#sexid#", $sexid )


	; Write file
	$fileoutput = StringReplace( $filetemplate, "_", "_"& _StringProper1($name & " " & $surname), -1 )
	if $typetext = 2 then
		$fileoutput = StringReplace( $filetemplate, "_", "_"& StringUpper($name & " " & $surname), -1 )
	EndIf
	if FileExists( $fileoutput ) then
		if MsgBox( 1, "Error", "File "& $fileoutput & " alredy exists. Overwite? ") = 2 then
			Return
		EndIf
		FileDelete( $fileoutput )

	EndIf

	Local $file = FileOpen( $fileoutput, 256+2  )

	if $file = -1 Then
			MsgBox( 0, "Error", "Can't open file " & $fileoutput )
			Return
	EndIf

	if FileWrite( $file, $sString ) = 0 then
			MsgBox( 0, "Error", "Can't write the file" )
			Return
	Endif

	FileClose($file)

	$sString  = ""
	$sString &= "Name : " & _StringProper1( $name & " " & $surname ) & @CRLF
	$sString &= "Fnr  : " & $fnr & @CRLF
	$sString &= "Fdato: " & $fdato & "  (" & $sex & $age& ")" & @CRLF
	;$sString &= "Guid : " & $id & @CRLF & @CRLF
	$sString &= @CRLF
	$sString &= "File : " & $fileoutput & @CRLF

	MsgBox( 0, "Pasient successfully created", $sString )
EndFunc



; -----------------------------------------------------------------------------
; Function: Exit handler
; -----------------------------------------------------------------------------
Func _GenerateGUID ()
Local $oScriptlet = ObjCreate ("Scriptlet.TypeLib")
Return $oScriptlet.Guid
EndFunc

Func Gender( $personalnummer)
Return BitAND( StringMid( $personalnummer, 3,1), 0x1)
EndFunc

; ===============================================================================================================================
Func _StringProper1($s_String)
	Local $iX = 0
	Local $CapNext = 1
	Local $s_nStr = ""
	Local $s_CurChar
	For $iX = 1 To StringLen($s_String)
		$s_CurChar = StringMid($s_String, $iX, 1)
		Select
			Case $CapNext = 1
				If StringRegExp($s_CurChar, '[a-zA-ZА-яљњћџ' & ChrW(198) & ChrW(230) & ChrW(216) & ChrW(248)& ChrW(197) & ChrW(229) & ']') Then
					$s_CurChar = StringUpper($s_CurChar)
					$CapNext = 0
				EndIf
			Case Not StringRegExp($s_CurChar, '[a-zA-ZА-яљњћџ' & ChrW(198) & ChrW(230) & ChrW(216) & ChrW(248)& ChrW(197) & ChrW(229) & ']')
				$CapNext = 1
			Case Else
				$s_CurChar = StringLower($s_CurChar)
		EndSelect
		$s_nStr &= $s_CurChar
	Next
	Return $s_nStr
EndFunc   ;==>_StringProper
