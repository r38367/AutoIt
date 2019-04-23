;~ #Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_outfile=GetALMTestReport.exe
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Run_Before=updversion.exe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****


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
Global	$birthdate
Global	$id


; Create space for input

GUICreate( "Create new pasient", 500, 500 )

; Add label
$left = 10
$down = 10
$vstep = 25
$hstep = 50

$ret = GUICtrlCreateLabel("Pasient", $left, $down )
	Global $ctrlName = GUICtrlCreateInput("navn etternavn f.nr", $left+$hstep, $down, 300 )
	GUICtrlSetFont( $ctrlName, 12, 600 )

GUICtrlCreateLabel("Navn", $left, $down+$vstep)
		Global $ctrlNameText = GUICtrlCreateLabel(" ", $left+$hstep, $down+$vstep, 300 )
GUICtrlCreateLabel("Etternavn", $left, $down+$vstep*2)
		Global $ctrlEtternavnText = GUICtrlCreateLabel(" ", $left+$hstep, $down+$vstep*2, 300 )
GUICtrlCreateLabel("F.nr", $left, $down+$vstep*3)
		Global $ctrlFnrText = GUICtrlCreateLabel(" ", $left+$hstep, $down+$vstep*3, 300 )
GUICtrlCreateLabel("F.dato", $left, $down+$vstep*4)
		Global $ctrlFdatoText = GUICtrlCreateLabel(" ", $left+$hstep, $down+$vstep*4, 300 )
GUICtrlCreateLabel("Alder", $left, $down+$vstep*5)
		Global $ctrlAlderText = GUICtrlCreateLabel("", $left+$hstep, $down+$vstep*5, 300 )
GUICtrlCreateLabel("UID", $left, $down+$vstep*6)
		Global $ctrlUidText = GUICtrlCreateLabel(StringMid($sString,2,StringLen($sString)-2), $left+$hstep, $down+$vstep*6, 300 )
GUICtrlCreateLabel("File", $left, $down+$vstep*7)
		Global $ctrlFileText = GUICtrlCreateInput(@WorkingDir, $left+$hstep, $down+$vstep*7, 300 )

; Add output text
Global $ctrlOutput = GUICtrlCreateLabel("", $left, $down+$vstep*8, 400,200)
GUICtrlSetData(-1, "<html> output </html>" )

; Set focus on input
GUICtrlSetState($ctrlName, $GUI_FOCUS)

GUISetState() ; will display an empty dialog box

Do
        $msg = GUIGetMsg()

		Select
			case $msg = $ctrlName

					$arrName = StringSplit( StringStripWS (GUICtrlRead($ctrlName), 7), " " )

					ParseInput( $arrName )
					GUICtrlSetState($ctrlName, $GUI_FOCUS)

			case $msg = $ctrlFileText
				$var = FileSelectFolder("Choose a folder.", "", 1, @WorkingDir )
				GUICtrlSetData( $ctrlFileText, $var & "\" & "auto_"&GUICtrlRead( $ctrlNameText )&" "&GUICtrlRead( $ctrlEtternavnText )&".xml" )

			case $msg = $GUI_EVENT_CLOSE
					Exit

		EndSelect

	;GUICtrlSetData( $ctrlOutput, GUICtrlRead($ctrlName))

Until $msg = -1  ;> 0 ;= $Button_Ok

;GUICtrlSetData( $ctrlOutput, $msg & @CRLF )
;GUICtrlSetState($ver, $GUI_HIDE)

GUICtrlSetState($ctrlOutput, $GUI_FOCUS)

GUIDelete ()


Exit

Func ParseInput( $arrName )

	; if less than 3
	if $arrName[0] < 3 then
		MsgBox( 0, "x", "Name Surname f;dselsnummer")
		Return
	endif

	; Navn og Fnr
	if StringIsDigit($arrName[1]) then
		GUICtrlSetData( $ctrlFnrText, $arrName[1] )
		GUICtrlSetData( $ctrlNameText, $arrName[2] )
		GUICtrlSetData( $ctrlEtternavnText, "" )
		for $i = 3 to $arrName[0]
			GUICtrlSetData( $ctrlEtternavnText, GUICtrlRead($ctrlEtternavnText) & " " & $arrName[$i])
		Next
	Elseif StringIsDigit($arrName[$arrName[0]]) then
		GUICtrlSetData( $ctrlNameText, $arrName[1] )
		GUICtrlSetData( $ctrlEtternavnText, "" )
		for $i = 2 to $arrName[0]-1
			GUICtrlSetData( $ctrlEtternavnText, GUICtrlRead( $ctrlEtternavnText) & " " & $arrName[$i])
		Next
		GUICtrlSetData( $ctrlFnrText, $arrName[$arrName[0]])
	Else
		MsgBox( 0, "x", "Name Surname fodselsnummer")
		Return
	EndIf

	; Check Fnr
	$sFnr = GUICtrlRead( $ctrlFnrText )

	; Get f.dato
	if StringLen( $sFnr ) = 11 then
		$fdato = StringLeft( $sFnr, 6 )
	elseif StringLen( $sFnr ) = 10 then
		$fdato = "0"&StringLeft( $sFnr, 5 )
	Else
		MsgBox( 0, "Error", "Ugyldig Fnr" )
		Return
	EndIf

	if StringLeft( $fdato, 2) = 0 then
		MsgBox( 0, "Error", "Ugyldig Fnr" )
		Return
	elseif StringLeft( $fdato, 2) > 71 then
		MsgBox( 0, "Error", "Ugyldig Fnr" )
		Return
	Elseif StringLeft( $fdato, 2) > 31 then
		$fdato = (StringLeft($fdato,1)-4)&StringRight( $fdato,5)

	EndIf

	if StringMid( $fdato, 3,2) > 12 then
		MsgBox( 0, "Error", "Ugyldig Fnr" )
		Return
	endif

	GUICtrlSetData( $ctrlFdatoText,  StringRegExpReplace($fdato, "(\d\d)(\d\d)(\d\d)", "$3-$2-$1") )
	If @error = 0 Then

    Else

    EndIf


	; Get Sex
	; xxxxxx xx0xx - 0 - kvinne, 1 - mann
	if mod( StringMid( $sFnr, StringLen( $sFnr)-2,1 ), 2 ) = 0 then
		$fsex = "K "
	else
		$fsex = "M "
	EndIf
	GUICtrlSetData( $ctrlAlderText,  $fsex)


	; Check Alder
	; 000-499 - 1900
	; 500-749 - 1854-1899
	; 500-999 - 2000-2039
	; 900-999 - 1940-1999
	$arrFnr = StringRegExp( $sFnr, "(\d?\d)(\d\d)(\d\d)(\d\d)(\d)(\d\d)", 1)
	$pers = StringMid( $sFnr, StringLen($sFnr)-4,3 )
	$yr = $arrFnr[2]
	if $pers < 500 Then
		$yr = 1900
	elseif $yr < 40 Then
		$yr = 2000
	Elseif $pers >899 Then
		$yr = 1900
	ElseIf $pers < 750 Then
		$yr = 1800
	Else
		MsgBox( 0, "Error", "Ugyldig Fnr" )
		Return
	endif
GUICtrlSetData( $ctrlFdatoText, $arrFnr[2]+$yr&"-"&$arrFnr[1]&"-"&$arrFnr[0] )
GUICtrlSetData( $ctrlAlderText,  $fsex & "*" & _DateDiff("Y", $arrFnr[2]+$yr&"/"&$arrFnr[1]&"/"&$arrFnr[0],_NowCalc()) )

	;_ArrayDisplay ( $arrFnr )

	; Check Sex


	; Set GUID
	$sString = _GenerateGUID()
	GUICtrlSetData( $ctrlUidText , StringMid($sString,2,StringLen($sString)-2))

	; REad file
	$filetemplate = @WorkingDir & "\auto_.xml"
	$sString = FileRead( $filetemplate )
	If @error = -1 Then
		MsgBox(0, "Error", "Can't open file" & $filetemplate )
		Exit
	endif
	$sString = StringReplace( $sString, "#name#", GUICtrlRead( $ctrlNameText ) )
	$sString = StringReplace( $sString, "#surname#", GUICtrlRead( $ctrlEtternavnText ) )
	$sString = StringReplace( $sString, "#birthdate#", GUICtrlRead( $ctrlFdatoText ) )
	$sString = StringReplace( $sString, "#fnr#", GUICtrlRead( $ctrlFnrText ) )
	$sString = StringReplace( $sString, "#id#", GUICtrlRead( $ctrlUidText ) )
	$sString = StringReplace( $sString, "#sex#", GUICtrlRead( $ctrlAlderText ) )

	GUICtrlSetData( $ctrlOutput, $sString)

; Write file
$fileoutput = StringReplace( $filetemplate, "_", "_"& StringUpper(GUICtrlRead($ctrlNameText) & " " & GUICtrlRead($ctrlEtternavnText)), -1 )
if FileExists( $fileoutput ) then
	if MsgBox( 1, "Error", "File "& $fileoutput & " alredy exists. Overwite? ") = 2 then
		Exit
	EndIf
	FileDelete( $fileoutput )

EndIf

if FileWrite( $fileoutput, $sString ) = 0 then
		MsgBox( 0, "Error", "Can't write the file" )
		Exit
Endif

EndFunc



; -----------------------------------------------------------------------------
; Function: Exit handler
; -----------------------------------------------------------------------------
Func _GenerateGUID ()
$oScriptlet = ObjCreate ("Scriptlet.TypeLib")
Return $oScriptlet.Guid
EndFunc

Func Gender( $personalnummer)
Return BitAND( StringMid( $personalnummer, 3,1), 0x1)
EndFunc
