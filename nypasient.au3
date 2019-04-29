#region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile=nypasient.exe
#AutoIt3Wrapper_Run_Before=updversion.exe
#endregion ;**** Directives created by AutoIt3Wrapper_GUI ****
;~ #Region ;**** Directives created by AutoIt3Wrapper_GUI ****


; ================================
; 13/04/19 - initial prototype
; 15/04/19 -
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

Global $sString ; temp string
Global $filetemplate ; filename for temaplte
Global $fileoutput ; filename for output

Global $name
Global $surname
Global $fnr
Global $sex
Global $sexid
Global $fdato
Global $id
Global $age

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

Global $msg
Global $arrName

Global $typetext = 1


Do
	$msg = GUIGetMsg()

	Switch $msg
		Case $ctrlName

			$arrName = StringSplit(StringStripWS(GUICtrlRead($ctrlName), 7), " ")
			ParseInput($arrName)
			GUICtrlSetState($ctrlName, $GUI_FOCUS)

		Case $GUI_EVENT_SECONDARYDOWN

			; Run the GUI until the dialog is closed
			While 1

				Switch GUIGetMsg()

					Case $properItem
						$typetext = 1
						ExitLoop
					Case $upperItem
						$typetext = 2
						ExitLoop
				EndSwitch

			WEnd

			Switch $typetext
				Case 1
					GUICtrlSetData($ctrlFile, _StringProper1(GUICtrlRead($ctrlFile)))
				Case 2
					GUICtrlSetData($ctrlFile, StringUpper(GUICtrlRead($ctrlFile)))

			EndSwitch

		Case $GUI_EVENT_CLOSE
			Exit

	EndSwitch

Until $msg = -1 ;> 0 ;= $Button_Ok

GUIDelete()

Exit

; -----------------------------------------------------------------------------
; Function: Parse Input
; -----------------------------------------------------------------------------
Func ParseInput($arrName)

	Local $dd, $mm, $yy, $pers, $dnr

	; if less than 3 parameters
	If $arrName[0] < 3 Then
		MsgBox(0, "Error", "Navn Etternavn fnr")
		Return
	EndIf


; possible input
; ddmmyyxx0xx
; 4dmmyyxx0xx
; ddmmyy -> ddmmyy00900
; ddmmyyK -> ddmmyy00000
; ddmmyyM -> ddmmyy00100
; ddmmyyyy -> ddmmyy20900
; ddmmyyyyK -> ddmmyy10000
; ddmmyyyyM -> ddmmyy20100

	; if fnr/dnr goes first
	if StringRegExp( $arrName[1], "^([04][1-9]|[1256][0-9]|[37][01])(0[1-9]|1[012])(\d){7}$") Then
		$fnr = $arrName[1]
		$name = $arrName[2]
		$surname = $arrName[3]
		for $i = 4 to $arrName[0]
			$surname = $surname & " " & $arrName[$i]
		Next

		; Get Sex: xxxxxx xx0xx - even - kvinne, odd - mann
		$sexid = 2 - mod( StringMid( $fnr, 9, 1), 2) ; 1->1 mann, 0->2 kvinne

		; get pers number for age
		$pers = StringMid( $fnr, 7, 3)

	; if fnr goes last
	Elseif StringRegExp( $arrName[$arrName[0]], "^([04][1-9]|[1256][0-9]|[37][01])(0[1-9]|1[012])(\d){7}$") then
		$name = $arrName[1]
		$surname = $arrName[2]
		for $i = 3 to $arrName[0]-1
			$surname = $surname & " " & $arrName[$i]
		Next

		$fnr = $arrName[$arrName[0]]

		$sexid = 2 - mod( StringMid( $fnr, 9, 1), 2) ; 1->1-mann, 0->2-kvinne

		; get pers number for age
		$pers = StringMid( $fnr, 7, 3)

	; if only f.date goes last with(out) sex
	Elseif StringRegExp( $arrName[$arrName[0]], "^(0[1-9]|[12][0-9]|3[01])(0[1-9]|1[012])[0-9][0-9][kKmM]?$") then
		$name = $arrName[1]
		$surname = $arrName[2]
		for $i = 3 to $arrName[0]-1
			$surname = $surname & " " & $arrName[$i]
		Next

		$fnr = StringLeft( $arrName[$arrName[0]], 6 )

		; get sex: Mann=1, Kvinne=2
		if StringRegExp( $arrName[$arrName[0]], "\d*[mM]$" ) Then
			$sexid = 1
		Elseif StringRegExp( $arrName[$arrName[0]], "\d*[kK]$" ) Then
			$sexid = 2
		Else
			$sexid = 9
		EndIf

		$pers = "000"

; if only f.date goes last with(out) sex
	Elseif StringRegExp( $arrName[$arrName[0]], "^(0[1-9]|[12][0-9]|3[01])(0[1-9]|1[012])(18|19|20)[0-9][0-9][kKmM]?$") then
		$name = $arrName[1]
		$surname = $arrName[2]
		for $i = 3 to $arrName[0]-1
			$surname = $surname & " " & $arrName[$i]
		Next

		; get sex: Mann=1, Kvinne=2
		if StringRegExp( $arrName[$arrName[0]], "\d*[mM]$" ) Then
			$sexid = 1
		Elseif StringRegExp( $arrName[$arrName[0]], "\d*[kK]$" ) Then
			$sexid = 2
		Else
			$sexid = 9
		EndIf

		$fnr = StringLeft( $arrName[$arrName[0]],4 ) & StringMid( $arrName[$arrName[0]], 7, 2 )

		$pers = "000"
		$yy = StringMid( $arrName[$arrName[0]], 5,4) 

	Else
		MsgBox( 0, "Error", "Ugyldig Fnr")
		Return
	EndIf

	$dnr = False

	$dd = StringMid( $fnr, 1, 2)
	If $dd > 40 Then
		$dnr = True
		$dd -= 40
		If $dd < 10 Then $dd = "0" & $dd
	EndIf

	Local $mm = StringMid( $fnr, 3, 2)
	Local $yy = StringMid( $fnr, 5, 2)

	; Get Sex
	; xxxxxx xx0xx - 0 - kvinne, 1 - mann
	switch $sexid
		case 1
			$sex = "Mann"
		case 2
			$sex = "Kvinne"
		case Else
			$sex = "Ukjent"
	EndSwitch

	; Check Alder
	; 000-499 - 1900
	; 500-749 - 1854-1899
	; 500-999 - 2000-2039
	; 900-999 - 1940-1999

	if $yy < 100 then 
		If $pers < 500 Then
			$yy = 1900 + $yy
		ElseIf $yy < 40 Then
			$yy = 2000 + $yy
		ElseIf $pers > 899 Then
			$yy = 1900 + $yy
		ElseIf $pers < 750 and $yy > 53 Then
			$yy = 1800 + $yy
		Else
			MsgBox(0, "Error", "Ugyldig Fnr")
			Return
		EndIf
	Endif 

	; Get fdato
	$fdato = $yy & "-" & $mm & "-" & $dd

	; Get Age
	$age = _DateDiff("Y", $yy & "/" & $mm & "/" & $dd, _NowCalc())
	if @error <> 0 Then
		MsgBox(0, "Error", "Ugyldig f"&ChrW(248)&"dselsdato" & $filetemplate)
		Return
	EndIf


	; Set GUID
	$id = _GenerateGUID()
	$id = StringMid($id, 2, StringLen($id) - 2)

	; Read file
	$filetemplate = @WorkingDir & "\auto_.xml"
	$sString = FileRead($filetemplate)
	If @error = -1 Then
		MsgBox(0, "Error", "Can't open file" & $filetemplate)
		Exit
	EndIf
	$sString = StringReplace($sString, "#name#", _StringProper1($name))
	$sString = StringReplace($sString, "#surname#", _StringProper1($surname))
	$sString = StringReplace($sString, "#birthdate#", $fdato)
	$sString = StringReplace($sString, "#fnr#", $fnr)
	$sString = StringReplace($sString, "#id#", $id)
	$sString = StringReplace($sString, "#sex#", $sex)
	$sString = StringReplace($sString, "#sexid#", $sexid)

	; change type
	if $dnr Then
		$sString = StringReplace($sString, '="FNR"', '="DNR"')
		$sString = StringReplace($sString, '="F'& ChrW(248) &'dselsnummer"', '="D-nummer"')

	Else
		$sString = StringReplace($sString, '="DNR"', '="FNR"')
		$sString = StringReplace($sString, '="D-nummer"', '="F'& ChrW(248) &'dselsnummer"')

	EndIf

	; Write file
	$fileoutput = StringReplace($filetemplate, "_", "_" & _StringProper1($name & " " & $surname), -1)
	If $typetext = 2 Then
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

	If FileWrite($file, $sString) = 0 Then
		MsgBox(0, "Error", "Can't write the file")
		Return
	EndIf

	FileClose($file)

	$sString = ""
	$sString &= "Name : " & _StringProper1($name & " " & $surname) & @CRLF
	$sString &= "Fnr  : " & $fnr & @CRLF
	$sString &= "Fdato: " & $fdato & "  (" & $sex & "-" & $age & ")" & @CRLF
	;$sString &= "Guid : " & $id & @CRLF & @CRLF
	$sString &= @CRLF
	$sString &= "File : " & $fileoutput & @CRLF

	MsgBox(0, "Pasient successfully created", $sString)
EndFunc   ;==>ParseInput



; -----------------------------------------------------------------------------
; Function: Exit handler
; -----------------------------------------------------------------------------
Func _GenerateGUID()
	Local $oScriptlet = ObjCreate("Scriptlet.TypeLib")
	Return $oScriptlet.Guid
EndFunc   ;==>_GenerateGUID

Func Gender($personalnummer)
	Return BitAND(StringMid($personalnummer, 3, 1), 0x1)
EndFunc   ;==>Gender
;~ ; -----------------------------------------------------------------------------
;~ ; Function: Exit handler
;~ ; -----------------------------------------------------------------------------
;~ Func MyExitFunc()

;~    ;Msgbox(0,"Exit","Exiting..." & @CRLF 	)

;~    if isObj($oIE)  Then
;~ 		ConsoleWrite( "*** Closing IE" & @CRLF )
;~ 		_IEQuit($oIE)
;~    EndIf

;~    if isObj($oSummary) then
;~ 		ConsoleWrite( "*** Closing Excel" & @CRLF )
;~ 		_ExcelBookClose( $oSummary, 0 )
;~    EndIf

;~    if isObj($tdc) > 0 Then
;~ 		ConsoleWrite( "*** Closing TDC" & @CRLF )
;~ 		$tdc.DisconnectProject()
;~ 		$tdc.ReleaseConnection()
;~    EndIf

;~ Endfunc
#cs
	namespace LIB.Validation
	{
	public class FodselsNummer
	{
	public enum Gender : int
	{
	Female = 0,
	Male = 1
	}

	public static Gender Check(string fnr)
	{
	if (fnr.Length != 11)
	throw new InvalidLengthException();

	// Valid date?  D-Number = +4
	string Date = (fnr[0] <= '3') ? fnr.Substring(0, 6) : ((fnr[0] - '4') + fnr.Substring(1, 5));
	DateTime tmp;
	if (DateTime.TryParseExact(Date, "ddMMyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out tmp) == false)
	throw new InvalidDateException();

	int[] n = new int[11];
	int tmp2 = 0;
	for (int i = 0; i < 11; i++)
	if (int.TryParse(fnr[i].ToString(), out tmp2))
	n[i] = tmp2;
	else
	throw new InvalidCharactersException();

	// Control number 1
	int k1 = 11 - (3 * n[0] + 7 * n[1] + 6 * n[2] + 1 * n[3] + 8 * n[4] + 9 * n[5] + 4 * n[6] + 5 * n[7] + 2 * n[8]) % 11;
	if (k1 == 11) k1 = 0;

	if (k1 == 10 || k1 != n[9])
	throw new InvalidControlNumberException();

	// Control number 2
	int k2 = 11 - (5 * n[0] + 4 * n[1] + 3 * n[2] + 2 * n[3] + 7 * n[4] + 6 * n[5] + 5 * n[6] + 4 * n[7] + 3 * n[8] + 2 * k1) % 11;
	if (k2 == 11) k2 = 0;

	if (k2 == 10 || k2 != n[10])
	throw new InvalidControlNumberException();

	// Sex
	return (Gender)(n[8] & 1);
	}

	public class InvalidFodselsNumberException : ApplicationException
	{
	}

	public class InvalidCharactersException : InvalidFodselsNumberException
	{
	}

	public class InvalidControlNumberException : InvalidFodselsNumberException
	{
	}

	public class InvalidDateException : InvalidFodselsNumberException
	{
	}

	public class InvalidLengthException : InvalidFodselsNumberException
	{
	}
	}
	}
#ce

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
				If StringRegExp($s_CurChar, '[a-zA-ZА-яљњћџ' & ChrW(198) & ChrW(230) & ChrW(216) & ChrW(248) & ChrW(197) & ChrW(229) & ']') Then
					$s_CurChar = StringUpper($s_CurChar)
					$CapNext = 0
				EndIf
			Case Not StringRegExp($s_CurChar, '[a-zA-ZА-яљњћџ' & ChrW(198) & ChrW(230) & ChrW(216) & ChrW(248) & ChrW(197) & ChrW(229) & ']')
				$CapNext = 1
			Case Else
				$s_CurChar = StringLower($s_CurChar)
		EndSelect
		$s_nStr &= $s_CurChar
	Next
	Return $s_nStr
EndFunc   ;==>_StringProper1
