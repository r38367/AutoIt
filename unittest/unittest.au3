#cs ----------------------------------------------------------------------------

 Micro unit test framework


#ce ----------------------------------------------------------------------------

#include "..\constants.au3"

Func UTAssertTrue(Const $bool, Const $msg = "Assert Failure", Const $erl = @ScriptLineNumber)
    If NOT $bool Then
        DoTestFail($msg, $erl)
    EndIf

EndFunc

Func UTAssertFalse(Const $bool, Const $msg = "Assert Failure", Const $erl = @ScriptLineNumber)
    Return UTAssertTrue( not $bool, $msg, $erl )
EndFunc

Func UTAssertEqual(Const $a, Const $b, Const $msg = "Assert Failure", Const $erl = @ScriptLineNumber)
   	If $a <> $b Then
		DoTestFail( StringFormat( $msg & " ( was=[%s], expected=[%s] )", $a, $b ), $erl)
	Else
		DoTestPass()
	EndIf

EndFunc

Func DoTestPass()
	;$iTotalPass += 1
	;ConsoleWrite( "." )
EndFunc

Func DoTestFail( $msg, $erl )
	;$iTotalFail += 1
	Local $message = @ScriptName & " (" & $erl & ") := " & $msg & @LF
	;StringFormat( "%-15s %s", $iTotalAssertions, "Error in " & $sCurrentTestName & " -> " & $msg  )
	;_ArrayAdd( $aResults, $message )
	ConsoleWrite( $message )
EndFunc


; ==============================
; Unis test runner
;
; NOTE: Unit test file shall be name <function under test>_UT.au3
;
; ==============================

Local const $function_under_test = StringReplace( @ScriptName, "_UT.au3", "")

If StringInStr(@ScriptName, $function_under_test ) Then
	ConsoleWrite( @CRLF & "Testing " & @ScriptFullPath & @CRLF  )
    testMyFunc1()
Endif

exit

