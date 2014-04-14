;ConsoleWrite(URLEncode("123abc!@#$%^&*()_+ ") & @crlf)

;ConsoleWrite(URLEncode("S,C,c,0")&"."&URLEncode("00,0")&"."&URLEncode("00") & @crlf)

ConsoleWrite(URLEncode("Univ. of Nebraska-Lincoln") & @crlf)

Func URLEncode($urlText)
    $url = ""
    For $i = 1 To StringLen($urlText)
        $acode = Asc(StringMid($urlText, $i, 1))
        Select
            Case ($acode >= 48 And $acode <= 57) Or _
                    ($acode >= 65 And $acode <= 90) Or _
                    ($acode >= 97 And $acode <= 122)
                $url = $url & StringMid($urlText, $i, 1)
            Case $acode = 32
                $url = $url & "+"
			Case $acode = 46
                $url = $url & "."
            Case Else
                $url = $url & "%" & Hex($acode, 2)
        EndSelect
    Next
    Return $url
EndFunc   ;==>URLEncode