Attribute VB_Name = "modJSON_Normalize"
'Attribute VB_Name = "modJSON_Normalize"
Option Explicit

' ============================================================
' JSON_Normalizar (FASE 2.2)
' ============================================================
Public Function JSON_Normalizar(ByRef prmTxt As String) As String
    Dim s As String
    Dim i As Long
    Dim c As String
    Dim dentroString As Boolean
    Dim resultado As String

    Dim strBuf As String

    s = prmTxt
    resultado = ""
    dentroString = False
    strBuf = ""

    For i = 1 To Len(s)
        c = Mid$(s, i, 1)

        If c = """" Then
            If dentroString Then
                If FechaStringProvavel(s, i) Then
                    dentroString = False

                    Dim nx As String
                    nx = NextNonWsChar(s, i + 1)

                    If nx = ":" Then
                        strBuf = Key_Normalizar(strBuf)
                    End If

                    resultado = resultado & """" & strBuf & """"
                    strBuf = ""
                Else
                    strBuf = strBuf & "'"
                End If
            Else
                dentroString = True
                strBuf = ""
            End If

        ElseIf dentroString Then
            strBuf = strBuf & c

        Else
            Select Case c
                Case " ", vbCr, vbLf, vbTab
                    ' ignora
                Case Else
                    resultado = resultado & c
            End Select
        End If
    Next i

    If dentroString Then
        resultado = resultado & """" & strBuf & """"
    End If

    resultado = MAP2_ConverterColchetesEmChaves(resultado)
    JSON_Normalizar = resultado
End Function

' ============================================================
' MAP-2 (tolerante)
' ============================================================
Private Function MAP2_ConverterColchetesEmChaves(ByVal txt As String) As String
    Dim i As Long
    Dim c As String
    Dim dentroString As Boolean

    Dim stackStart() As Long
    Dim stackHasPair() As Boolean
    Dim stackFirstTokenIsObj() As Boolean
    Dim stackFirstTokenSeen() As Boolean
    Dim top As Long

    Dim chars() As String

    If Len(txt) = 0 Then
        MAP2_ConverterColchetesEmChaves = txt
        Exit Function
    End If

    ReDim chars(1 To Len(txt))
    For i = 1 To Len(txt)
        chars(i) = Mid$(txt, i, 1)
    Next i

    ReDim stackStart(0 To 200)
    ReDim stackHasPair(0 To 200)
    ReDim stackFirstTokenIsObj(0 To 200)
    ReDim stackFirstTokenSeen(0 To 200)
    top = -1

    dentroString = False

    For i = 1 To Len(txt)
        c = chars(i)

        If c = """" Then
            dentroString = Not dentroString
            GoTo ContinueFor
        End If

        If dentroString Then GoTo ContinueFor

        If c = "[" Then
            top = top + 1
            stackStart(top) = i
            stackHasPair(top) = False
            stackFirstTokenIsObj(top) = False
            stackFirstTokenSeen(top) = False
            GoTo ContinueFor
        End If

        If top >= 0 Then
            If Not stackFirstTokenSeen(top) Then
                If c <> " " And c <> vbTab And c <> vbCr And c <> vbLf Then
                    stackFirstTokenSeen(top) = True
                    If c = "{" Then stackFirstTokenIsObj(top) = True
                End If
            End If
        End If

        If c = ":" Then
            If top >= 0 Then
                If MAP2_ColonEhParDeObjeto(chars, i) Then
                    stackHasPair(top) = True
                End If
            End If
            GoTo ContinueFor
        End If

        If c = "]" Then
            If top >= 0 Then
                If stackHasPair(top) And (Not stackFirstTokenIsObj(top)) Then
                    chars(stackStart(top)) = "{"
                    chars(i) = "}"
                End If
                top = top - 1
            End If
            GoTo ContinueFor
        End If

ContinueFor:
    Next i

    Dim out As String
    out = ""
    For i = 1 To Len(txt)
        out = out & chars(i)
    Next i

    MAP2_ConverterColchetesEmChaves = out
End Function

Private Function MAP2_ColonEhParDeObjeto(ByRef chars() As String, ByVal posColon As Long) As Boolean
    Dim j As Long

    j = posColon - 1
    Do While j >= LBound(chars)
        If chars(j) = " " Or chars(j) = vbTab Or chars(j) = vbCr Or chars(j) = vbLf Then
            j = j - 1
        Else
            Exit Do
        End If
    Loop

    If j < LBound(chars) Or chars(j) <> """" Then
        MAP2_ColonEhParDeObjeto = False
        Exit Function
    End If

    j = j - 1
    Do While j >= LBound(chars)
        If chars(j) = """" Then
            MAP2_ColonEhParDeObjeto = True
            Exit Function
        End If
        j = j - 1
    Loop

    MAP2_ColonEhParDeObjeto = False
End Function

Private Function FechaStringProvavel(ByRef s As String, ByVal posAspa As Long) As Boolean
    Dim nx As String
    nx = NextNonWsChar(s, posAspa + 1)

    Select Case nx
        Case ",", "}", "]", ":", vbNullString
            FechaStringProvavel = True
        Case Else
            FechaStringProvavel = False
    End Select
End Function

Private Function NextNonWsChar(ByRef s As String, ByVal startPos As Long) As String
    Dim j As Long
    Dim ch As String

    For j = startPos To Len(s)
        ch = Mid$(s, j, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then
            NextNonWsChar = ch
            Exit Function
        End If
    Next j

    NextNonWsChar = vbNullString
End Function

' ============================================================
' Normalizaзгo de chaves (ASCII + "_")
' ============================================================
Private Function Key_Normalizar(ByVal k As String) As String
    Dim s As String
    Dim i As Long
    Dim c As String
    Dim out As String
    Dim lastWasUnd As Boolean

    s = Trim$(k)
    s = Key_RemoverAcentos(s)

    out = ""
    lastWasUnd = False

    For i = 1 To Len(s)
        c = Mid$(s, i, 1)

        If c Like "[A-Za-z0-9]" Then
            out = out & c
            lastWasUnd = False

        ElseIf c = " " Or c = "-" Or c = "." Or c = "/" Or c = "\" Or c = "_" Then
            If Len(out) > 0 Then
                If Not lastWasUnd Then
                    out = out & "_"
                    lastWasUnd = True
                End If
            End If

        Else
            ' ignora
        End If
    Next i

    Do While Len(out) > 0 And Right$(out, 1) = "_"
        out = Left$(out, Len(out) - 1)
    Loop

    If out = "" Then out = "K"
    If Left$(out, 1) Like "[0-9]" Then out = "_" & out

    Key_Normalizar = out
End Function

Private Function Key_RemoverAcentos(ByVal s As String) As String
    Dim a As Variant, b As Variant
    Dim i As Long

    a = Array("б", "а", "г", "в", "д", "Б", "А", "Г", "В", "Д", _
              "й", "и", "к", "л", "Й", "И", "К", "Л", _
              "н", "м", "о", "п", "Н", "М", "О", "П", _
              "у", "т", "х", "ф", "ц", "У", "Т", "Х", "Ф", "Ц", _
              "ъ", "щ", "ы", "ь", "Ъ", "Щ", "Ы", "Ь", _
              "з", "З", "с", "С")

    b = Array("a", "a", "a", "a", "a", "A", "A", "A", "A", "A", _
              "e", "e", "e", "e", "E", "E", "E", "E", _
              "i", "i", "i", "i", "I", "I", "I", "I", _
              "o", "o", "o", "o", "o", "O", "O", "O", "O", "O", _
              "u", "u", "u", "u", "U", "U", "U", "U", _
              "c", "C", "n", "N")

    For i = LBound(a) To UBound(a)
        s = Replace(s, a(i), b(i))
    Next i

    Key_RemoverAcentos = s
End Function

