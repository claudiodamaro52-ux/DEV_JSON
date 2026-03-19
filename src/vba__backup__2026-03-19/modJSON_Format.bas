Attribute VB_Name = "modJSON_Format"
'Attribute VB_Name = "modJSON_Format"
Option Explicit

' ============================================================
' JSON_Indentar
' ============================================================
Public Function JSON_Indentar(ByRef prmTxt As String) As String
    Dim s As String
    Dim i As Long
    Dim c As String
    Dim resultado As String
    Dim nivel As Long
    Dim dentroString As Boolean

    s = prmTxt
    resultado = ""
    nivel = 0
    dentroString = False

    For i = 1 To Len(s)
        c = Mid$(s, i, 1)

        If c = """" Then
            dentroString = Not dentroString
            resultado = resultado & c

        ElseIf dentroString Then
            resultado = resultado & c

        Else
            Select Case c
                Case "{", "["
                    resultado = resultado & c & vbCrLf
                    nivel = nivel + 1
                    resultado = resultado & String(nivel * 4, " ")

                Case "}", "]"
                    resultado = resultado & vbCrLf
                    nivel = nivel - 1
                    If nivel < 0 Then nivel = 0
                    resultado = resultado & String(nivel * 4, " ") & c

                Case ","
                    resultado = resultado & c & vbCrLf & String(nivel * 4, " ")

                Case ":"
                    resultado = resultado & ": "

                Case " ", vbCr, vbLf, vbTab
                    ' ignora

                Case Else
                    resultado = resultado & c
            End Select
        End If
    Next i

    JSON_Indentar = resultado
End Function
