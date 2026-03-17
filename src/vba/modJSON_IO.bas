Attribute VB_Name = "modJSON_IO"
'Attribute VB_Name = "modJSON_IO"
Option Explicit

' ============================================================
'  NormalizarEntradaExterna
' ============================================================
Public Function NormalizarEntradaExterna(ByRef prmTxt As String) As String
    Dim s As String
    Dim i As Long
    Dim ch As Integer
    Dim out As String
    Dim prevWasCR As Boolean

    s = prmTxt
    out = ""
    prevWasCR = False

    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))

        ' "|" -> "/" (pipe é reservado)
        If ch = 124 Then
            out = out & "/"
            prevWasCR = False
            GoTo Proximo
        End If

        ' TAB -> espaço
        If ch = 9 Then
            out = out & " "
            prevWasCR = False
            GoTo Proximo
        End If

        ' Quebras: CR/LF -> CRLF
        If ch = 13 Then
            out = out & vbCrLf
            prevWasCR = True
            GoTo Proximo
        End If

        If ch = 10 Then
            If Not prevWasCR Then out = out & vbCrLf
            prevWasCR = False
            GoTo Proximo
        End If

        prevWasCR = False

        ' Remove controle invisível (obrigatório): < 32
        If ch < 32 Then GoTo Proximo

        out = out & ChrW(ch)

Proximo:
    Next i

    NormalizarEntradaExterna = Trim$(out)
End Function

