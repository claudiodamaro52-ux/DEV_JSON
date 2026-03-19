Attribute VB_Name = "modJSON_Utils"
'Attribute VB_Name = "modJSON_Utils"
Option Explicit

' ============================================================
'  modJSON_Utils
'  Helpers genéricos compartilhados
' ============================================================

Public Function JSON_Indent(ByVal level As Long) As String
    If level < 0 Then level = 0
    JSON_Indent = String(level * 4, " ")
End Function

' Inferência simples DEV_JSON:
' - "" => ""
' - "true/false/null" => literal
' - numeric => literal
' - caso contrário => string com aspas; aspas internas viram '
Public Function JSON_ValorJSON(ByVal rawValue As String) As String
    Dim v As String
    v = Trim$(rawValue)

    If v = "" Then
        JSON_ValorJSON = """"""
        Exit Function
    End If

    If Left$(v, 1) = """" And Right$(v, 1) = """" And Len(v) >= 2 Then
        JSON_ValorJSON = v
        Exit Function
    End If

    Select Case LCase$(v)
        Case "true", "false", "null"
            JSON_ValorJSON = LCase$(v)
            Exit Function
    End Select

    If IsNumeric(v) Then
        JSON_ValorJSON = v
        Exit Function
    End If

    v = Replace(v, """", "'")
    JSON_ValorJSON = """" & v & """"
End Function

