Attribute VB_Name = "modJSON_Templates"
Option Explicit

' ============================================================
' JSON_TemplatePares
' Entrada: texto com linhas "caminho = valor"
' Saída:  texto com linhas "caminho = \"...\""
'
' modo:
'   - "vazios":   sempre "" (aspas duplas)
'   - "variavel": sempre "v_<chaveFinal>"
'   - "mascara":  tenta inferir máscara pelo valor; se falhar, vira "v_<chaveFinal>"
'
' Regras:
' - null -> "" (como você pediu)
' - sempre retorna valor entre aspas duplas no output
' ============================================================
Public Function JSON_TemplatePares(ByRef prmPares As String, ByRef prmModo As String) As String
    Dim linhas() As String
    Dim i As Long
    Dim ln As String
    Dim p As Long

    Dim caminho As String
    Dim valor As String
    Dim chaveFinal As String

    Dim out As String
    Dim modo As String
    Dim exemplo As String

    modo = LCase$(Trim$(prmModo))
    If modo = "" Then modo = "vazios"

    linhas = Split(prmPares, vbCrLf)
    out = ""

    For i = 0 To UBound(linhas)
        ln = Trim$(linhas(i))
        If ln = "" Then GoTo Prox

        p = InStr(1, ln, "=", vbBinaryCompare)
        If p <= 0 Then GoTo Prox

        caminho = Trim$(Left$(ln, p - 1))
        valor = Trim$(Mid$(ln, p + 1))

        chaveFinal = JSON_GetChaveFinal(caminho)

        ' null: no modo mascara, retornar "" direto (sem cair no fallback variavel)
        If LCase$(valor) = "null" Then
            If modo = "mascara" Or modo = "vazios" Then
                out = out & caminho & " = " & """" & "" & """" & vbCrLf
                GoTo Prox
            ElseIf modo = "variavel" Then
                out = out & caminho & " = " & """" & ("v_" & chaveFinal) & """" & vbCrLf
                GoTo Prox
            End If
        End If

        Select Case modo
            Case "vazios"
                exemplo = ""

            Case "variavel"
                exemplo = "v_" & chaveFinal

            Case "mascara"
                exemplo = JSON_InferirMascaraExemploComChave(valor, chaveFinal)
                If exemplo = "" Then
                    exemplo = "v_" & chaveFinal
                End If

            Case Else
                exemplo = ""
        End Select

        out = out & caminho & " = " & """" & exemplo & """" & vbCrLf

Prox:
    Next i

    JSON_TemplatePares = out
End Function


' pega a "chave final" do caminho:
' - pedido.itens[0].sku  -> sku
' - pedido.valores[1]    -> valores (não tem .; nesse caso usa o nome antes do [)
Public Function JSON_GetChaveFinal(ByRef caminho As String) As String
    Dim s As String
    Dim pDot As Long
    Dim pBr As Long

    s = caminho

    pDot = InStrRev(s, ".")
    If pDot > 0 Then s = Mid$(s, pDot + 1)

    pBr = InStr(1, s, "[", vbBinaryCompare)
    If pBr > 0 Then s = Left$(s, pBr - 1)

    JSON_GetChaveFinal = s
End Function
' ============================================================
' JSON_InferirMascaraExemplo
' Recebe valor "bruto" do flatten (pode ser: 10, 9.9, "abc", true, false, "")
' Retorna um exemplo preenchido (SEM aspas), ou "" se não inferir.
' ============================================================



Private Function JSON_OnlyDigits(ByRef s As String) As String
    Dim i As Long
    Dim c As String
    Dim out As String

    out = ""
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c Like "[0-9]" Then out = out & c
    Next i

    JSON_OnlyDigits = out
End Function

Private Function JSON_LooksLikeDateYYYYMMDD(ByRef s As String) As Boolean
    ' bem simples: ####-##-##
    If Len(s) <> 10 Then
        JSON_LooksLikeDateYYYYMMDD = False
        Exit Function
    End If

    If Mid$(s, 5, 1) <> "-" Then GoTo Nope
    If Mid$(s, 8, 1) <> "-" Then GoTo Nope

    If Not (Mid$(s, 1, 4) Like "####") Then GoTo Nope
    If Not (Mid$(s, 6, 2) Like "##") Then GoTo Nope
    If Not (Mid$(s, 9, 2) Like "##") Then GoTo Nope

    JSON_LooksLikeDateYYYYMMDD = True
    Exit Function
Nope:
    JSON_LooksLikeDateYYYYMMDD = False
End Function

Private Function JSON_LooksLikeDateTimeISO(ByRef s As String) As Boolean
    ' bem simples: começa com data e tem "T"
    If Len(s) < 19 Then
        JSON_LooksLikeDateTimeISO = False
        Exit Function
    End If

    If Not JSON_LooksLikeDateYYYYMMDD(Left$(s, 10)) Then
        JSON_LooksLikeDateTimeISO = False
        Exit Function
    End If

    If Mid$(s, 11, 1) <> "T" Then
        JSON_LooksLikeDateTimeISO = False
        Exit Function
    End If

    JSON_LooksLikeDateTimeISO = True
End Function

Private Function JSON_IsNumericLoose(ByRef s As String) As Boolean
    Dim t As String
    t = Trim$(s)

    ' remove vírgula final (caso venha do JSON indentado) e aspas
    If Right$(t, 1) = "," Then t = Left$(t, Len(t) - 1)
    t = Replace(t, """", "")

    JSON_IsNumericLoose = IsNumeric(t)
End Function
Private Function JSON_InferirMascaraExemploComChave(ByRef valorBruto As String, ByRef chaveFinal As String) As String
    Dim v As String
    Dim s As String
    Dim digits As String
    Dim k As String

    v = Trim$(valorBruto)

    ' extrai string (se vier "...")
    If Len(v) >= 2 Then
        If Left$(v, 1) = """" And Right$(v, 1) = """" Then
            s = Mid$(v, 2, Len(v) - 2)
        Else
            s = v
        End If
    Else
        s = v
    End If

    If s = "" Then
        JSON_InferirMascaraExemploComChave = ""
        Exit Function
    End If

    k = LCase$(Trim$(chaveFinal))

    ' 1) PRIMEIRO: formatos ISO
    If JSON_LooksLikeDateTimeISO(s) Then
        JSON_InferirMascaraExemploComChave = "2000-01-01T00:00:00"
        Exit Function
    End If

    If JSON_LooksLikeDateYYYYMMDD(s) Then
        JSON_InferirMascaraExemploComChave = "2000-01-01"
        Exit Function
    End If

    digits = JSON_OnlyDigits(s)

    ' 2) Telefones por chave (prioridade)
    ' Se a chave indicar telefone/celular/whatsapp e tiver 10 ou 11 dígitos, retorna máscara de telefone.
    If (InStr(1, k, "fone", vbBinaryCompare) > 0) Or _
       (InStr(1, k, "tel", vbBinaryCompare) > 0) Or _
       (InStr(1, k, "cel", vbBinaryCompare) > 0) Or _
       (InStr(1, k, "whats", vbBinaryCompare) > 0) Or _
       (InStr(1, k, "zap", vbBinaryCompare) > 0) Or _
       (InStr(1, k, "contato", vbBinaryCompare) > 0) Then

        If Len(digits) = 10 Then
            JSON_InferirMascaraExemploComChave = "(00) 0000-0000"
            Exit Function
        End If

        If Len(digits) = 11 Then
            JSON_InferirMascaraExemploComChave = "(00) 00000-0000"
            Exit Function
        End If
    End If

    ' 3) Documentos
    If Len(digits) = 11 Then
        JSON_InferirMascaraExemploComChave = "000.000.000-00"  ' CPF
        Exit Function
    End If

    If Len(digits) = 14 Then
        JSON_InferirMascaraExemploComChave = "00.000.000/0000-00"  ' CNPJ
        Exit Function
    End If

    If Len(digits) = 8 Then
        JSON_InferirMascaraExemploComChave = "00000-000"  ' CEP
        Exit Function
    End If

    ' 4) Telefone fixo (heurística final, sem depender da chave)
    If Len(digits) = 10 Then
        JSON_InferirMascaraExemploComChave = "(00) 0000-0000"
        Exit Function
    End If

    ' 5) Numérico genérico
    If JSON_IsNumericLoose(v) Then
        JSON_InferirMascaraExemploComChave = "0.00"
        Exit Function
    End If

    JSON_InferirMascaraExemploComChave = ""
End Function

