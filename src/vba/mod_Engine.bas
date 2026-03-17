Option Explicit
' ============================================================
'  MÓDULO: modJSON_Engine  (FASE 1 - CONTRATO v0.6)
'  Objetivo desta fase:
'   - Garantir contrato Texto->Texto nas functions do engine
'   - Entrada externa SEMPRE passa por normalização
'   - "Pipe" (|) é reservado do Integrador: remove/substitui na entrada
'   - Remover caracteres de controle invisíveis (obrigatório)
'
'  Regra: As functions abaixo retornam SOMENTE PAYLOAD (sem "OP|...|")
' ============================================================

' ============================================================
' NormalizarEntradaExterna
'  - Obrigatório para toda entrada externa (usuário/API/arquivo)
'  - Padroniza CRLF, TAB->espaço, remove chars de controle invisíveis
'  - Substitui "|" por "/" (pipe é reservado)
'  - Não usa objetos; tudo texto
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

        ' "|" -> "/"
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


' ============================================================
' JSON_Normalizar (FASE 2.2)
'  - Minifica JSON removendo espaços/quebras fora de strings
'  - Dentro de strings:
'       * mantém conteúdo
'       * corrige aspas internas: " -> ' (heurística)
'  - Normaliza CHAVES (strings em posição de chave) para ASCII + "_"
'  - Não reordena (preserva ordem)
'  - Retorna SOMENTE payload
' ============================================================
Public Function JSON_Normalizar(ByRef prmTxt As String) As String
    Dim s As String
    Dim i As Long
    Dim c As String
    Dim dentroString As Boolean
    Dim resultado As String

    ' buffer da string atual (sem aspas)
    Dim strBuf As String
    Dim strStartedAt As Long

    s = prmTxt
    resultado = ""
    dentroString = False
    strBuf = ""
    strStartedAt = 0

    For i = 1 To Len(s)
        c = Mid$(s, i, 1)

        If c = """" Then
            If dentroString Then
                ' Estamos dentro de string: decidir se é fechamento ou aspa interna
                If FechaStringProvavel(s, i) Then
                    ' Fechou string: decidir se é "chave" (próximo token = :)
                    dentroString = False

                    Dim nx As String
                    nx = NextNonWsChar(s, i + 1)

                    If nx = ":" Then
                        ' é chave
                        strBuf = Key_Normalizar(strBuf)
                    End If

                    ' escreve string final
                    resultado = resultado & """" & strBuf & """"

                    ' limpa buffer
                    strBuf = ""
                    strStartedAt = 0
                Else
                    ' Aspas internas => simplificação DEV_JSON
                    strBuf = strBuf & "'"
                End If
            Else
                ' Abre string: não escreve agora; bufferiza
                dentroString = True
                strBuf = ""
                strStartedAt = i
            End If

        ElseIf dentroString Then
            ' Dentro de string: bufferiza (não escreve direto)
            strBuf = strBuf & c

        Else
            ' Fora de string: minifica
            Select Case c
                Case " ", vbCr, vbLf, vbTab
                    ' ignora
                Case Else
                    resultado = resultado & c
            End Select
        End If
    Next i

    ' Se terminou ainda dentro de string (entrada quebrada), fecha do jeito mais seguro
    If dentroString Then
        ' Aplica correção mínima: não considera chave aqui (não há como ter :)
        resultado = resultado & """" & strBuf & """"
    End If
    resultado = MAP2_ConverterColchetesEmChaves(resultado)
    JSON_Normalizar = resultado

End Function

' ============================================================
' MAP-2 (tolerante):
'  Converte blocos [ ... ] em { ... } quando o conteúdo parece
'  um "objeto escrito com colchetes", ex:
'    ["Altura":"1,20","Largura":"2,40"]
'  Proteção (decisão A):
'   - NÃO converte se o primeiro token após "[" for "{"
'     (arrays de objetos válidos)
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

        ' marca primeiro token do array
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

    ' pula whitespace (por segurança)
    j = posColon - 1
    Do While j >= LBound(chars)
        If chars(j) = " " Or chars(j) = vbTab Or chars(j) = vbCr Or chars(j) = vbLf Then
            j = j - 1
        Else
            Exit Do
        End If
    Loop

    ' precisa haver " imediatamente antes do :
    If j < LBound(chars) Or chars(j) <> """" Then
        MAP2_ColonEhParDeObjeto = False
        Exit Function
    End If

    ' procura a aspa de abertura anterior
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
' ============================================================
' Helpers - detecção de fechamento de string (heurística)
' ============================================================
Private Function FechaStringProvavel(ByRef s As String, ByVal posAspa As Long) As Boolean
    Dim nx As String
    nx = NextNonWsChar(s, posAspa + 1)

    ' tokens típicos após o fechamento de string em JSON
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
' Normalização de chaves (ASCII + "_", sem forçar caixa)
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

        ElseIf c = " " Or c = "-" Or c = "." Or c = "/" Or c = "\" Then
            If Len(out) > 0 Then
                If Not lastWasUnd Then
                    out = out & "_"
                    lastWasUnd = True
                End If
            End If

        ElseIf c = "_" Then
            If Len(out) > 0 Then
                If Not lastWasUnd Then
                    out = out & "_"
                    lastWasUnd = True
                End If
            End If

        Else
            ' ignora caracteres não permitidos
        End If
    Next i

    ' remove underscores finais
    Do While Len(out) > 0 And Right$(out, 1) = "_"
        out = Left$(out, Len(out) - 1)
    Loop

    If out = "" Then out = "K"

    ' se começar por número, prefixa "_"
    If Left$(out, 1) Like "[0-9]" Then out = "_" & out

    Key_Normalizar = out
End Function

Private Function Key_RemoverAcentos(ByVal s As String) As String
    Dim a As Variant, b As Variant
    Dim i As Long

    a = Array("á", "à", "ã", "â", "ä", "Á", "À", "Ã", "Â", "Ä", _
              "é", "è", "ê", "ë", "É", "È", "Ê", "Ë", _
              "í", "ì", "î", "ï", "Í", "Ì", "Î", "Ï", _
              "ó", "ò", "õ", "ô", "ö", "Ó", "Ò", "Õ", "Ô", "Ö", _
              "ú", "ù", "û", "ü", "Ú", "Ù", "Û", "Ü", _
              "ç", "Ç", "ñ", "Ñ")

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


' ============================================================
' JSON_Indentar
'  - Indenta com 4 espaços por nível (texto puro)
'  - Não usa parser estrutural
'  - Retorna SOMENTE o JSON indentado (payload)
'  Observação: Mantém ": " (com espaço) por legibilidade.
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
                Case "{"
                    resultado = resultado & c & vbCrLf
                    nivel = nivel + 1
                    resultado = resultado & String(nivel * 4, " ")

                Case "}"
                    resultado = resultado & vbCrLf
                    nivel = nivel - 1
                    If nivel < 0 Then nivel = 0
                    resultado = resultado & String(nivel * 4, " ") & c

                Case "["
                    resultado = resultado & c & vbCrLf
                    nivel = nivel + 1
                    resultado = resultado & String(nivel * 4, " ")

                Case "]"
                    resultado = resultado & vbCrLf
                    nivel = nivel - 1
                    If nivel < 0 Then nivel = 0
                    resultado = resultado & String(nivel * 4, " ") & c

                Case ","
                    resultado = resultado & c & vbCrLf & String(nivel * 4, " ")

                Case ":"
                    resultado = resultado & ": "

                Case " ", vbCr, vbLf, vbTab
                    ' ignora fora de strings

                Case Else
                    resultado = resultado & c
            End Select
        End If
    Next i

    JSON_Indentar = resultado
End Function

' ============================================================
' JSON_FlattenPares
'  - Normaliza + Indenta internamente e percorre linha a linha
'  - Gera pares no formato: a.b[0].c = valor
'  - Retorna SOMENTE os pares (payload)
' ============================================================
Public Function JSON_FlattenPares(ByRef prmTxt As String) As String
    Dim txt As String
    Dim linhas() As String
    Dim i As Long
    Dim linha As String
    Dim caminho As String
    Dim pilha(0 To 50) As String
    Dim nivel As Long
    Dim pares As String
    Dim chave As String
    Dim valor As String
    Dim idx As Long
    Dim k As Long

    txt = JSON_Normalizar(prmTxt)
    txt = JSON_Indentar(txt)

    linhas = Split(txt, vbCrLf)

    nivel = 0
    pares = ""
    idx = 0

    For i = 0 To UBound(linhas)
        linha = Trim$(linhas(i))
        If linha = "" Then GoTo Proximo

        ' fechamento de objeto/array
        If linha = "}" Or linha = "}," Or linha = "]" Or linha = "]," Then
            If nivel > 0 Then nivel = nivel - 1
            GoTo Proximo
        End If

        ' abertura de objeto:  "x": {
        If Right$(linha, 1) = "{" Then
            chave = Replace(linha, "{", "")
            chave = Replace(chave, """", "")
            chave = Replace(chave, ":", "")
            chave = Trim$(chave)

            pilha(nivel) = chave
            nivel = nivel + 1
            GoTo Proximo
        End If

        ' abertura de array: "x": [
        If Right$(linha, 1) = "[" Then
            chave = Replace(linha, "[", "")
            chave = Replace(chave, """", "")
            chave = Replace(chave, ":", "")
            chave = Trim$(chave)

            pilha(nivel) = chave
            nivel = nivel + 1
            idx = 0
            GoTo Proximo
        End If

        ' linha "chave": valor
        If InStr(1, linha, ":", vbBinaryCompare) > 0 Then
            chave = Trim$(Left$(linha, InStr(1, linha, ":", vbBinaryCompare) - 1))
            chave = Replace(chave, """", "")
            chave = Trim$(chave)

            valor = Trim$(Mid$(linha, InStr(1, linha, ":", vbBinaryCompare) + 1))
            valor = Replace(valor, ",", "")
            valor = Trim$(valor)

            caminho = ""
            For k = 0 To nivel - 1
                If pilha(k) <> "" Then
                    If caminho = "" Then
                        caminho = pilha(k)
                    Else
                        caminho = caminho & "." & pilha(k)
                    End If
                End If
            Next k

            If caminho = "" Then
                caminho = chave
            Else
                caminho = caminho & "." & chave
            End If

            pares = pares & caminho & " = " & valor & vbCrLf
            GoTo Proximo
        End If

        ' elemento simples de array (sem chave)
        valor = Replace(linha, ",", "")
        valor = Trim$(valor)

        caminho = ""
        For k = 0 To nivel - 1
            If pilha(k) <> "" Then
                If caminho = "" Then
                    caminho = pilha(k)
                Else
                    caminho = caminho & "." & pilha(k)
                End If
            End If
        Next k

        caminho = caminho & "[" & idx & "]"
        idx = idx + 1

        pares = pares & caminho & " = " & valor & vbCrLf

Proximo:
    Next i

    JSON_FlattenPares = pares
End Function


' ============================================================
' STUBS (FASE 1): mantêm integrador funcionando sem prometer lógica
' ============================================================
Public Function JSON_ExpandInlineArrays(ByRef prmTxt As String) As String
    JSON_ExpandInlineArrays = prmTxt
End Function

Public Function JSON_Validar(ByRef prmTxt As String) As String
    ' Nesta fase, validação completa ainda não implementada.
    ' Retorna a entrada para não quebrar o fluxo.
    JSON_Validar = prmTxt
End Function


Public Function CSV_Para_JSON(ByRef prmTxt As String) As String
    CSV_Para_JSON = prmTxt
End Function

Public Function CSV_Para_JSON_Tabela(ByRef prmTxt As String) As String
    CSV_Para_JSON_Tabela = prmTxt
End Function

Public Function TAB_Para_CSV(ByRef prmTxt As String) As String
    TAB_Para_CSV = prmTxt
End Function

Public Function TAB_Para_JSON_Tabela(ByRef prmTxt As String) As String
    TAB_Para_JSON_Tabela = prmTxt
End Function
