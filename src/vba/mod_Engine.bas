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


' ------------------------
' Tokenização do path
' ------------------------
Private Sub TokenizarPath(ByVal path As String, _
                          ByRef tokType() As String, ByRef tokName() As String, ByRef tokCount As Long, _
                          ByRef lastIsArrayItem As Boolean, ByRef lastArrIdx As Long, ByRef lastKey As String)
    Dim segs() As String
    Dim k As Long
    Dim seg As String

    segs = Split(path, ".")

    ReDim tokType(0 To 200)
    ReDim tokName(0 To 200)
    tokCount = 0

    lastIsArrayItem = False
    lastArrIdx = -1
    lastKey = ""

    ' Segmentos intermediários viram contextos "obj" (nesta fase)
    If UBound(segs) >= 1 Then
        For k = 0 To UBound(segs) - 1
            seg = segs(k)

            If InStr(1, seg, "[", vbBinaryCompare) > 0 Then
                ' nesta fase, não suportamos a.b[0].c
                tokType(tokCount) = "arr"
                tokName(tokCount) = Left$(seg, InStr(1, seg, "[", vbBinaryCompare) - 1)
                tokCount = tokCount + 1
                Exit For
            Else
                tokType(tokCount) = "obj"
                tokName(tokCount) = seg
                tokCount = tokCount + 1
            End If
        Next k
    End If

    ' Último segmento: chave simples ou array item
    seg = segs(UBound(segs))

    If InStr(1, seg, "[", vbBinaryCompare) > 0 Then
        lastIsArrayItem = True
        lastKey = Left$(seg, InStr(1, seg, "[", vbBinaryCompare) - 1)
        lastArrIdx = CLng(Replace(Mid$(seg, InStr(1, seg, "[", vbBinaryCompare) + 1), "]", ""))

        tokType(tokCount) = "arr"
        tokName(tokCount) = lastKey
        tokCount = tokCount + 1
    Else
        lastIsArrayItem = False
        lastKey = seg
    End If

    ' **correção principal**: só redimensiona se houver tokens
    If tokCount > 0 Then
        ReDim Preserve tokType(0 To tokCount - 1)
        ReDim Preserve tokName(0 To tokCount - 1)
    Else
        ' nenhum contexto (ex: "a"), deixa arrays vazios
        ReDim tokType(0 To 0)
        ReDim tokName(0 To 0)
    End If
End Sub

' ------------------------
' Stack de contextos
' ------------------------
Private Sub PushCtx(ByRef ctxType() As String, ByRef ctxName() As String, ByRef ctxArrIndex() As Long, _
                    ByRef top As Long, ByVal t As String, ByVal n As String, ByVal arrIdx As Long)
    top = top + 1
    ctxType(top) = t
    ctxName(top) = n
    ctxArrIndex(top) = arrIdx
End Sub

Private Function Indent(ByVal level As Long) As String
    If level < 0 Then level = 0
    Indent = String(level * 4, " ")
End Function

Private Sub FecharCtx(ByRef ctxType() As String, ByRef ctxName() As String, ByRef ctxArrIndex() As Long, _
                      ByRef top As Long, ByRef json As String)
    If Right$(json, 1) = "," Then json = Left$(json, Len(json) - 1)

    If top < 0 Then Exit Sub

    If ctxType(top) = "arr" Then
        json = json & vbCrLf & Indent(top - 1) & "],"
    Else
        If ctxName(top) = "$" Then
            json = json & vbCrLf & "}"
        Else
            json = json & vbCrLf & Indent(top - 1) & "},"
        End If
    End If

    top = top - 1
End Sub

Private Sub xFecharAtePrefixoComum(ByRef tokType() As String, ByRef tokName() As String, ByVal tokCount As Long, _
                                  ByRef ctxType() As String, ByRef ctxName() As String, ByRef ctxArrIndex() As Long, _
                                  ByRef top As Long, ByRef json As String)
    ' common = nível de ctx (top) que pode permanecer aberto.
    ' ctx(0) = "$"
    ' tok(0) <-> ctx(1)
    Dim common As Long
    Dim t As Long

    common = 0 ' sempre mantém "$"

    For t = 0 To tokCount - 1
        If (t + 1) <= top Then
            If ctxType(t + 1) = tokType(t) And ctxName(t + 1) = tokName(t) Then
                common = t + 1
            Else
                Exit For
            End If
        Else
            Exit For
        End If
    Next t

    Do While top > common
        FecharCtx ctxType, ctxName, ctxArrIndex, top, json
    Loop
End Sub
Private Sub AbrirContextos(ByRef tokType() As String, ByRef tokName() As String, ByVal tokCount As Long, _
                           ByRef ctxType() As String, ByRef ctxName() As String, ByRef ctxArrIndex() As Long, _
                           ByRef top As Long, ByRef json As String)
    ' ctx(0) = "$"
    ' tokens abertos atualmente = top  (top=0 => nenhum token aberto)
    Dim opened As Long
    Dim t As Long

    opened = top ' número de níveis além do "$" já abertos

    For t = opened To tokCount - 1
        If tokType(t) = "obj" Then
            json = json & vbCrLf & Indent(top) & """" & tokName(t) & """: {"
            PushCtx ctxType, ctxName, ctxArrIndex, top, "obj", tokName(t), -1
        ElseIf tokType(t) = "arr" Then
            json = json & vbCrLf & Indent(top) & """" & tokName(t) & """: ["
            PushCtx ctxType, ctxName, ctxArrIndex, top, "arr", tokName(t), -1
        End If
    Next t
End Sub
' ------------------------
' ValorJSON (inferência)
' ------------------------
Private Function ValorJSON(ByVal rawValue As String) As String
    Dim v As String
    v = Trim$(rawValue)

    If v = "" Then
        ValorJSON = """"""
        Exit Function
    End If

    If Left$(v, 1) = """" And Right$(v, 1) = """" And Len(v) >= 2 Then
        ValorJSON = v
        Exit Function
    End If

    Select Case LCase$(v)
        Case "true", "false", "null"
            ValorJSON = LCase$(v)
            Exit Function
    End Select

    If IsNumeric(v) Then
        ValorJSON = v
        Exit Function
    End If

    v = Replace(v, """", "'")
    ValorJSON = """" & v & """"
End Function

' ============================================================
' Stack helpers
' ============================================================
Private Sub PushStack(ByRef stackPath() As String, ByRef stackType() As String, ByRef stackArrName() As String, ByRef stackArrIndex() As Long, _
                      ByRef top As Long, ByVal p As String, ByVal t As String, ByVal arrName As String, ByVal arrIdx As Long)
    top = top + 1
    stackPath(top) = p
    stackType(top) = t
    stackArrName(top) = arrName
    stackArrIndex(top) = arrIdx
End Sub


' ============================================================
' AjustarParaCaminho
'  - fecha níveis até encontrar o maior prefixo comum
' ============================================================
Private Sub AjustarParaCaminho(ByRef segs() As String, ByVal segCount As Long, _
                               ByRef stackPath() As String, ByRef stackType() As String, ByRef stackArrName() As String, ByRef stackArrIndex() As Long, _
                               ByRef top As Long, ByRef json As String)
    ' stackPath guarda caminhos como "$.a.b" ou "$.a.b.d["
    ' Aqui vamos comparar por prefixo de chaves (somente objetos),
    ' e para arrays simples só garantimos que o array atual é o mesmo.

    Dim targetObjPath As String
    targetObjPath = "$"

    Dim k As Long
    For k = 0 To segCount - 2
        ' tudo exceto o último segmento, porque o último pode ser valor direto
        If InStr(1, segs(k), "[", vbBinaryCompare) > 0 Then Exit For ' não suportamos objeto dentro de array nesta fase
        targetObjPath = targetObjPath & "." & segs(k)
    Next k

    ' fecha enquanto o topo não for prefixo do target
    Do While top >= 0
        If stackType(top) = "obj" Then
            If stackPath(top) = targetObjPath Or stackPath(top) = "$" Then Exit Do
        End If
        FecharNivel stackPath, stackType, stackArrName, stackArrIndex, top, json
    Loop
End Sub

' ============================================================
' ProcessarSegmentos
'  - abre objetos e arrays necessários e escreve o valor final
' ============================================================
Private Sub ProcessarSegmentos(ByRef segs() As String, ByVal segCount As Long, _
                               ByRef stackPath() As String, ByRef stackType() As String, ByRef stackArrName() As String, ByRef stackArrIndex() As Long, _
                               ByRef top As Long, ByRef json As String, ByVal rawValue As String)

    Dim curPath As String
    curPath = stackPath(top)

    Dim k As Long
    For k = PathDepthFromStack(curPath) To segCount - 2
        ' abre objetos para segmentos intermediários (somente "a", "b")
        If InStr(1, segs(k), "[", vbBinaryCompare) > 0 Then Exit For

        json = json & vbCrLf & Indent(top) & """" & segs(k) & """: {"
        curPath = curPath & "." & segs(k)
        PushStack stackPath, stackType, stackArrName, stackArrIndex, top, curPath, "obj", "", -1
    Next k

    ' último segmento (pode ser chave simples ou array)
    Dim lastSeg As String
    lastSeg = segs(segCount - 1)

    If InStr(1, lastSeg, "[", vbBinaryCompare) > 0 Then
        ' array de valores simples: nome[idx]
        Dim arrName As String, idx As Long
        arrName = Left$(lastSeg, InStr(1, lastSeg, "[", vbBinaryCompare) - 1)
        idx = CLng(Replace(Mid$(lastSeg, InStr(1, lastSeg, "[", vbBinaryCompare) + 1), "]", ""))

        ' se o topo não é este array aberto, abre agora
        If Not (stackType(top) = "arr" And stackArrName(top) = arrName) Then
            json = json & vbCrLf & Indent(top) & """" & arrName & """: ["
            PushStack stackPath, stackType, stackArrName, stackArrIndex, top, stackPath(top) & "." & arrName & "[", "arr", arrName, -1
        End If

        ' preenche índices faltantes com null (simples e determinístico)
        Dim nextIdx As Long
        nextIdx = stackArrIndex(top) + 1

        Do While nextIdx < idx
            json = json & vbCrLf & Indent(top) & "null,"
            stackArrIndex(top) = nextIdx
            nextIdx = nextIdx + 1
        Loop

        ' escreve valor no índice
        json = json & vbCrLf & Indent(top) & ValorJSON(rawValue) & ","
        stackArrIndex(top) = idx

    Else
        ' chave simples
        json = json & vbCrLf & Indent(top) & """" & lastSeg & """: " & ValorJSON(rawValue) & ","
    End If
End Sub

Private Function PathDepthFromStack(ByVal p As String) As Long
    ' "$" => 0
    ' "$.a" => 1
    ' "$.a.b" => 2
    If p = "$" Then
        PathDepthFromStack = 0
    Else
        PathDepthFromStack = UBound(Split(Mid$(p, 3), ".")) + 1
    End If
End Function

' ============================================================
' FecharNivel
'  - fecha objeto/array e remove vírgula final antes de fechar
' ============================================================
Private Sub FecharNivel(ByRef stackPath() As String, ByRef stackType() As String, ByRef stackArrName() As String, ByRef stackArrIndex() As Long, _
                        ByRef top As Long, ByRef json As String)
    ' remove vírgula final se existir
    If Right$(json, 1) = "," Then json = Left$(json, Len(json) - 1)

    If top < 0 Then Exit Sub

    If stackType(top) = "arr" Then
        json = json & vbCrLf & Indent(top - 1) & "],"
    ElseIf stackType(top) = "obj" Then
        If stackPath(top) = "$" Then
            json = json & vbCrLf & "}"
        Else
            json = json & vbCrLf & Indent(top - 1) & "},"
        End If
    End If

    top = top - 1
End Sub

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
