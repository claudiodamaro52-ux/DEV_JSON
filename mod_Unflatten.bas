Option Explicit
' ============================================================
' JSON_Unflatten (FASE 3.2)
'  Entrada: linhas "path = value"
'   - path: "a.b.c", "x.y[0]" e agora também "itens[0].sku"
'  Saída: JSON indentado (4 espaços)
'
' Regras:
'  - inferência: number / true / false / null; senão string
'  - assume pares em ordem (por enquanto)
'  - arrays:
'     * array de valores: d[2] = 3
'     * array de objetos: itens[0].sku = 123  (A1)
'       - cria itens faltantes como {} (2A)
' ============================================================
Public Function JSON_Unflatten(ByRef prmTxt As String) As String
    Dim linhas() As String
    Dim i As Long, j As Long
    Dim linha As String
    Dim path As String, rawValue As String

    Dim json As String

    ' context stack
    Dim ctxType() As String   ' "obj" or "arr"
    Dim ctxName() As String   ' key name (or "$")
    Dim ctxArrIndex() As Long ' last written index (arr only)
    Dim top As Long

    ' path tokens (contexts only)
    Dim tokType() As String      ' "obj" / "arrVal" / "arrObj"
    Dim tokName() As String      ' key name
    Dim tokIdx() As Long         ' index for arrVal/arrObj, else -1
    Dim tokCount As Long

    ' final write target
    Dim lastKind As String       ' "key" / "arrVal" (write item)
    Dim lastKey As String
    Dim lastIdx As Long

    linhas = Split(prmTxt, vbCrLf)

    ReDim ctxType(0 To 400)
    ReDim ctxName(0 To 400)
    ReDim ctxArrIndex(0 To 400)
    top = -1

    json = "{"
    PushCtx ctxType, ctxName, ctxArrIndex, top, "obj", "$", -1

    For i = 0 To UBound(linhas)
        linha = Trim$(linhas(i))
        If linha = "" Then GoTo Proximo
        If InStr(1, linha, "=", vbBinaryCompare) = 0 Then GoTo Proximo

        path = Trim$(Left$(linha, InStr(1, linha, "=", vbBinaryCompare) - 1))
        rawValue = Trim$(Mid$(linha, InStr(1, linha, "=", vbBinaryCompare) + 1))

        TokenizarPath_32 path, tokType, tokName, tokIdx, tokCount, lastKind, lastKey, lastIdx

        ' 1) close to common prefix
        FecharAtePrefixoComum_32 tokType, tokName, tokIdx, tokCount, ctxType, ctxName, ctxArrIndex, top, json

        ' 2) open missing contexts
        AbrirContextos_32 tokType, tokName, tokIdx, tokCount, ctxType, ctxName, ctxArrIndex, top, json

        ' 3) emit final value
        If lastKind = "key" Then
            json = json & vbCrLf & Indent(top) & """" & lastKey & """: " & ValorJSON(rawValue) & ","
ElseIf lastKind = "arrValItem" Then
    ' estamos dentro do array (ctx "arr") com o nome lastKey
    If Not (top >= 0 And ctxType(top) = "arr" And ctxName(top) = lastKey) Then
        ' fallback: abre/entra no array
        AbrirOuEntrarArrayValores lastKey, ctxType, ctxName, ctxArrIndex, top, json
    End If

    j = ctxArrIndex(top) + 1
    Do While j < lastIdx
        json = json & vbCrLf & Indent(top) & "null,"
        ctxArrIndex(top) = j
        j = j + 1
    Loop

    json = json & vbCrLf & Indent(top) & ValorJSON(rawValue) & ","
    ctxArrIndex(top) = lastIdx
End If

Proximo:
    Next i

    Do While top >= 0
        FecharCtx ctxType, ctxName, ctxArrIndex, top, json
    Loop

    JSON_Unflatten = json
End Function

' ============================================================
' Tokenização (3.2)
'  Produz tokens de contexto e descreve o alvo final:
'    - lastKind="key": última parte é chave simples (emitir "k": v)
'    - lastKind="arrVal": última parte é array de valores (emitir item)
'
'  Tokens possíveis:
'    obj:   a / b
'    arrVal: d[idx]    (se for o último segmento)
'    arrObj: itens[idx] (se houver segmentos depois dele)
'      => abre array itens e entra no objeto do item idx
' ============================================================
Private Sub TokenizarPath_32(ByVal path As String, _
                             ByRef tokType() As String, ByRef tokName() As String, ByRef tokIdx() As Long, ByRef tokCount As Long, _
                             ByRef lastKind As String, ByRef lastKey As String, ByRef lastIdx As Long)
    Dim segs() As String
    Dim k As Long
    Dim seg As String
    Dim nm As String
    Dim idx As Long

    segs = Split(path, ".")

    ReDim tokType(0 To 400)
    ReDim tokName(0 To 400)
    ReDim tokIdx(0 To 400)
    tokCount = 0

    lastKind = "key"
    lastKey = ""
    lastIdx = -1

    For k = 0 To UBound(segs)
        seg = segs(k)

        If InStr(1, seg, "[", vbBinaryCompare) > 0 Then
            nm = Left$(seg, InStr(1, seg, "[", vbBinaryCompare) - 1)
            idx = CLng(Replace(Mid$(seg, InStr(1, seg, "[", vbBinaryCompare) + 1), "]", ""))

            If k < UBound(segs) Then
                ' itens[0].sku => array de objetos (contexto)
                tokType(tokCount) = "arrObj"
                tokName(tokCount) = nm
                tokIdx(tokCount) = idx
                tokCount = tokCount + 1
            Else
                ' d[2] = 3 => array de valores (contexto + write item)
                tokType(tokCount) = "arrVal"
                tokName(tokCount) = nm
                tokIdx(tokCount) = idx
                tokCount = tokCount + 1

                lastKind = "arrValItem"
                lastKey = nm
                lastIdx = idx
            End If

        Else
            If k < UBound(segs) Then
                tokType(tokCount) = "obj"
                tokName(tokCount) = seg
                tokIdx(tokCount) = -1
                tokCount = tokCount + 1
            Else
                lastKind = "key"
                lastKey = seg
                lastIdx = -1
            End If
        End If
    Next k

    If tokCount > 0 Then
        ReDim Preserve tokType(0 To tokCount - 1)
        ReDim Preserve tokName(0 To tokCount - 1)
        ReDim Preserve tokIdx(0 To tokCount - 1)
    Else
        ReDim tokType(0 To 0)
        ReDim tokName(0 To 0)
        ReDim tokIdx(0 To 0)
    End If
End Sub

' ============================================================
' Prefixo comum (3.2)
'  ctx(0)="$"
'  Tokens correspondem a ctx(1..tokCount)
'  Para arrObj, o prefixo inclui TAMBÉM o índice (tem que bater)
' ============================================================
Private Sub FecharAtePrefixoComum_32(ByRef tokType() As String, ByRef tokName() As String, ByRef tokIdx() As Long, ByVal tokCount As Long, _
                                     ByRef ctxType() As String, ByRef ctxName() As String, ByRef ctxArrIndex() As Long, _
                                     ByRef top As Long, ByRef json As String)
    Dim common As Long
    Dim t As Long
    Dim ctxPos As Long

    common = 0        ' ctx(0)="$"
    ctxPos = 1

    For t = 0 To tokCount - 1
        If tokType(t) = "obj" Then
            If ctxPos <= top And ctxType(ctxPos) = "obj" And ctxName(ctxPos) = tokName(t) Then
                common = ctxPos
                ctxPos = ctxPos + 1
            Else
                Exit For
            End If
        
        ElseIf tokType(t) = "arrVal" Then
            If ctxPos <= top And ctxType(ctxPos) = "arr" And ctxName(ctxPos) = tokName(t) Then
                common = ctxPos
                ctxPos = ctxPos + 1
            Else
                Exit For
            End If
        
        ElseIf tokType(t) = "arrObj" Then
            ' Primeiro, tenta manter o array como prefixo comum
            If ctxPos <= top And ctxType(ctxPos) = "arr" And ctxName(ctxPos) = tokName(t) Then
                common = ctxPos   ' mantém o array, mesmo que o índice mude
                ' Depois, só mantém o objeto do índice se também bater
                If (ctxPos + 1) <= top Then
                    If ctxType(ctxPos + 1) = "obj" And ctxName(ctxPos + 1) = ("#idx:" & CStr(tokIdx(t))) Then
                        common = ctxPos + 1
                        ctxPos = ctxPos + 2
                    Else
                        ' índice diferente: não é prefixo comum além do array
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Else
                Exit For
            End If
        End If
    Next t

    Do While top > common
        FecharCtx ctxType, ctxName, ctxArrIndex, top, json
    Loop
End Sub
' ============================================================
' Abrir contextos (3.2)
'  - obj: abre {"name":{ ... }}
'  - arrObj: abre "itens":[ ... ] e entra no objeto do item idx
' ============================================================
Private Sub AbrirContextos_32(ByRef tokType() As String, ByRef tokName() As String, ByRef tokIdx() As Long, ByVal tokCount As Long, _
                              ByRef ctxType() As String, ByRef ctxName() As String, ByRef ctxArrIndex() As Long, _
                              ByRef top As Long, ByRef json As String)
    Dim t As Long
    Dim ctxPos As Long

    ' ctxPos aponta para o próximo nível esperado no stack (ctx(0)="$")
    ctxPos = 1

    For t = 0 To tokCount - 1
        If tokType(t) = "obj" Then
            ' Se já existe o obj correto nesse nível, só avança
            If ctxPos <= top Then
                If ctxType(ctxPos) = "obj" And ctxName(ctxPos) = tokName(t) Then
                    ctxPos = ctxPos + 1
                Else
                    ' stack divergiu: fecha até antes desse nível e abre correto
                    Do While top >= ctxPos
                        FecharCtx ctxType, ctxName, ctxArrIndex, top, json
                    Loop
                    json = json & vbCrLf & Indent(top) & """" & tokName(t) & """: {"
                    PushCtx ctxType, ctxName, ctxArrIndex, top, "obj", tokName(t), -1
                    ctxPos = ctxPos + 1
                End If
            Else
                json = json & vbCrLf & Indent(top) & """" & tokName(t) & """: {"
                PushCtx ctxType, ctxName, ctxArrIndex, top, "obj", tokName(t), -1
                ctxPos = ctxPos + 1
            End If

        ElseIf tokType(t) = "arrObj" Then
            ' Abre/entra no array e no item idx
            AbrirOuEntrarArrayObjeto tokName(t), tokIdx(t), ctxType, ctxName, ctxArrIndex, top, json
            ' arrObj consome 2 níveis no stack: arr + obj(#idx:n)
            ctxPos = top + 1

        ElseIf tokType(t) = "arrVal" Then
            ' Abre/entra no array de valores
            AbrirOuEntrarArrayValores tokName(t), ctxType, ctxName, ctxArrIndex, top, json
            ' arrVal consome 1 nível no stack: arr
            ctxPos = top + 1
        End If
    Next t
End Sub



Private Sub AbrirOuEntrarArrayObjeto(ByVal arrName As String, ByVal idx As Long, _
                                     ByRef ctxType() As String, ByRef ctxName() As String, ByRef ctxArrIndex() As Long, _
                                     ByRef top As Long, ByRef json As String)
    Dim j As Long

    ' Se já estamos no objeto do índice certo, não faz nada
    If top >= 1 Then
        If ctxType(top) = "obj" And ctxName(top) = ("#idx:" & CStr(idx)) Then
            If ctxType(top - 1) = "arr" And ctxName(top - 1) = arrName Then
                Exit Sub
            End If
        End If
    End If

    ' Se estamos dentro de um #idx: (objeto de item), fecha só esse objeto
    If top >= 0 Then
        If ctxType(top) = "obj" And Left$(ctxName(top), 5) = "#idx:" Then
            FecharCtx ctxType, ctxName, ctxArrIndex, top, json
        End If
    End If

    ' Se o topo é um array mas não é o array alvo, fecha só esse array (não fecha o pai)
    If top >= 0 Then
        If ctxType(top) = "arr" And ctxName(top) <> arrName Then
            FecharCtx ctxType, ctxName, ctxArrIndex, top, json
        End If
    End If

    ' Se já estamos no array certo, ok; senão abre como propriedade do contexto atual
    If Not (top >= 0 And ctxType(top) = "arr" And ctxName(top) = arrName) Then
        json = json & vbCrLf & Indent(top) & """" & arrName & """: ["
        PushCtx ctxType, ctxName, ctxArrIndex, top, "arr", arrName, -1
    End If

    ' Preenche buracos com {} até chegar no índice desejado
    j = ctxArrIndex(top) + 1
    Do While j < idx
        json = json & vbCrLf & Indent(top) & "{},"
        ctxArrIndex(top) = j
        j = j + 1
    Loop

    ' Abre o objeto do item idx
    json = json & vbCrLf & Indent(top) & "{"
    PushCtx ctxType, ctxName, ctxArrIndex, top, "obj", "#idx:" & CStr(idx), -1
    ctxArrIndex(top - 1) = idx
End Sub

Private Sub AbrirOuEntrarArrayValores(ByVal arrName As String, _
                                      ByRef ctxType() As String, ByRef ctxName() As String, ByRef ctxArrIndex() As Long, _
                                      ByRef top As Long, ByRef json As String)
    ' Array de valores deve ser aberto COMO propriedade do objeto atual.
    ' Não pode fechar objetos pai procurando um array com o mesmo nome.

    ' Se já estamos nesse array, apenas reutiliza
    If top >= 0 Then
        If ctxType(top) = "arr" And ctxName(top) = arrName Then Exit Sub
    End If

    ' Se estamos dentro de objeto de item de array (#idx:), fecha só esse item (não fecha o pai)
    If top >= 0 Then
        If ctxType(top) = "obj" And Left$(ctxName(top), 5) = "#idx:" Then
            FecharCtx ctxType, ctxName, ctxArrIndex, top, json
        End If
    End If

    ' Se o topo atual for outro array, fecha só esse array
    If top >= 0 Then
        If ctxType(top) = "arr" And ctxName(top) <> arrName Then
            FecharCtx ctxType, ctxName, ctxArrIndex, top, json
        End If
    End If

    ' Agora abre o array dentro do contexto atual (normalmente obj)
    json = json & vbCrLf & Indent(top) & """" & arrName & """: ["
    PushCtx ctxType, ctxName, ctxArrIndex, top, "arr", arrName, -1
End Sub
' ============================================================
' Shared helpers (iguais à 3.1b)
' ============================================================
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
        ElseIf Left$(ctxName(top), 5) = "#idx:" Then
            ' fecha objeto dentro de array: precisa fechar com "}," (ou "}" e vírgula gerida pelo chamador)
            json = json & vbCrLf & Indent(top - 1) & "},"
        Else
            json = json & vbCrLf & Indent(top - 1) & "},"
        End If
    End If

    top = top - 1
End Sub

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

