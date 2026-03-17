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
