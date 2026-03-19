Attribute VB_Name = "modJSON_Unflatten"
'Attribute VB_Name = "modJSON_Unflatten"
Option Explicit

' ============================================================
' JSON_Unflatten (FASE 3.2)
' ============================================================
Public Function JSON_Unflatten(ByRef prmTxt As String) As String
    Dim linhas() As String
    Dim i As Long, j As Long
    Dim linha As String
    Dim path As String, rawValue As String

    Dim json As String

    ' context stack
    Dim ctxType() As String   ' "obj" or "arr"
    Dim ctxName() As String   ' key name (or "$" / "#idx:n")
    Dim ctxArrIndex() As Long ' last written index (arr only)
    Dim top As Long

    ' path tokens (contexts only)
    Dim tokType() As String      ' "obj" / "arrVal" / "arrObj"
    Dim tokName() As String      ' key name
    Dim tokIdx() As Long         ' index for arrVal/arrObj, else -1
    Dim tokCount As Long

    ' final write target
    Dim lastKind As String       ' "key" / "arrValItem"
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

        FecharAtePrefixoComum_32 tokType, tokName, tokIdx, tokCount, ctxType, ctxName, ctxArrIndex, top, json
        AbrirContextos_32 tokType, tokName, tokIdx, tokCount, ctxType, ctxName, ctxArrIndex, top, json

        If lastKind = "key" Then
            json = json & vbCrLf & JSON_Indent(top) & """" & lastKey & """: " & JSON_ValorJSON(rawValue) & ","

        ElseIf lastKind = "arrValItem" Then
            If Not (top >= 0 And ctxType(top) = "arr" And ctxName(top) = lastKey) Then
                AbrirOuEntrarArrayValores lastKey, ctxType, ctxName, ctxArrIndex, top, json
            End If

            j = ctxArrIndex(top) + 1
            Do While j < lastIdx
                json = json & vbCrLf & JSON_Indent(top) & "null,"
                ctxArrIndex(top) = j
                j = j + 1
            Loop

            json = json & vbCrLf & JSON_Indent(top) & JSON_ValorJSON(rawValue) & ","
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
                tokType(tokCount) = "arrObj"
                tokName(tokCount) = nm
                tokIdx(tokCount) = idx
                tokCount = tokCount + 1
            Else
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
' ============================================================
Private Sub FecharAtePrefixoComum_32(ByRef tokType() As String, ByRef tokName() As String, ByRef tokIdx() As Long, ByVal tokCount As Long, _
                                     ByRef ctxType() As String, ByRef ctxName() As String, ByRef ctxArrIndex() As Long, _
                                     ByRef top As Long, ByRef json As String)
    Dim common As Long
    Dim t As Long
    Dim ctxPos As Long

    common = 0
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
            If ctxPos <= top And ctxType(ctxPos) = "arr" And ctxName(ctxPos) = tokName(t) Then
                common = ctxPos
                If (ctxPos + 1) <= top Then
                    If ctxType(ctxPos + 1) = "obj" And ctxName(ctxPos + 1) = ("#idx:" & CStr(tokIdx(t))) Then
                        common = ctxPos + 1
                        ctxPos = ctxPos + 2
                    Else
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
' ============================================================
Private Sub AbrirContextos_32(ByRef tokType() As String, ByRef tokName() As String, ByRef tokIdx() As Long, ByVal tokCount As Long, _
                              ByRef ctxType() As String, ByRef ctxName() As String, ByRef ctxArrIndex() As Long, _
                              ByRef top As Long, ByRef json As String)
    Dim t As Long
    Dim ctxPos As Long

    ctxPos = 1

    For t = 0 To tokCount - 1
        If tokType(t) = "obj" Then
            If ctxPos <= top Then
                If ctxType(ctxPos) = "obj" And ctxName(ctxPos) = tokName(t) Then
                    ctxPos = ctxPos + 1
                Else
                    Do While top >= ctxPos
                        FecharCtx ctxType, ctxName, ctxArrIndex, top, json
                    Loop
                    json = json & vbCrLf & JSON_Indent(top) & """" & tokName(t) & """: {"
                    PushCtx ctxType, ctxName, ctxArrIndex, top, "obj", tokName(t), -1
                    ctxPos = ctxPos + 1
                End If
            Else
                json = json & vbCrLf & JSON_Indent(top) & """" & tokName(t) & """: {"
                PushCtx ctxType, ctxName, ctxArrIndex, top, "obj", tokName(t), -1
                ctxPos = ctxPos + 1
            End If

        ElseIf tokType(t) = "arrObj" Then
            AbrirOuEntrarArrayObjeto tokName(t), tokIdx(t), ctxType, ctxName, ctxArrIndex, top, json
            ctxPos = top + 1

        ElseIf tokType(t) = "arrVal" Then
            AbrirOuEntrarArrayValores tokName(t), ctxType, ctxName, ctxArrIndex, top, json
            ctxPos = top + 1
        End If
    Next t
End Sub

Private Sub AbrirOuEntrarArrayObjeto(ByVal arrName As String, ByVal idx As Long, _
                                     ByRef ctxType() As String, ByRef ctxName() As String, ByRef ctxArrIndex() As Long, _
                                     ByRef top As Long, ByRef json As String)
    Dim j As Long

    If top >= 1 Then
        If ctxType(top) = "obj" And ctxName(top) = ("#idx:" & CStr(idx)) Then
            If ctxType(top - 1) = "arr" And ctxName(top - 1) = arrName Then
                Exit Sub
            End If
        End If
    End If

    If top >= 0 Then
        If ctxType(top) = "obj" And Left$(ctxName(top), 5) = "#idx:" Then
            FecharCtx ctxType, ctxName, ctxArrIndex, top, json
        End If
    End If

    If top >= 0 Then
        If ctxType(top) = "arr" And ctxName(top) <> arrName Then
            FecharCtx ctxType, ctxName, ctxArrIndex, top, json
        End If
    End If

    If Not (top >= 0 And ctxType(top) = "arr" And ctxName(top) = arrName) Then
        json = json & vbCrLf & JSON_Indent(top) & """" & arrName & """: ["
        PushCtx ctxType, ctxName, ctxArrIndex, top, "arr", arrName, -1
    End If

    j = ctxArrIndex(top) + 1
    Do While j < idx
        json = json & vbCrLf & JSON_Indent(top) & "{},"
        ctxArrIndex(top) = j
        j = j + 1
    Loop

    json = json & vbCrLf & JSON_Indent(top) & "{"
    PushCtx ctxType, ctxName, ctxArrIndex, top, "obj", "#idx:" & CStr(idx), -1
    ctxArrIndex(top - 1) = idx
End Sub

Private Sub AbrirOuEntrarArrayValores(ByVal arrName As String, _
                                      ByRef ctxType() As String, ByRef ctxName() As String, ByRef ctxArrIndex() As Long, _
                                      ByRef top As Long, ByRef json As String)
    If top >= 0 Then
        If ctxType(top) = "arr" And ctxName(top) = arrName Then Exit Sub
    End If

    If top >= 0 Then
        If ctxType(top) = "obj" And Left$(ctxName(top), 5) = "#idx:" Then
            FecharCtx ctxType, ctxName, ctxArrIndex, top, json
        End If
    End If

    If top >= 0 Then
        If ctxType(top) = "arr" And ctxName(top) <> arrName Then
            FecharCtx ctxType, ctxName, ctxArrIndex, top, json
        End If
    End If

    json = json & vbCrLf & JSON_Indent(top) & """" & arrName & """: ["
    PushCtx ctxType, ctxName, ctxArrIndex, top, "arr", arrName, -1
End Sub

' ============================================================
' Stack helpers (internos do Unflatten)
' ============================================================
Private Sub PushCtx(ByRef ctxType() As String, ByRef ctxName() As String, ByRef ctxArrIndex() As Long, _
                    ByRef top As Long, ByVal t As String, ByVal n As String, ByVal arrIdx As Long)
    top = top + 1
    ctxType(top) = t
    ctxName(top) = n
    ctxArrIndex(top) = arrIdx
End Sub

Private Sub FecharCtx(ByRef ctxType() As String, ByRef ctxName() As String, ByRef ctxArrIndex() As Long, _
                      ByRef top As Long, ByRef json As String)
    If Right$(json, 1) = "," Then json = Left$(json, Len(json) - 1)
    If top < 0 Then Exit Sub

    If ctxType(top) = "arr" Then
        json = json & vbCrLf & JSON_Indent(top - 1) & "],"
    Else
        If ctxName(top) = "$" Then
            json = json & vbCrLf & "}"
        Else
            json = json & vbCrLf & JSON_Indent(top - 1) & "},"
        End If
    End If

    top = top - 1
End Sub
