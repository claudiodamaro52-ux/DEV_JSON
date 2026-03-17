Attribute VB_Name = "modJSON_Flatten"
'Attribute VB_Name = "modJSON_Flatten"
Option Explicit

' ============================================================
' JSON_FlattenPares
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

        If linha = "}" Or linha = "}," Or linha = "]" Or linha = "]," Then
            If nivel > 0 Then nivel = nivel - 1
            GoTo Proximo
        End If

        If Right$(linha, 1) = "{" Then
            chave = Replace(linha, "{", "")
            chave = Replace(chave, """", "")
            chave = Replace(chave, ":", "")
            chave = Trim$(chave)

            pilha(nivel) = chave
            nivel = nivel + 1
            GoTo Proximo
        End If

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

