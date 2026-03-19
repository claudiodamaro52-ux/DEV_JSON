Attribute VB_Name = "mod_JSON_Flatten"
' ============================================================
' MODULO: modJSON_Flatten.bas
' ============================================================
Option Explicit

' ============================================================
' JSON_FlattenPares_Raw  (CORRIGIDA)
' - Entrada/saída: sempre texto
' - Normaliza e indenta
' - Arrays de objetos: inclui [idx] no caminho
' - Corrige o detalhe "itens.[0]" -> "itens[0]" (sem ponto antes do índice)
' - Índice 0-based (compatível com Python)
' ============================================================
Public Function JSON_FlattenPares_Raw(ByRef prmTxt As String) As String
    Dim txt As String
    Dim linhas() As String
    Dim i As Long
    Dim linha As String

    Dim pilha(0 To 100) As String
    Dim nivel As Long

    Dim inArray(0 To 100) As Boolean
    Dim arrayIdx(0 To 100) As Long

    Dim pares As String
    Dim chave As String
    Dim valor As String
    Dim caminho As String
    Dim k As Long

    txt = JSON_Normalizar(prmTxt)
    txt = JSON_Indentar(txt)

    linhas = Split(txt, vbCrLf)

    nivel = 0
    pares = ""

    For i = 0 To UBound(linhas)
        linha = Trim$(linhas(i))
        If linha = "" Then GoTo Proximo

        ' fecha bloco
        If linha = "}" Or linha = "}," Or linha = "]" Or linha = "]," Then
            If nivel > 0 Then nivel = nivel - 1
            GoTo Proximo
        End If

        ' abre objeto:  "pedido": {
        If Right$(linha, 1) = "{" Then
            chave = Trim$(Replace(Replace(Replace(linha, "{", ""), """", ""), ":", ""))

            ' se o pai for array, aplicar índice no item objeto
            If nivel > 0 Then
                If inArray(nivel - 1) Then
                    pilha(nivel) = chave & "[" & arrayIdx(nivel - 1) & "]"
                    arrayIdx(nivel - 1) = arrayIdx(nivel - 1) + 1
                Else
                    pilha(nivel) = chave
                End If
            Else
                pilha(nivel) = chave
            End If

            inArray(nivel) = False
            arrayIdx(nivel) = 0

            nivel = nivel + 1
            GoTo Proximo
        End If

        ' abre array: "itens": [
        If Right$(linha, 1) = "[" Then
            chave = Trim$(Replace(Replace(Replace(linha, "[", ""), """", ""), ":", ""))

            pilha(nivel) = chave
            inArray(nivel) = True
            arrayIdx(nivel) = 0

            nivel = nivel + 1
            GoTo Proximo
        End If

        ' linha com par:  "id": 10
        If InStr(1, linha, ":", vbBinaryCompare) > 0 Then
            chave = Trim$(Left$(linha, InStr(1, linha, ":", vbBinaryCompare) - 1))
            chave = Trim$(Replace(chave, """", ""))

            valor = Trim$(Mid$(linha, InStr(1, linha, ":", vbBinaryCompare) + 1))
            valor = Trim$(Replace(valor, ",", ""))

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

        ' valor solto dentro de array (ex.: 9.9)
        valor = Trim$(Replace(linha, ",", ""))

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

        If nivel > 0 Then
            If inArray(nivel - 1) Then
                caminho = caminho & "[" & arrayIdx(nivel - 1) & "]"
                arrayIdx(nivel - 1) = arrayIdx(nivel - 1) + 1
            End If
        End If

        pares = pares & caminho & " = " & valor & vbCrLf

Proximo:
    Next i

    ' Corrige "pedido.itens.[0]" -> "pedido.itens[0]" (sem ponto antes do índice)
    pares = Replace(pares, ".[", "[")

    JSON_FlattenPares_Raw = pares
End Function

