Option Compare Database
Option Explicit

' =========================================
' Test Runner: DEV_JSON (Fase 3.2)
' Compatível: MS Access 2016 (VBA)
'
' Como rodar:
'   Immediate Window (Ctrl+G):
'     RunAllTests_32
' =========================================

Public Function RunAllTests_32()
    Dim passed As Long
    Dim failed As Long

    Debug.Print String(60, "=")
    Debug.Print "RunAllTests_32 - inicio"
    Debug.Print String(60, "=")

    Call RunOneTest_32( _
        "Caso1_obj_simples", _
        "a = 1" & vbCrLf & _
        "b = ""x""", _
        "{""a"":1,""b"":""x""}", _
        passed, failed _
    )

    Call RunOneTest_32( _
        "Caso2_obj_aninhado", _
        "b.c = 2" & vbCrLf & _
        "b.d = 3", _
        "{""b"":{""c"":2,""d"":3}}", _
        passed, failed _
    )

    Call RunOneTest_32( _
        "Caso3_array_valores", _
        "a = 1" & vbCrLf & _
        "b.c = 2" & vbCrLf & _
        "b.d[0] = 1" & vbCrLf & _
        "b.d[1] = 2" & vbCrLf & _
        "b.d[2] = 3", _
        "{""a"":1,""b"":{""c"":2,""d"":[1,2,3]}}", _
        passed, failed _
    )

    Call RunOneTest_32( _
        "Caso4_array_objetos", _
        "itens[0].sku = 123" & vbCrLf & _
        "itens[0].preco = 9.9" & vbCrLf & _
        "itens[1].sku = 456", _
        "{""itens"":[{""sku"":123,""preco"":9.9},{""sku"":456}]}", _
        passed, failed _
    )

    Call RunOneTest_32( _
        "Caso5_mix_obj_arrObj_arrVal", _
        "pedido.id = 10" & vbCrLf & _
        "pedido.itens[0].sku = 123" & vbCrLf & _
        "pedido.itens[0].qtd = 2" & vbCrLf & _
        "pedido.itens[1].sku = 456" & vbCrLf & _
        "pedido.valores[0] = 9.9" & vbCrLf & _
        "pedido.valores[1] = 1.5", _
        "{""pedido"":{""id"":10,""itens"":[{""sku"":123,""qtd"":2},{""sku"":456}],""valores"":[9.9,1.5]}}", _
        passed, failed _
    )

    Call RunOneTest_32( _
        "Caso6_buraco_array_objetos", _
        "itens[2].sku = 999", _
        "{""itens"":[{},{},{""sku"":999}]}", _
        passed, failed _
    )

    Debug.Print String(60, "-")
    Debug.Print "RunAllTests_32 - fim"
    Debug.Print "Passed=" & passed & " Failed=" & failed
    Debug.Print String(60, "=")

    If failed > 0 Then
        Err.Raise vbObjectError + 32001, "RunAllTests_32", "Ha testes falhando: " & failed
    End If
End Function

Private Sub RunOneTest_32(ByVal testName As String, ByVal inputPairs As String, ByVal expectedNormalized As String, _
                          ByRef passed As Long, ByRef failed As Long)
    On Error GoTo EH

    Dim jsonPretty As String
    Dim jsonNormalized As String

    jsonPretty = JSON_Unflatten(inputPairs)
    jsonNormalized = JSON_Normalizar(jsonPretty)

    If jsonNormalized <> expectedNormalized Then
        failed = failed + 1
        Debug.Print ""
        Debug.Print "FAIL: " & testName
        Debug.Print "Expected: " & expectedNormalized
        Debug.Print "Actual  : " & jsonNormalized
        Debug.Print "Pretty  :"
        Debug.Print jsonPretty
    Else
        passed = passed + 1
        Debug.Print "OK  : " & testName
    End If

    Exit Sub

EH:
    failed = failed + 1
    Debug.Print ""
    Debug.Print "ERROR: " & testName
    Debug.Print "Err " & Err.Number & ": " & Err.Description
End Sub

