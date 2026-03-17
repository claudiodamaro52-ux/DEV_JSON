Option Compare Database
Option Explicit

' =========================================
' Test Runner: DEV_JSON (Fase 3.2)
' Valida: JSON_Unflatten + JSON_Normalizar
' Comparação: string exata (normalizada)
' =========================================

Public Sub RunAllTests_32()
    Dim passed As Long, failed As Long

    Debug.Print String(60, "=")
    Debug.Print "RunAllTests_32 - início"
    Debug.Print String(60, "=")

    RunOneTest_32 "Caso1_obj_simples", _
        JoinLines(Array( _
            "a = 1", _
            "b = \"x\"" _
        )), _
        "{\"a\":1,\"b\":\"x\"}", _
        passed, failed

    RunOneTest_32 "Caso2_obj_aninhado", _
        JoinLines(Array( _
            "b.c = 2", _
            "b.d = 3" _
        )), _
        "{\"b\":{\"c\":2,\"d\":3}}", _
        passed, failed

    RunOneTest_32 "Caso3_array_valores", _
        JoinLines(Array( _
            "a = 1", _
            "b.c = 2", _
            "b.d[0] = 1", _
            "b.d[1] = 2", _
            "b.d[2] = 3" _
        )), _
        "{\"a\":1,\"b\":{\"c\":2,\"d\":[1,2,3]}}", _
        passed, failed

    RunOneTest_32 "Caso4_array_objetos", _
        JoinLines(Array( _
            "itens[0].sku = 123", _
            "itens[0].preco = 9.9", _
            "itens[1].sku = 456" _
        )), _
        "{\"itens\":[{\"sku\":123,\"preco\":9.9},{\"sku\":456}]} ", _
        passed, failed

    RunOneTest_32 "Caso5_mix_obj_arrObj_arrVal", _
        JoinLines(Array( _
            "pedido.id = 10", _
            "pedido.itens[0].sku = 123", _
            "pedido.itens[0].qtd = 2", _
            "pedido.itens[1].sku = 456", _
            "pedido.valores[0] = 9.9", _
            "pedido.valores[1] = 1.5" _
        )), _
        "{\"pedido\":{\"id\":10,\"itens\":[{\"sku\":123,\"qtd\":2},{\"sku\":456}],\"valores\":[9.9,1.5]}}", _
        passed, failed

    RunOneTest_32 "Caso6_buraco_array_objetos", _
        JoinLines(Array( _
            "itens[2].sku = 999" _
        )), _
        "{\"itens\":[{},{},{\"sku\":999}]} ", _
        passed, failed

    Debug.Print String(60, "-")
    Debug.Print "RunAllTests_32 - fim"
    Debug.Print "Passed=" & passed & " Failed=" & failed
    Debug.Print String(60, "=")

    If failed > 0 Then
        Err.Raise vbObjectError + 32001, "RunAllTests_32", "Há testes falhando: " & failed
    End If
End Sub

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

Private Function JoinLines(ByVal a As Variant) As String
    Dim i As Long
    Dim s As String

    s = ""
    For i = LBound(a) To UBound(a)
        If s <> "" Then s = s & vbCrLf
        s = s & CStr(a(i))
    Next i

    JoinLines = s
End Function
