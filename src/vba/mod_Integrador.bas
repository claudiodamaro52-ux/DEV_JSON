Attribute VB_Name = "mod_Integrador"
Option Explicit
' ============================================================
'  MÓDULO: modIntegrador (FASE 1 - CONTRATO v0.6)
'  - Integrador é o ÚNICO lugar onde existe "OP|PAYLOAD|FASES"
'  - Não faz Split de payload (payload não pode conter "|")
' ============================================================

Public Function FverSaida(Optional prmTxt As String = "") As String
    Dim ctrl As Control
    Set ctrl = Forms("frmIntegrador").Controls("ctAndamento")

    ctrl = ctrl & vbCrLf & prmTxt
    ctrl.Requery

    FverSaida = prmTxt
End Function

Public Function fFase(Optional prmFase As String = "- ", Optional prmFaseAnt As String = "") As String
    Nf = Nf + 1
    fFase = Nf & " " & prmFase & vbCrLf & String(30, "-") & vbCrLf & prmFaseAnt
End Function

Public Function fListaCases() As String
    Dim lista As String

    lista = lista & vbCrLf & "JSON_Normalizar"
    lista = lista & vbCrLf & "JSON_Indentar"
    lista = lista & vbCrLf & "JSON_ExpandInlineArrays"
    lista = lista & vbCrLf & "JSON_Unflatten"
    lista = lista & vbCrLf & "JSON_Validar"

    lista = lista & vbCrLf & "CSV para JSON"
    lista = lista & vbCrLf & "CSV para JSON Tabela"
    lista = lista & vbCrLf & "TAB para CSV"
    lista = lista & vbCrLf & "TAB para JSON Tabela"

    lista = lista & vbCrLf & "JSON Flatten (csv)"
    lista = lista & vbCrLf & "JSON Flatten (pares)"

    lista = lista & vbCrLf & "JSON para CSV"
    lista = lista & vbCrLf & "JSON para CSV (híbrido)"

    lista = lista & vbCrLf & "JSON para Matriz"
    lista = lista & vbCrLf & "JSON para SQL INSERT"
    lista = lista & vbCrLf & "JSON para Template CSV"
    lista = lista & vbCrLf & "JSON para Template JSON"

    fListaCases = lista
End Function

Public Function IntegradorPipeline(Optional ByVal prmTxt As String = "") As String
On Error GoTo SaidaErr

    Dim vFvalor As String
    Dim Fases As String

    Fases = ""
    Nf = 0

    vFvalor = prmTxt
    Fases = fFase("Inicio", Fases)

    Select Case gblOpc

        Case "JSON_Indentar"
            vFvalor = JSON_Indentar(vFvalor)
            Fases = fFase("Indentar", Fases)

        Case "JSON_Normalizar"
            vFvalor = JSON_Normalizar(vFvalor)
            Fases = fFase("Normalizar", Fases)

        Case "JSON_ExpandInlineArrays"
            vFvalor = JSON_ExpandInlineArrays(vFvalor)
            Fases = fFase("ExpandInlineArrays", Fases)

        Case "JSON_Unflatten"
            vFvalor = JSON_Unflatten(vFvalor)
            Fases = fFase("Unflatten JSON", Fases)

        Case "JSON_Validar"
            vFvalor = JSON_Validar(vFvalor)
            Fases = fFase("Validar", Fases)

        Case "CSV para JSON"
            vFvalor = CSV_Para_JSON(vFvalor)
            Fases = fFase("CSV para JSON", Fases)

        Case "CSV para JSON Tabela"
            vFvalor = CSV_Para_JSON_Tabela(vFvalor)
            Fases = fFase("CSV para JSON Tabela", Fases)

        Case "TAB para CSV"
            vFvalor = TAB_Para_CSV(vFvalor)
            Fases = fFase("TAB para CSV", Fases)

        Case "TAB para JSON Tabela"
            vFvalor = TAB_Para_JSON_Tabela(vFvalor)
            Fases = fFase("TAB para JSON Tabela", Fases)

        Case "JSON Flatten (csv)"
            Fases = fFase("Flatten CSV (não implementado)", Fases)

        Case "JSON Flatten (pares)"
            vFvalor = JSON_FlattenPares(vFvalor)
            Fases = fFase("Flatten Pares", Fases)

        Case Else
            vFvalor = prmTxt
            Fases = fFase("Operação não encontrada", Fases)

    End Select

Finalizar:
    FverSaida vFvalor
    IntegradorPipeline = gblOpc & "|" & vFvalor & "|" & Fases
    Exit Function

SaidaErr:
    Fases = "*** " & Err.Description & vbCrLf & fFase("Erro", Fases)
    IntegradorPipeline = gblOpc & "|" & vFvalor & "|" & Fases
End Function
