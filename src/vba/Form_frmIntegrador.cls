VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmIntegrador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ============================================================
'  FORM: frmIntegrador (FASE 1 - CONTRATO v0.6)
'  - Toda entrada externa passa por NormalizarEntradaExterna
'  - Integrador devolve OP|PAYLOAD|FASES
' ============================================================

Private Sub cmdTeste_Click()
On Error GoTo SaidaErr

    Dim txtEntrada As String
    Dim txtResp As String
    Dim aResp As Variant
    Dim vMsg As String

    vMsg = ""

    Me!Msg = ""
    Me!ctSaida = ""
    Me!ctAndamento = ""

    txtEntrada = Nz(Me!ctEntrada, "")
    If Trim$(txtEntrada) = "" Then
        Me!Msg = "Entrada vazia."
        Exit Sub
    End If

    gblOpc = Nz(Me!cbOpc, "")
    If Trim$(gblOpc) = "" Then
        Me!Msg = "Selecione uma operação."
        Exit Sub
    End If

    ' Obrigatório no DEV_JSON: normaliza entrada externa
    txtEntrada = NormalizarEntradaExterna(txtEntrada)

    txtResp = IntegradorPipeline(txtEntrada)

    aResp = Split(txtResp, "|")

    If UBound(aResp) >= 1 Then Me!ctSaida = aResp(1)
    If UBound(aResp) >= 2 Then Me!ctAndamento = Trim$(aResp(2))
    Me!Timer = Now()

    vMsg = vMsg & "Teste de: " & Me!cbOpc & vbCrLf
    vMsg = vMsg & "Entrada:" & vbCrLf & Me!ctEntrada & vbCrLf
    vMsg = vMsg & "Andamento:" & vbCrLf & Me!ctAndamento & vbCrLf
    vMsg = vMsg & "Saida:" & vbCrLf & Me!ctSaida

    Me!Msg = vMsg
    Exit Sub

SaidaErr:
    Me!Msg = "Erro: " & Err.Description
End Sub
Private Sub popCbOpc()
    Dim lista As String
    lista = fListaCases()
    Me!cbOpc.RowSource = Replace(lista, vbCrLf, ";")
End Sub

Private Sub Form_Open(Cancel As Integer)
    popCbOpc
    InitGlobals
    DoCmd.GoToRecord acDataForm, Me.Name, acLast
End Sub

Private Sub Form_Current()
    If Me!cbOpc.ListCount = 0 Then popCbOpc
End Sub
