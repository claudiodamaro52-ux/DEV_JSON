Option Explicit
' ============================================================
'  MÓDULO: modGlobals
'  - Variáveis globais usadas pelo integrador
' ============================================================

Public gblOpc As String      ' operação selecionada no formulário
Public Nf As Long            ' contador de fases do pipeline

' Inicialização padrão
Public Sub InitGlobals()
    gblOpc = ""
    Nf = 0
End Sub

