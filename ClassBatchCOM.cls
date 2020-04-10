VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBatchCOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Const TITULO_TELABATCH_BAIXA_REQCOMPRAS = "Baixa de Requisição de Compras"
Const ROTINA_REQCOMPRAS_BAIXAR = 2

Function Rotina_ReqComprasBaixar_Batch(ByVal sNomeArqParam As String, ByVal colReqComprasInfo As Collection) As Long

On Error GoTo Erro_Rotina_ReqComprasBaixar_Batch

    TelaAcompanhaBatchCOM.Caption = TITULO_TELABATCH_BAIXA_REQCOMPRAS
    TelaAcompanhaBatchCOM.sNomeArqParam = sNomeArqParam
    TelaAcompanhaBatchCOM.iRotinaBatch = ROTINA_REQCOMPRAS_BAIXAR
    Set TelaAcompanhaBatchCOM.colReqComprasInfo = colReqComprasInfo

    'Exibe a tela de acompanhamento do batch
    TelaAcompanhaBatchCOM.Show

    Rotina_ReqComprasBaixar_Batch = SUCESSO

    Exit Function

Erro_Rotina_ReqComprasBaixar_Batch:

    Rotina_ReqComprasBaixar_Batch = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144646)

    End Select

    Exit Function

End Function

Function Rotina_CalculaPtoPedido(ByVal sNomeArqParam As String) As Long

On Error GoTo Erro_Rotina_CalculaPtoPedido

    TelaAcompanhaBatchCOM.Caption = TITULO_TELA_BACH_CALCULO_PTOPEDIDO
    TelaAcompanhaBatchCOM.sNomeArqParam = sNomeArqParam
    TelaAcompanhaBatchCOM.iRotinaBatch = ROTINA_CALCULO_PTOPEDIDO

    'Exibe a tela de acompanhamento do batch
    TelaAcompanhaBatchCOM.Show

    Rotina_CalculaPtoPedido = SUCESSO

    Exit Function

Erro_Rotina_CalculaPtoPedido:

    Rotina_CalculaPtoPedido = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144647)

    End Select

    Exit Function

End Function


