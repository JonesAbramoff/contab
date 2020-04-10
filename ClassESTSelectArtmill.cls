VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassESTSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const TAXA_CONSUMO_TEMPO_PADRAO_ARTMILL = "min"

Function Remove_Pedido_Compra(ByVal objPedidoCompras As ClassPedidoCompras, ByVal colPedidoCompra As Collection, ByVal iIndice As Integer) As Long

On Error GoTo Erro_Remove_Pedido_Compra

    Remove_Pedido_Compra = SUCESSO

    Exit Function

Erro_Remove_Pedido_Compra:

    Remove_Pedido_Compra = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149709)

    End Select

    Exit Function

End Function

Function PrecoPrazo_Customizado(dPrecoPrazo As Double) As Long
    dPrecoPrazo = 0
    PrecoPrazo_Customizado = SUCESSO
End Function

Public Function Taxa_Producao_UM_Padrao_Obtem(sUM As String) As Long

    sUM = TAXA_CONSUMO_TEMPO_PADRAO_ARTMILL

    Taxa_Producao_UM_Padrao_Obtem = SUCESSO
End Function
