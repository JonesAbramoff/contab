VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPPItemPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'classe para importacao de item de pedido do aplicativo de pesquisa de preços

Option Explicit

Public Seq As Integer
Public ProdutoEan As String
Public ProdutoCodigo As String
Public Descricao As String
Public Quantidade As Double
Public PrecoUnitario As Double
Public Observacao As String

