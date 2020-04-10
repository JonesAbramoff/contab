VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemPedidoWB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colReserva"
Option Explicit

Private mvardPrecoTabela As Double 'preco unitario da tabela de precos, que inclui o IPI
Private mvardDescTabela As Double 'percentual de desconto sobre o dPrecoTabela para chegar ao preço ario final

Public Property Let dDescTabela(ByVal vData As Double)
    mvardDescTabela = vData
End Property

Public Property Get dDescTabela() As Double
    dDescTabela = mvardDescTabela
End Property

Public Property Let dPrecoTabela(ByVal vData As Double)
    mvardPrecoTabela = vData
End Property

Public Property Get dPrecoTabela() As Double
    dPrecoTabela = mvardPrecoTabela
End Property


