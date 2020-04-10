VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassTipoFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvarsDescricao As String
Private mvardDesconto As Double
Private mvarsObservacao As String
Private mvariHistPadraoDespesa As Integer
Private mvarsContaDespesa As String
Private mvariCondicaoPagto As Integer

Public Property Let iCondicaoPagto(ByVal vData As Integer)
    mvariCondicaoPagto = vData
End Property

Public Property Get iCondicaoPagto() As Integer
    iCondicaoPagto = mvariCondicaoPagto
End Property

Public Property Let sContaDespesa(ByVal vData As String)
    mvarsContaDespesa = vData
End Property

Public Property Get sContaDespesa() As String
    sContaDespesa = mvarsContaDespesa
End Property

Public Property Let iHistPadraoDespesa(ByVal vData As Integer)
    mvariHistPadraoDespesa = vData
End Property

Public Property Get iHistPadraoDespesa() As Integer
    iHistPadraoDespesa = mvariHistPadraoDespesa
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let dDesconto(ByVal vData As Double)
    mvardDesconto = vData
End Property

Public Property Get dDesconto() As Double
    dDesconto = mvardDesconto
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property



