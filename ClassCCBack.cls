VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCCBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mvarsProduto As String
Private mvarsDescProd As String
Private mvardQuantEst As Double 'quantidade total do produto na unidade de estoque
Private mvardValor As Double 'valor faturado (preço - desconto)
Private mvarlNumIntMovEst As Long
Private mvarsUnidadeMed As String
Private mvariClasseUM As Integer
Private mvariControleEstoque As Integer
Private mvarsUMEstoque As String

Public Property Let sUMEstoque(ByVal vData As String)
    mvarsUMEstoque = vData
End Property

Public Property Get sUMEstoque() As String
    sUMEstoque = mvarsUMEstoque
End Property

Public Property Let iControleEstoque(ByVal vData As Integer)
    mvariControleEstoque = vData
End Property

Public Property Get iControleEstoque() As Integer
    iControleEstoque = mvariControleEstoque
End Property

Public Property Let iClasseUM(ByVal vData As Integer)
    mvariClasseUM = vData
End Property

Public Property Get iClasseUM() As Integer
    iClasseUM = mvariClasseUM
End Property

Public Property Let sUnidadeMed(ByVal vData As String)
    mvarsUnidadeMed = vData
End Property

Public Property Get sUnidadeMed() As String
    sUnidadeMed = mvarsUnidadeMed
End Property

Public Property Let lNumIntMovEst(ByVal vData As Long)
    mvarlNumIntMovEst = vData
End Property

Public Property Get lNumIntMovEst() As Long
    lNumIntMovEst = mvarlNumIntMovEst
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dQuantEst(ByVal vData As Double)
    mvardQuantEst = vData
End Property

Public Property Get dQuantEst() As Double
    dQuantEst = mvardQuantEst
End Property

Public Property Let sDescProd(ByVal vData As String)
    mvarsDescProd = vData
End Property

Public Property Get sDescProd() As String
    sDescProd = mvarsDescProd
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

