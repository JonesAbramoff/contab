VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProdutoCusto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsCodProduto As String
Private mvardCusto As Double
Private mvardCustoMesAnterior As Double
Private mvarsDescProduto As String
Private mvarsSiglaUMEstoque As String

Public Property Let sSiglaUMEstoque(ByVal vData As String)
    mvarsSiglaUMEstoque = vData
End Property

Public Property Get sSiglaUMEstoque() As String
    sSiglaUMEstoque = mvarsSiglaUMEstoque
End Property

Public Property Let sDescProduto(ByVal vData As String)
    mvarsDescProduto = vData
End Property

Public Property Get sDescProduto() As String
    sDescProduto = mvarsDescProduto
End Property

Public Property Let dCusto(ByVal vData As Double)
    mvardCusto = vData
End Property

Public Property Get dCusto() As Double
    dCusto = mvardCusto
End Property

Public Property Let dCustoMesAnterior(ByVal vData As Double)
    mvardCustoMesAnterior = vData
End Property

Public Property Get dCustoMesAnterior() As Double
    dCustoMesAnterior = mvardCustoMesAnterior
End Property

Public Property Let sCodProduto(ByVal vData As String)
    mvarsCodProduto = vData
End Property

Public Property Get sCodProduto() As String
    sCodProduto = mvarsCodProduto
End Property



