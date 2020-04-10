VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEmbProdItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Janaina
Private mvariEmbalagem As Integer
Private mvarsUMEmbalagem As String
Private mvarsUMPeso As String
Private mvardCapacidade As Double
Private mvariTipoDocOrigem As Integer
Private mvardQuantEmb As Double
Private mvardQuantProduto As Double
Private mvardPesoLiqTotal As Double
Private mvardPesoBruto As Double
Private mvarsProduto As String
Private mvariItem As String

Public Property Let iTipoDocOrigem(ByVal vData As Integer)
    mvariTipoDocOrigem = vData
End Property

Public Property Get iTipoDocOrigem() As Integer
    iTipoDocOrigem = mvariTipoDocOrigem
End Property

Public Property Let iItem(ByVal vData As Integer)
    mvariItem = vData
End Property

Public Property Get iItem() As Integer
    iItem = mvariItem
End Property

Public Property Let iEmbalagem(ByVal vData As Integer)
    mvariEmbalagem = vData
End Property

Public Property Get iEmbalagem() As Integer
    iEmbalagem = mvariEmbalagem
End Property

Public Property Let sUMPeso(ByVal vData As String)
    mvarsUMPeso = vData
End Property

Public Property Get sUMPeso() As String
    sUMPeso = mvarsUMPeso
End Property

Public Property Let sUMEmbalagem(ByVal vData As String)
    mvarsUMEmbalagem = vData
End Property

Public Property Get sUMEmbalagem() As String
    sUMEmbalagem = mvarsUMEmbalagem
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let dCapacidade(ByVal vData As Double)
    mvardCapacidade = vData
End Property

Public Property Get dCapacidade() As Double
    dCapacidade = mvardCapacidade
End Property

Public Property Let dQuantEmb(ByVal vData As Double)
    mvardQuantEmb = vData
End Property

Public Property Get dQuantEmb() As Double
    dQuantEmb = mvardQuantEmb
End Property

Public Property Let dQuantProduto(ByVal vData As Double)
    mvardQuantProduto = vData
End Property

Public Property Get dQuantProduto() As Double
    dQuantProduto = mvardQuantProduto
End Property

Public Property Let dPesoLiqTotal(ByVal vData As Double)
    mvardPesoLiqTotal = vData
End Property

Public Property Get dPesoLiqTotal() As Double
    dPesoLiqTotal = mvardPesoLiqTotal
End Property

Public Property Let dPesoBruto(ByVal vData As Double)
    mvardPesoBruto = vData
End Property

Public Property Get dPesoBruto() As Double
    dPesoBruto = mvardPesoBruto
End Property
'Janaina
