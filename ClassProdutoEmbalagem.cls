VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProdutoEmbalagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsProduto As String
Private mvariEmbalagem As Integer
Private mvariSeqGrid As Integer
Private mvariPadrao As Integer
Private mvarsNomeProdEmb As String
Private mvarsUMEmbalagem As String
Private mvardCapacidade As Double
Private mvarsUMPeso As String
Private mvardPesoLiqTotal As Double
Private mvardPesoEmbalagem As Double
Private mvardPesoBruto As Double

Public Property Let dPesoBruto(ByVal vData As Double)
    mvardPesoBruto = vData
End Property

Public Property Get dPesoBruto() As Double
    dPesoBruto = mvardPesoBruto
End Property

Public Property Let dPesoEmbalagem(ByVal vData As Double)
    mvardPesoEmbalagem = vData
End Property

Public Property Get dPesoEmbalagem() As Double
    dPesoEmbalagem = mvardPesoEmbalagem
End Property

Public Property Let dPesoLiqTotal(ByVal vData As Double)
    mvardPesoLiqTotal = vData
End Property

Public Property Get dPesoLiqTotal() As Double
    dPesoLiqTotal = mvardPesoLiqTotal
End Property

Public Property Let sUMPeso(ByVal vData As String)
    mvarsUMPeso = vData
End Property

Public Property Get sUMPeso() As String
    sUMPeso = mvarsUMPeso
End Property

Public Property Let dCapacidade(ByVal vData As Double)
    mvardCapacidade = vData
End Property

Public Property Get dCapacidade() As Double
    dCapacidade = mvardCapacidade
End Property

Public Property Let sUMEmbalagem(ByVal vData As String)
    mvarsUMEmbalagem = vData
End Property

Public Property Get sUMEmbalagem() As String
    sUMEmbalagem = mvarsUMEmbalagem
End Property

Public Property Let sNomeProdEmb(ByVal vData As String)
    mvarsNomeProdEmb = vData
End Property

Public Property Get sNomeProdEmb() As String
    sNomeProdEmb = mvarsNomeProdEmb
End Property

Public Property Let iEmbalagem(ByVal vData As Integer)
    mvariEmbalagem = vData
End Property

Public Property Get iEmbalagem() As Integer
    iEmbalagem = mvariEmbalagem
End Property

Public Property Let iPadrao(ByVal vData As Integer)
    mvariPadrao = vData
End Property

Public Property Get iPadrao() As Integer
    iPadrao = mvariPadrao
End Property

Public Property Let iSeqGrid(ByVal vData As Integer)
    mvariSeqGrid = vData
End Property

Public Property Get iSeqGrid() As Integer
    iSeqGrid = mvariSeqGrid
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

