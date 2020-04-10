VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelRoteiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

Private mvariNivel As Integer
Private mvarobjRoteiro As New ClassRoteirosDeFabricacao
Private mvarobjProduto As New ClassProduto
Private mvarcolRoteirosFilhos As New Collection
Private mvardFatorMultiplicacao As Double

Public Property Let dFatorMultiplicacao(ByVal vData As Double)
    mvardFatorMultiplicacao = vData
End Property

Public Property Get dFatorMultiplicacao() As Double
    dFatorMultiplicacao = mvardFatorMultiplicacao
End Property

Public Property Set colRoteirosFilhos(ByVal vData As Collection)
    Set mvarcolRoteirosFilhos = vData
End Property

Public Property Get colRoteirosFilhos() As Collection
    Set colRoteirosFilhos = mvarcolRoteirosFilhos
End Property

Public Property Set objRoteiro(ByVal vData As ClassRoteirosDeFabricacao)
    Set mvarobjRoteiro = vData
End Property

Public Property Get objRoteiro() As ClassRoteirosDeFabricacao
    Set objRoteiro = mvarobjRoteiro
End Property

Public Property Set objProduto(ByVal vData As ClassProduto)
    Set mvarobjProduto = vData
End Property

Public Property Get objProduto() As ClassProduto
    Set objProduto = mvarobjProduto
End Property

Public Property Let iNivel(ByVal vData As Integer)
    mvariNivel = vData
End Property

Public Property Get iNivel() As Integer
    iNivel = mvariNivel
End Property
