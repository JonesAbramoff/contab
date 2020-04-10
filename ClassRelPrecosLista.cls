VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelPrecosLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntRel As Long
Private mvarsProduto As String
Private mvarsDescricao As String
Private mvardValor1 As Double
Private mvardValor2 As Double
Private mvardValor3 As Double
Private mvardValor4 As Double
Private mvardEmbPadrao As Double
Private mvardKGsOrcados As Double

Public Property Let dKGsOrcados(ByVal vData As Double)
    mvardKGsOrcados = vData
End Property

Public Property Get dKGsOrcados() As Double
    dKGsOrcados = mvardKGsOrcados
End Property

Public Property Let dEmbPadrao(ByVal vData As Double)
    mvardEmbPadrao = vData
End Property

Public Property Get dEmbPadrao() As Double
    dEmbPadrao = mvardEmbPadrao
End Property

Public Property Let dValor4(ByVal vData As Double)
    mvardValor4 = vData
End Property

Public Property Get dValor4() As Double
    dValor4 = mvardValor4
End Property

Public Property Let dvalor3(ByVal vData As Double)
    mvardValor3 = vData
End Property

Public Property Get dvalor3() As Double
    dvalor3 = mvardValor3
End Property

Public Property Let dvalor2(ByVal vData As Double)
    mvardValor2 = vData
End Property

Public Property Get dvalor2() As Double
    dvalor2 = mvardValor2
End Property

Public Property Let dvalor1(ByVal vData As Double)
    mvardValor1 = vData
End Property

Public Property Get dvalor1() As Double
    dvalor1 = mvardValor1
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let lNumIntRel(ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel() As Long
    lNumIntRel = mvarlNumIntRel
End Property
