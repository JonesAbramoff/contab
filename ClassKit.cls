VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassKit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarsProdutoRaiz As String
Private mvarsVersao As String
Private mvardtData As Date
Private mvarsObservacao As String
Private mvarcolComponentes As New Collection
Private mvariSituacao As Integer
Private mvardPesoFator1 As Double
Private mvardPesoFator2 As Double
Private mvardPesoFator3 As Double
Private mvardPesoFator4 As Double
Private mvardPesoFator5 As Double
Private mvardPesoFator6 As Double
Private mvariVersaoFormPreco As Integer

Public Property Let iVersaoFormPreco(ByVal vData As Integer)
    mvariVersaoFormPreco = vData
End Property

Public Property Get iVersaoFormPreco() As Integer
    iVersaoFormPreco = mvariVersaoFormPreco
End Property

Public Property Let dPesoFator1(ByVal vData As Double)
    mvardPesoFator1 = vData
End Property

Public Property Get dPesoFator1() As Double
    dPesoFator1 = mvardPesoFator1
End Property

Public Property Let dPesoFator2(ByVal vData As Double)
    mvardPesoFator2 = vData
End Property

Public Property Get dPesoFator2() As Double
    dPesoFator2 = mvardPesoFator2
End Property

Public Property Let dPesoFator3(ByVal vData As Double)
    mvardPesoFator3 = vData
End Property

Public Property Get dPesoFator3() As Double
    dPesoFator3 = mvardPesoFator3
End Property

Public Property Let dPesoFator4(ByVal vData As Double)
    mvardPesoFator4 = vData
End Property

Public Property Get dPesoFator4() As Double
    dPesoFator4 = mvardPesoFator4
End Property

Public Property Let iSituacao(ByVal vData As Integer)
    mvariSituacao = vData
End Property

Public Property Get iSituacao() As Integer
    iSituacao = mvariSituacao
End Property

Public Property Set colComponentes(ByVal vData As Object)
    Set mvarcolComponentes = vData
End Property

Public Property Get colComponentes() As Collection
    Set colComponentes = mvarcolComponentes
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let sVersao(ByVal vData As String)
    mvarsVersao = vData
End Property

Public Property Get sVersao() As String
    sVersao = mvarsVersao
End Property

Public Property Let sProdutoRaiz(ByVal vData As String)
    mvarsProdutoRaiz = vData
End Property

Public Property Get sProdutoRaiz() As String
    sProdutoRaiz = mvarsProdutoRaiz
End Property

Public Property Let dPesoFator5(ByVal vData As Double)
    mvardPesoFator5 = vData
End Property

Public Property Get dPesoFator5() As Double
    dPesoFator5 = mvardPesoFator5
End Property

Public Property Let dPesoFator6(ByVal vData As Double)
    mvardPesoFator6 = vData
End Property

Public Property Get dPesoFator6() As Double
    dPesoFator6 = mvardPesoFator6
End Property


