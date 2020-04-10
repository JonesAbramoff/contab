VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCategoriaProdutoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsCategoria As String
Private mvarsItem As String
Private mvariOrdem As Integer
Private mvarsDescricao As String
Private mvardValor1 As Double
Private mvardValor2 As Double
Private mvardValor3 As Double
Private mvardValor4 As Double
Private mvardValor5 As Double
Private mvardValor6 As Double
Private mvardValor7 As Double
Private mvardValor8 As Double

'variaveis auxiliares, usadas no rateio de custo fixo
Private mvardValorAux1 As Double
Private mvardValorAux2 As Double

Public Property Let iOrdem(ByVal vData As Integer)
    mvariOrdem = vData
End Property

Public Property Get iOrdem() As Integer
    iOrdem = mvariOrdem
End Property

Public Property Let sItem(ByVal vData As String)
    mvarsItem = vData
End Property

Public Property Get sItem() As String
    sItem = mvarsItem
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sCategoria(ByVal vData As String)
    mvarsCategoria = vData
End Property

Public Property Get sCategoria() As String
    sCategoria = mvarsCategoria
End Property


Public Property Let dvalor1(ByVal vData As Double)
    mvardValor1 = vData
End Property

Public Property Get dvalor1() As Double
    dvalor1 = mvardValor1
End Property

Public Property Let dvalor2(ByVal vData As Double)
    mvardValor2 = vData
End Property

Public Property Get dvalor2() As Double
    dvalor2 = mvardValor2
End Property

Public Property Let dvalor3(ByVal vData As Double)
    mvardValor3 = vData
End Property

Public Property Get dvalor3() As Double
    dvalor3 = mvardValor3
End Property

Public Property Let dvalor4(ByVal vData As Double)
    mvardValor4 = vData
End Property

Public Property Get dvalor4() As Double
    dvalor4 = mvardValor4
End Property

Public Property Let dvalor5(ByVal vData As Double)
    mvardValor5 = vData
End Property

Public Property Get dvalor5() As Double
    dvalor5 = mvardValor5
End Property

Public Property Let dvalor6(ByVal vData As Double)
    mvardValor6 = vData
End Property

Public Property Get dvalor6() As Double
    dvalor6 = mvardValor6
End Property

Public Property Let dvalor7(ByVal vData As Double)
    mvardValor7 = vData
End Property

Public Property Get dvalor7() As Double
    dvalor7 = mvardValor7
End Property

Public Property Let dvalor8(ByVal vData As Double)
    mvardValor8 = vData
End Property

Public Property Get dvalor8() As Double
    dvalor8 = mvardValor8
End Property

Public Property Let dvalorAux1(ByVal vData As Double)
    mvardValorAux1 = vData
End Property

Public Property Get dvalorAux1() As Double
    dvalorAux1 = mvardValorAux1
End Property

Public Property Let dvalorAux2(ByVal vData As Double)
    mvardValorAux2 = vData
End Property

Public Property Get dvalorAux2() As Double
    dvalorAux2 = mvardValorAux2
End Property

