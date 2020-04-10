VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCustoDirFabr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarobjTela As Object

Private mvariFilialEmpresa As Integer
Private mvariAno As Integer
Private mvarsCodigoPrevVenda As String
Private mvariMesIni As Integer
Private mvariMesFim As Integer
Private mvardtData As Date
Private mvardCustoTotal As Double
Private mvardQuantFator1 As Double
Private mvardCustoFator1 As Double
Private mvardQuantFator2 As Double
Private mvardCustoFator2 As Double
Private mvardQuantFator3 As Double
Private mvardCustoFator3 As Double
Private mvardQuantFator4 As Double
Private mvardCustoFator4 As Double
Private mvardQuantFator5 As Double
Private mvardCustoFator5 As Double
Private mvardQuantFator6 As Double
Private mvardCustoFator6 As Double
Private mvarsProduto As String
Private mvardCustoOutros As Double

Private mvariRotinaOrigem As Integer

Public Property Let iMesFim(ByVal vData As Integer)
    mvariMesFim = vData
End Property

Public Property Get iMesFim() As Integer
    iMesFim = mvariMesFim
End Property

Public Property Let iMesIni(ByVal vData As Integer)
    mvariMesIni = vData
End Property

Public Property Get iMesIni() As Integer
    iMesIni = mvariMesIni
End Property

Public Property Let iRotinaOrigem(ByVal vData As Integer)
    mvariRotinaOrigem = vData
End Property

Public Property Get iRotinaOrigem() As Integer
    iRotinaOrigem = mvariRotinaOrigem
End Property

Public Property Let dCustoOutros(ByVal vData As Double)
    mvardCustoOutros = vData
End Property

Public Property Get dCustoOutros() As Double
    dCustoOutros = mvardCustoOutros
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let sCodigoPrevVenda(ByVal vData As String)
    mvarsCodigoPrevVenda = vData
End Property

Public Property Get sCodigoPrevVenda() As String
    sCodigoPrevVenda = mvarsCodigoPrevVenda
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let dQuantFator4(ByVal vData As Double)
    mvardQuantFator4 = vData
End Property

Public Property Get dQuantFator4() As Double
    dQuantFator4 = mvardQuantFator4
End Property

Public Property Let dCustoFator4(ByVal vData As Double)
    mvardCustoFator4 = vData
End Property

Public Property Get dCustoFator4() As Double
    dCustoFator4 = mvardCustoFator4
End Property

Public Property Let dQuantFator3(ByVal vData As Double)
    mvardQuantFator3 = vData
End Property

Public Property Get dQuantFator3() As Double
    dQuantFator3 = mvardQuantFator3
End Property

Public Property Let dCustoFator3(ByVal vData As Double)
    mvardCustoFator3 = vData
End Property

Public Property Get dCustoFator3() As Double
    dCustoFator3 = mvardCustoFator3
End Property

Public Property Let dQuantFator2(ByVal vData As Double)
    mvardQuantFator2 = vData
End Property

Public Property Get dQuantFator2() As Double
    dQuantFator2 = mvardQuantFator2
End Property

Public Property Let dCustoFator2(ByVal vData As Double)
    mvardCustoFator2 = vData
End Property

Public Property Get dCustoFator2() As Double
    dCustoFator2 = mvardCustoFator2
End Property

Public Property Let dQuantFator1(ByVal vData As Double)
    mvardQuantFator1 = vData
End Property

Public Property Get dQuantFator1() As Double
    dQuantFator1 = mvardQuantFator1
End Property

Public Property Let dCustoFator1(ByVal vData As Double)
    mvardCustoFator1 = vData
End Property

Public Property Get dCustoFator1() As Double
    dCustoFator1 = mvardCustoFator1
End Property

Public Property Set objTela(ByVal vData As Object)
    Set mvarobjTela = vData
End Property

Public Property Get objTela() As Object
    Set objTela = mvarobjTela
End Property

Public Property Let dCustoTotal(ByVal vData As Double)
    mvardCustoTotal = vData
End Property

Public Property Get dCustoTotal() As Double
    dCustoTotal = mvardCustoTotal
End Property

Public Property Let iAno(ByVal vData As Integer)
    mvariAno = vData
End Property

Public Property Get iAno() As Integer
    iAno = mvariAno
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let dQuantFator5(ByVal vData As Double)
    mvardQuantFator5 = vData
End Property

Public Property Get dQuantFator5() As Double
    dQuantFator5 = mvardQuantFator5
End Property

Public Property Let dCustoFator5(ByVal vData As Double)
    mvardCustoFator5 = vData
End Property

Public Property Get dCustoFator5() As Double
    dCustoFator5 = mvardCustoFator5
End Property

Public Property Let dQuantFator6(ByVal vData As Double)
    mvardQuantFator6 = vData
End Property

Public Property Get dQuantFator6() As Double
    dQuantFator6 = mvardQuantFator6
End Property

Public Property Let dCustoFator6(ByVal vData As Double)
    mvardCustoFator6 = vData
End Property

Public Property Get dCustoFator6() As Double
    dCustoFator6 = mvardCustoFator6
End Property


