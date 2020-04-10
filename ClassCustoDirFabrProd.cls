VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCustoDirFabrProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvariAno As Integer
Private mvarsProduto As String
Private mvardtData As Date
Private mvardQuantPrevista As Double
Private mvardQuantFator1 As Double
Private mvardQuantFator2 As Double
Private mvardQuantFator3 As Double
Private mvardQuantFator4 As Double
Private mvardQuantFator5 As Double
Private mvardQuantFator6 As Double

Private mvarcolPlan As New Collection

Private mvardCusto As Double 'parte do custo total de fabricacao rateado para o produto

Public Property Let dCusto(ByVal vData As Double)
    mvardCusto = vData
End Property

Public Property Get dCusto() As Double
    dCusto = mvardCusto
End Property

Public Property Let dQuantFator4(ByVal vData As Double)
    mvardQuantFator4 = vData
End Property

Public Property Get dQuantFator4() As Double
    dQuantFator4 = mvardQuantFator4
End Property

Public Property Let dQuantFator3(ByVal vData As Double)
    mvardQuantFator3 = vData
End Property

Public Property Get dQuantFator3() As Double
    dQuantFator3 = mvardQuantFator3
End Property

Public Property Let dQuantFator2(ByVal vData As Double)
    mvardQuantFator2 = vData
End Property

Public Property Get dQuantFator2() As Double
    dQuantFator2 = mvardQuantFator2
End Property

Public Property Let dQuantFator1(ByVal vData As Double)
    mvardQuantFator1 = vData
End Property

Public Property Get dQuantFator1() As Double
    dQuantFator1 = mvardQuantFator1
End Property

Public Property Let dQuantPrevista(ByVal vData As Double)
    mvardQuantPrevista = vData
End Property

Public Property Get dQuantPrevista() As Double
    dQuantPrevista = mvardQuantPrevista
End Property

Public Property Set colPlan(ByVal vData As Object)
    Set mvarcolPlan = vData
End Property

Public Property Get colPlan() As Collection
    Set colPlan = mvarcolPlan
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
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

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let dQuantFator5(ByVal vData As Double)
    mvardQuantFator5 = vData
End Property

Public Property Get dQuantFator5() As Double
    dQuantFator5 = mvardQuantFator5
End Property

Public Property Let dQuantFator6(ByVal vData As Double)
    mvardQuantFator6 = vData
End Property

Public Property Get dQuantFator6() As Double
    dQuantFator6 = mvardQuantFator6
End Property


