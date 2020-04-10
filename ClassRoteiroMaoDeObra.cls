VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRoteiroMaoDeObra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariCodMO As Integer
Private mvarsUM As String
Private mvardHoras As Double
Private mvardCustoTotal As Double
Private mvardCustoUnitario As Double
Private mvariQuantidade As Integer
Private mvarobjMaoDeObra As New ClassTiposDeMaodeObra

Public Property Set objMaoDeObra(ByVal vData As ClassTiposDeMaodeObra)
    Set mvarobjMaoDeObra = vData
End Property

Public Property Get objMaoDeObra() As ClassTiposDeMaodeObra
    Set objMaoDeObra = mvarobjMaoDeObra
End Property

Public Property Let iQuantidade(ByVal vData As Integer)
    mvariQuantidade = vData
End Property

Public Property Get iQuantidade() As Integer
    iQuantidade = mvariQuantidade
End Property

Public Property Let dCustoUnitario(ByVal vData As Double)
    mvardCustoUnitario = vData
End Property

Public Property Get dCustoUnitario() As Double
    dCustoUnitario = mvardCustoUnitario
End Property

Public Property Let dCustoTotal(ByVal vData As Double)
    mvardCustoTotal = vData
End Property

Public Property Get dCustoTotal() As Double
    dCustoTotal = mvardCustoTotal
End Property

Public Property Let dHoras(ByVal vData As Double)
    mvardHoras = vData
End Property

Public Property Get dHoras() As Double
    dHoras = mvardHoras
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let iCodMO(ByVal vData As Integer)
    mvariCodMO = vData
End Property

Public Property Get iCodMO() As Integer
    iCodMO = mvariCodMO
End Property
