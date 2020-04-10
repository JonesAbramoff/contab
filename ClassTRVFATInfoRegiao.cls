VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVFATInfoRegiao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarcolInfoClientes As New Collection
Private mvariRegiao As Integer
Private mvardValor As Double
Private mvardValorS As Double
Private mvariNumItens As Integer
Private mvariNumItensS As Integer
Private mvariNumFatS As Integer

Public Property Let iNumFatS(ByVal vData As Integer)
    mvariNumFatS = vData
End Property

Public Property Get iNumFatS() As Integer
    iNumFatS = mvariNumFatS
End Property

Public Property Let iNumItensS(ByVal vData As Integer)
    mvariNumItensS = vData
End Property

Public Property Get iNumItensS() As Integer
    iNumItensS = mvariNumItensS
End Property

Public Property Let iNumItens(ByVal vData As Integer)
    mvariNumItens = vData
End Property

Public Property Get iNumItens() As Integer
    iNumItens = mvariNumItens
End Property

Public Property Let dValorS(ByVal vData As Double)
    mvardValorS = vData
End Property

Public Property Get dValorS() As Double
    dValorS = mvardValorS
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let iRegiao(ByVal vData As Integer)
    mvariRegiao = vData
End Property

Public Property Get iRegiao() As Integer
    iRegiao = mvariRegiao
End Property

Public Property Set colInfoClientes(ByVal vData As Collection)
    Set mvarcolInfoClientes = vData
End Property

Public Property Get colInfoClientes() As Collection
    Set colInfoClientes = mvarcolInfoClientes
End Property
