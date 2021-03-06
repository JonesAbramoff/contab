VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVFATInfoFilial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarcolInfoClientes As New Collection
Private mvariFilialEmpresa As Integer
Private mvardValor As Double
Private mvardValorS As Double
Private mvarlNumItens As Long
Private mvarlNumItensS As Long
Private mvarlNumFatS As Long
Private mvarsNome As String

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let lNumFatS(ByVal vData As Long)
    mvarlNumFatS = vData
End Property

Public Property Get lNumFatS() As Long
    lNumFatS = mvarlNumFatS
End Property

Public Property Let lNumItensS(ByVal vData As Long)
    mvarlNumItensS = vData
End Property

Public Property Get lNumItensS() As Long
    lNumItensS = mvarlNumItensS
End Property

Public Property Let lNumItens(ByVal vData As Long)
    mvarlNumItens = vData
End Property

Public Property Get lNumItens() As Long
    lNumItens = mvarlNumItens
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

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Set colInfoClientes(ByVal vData As Collection)
    Set mvarcolInfoClientes = vData
End Property

Public Property Get colInfoClientes() As Collection
    Set colInfoClientes = mvarcolInfoClientes
End Property
