VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRPFATInfoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarcolInfoFaturas As New Collection
Private mvarlCliente As Long
Private mvariFilialEmpresa As Integer
Private mvariCondPagto As Integer
Private mvardValor As Double
Private mvardValorS As Double
Private mvariNumItens As Integer
Private mvariNumItensS As Integer
Private mvariMarcado As Integer
Private mvarsNome As String
Private mvariLinha As Integer
Private mvariTipo As Integer

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let iLinha(ByVal vData As Integer)
    mvariLinha = vData
End Property

Public Property Get iLinha() As Integer
    iLinha = mvariLinha
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let iMarcado(ByVal vData As Integer)
    mvariMarcado = vData
End Property

Public Property Get iMarcado() As Integer
    iMarcado = mvariMarcado
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

Public Property Let iCondPagto(ByVal vData As Integer)
    mvariCondPagto = vData
End Property

Public Property Get iCondPagto() As Integer
    iCondPagto = mvariCondPagto
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Set colInfoFaturas(ByVal vData As Collection)
    Set mvarcolInfoFaturas = vData
End Property

Public Property Get colInfoFaturas() As Collection
    Set colInfoFaturas = mvarcolInfoFaturas
End Property
