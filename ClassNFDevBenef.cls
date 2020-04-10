VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFDevBenef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarlCliente As Long
Private mvariFilialCliente As Integer
Private mvardtDataUltDev As Date
Private mvardtData As Date
Private mvariRetorno As Integer

Private mvarcolOP As New Collection 'Contém as OPs a serem consideradas
Private mvarcolEntProd As New Collection 'Detalha todas as entradas de produção a serem consideradas
Private mvarcolReqProd As New Collection 'Detalha todas as requisições de produção a serem consideradas
Private mvarcolItemOP As New Collection 'Dados das produções entradas resumidos por item de OP

Private mvarcolItens As New Collection 'Guardará os produtos e quantidade calculados com os seus devidos preços
Private mvarcolItensDev As New Collection 'Guardará a quantidade total a ser devolvida sem considerar os preços

Public Property Set colItemOP(ByVal vData As Object)
    Set mvarcolItemOP = vData
End Property

Public Property Get colItemOP() As Collection
    Set colItemOP = mvarcolItemOP
End Property

Public Property Set colReqProd(ByVal vData As Object)
    Set mvarcolReqProd = vData
End Property

Public Property Get colReqProd() As Collection
    Set colReqProd = mvarcolReqProd
End Property

Public Property Set colItensDev(ByVal vData As Object)
    Set mvarcolItensDev = vData
End Property

Public Property Get colItensDev() As Collection
    Set colItensDev = mvarcolItensDev
End Property

Public Property Set colItens(ByVal vData As Object)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Set colEntProd(ByVal vData As Object)
    Set mvarcolEntProd = vData
End Property

Public Property Get colEntProd() As Collection
    Set colEntProd = mvarcolEntProd
End Property

Public Property Set colOP(ByVal vData As Object)
    Set mvarcolOP = vData
End Property

Public Property Get colOP() As Collection
    Set colOP = mvarcolOP
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let dtDataUltDev(ByVal vData As Date)
    mvardtDataUltDev = vData
End Property

Public Property Get dtDataUltDev() As Date
    dtDataUltDev = mvardtDataUltDev
End Property

Public Property Let iFilialCliente(ByVal vData As Integer)
    mvariFilialCliente = vData
End Property

Public Property Get iFilialCliente() As Integer
    iFilialCliente = mvariFilialCliente
End Property

Public Property Let iRetorno(ByVal vData As Integer)
    mvariRetorno = vData
End Property

Public Property Get iRetorno() As Integer
    iRetorno = mvariRetorno
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Private Sub Class_Initialize()
    mvardtData = DATA_NULA
    mvardtDataUltDev = DATA_NULA
End Sub
