VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassOSAp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarcolMO As New Collection
Private mvarcolMaq As New Collection
Private mvariFilialEmpresa As Integer
Private mvarlCodigo As Long
Private mvarlCodigoMovEst As Long
Private mvardtData As Date
Private mvarlCliente As Long
Private mvardtHora As Date
Private mvarlNumIntDoc As Long

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let dtHora(ByVal vData As Date)
    mvardtHora = vData
End Property

Public Property Get dtHora() As Date
    dtHora = mvardtHora
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

Public Property Let lCodigoMovEst(ByVal vData As Long)
    mvarlCodigoMovEst = vData
End Property

Public Property Get lCodigoMovEst() As Long
    lCodigoMovEst = mvarlCodigoMovEst
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Set colMaq(ByVal vData As Collection)
    Set mvarcolMaq = vData
End Property

Public Property Get colMaq() As Collection
    Set colMaq = mvarcolMaq
End Property

Public Property Set colMO(ByVal vData As Collection)
    Set mvarcolMO = vData
End Property

Public Property Get colMO() As Collection
    Set colMO = mvarcolMO
End Property

