VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelacClientesCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlCodigoDe As Long
Private mvarlCodigoAte As Long
Private mvariFilialEmpresa As Integer
Private mvariOrigem As Integer
Private mvardtDataDe As Date
Private mvardtDataAte As Date
Private mvarlTipo As Long
Private mvarlCliente As Long
Private mvariFilialCliente As Integer
Private mvariAtendenteDe As Integer
Private mvariAtendenteAte As Integer
Private mvariStatus As Integer
Private mvarcolRelacionamentoClientes As New Collection

Public Property Let lCodigoDe(ByVal vData As Long)
    mvarlCodigoDe = vData
End Property

Public Property Get lCodigoDe() As Long
    lCodigoDe = mvarlCodigoDe
End Property

Public Property Let lCodigoAte(ByVal vData As Long)
    mvarlCodigoAte = vData
End Property

Public Property Get lCodigoAte() As Long
    lCodigoAte = mvarlCodigoAte
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iOrigem(ByVal vData As Integer)
    mvariOrigem = vData
End Property

Public Property Get iOrigem() As Integer
    iOrigem = mvariOrigem
End Property

Public Property Let dtDataDe(ByVal vData As Date)
    mvardtDataDe = vData
End Property

Public Property Get dtDataDe() As Date
    dtDataDe = mvardtDataDe
End Property

Public Property Let dtDataAte(ByVal vData As Date)
    mvardtDataAte = vData
End Property

Public Property Get dtDataAte() As Date
    dtDataAte = mvardtDataAte
End Property

Public Property Let lTipo(ByVal vData As Long)
    mvarlTipo = vData
End Property

Public Property Get lTipo() As Long
    lTipo = mvarlTipo
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let iFilialCliente(ByVal vData As Integer)
    mvariFilialCliente = vData
End Property

Public Property Get iFilialCliente() As Integer
    iFilialCliente = mvariFilialCliente
End Property

Public Property Let iAtendenteDe(ByVal vData As Integer)
    mvariAtendenteDe = vData
End Property

Public Property Get iAtendenteDe() As Integer
    iAtendenteDe = mvariAtendenteDe
End Property

Public Property Let iAtendenteAte(ByVal vData As Integer)
    mvariAtendenteAte = vData
End Property

Public Property Get iAtendenteAte() As Integer
    iAtendenteAte = mvariAtendenteAte
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Set colRelacionamentoClientes(ByVal vData As Object)
    Set mvarcolRelacionamentoClientes = vData
End Property

Public Property Get colRelacionamentoClientes() As Collection
    Set colRelacionamentoClientes = mvarcolRelacionamentoClientes
End Property

