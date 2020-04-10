VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSolicSRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarlCodigo As Long
Private mvardtData As Date
Private mvardtHora As Date
Private mvarlCliente As Long
Private mvariFilial As Integer
Private mvarlNumIntDoc As Long
Private mvariVendedor As Integer
Private mvariAtendente As Integer
Private mvarcolItens As New Collection
Private mvarobjTela As Object
Private mvariStatusItem As Integer
Private mvarlClienteBenef As Long
Private mvariFilialClienteBenef As Integer
Private mvariPrazo As Integer
Private mvardtDataEntrega As Date
Private mvariPrazoTipo As Integer
Private mvarobjRelacCli As Object
Private mvarsObs As String
Private mvariGravarCRM As Integer
Private mvarlTipo As Long
Private mvarlFase As Long

Public Property Let lTipo(ByVal vData As Long)
    mvarlTipo = vData
End Property

Public Property Get lTipo() As Long
    lTipo = mvarlTipo
End Property

Public Property Let lFase(ByVal vData As Long)
    mvarlFase = vData
End Property

Public Property Get lFase() As Long
    lFase = mvarlFase
End Property

Public Property Let iGravarCRM(ByVal vData As Integer)
    mvariGravarCRM = vData
End Property

Public Property Get iGravarCRM() As Integer
    iGravarCRM = mvariGravarCRM
End Property

Public Property Let sObs(ByVal vData As String)
    mvarsObs = vData
End Property

Public Property Get sObs() As String
    sObs = mvarsObs
End Property

Public Property Set objRelacCli(ByVal vData As Object)
    Set mvarobjRelacCli = vData
End Property

Public Property Get objRelacCli() As Object
    Set objRelacCli = mvarobjRelacCli
End Property

Public Property Let iPrazoTipo(ByVal vData As Integer)
    mvariPrazoTipo = vData
End Property

Public Property Get iPrazoTipo() As Integer
    iPrazoTipo = mvariPrazoTipo
End Property

Public Property Let dtDataEntrega(ByVal vDataEntrega As Date)
    mvardtDataEntrega = vDataEntrega
End Property

Public Property Get dtDataEntrega() As Date
    dtDataEntrega = mvardtDataEntrega
End Property

Public Property Let iPrazo(ByVal vData As Integer)
    mvariPrazo = vData
End Property

Public Property Get iPrazo() As Integer
    iPrazo = mvariPrazo
End Property

Public Property Let iFilialClienteBenef(ByVal vData As Integer)
    mvariFilialClienteBenef = vData
End Property

Public Property Get iFilialClienteBenef() As Integer
    iFilialClienteBenef = mvariFilialClienteBenef
End Property

Public Property Let lClienteBenef(ByVal vData As Long)
    mvarlClienteBenef = vData
End Property

Public Property Get lClienteBenef() As Long
    lClienteBenef = mvarlClienteBenef
End Property

Public Property Let iStatusItem(ByVal vData As Integer)
    mvariStatusItem = vData
End Property

Public Property Get iStatusItem() As Integer
    iStatusItem = mvariStatusItem
End Property

Public Property Set objTela(ByVal vData As Object)
    Set mvarobjTela = vData
End Property

Public Property Get objTela() As Object
    Set objTela = mvarobjTela
End Property

Public Property Set colItens(ByVal vData As Collection)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Let iAtendente(ByVal vData As Integer)
    mvariAtendente = vData
End Property

Public Property Get iAtendente() As Integer
    iAtendente = mvariAtendente
End Property

Public Property Let iVendedor(ByVal vData As Integer)
    mvariVendedor = vData
End Property

Public Property Get iVendedor() As Integer
    iVendedor = mvariVendedor
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iFilial(ByVal vData As Integer)
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let dtHora(ByVal vData As Date)
    mvardtHora = vData
End Property

Public Property Get dtHora() As Date
    dtHora = mvardtHora
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
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

