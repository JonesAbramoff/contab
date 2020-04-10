VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRecebimentoCarne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlCodCliente As Long
Private mvarsCodCarneDe As String
Private mvarsCodCarneAte As String
Private mvardtDataVenctoDe As Date
Private mvardtDataVenctoAte As Date
Private mvarcolParcelas As New Collection
Private mvariStatus As Integer
Private mvariFilialEmpresa As Integer
Private mvariOpcao As Integer

Public Property Let iStatus(ByVal vdata As Integer)
    mvariStatus = vdata
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let iOpcao(ByVal vdata As Integer)
    mvariOpcao = vdata
End Property

Public Property Get iOpcao() As Integer
    iOpcao = mvariOpcao
End Property

Public Property Let iFilialEmpresa(ByVal vdata As Integer)
    mvariFilialEmpresa = vdata
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Set colParcelas(ByVal vdata As Collection)
    Set mvarcolParcelas = vdata
End Property

Public Property Get colParcelas() As Collection
    Set colParcelas = mvarcolParcelas
End Property

Public Property Let lCodCliente(ByVal vdata As Long)
    mvarlCodCliente = vdata
End Property

Public Property Get lCodCliente() As Long
    lCodCliente = mvarlCodCliente
End Property

Public Property Let lNumIntDoc(ByVal vdata As Long)
    mvarlNumIntDoc = vdata
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sCodCarneAte(ByVal vdata As String)
    mvarsCodCarneAte = vdata
End Property

Public Property Get sCodCarneAte() As String
    sCodCarneAte = mvarsCodCarneAte
End Property

Public Property Let sCodCarneDe(ByVal vdata As String)
    mvarsCodCarneDe = vdata
End Property

Public Property Get sCodCarneDe() As String
    sCodCarneDe = mvarsCodCarneDe
End Property

Public Property Let dtDataVenctoAte(ByVal vdata As Date)
    mvardtDataVenctoAte = vdata
End Property

Public Property Get dtDataVenctoAte() As Date
    dtDataVenctoAte = mvardtDataVenctoAte
End Property

Public Property Let dtDataVenctoDe(ByVal vdata As Date)
    mvardtDataVenctoDe = vdata
End Property

Public Property Get dtDataVenctoDe() As Date
    dtDataVenctoDe = mvardtDataVenctoDe
End Property
