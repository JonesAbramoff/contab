VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCarne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDoc As Long
Private mvariStatus As Integer
Private mvariFilialEmpresa As Integer
Private mvardtDataReferencia As Date
Private mvarlCupomFiscal As Long
Private mvarcolParcelas As New Collection 'local copy
Private mvarlCliente As Long
Private mvarsAutorizacao As String
Private mvarsCodBarrasCarne As String
Private mvarlNumIntExt As Long
Private mvariParcelamento As Integer

Public Property Let iParcelamento(ByVal vData As Integer)
    mvariParcelamento = vData
End Property

Public Property Get iParcelamento() As Integer
    iParcelamento = mvariParcelamento
End Property

Public Property Let lNumIntExt(ByVal vData As Long)
    mvarlNumIntExt = vData
End Property

Public Property Get lNumIntExt() As Long
    lNumIntExt = mvarlNumIntExt
End Property

Public Property Let sCodBarrasCarne(ByVal vData As String)
    mvarsCodBarrasCarne = vData
End Property

Public Property Get sCodBarrasCarne() As String
    sCodBarrasCarne = mvarsCodBarrasCarne
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let sAutorizacao(ByVal vData As String)
    mvarsAutorizacao = vData
End Property

Public Property Get sAutorizacao() As String
    sAutorizacao = mvarsAutorizacao
End Property

Public Property Let ifilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get ifilialEmpresa() As Integer
    ifilialEmpresa = mvariFilialEmpresa
End Property

Public Property Set colParcelas(ByVal vData As Collection)
    Set mvarcolParcelas = vData
End Property

Public Property Get colParcelas() As Collection
    Set colParcelas = mvarcolParcelas
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lCupomFiscal(ByVal vData As Long)
    mvarlCupomFiscal = vData
End Property

Public Property Get lCupomFiscal() As Long
    lCupomFiscal = mvarlCupomFiscal
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let dtDataReferencia(ByVal vData As Date)
    mvardtDataReferencia = vData
End Property

Public Property Get dtDataReferencia() As Date
    dtDataReferencia = mvardtDataReferencia
End Property

