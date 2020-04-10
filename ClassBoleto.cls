VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBoleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvariFilialEmpresa As Integer
Private mvariECF As Integer
Private mvariOperador As Integer
Private mvardtDataTransacao As Date
Private mvarsAdmMeioPagto As String
Private mvariTipo As Integer
Private mvarsTipoTransacao As String
Private mvarsPOS As String
Private mvarsNSU As String
Private mvarsLote As String
Private mvarsAutorizacao As String
Private mvardValor As Double
Private mvariNumParcelas As Integer
Private mvariTipoParcelamento As Integer
Private mvarlBordero As Long
Private mvarlFechamento As Long
Private mvariStatus As Integer
Private mvarcolBoletoParcela As New Collection

Public Property Set colBoletoParcela(ByVal vData As Collection)
    Set mvarcolBoletoParcela = vData
End Property

Public Property Get colBoletoParcela() As Collection
    Set colBoletoParcela = mvarcolBoletoParcela
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let lFechamento(ByVal vData As Long)
    mvarlFechamento = vData
End Property

Public Property Get lFechamento() As Long
    lFechamento = mvarlFechamento
End Property

Public Property Let lBordero(ByVal vData As Long)
    mvarlBordero = vData
End Property

Public Property Get lBordero() As Long
    lBordero = mvarlBordero
End Property

Public Property Let iTipoParcelamento(ByVal vData As Integer)
    mvariTipoParcelamento = vData
End Property

Public Property Get iTipoParcelamento() As Integer
    iTipoParcelamento = mvariTipoParcelamento
End Property

Public Property Let iNumParcelas(ByVal vData As Integer)
    mvariNumParcelas = vData
End Property

Public Property Get iNumParcelas() As Integer
    iNumParcelas = mvariNumParcelas
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let sAutorizacao(ByVal vData As String)
    mvarsAutorizacao = vData
End Property

Public Property Get sAutorizacao() As String
    sAutorizacao = mvarsAutorizacao
End Property

Public Property Let sLote(ByVal vData As String)
    mvarsLote = vData
End Property

Public Property Get sLote() As String
    sLote = mvarsLote
End Property

Public Property Let sNSU(ByVal vData As String)
    mvarsNSU = vData
End Property

Public Property Get sNSU() As String
    sNSU = mvarsNSU
End Property

Public Property Let sPOS(ByVal vData As String)
    mvarsPOS = vData
End Property

Public Property Get sPOS() As String
    sPOS = mvarsPOS
End Property

Public Property Let sTipoTransacao(ByVal vData As String)
    mvarsTipoTransacao = vData
End Property

Public Property Get sTipoTransacao() As String
    sTipoTransacao = mvarsTipoTransacao
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let sAdmMeioPagto(ByVal vData As String)
    mvarsAdmMeioPagto = vData
End Property

Public Property Get sAdmMeioPagto() As String
    sAdmMeioPagto = mvarsAdmMeioPagto
End Property

Public Property Let dtDataTransacao(ByVal vData As Date)
    mvardtDataTransacao = vData
End Property

Public Property Get dtDataTransacao() As Date
    dtDataTransacao = mvardtDataTransacao
End Property

Public Property Let iOperador(ByVal vData As Integer)
    mvariOperador = vData
End Property

Public Property Get iOperador() As Integer
    iOperador = mvariOperador
End Property

Public Property Let iECF(ByVal vData As Integer)
    mvariECF = vData
End Property

Public Property Get iECF() As Integer
    iECF = mvariECF
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property
