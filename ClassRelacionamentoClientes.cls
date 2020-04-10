VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelacClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlCodigo As Long
Private mvariFilialEmpresa As Integer
Private mvariOrigem As Integer
Private mvardtData As Date
Private mvardtHora As Date
Private mvarlTipo As Long
Private mvarlCliente As Long
Private mvariFilialCliente As Integer
Private mvariContato As Integer
Private mvariAtendente As Integer
Private mvarlRelacionamentoAnt As Long
Private mvarsAssunto1 As String
Private mvarsAssunto2 As String
Private mvariStatus As Integer
Private mvardtDataProxCobr As Date
Private mvardtDataPrevReceb As Date
Private mvarlNumIntParcRec As Long
Private mvariTipoDoc As Integer
Private mvarlNumIntDocOrigem As Long
Private mvariStatusCG As Integer
Private mvardtDataFim As Date
Private mvardtHoraFim As Date
Private mvarlMotivo As Long
Private mvarlStatusTipoDoc As Long
Private mvarlSatisfacao As Long

Public Property Let lSatisfacao(ByVal vData As Long)
    mvarlSatisfacao = vData
End Property

Public Property Get lSatisfacao() As Long
    lSatisfacao = mvarlSatisfacao
End Property

Public Property Let lStatusTipoDoc(ByVal vData As Long)
    mvarlStatusTipoDoc = vData
End Property

Public Property Get lStatusTipoDoc() As Long
    lStatusTipoDoc = mvarlStatusTipoDoc
End Property

Public Property Let iTipoDoc(ByVal vData As Integer)
    mvariTipoDoc = vData
End Property

Public Property Get iTipoDoc() As Integer
    iTipoDoc = mvariTipoDoc
End Property

Public Property Let lNumIntDocOrigem(ByVal vData As Long)
    mvarlNumIntDocOrigem = vData
End Property

Public Property Get lNumIntDocOrigem() As Long
    lNumIntDocOrigem = mvarlNumIntDocOrigem
End Property

Public Property Let dtHoraFim(ByVal vData As Date)
    mvardtHoraFim = vData
End Property

Public Property Get dtHoraFim() As Date
    dtHoraFim = mvardtHoraFim
End Property

Public Property Let dtDataFim(ByVal vData As Date)
    mvardtDataFim = vData
End Property

Public Property Get dtDataFim() As Date
    dtDataFim = mvardtDataFim
End Property

Public Property Let lMotivo(ByVal vData As Long)
    mvarlMotivo = vData
End Property

Public Property Get lMotivo() As Long
    lMotivo = mvarlMotivo
End Property

Public Property Let iStatusCG(ByVal vData As Integer)
    mvariStatusCG = vData
End Property

Public Property Get iStatusCG() As Integer
    iStatusCG = mvariStatusCG
End Property

Public Property Let lNumIntParcRec(ByVal vData As Long)
    mvarlNumIntParcRec = vData
End Property

Public Property Get lNumIntParcRec() As Long
    lNumIntParcRec = mvarlNumIntParcRec
End Property

Public Property Let dtDataPrevReceb(ByVal vDataPrevReceb As Date)
    mvardtDataPrevReceb = vDataPrevReceb
End Property

Public Property Get dtDataPrevReceb() As Date
    dtDataPrevReceb = mvardtDataPrevReceb
End Property

Public Property Let dtDataProxCobr(ByVal vDataProxCobr As Date)
    mvardtDataProxCobr = vDataProxCobr
End Property

Public Property Get dtDataProxCobr() As Date
    dtDataProxCobr = mvardtDataProxCobr
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

Public Property Let iOrigem(ByVal vData As Integer)
    mvariOrigem = vData
End Property

Public Property Get iOrigem() As Integer
    iOrigem = mvariOrigem
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

Public Property Let iContato(ByVal vData As Integer)
    mvariContato = vData
End Property

Public Property Get iContato() As Integer
    iContato = mvariContato
End Property

Public Property Let iAtendente(ByVal vData As Integer)
    mvariAtendente = vData
End Property

Public Property Get iAtendente() As Integer
    iAtendente = mvariAtendente
End Property

Public Property Let lRelacionamentoAnt(ByVal vData As Long)
    mvarlRelacionamentoAnt = vData
End Property

Public Property Get lRelacionamentoAnt() As Long
    lRelacionamentoAnt = mvarlRelacionamentoAnt
End Property

Public Property Let sAssunto1(ByVal vData As String)
    mvarsAssunto1 = vData
End Property

Public Property Get sAssunto1() As String
    sAssunto1 = mvarsAssunto1
End Property

Public Property Let sAssunto2(ByVal vData As String)
    mvarsAssunto2 = vData
End Property

Public Property Get sAssunto2() As String
    sAssunto2 = mvarsAssunto2
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

