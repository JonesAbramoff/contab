VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassTRVOcrExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc AS Long
Private mvarsUsuario AS String
Private mvardtData AS Date
Private mvardHora AS Double
Private mvarlNumIntDocOCR AS Long
Private mvarlNumVou AS Long
Private mvarsTipoDoc AS String
Private mvarsSerie AS String
Private mvarlCodigo AS Long
Private mvarlCliente AS Long
Private mvardtDataEmissao AS Date
Private mvarsObservacao AS String
Private mvariStatus AS Integer
Private mvariOrigem AS Integer
Private mvarsHistorico AS String
Private mvariFormaPagto AS Integer
Private mvardValorTotal AS Double
Private mvariExcluido AS Integer
Private mvariExportado AS Integer
Private mvarsUsuarioExportacao AS String
Private mvardtDataExportacao AS Date
Private mvardHoraExportacao AS Double
Private mvarsNomeArq AS String

Public Property Let lNumIntDoc (ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc () AS Long
    lNumIntDoc= mvarlNumIntDoc
End Property

Public Property Let sUsuario (ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario () AS String
    sUsuario= mvarsUsuario
End Property

Public Property Let dtData (ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData () AS Date
    dtData= mvardtData
End Property

Public Property Let dHora (ByVal vData As Double)
    mvardHora = vData
End Property

Public Property Get dHora () AS Double
    dHora= mvardHora
End Property

Public Property Let lNumIntDocOCR (ByVal vData As Long)
    mvarlNumIntDocOCR = vData
End Property

Public Property Get lNumIntDocOCR () AS Long
    lNumIntDocOCR= mvarlNumIntDocOCR
End Property

Public Property Let lNumVou (ByVal vData As Long)
    mvarlNumVou = vData
End Property

Public Property Get lNumVou () AS Long
    lNumVou= mvarlNumVou
End Property

Public Property Let sTipoDoc (ByVal vData As String)
    mvarsTipoDoc = vData
End Property

Public Property Get sTipoDoc () AS String
    sTipoDoc= mvarsTipoDoc
End Property

Public Property Let sSerie (ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie () AS String
    sSerie= mvarsSerie
End Property

Public Property Let lCodigo (ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo () AS Long
    lCodigo= mvarlCodigo
End Property

Public Property Let lCliente (ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente () AS Long
    lCliente= mvarlCliente
End Property

Public Property Let dtDataEmissao (ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao () AS Date
    dtDataEmissao= mvardtDataEmissao
End Property

Public Property Let sObservacao (ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao () AS String
    sObservacao= mvarsObservacao
End Property

Public Property Let iStatus (ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus () AS Integer
    iStatus= mvariStatus
End Property

Public Property Let iOrigem (ByVal vData As Integer)
    mvariOrigem = vData
End Property

Public Property Get iOrigem () AS Integer
    iOrigem= mvariOrigem
End Property

Public Property Let sHistorico (ByVal vData As String)
    mvarsHistorico = vData
End Property

Public Property Get sHistorico () AS String
    sHistorico= mvarsHistorico
End Property

Public Property Let iFormaPagto (ByVal vData As Integer)
    mvariFormaPagto = vData
End Property

Public Property Get iFormaPagto () AS Integer
    iFormaPagto= mvariFormaPagto
End Property

Public Property Let dValorTotal (ByVal vData As Double)
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal () AS Double
    dValorTotal= mvardValorTotal
End Property

Public Property Let iExcluido (ByVal vData As Integer)
    mvariExcluido = vData
End Property

Public Property Get iExcluido () AS Integer
    iExcluido= mvariExcluido
End Property

Public Property Let iExportado (ByVal vData As Integer)
    mvariExportado = vData
End Property

Public Property Get iExportado () AS Integer
    iExportado= mvariExportado
End Property

Public Property Let sUsuarioExportacao (ByVal vData As String)
    mvarsUsuarioExportacao = vData
End Property

Public Property Get sUsuarioExportacao () AS String
    sUsuarioExportacao= mvarsUsuarioExportacao
End Property

Public Property Let dtDataExportacao (ByVal vData As Date)
    mvardtDataExportacao = vData
End Property

Public Property Get dtDataExportacao () AS Date
    dtDataExportacao= mvardtDataExportacao
End Property

Public Property Let dHoraExportacao (ByVal vData As Double)
    mvardHoraExportacao = vData
End Property

Public Property Get dHoraExportacao () AS Double
    dHoraExportacao= mvardHoraExportacao
End Property

Public Property Let sNomeArq (ByVal vData As String)
    mvarsNomeArq = vData
End Property

Public Property Get sNomeArq () AS String
    sNomeArq= mvarsNomeArq
End Property
