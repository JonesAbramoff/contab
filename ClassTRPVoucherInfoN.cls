VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRPVoucherInfoN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarsTipoDoc As String
Private mvarsTipVou As String
Private mvarsSerie As String
Private mvarlNumVou As Long
Private mvariSeq As Integer
Private mvardtData As Date
Private mvariTipoDocDestino As Integer
Private mvarlNumIntDocDestino As Long
Private mvarlNumIntDocComiInt As Long
Private mvardValor As Double
Private mvarsHistorico As String
Private mvarlNumTitulo As Long
Private mvariStatus As Integer
Private mvariTipoLiberacao As Integer
Private mvariManual As Integer
Private mvariTipoCliForn As Integer
Private mvariEstorno As Integer
Private mvarlNumIntDocLiberacao As Long
Private mvarlCliForn As Long
Private mvarlNumIntDocOcr As Long
Private mvarlNumIntDocEstorno As Long

Public Property Let lNumIntDocEstorno(ByVal vData As Long)
    mvarlNumIntDocEstorno = vData
End Property

Public Property Get lNumIntDocEstorno() As Long
    lNumIntDocEstorno = mvarlNumIntDocEstorno
End Property

Public Property Let lNumIntDocOcr(ByVal vData As Long)
    mvarlNumIntDocOcr = vData
End Property

Public Property Get lNumIntDocOcr() As Long
    lNumIntDocOcr = mvarlNumIntDocOcr
End Property

Public Property Let lCliForn(ByVal vData As Long)
    mvarlCliForn = vData
End Property

Public Property Get lCliForn() As Long
    lCliForn = mvarlCliForn
End Property

Public Property Let lNumIntDocLiberacao(ByVal vData As Long)
    mvarlNumIntDocLiberacao = vData
End Property

Public Property Get lNumIntDocLiberacao() As Long
    lNumIntDocLiberacao = mvarlNumIntDocLiberacao
End Property

Public Property Let iEstorno(ByVal vData As Integer)
    mvariEstorno = vData
End Property

Public Property Get iEstorno() As Integer
    iEstorno = mvariEstorno
End Property

Public Property Let iTipoCliForn(ByVal vData As Integer)
    mvariTipoCliForn = vData
End Property

Public Property Get iTipoCliForn() As Integer
    iTipoCliForn = mvariTipoCliForn
End Property

Public Property Let iManual(ByVal vData As Integer)
    mvariManual = vData
End Property

Public Property Get iManual() As Integer
    iManual = mvariManual
End Property

Public Property Let iTipoLiberacao(ByVal vData As Integer)
    mvariTipoLiberacao = vData
End Property

Public Property Get iTipoLiberacao() As Integer
    iTipoLiberacao = mvariTipoLiberacao
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let lNumTitulo(ByVal vData As Long)
    mvarlNumTitulo = vData
End Property

Public Property Get lNumTitulo() As Long
    lNumTitulo = mvarlNumTitulo
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sTipoDoc(ByVal vData As String)
    mvarsTipoDoc = vData
End Property

Public Property Get sTipoDoc() As String
    sTipoDoc = mvarsTipoDoc
End Property

Public Property Let sTipVou(ByVal vData As String)
    mvarsTipVou = vData
End Property

Public Property Get sTipVou() As String
    sTipVou = mvarsTipVou
End Property

Public Property Let sSerie(ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let lNumVou(ByVal vData As Long)
    mvarlNumVou = vData
End Property

Public Property Get lNumVou() As Long
    lNumVou = mvarlNumVou
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let iTipoDocDestino(ByVal vData As Integer)
    mvariTipoDocDestino = vData
End Property

Public Property Get iTipoDocDestino() As Integer
    iTipoDocDestino = mvariTipoDocDestino
End Property

Public Property Let lNumIntDocDestino(ByVal vData As Long)
    mvarlNumIntDocDestino = vData
End Property

Public Property Get lNumIntDocDestino() As Long
    lNumIntDocDestino = mvarlNumIntDocDestino
End Property

Public Property Let lNumIntDocComiInt(ByVal vData As Long)
    mvarlNumIntDocComiInt = vData
End Property

Public Property Get lNumIntDocComiInt() As Long
    lNumIntDocComiInt = mvarlNumIntDocComiInt
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let sHistorico(ByVal vData As String)
    mvarsHistorico = vData
End Property

Public Property Get sHistorico() As String
    sHistorico = mvarsHistorico
End Property
