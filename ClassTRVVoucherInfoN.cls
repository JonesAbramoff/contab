VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVVoucherInfoN"
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
Private mvarlNumIntDocOCR As Long
Private mvarlNumIntDocEstorno As Long
Private mvariIndireta As Integer
Private mvardtDataRegistro As Date
Private mvardHoraRegistro As Double
Private mvarsUsuario As String

'##################################
'AUXILIAR
Private mvarlNumIntDocNF As Long
Private mvarlNumNF As Long
Private mvardValorNF As Double
Private mvardtDataNF As Date
Private mvarlFornNF As Long
Private mvarsFornNF As String
Private mvardValorNC As Double
Private mvardtDataNC As Date
Private mvarlFornNC As Long
Private mvarsFornNC As String

Public Property Let sFornNC(ByVal vData As String)
    mvarsFornNC = vData
End Property

Public Property Get sFornNC() As String
    sFornNC = mvarsFornNC
End Property

Public Property Let lFornNC(ByVal vData As Long)
    mvarlFornNC = vData
End Property

Public Property Get lFornNC() As Long
    lFornNC = mvarlFornNC
End Property

Public Property Let dtDataNC(ByVal vDataNC As Date)
    mvardtDataNC = vDataNC
End Property

Public Property Get dtDataNC() As Date
    dtDataNC = mvardtDataNC
End Property

Public Property Let dValorNC(ByVal vData As Double)
    mvardValorNC = vData
End Property

Public Property Get dValorNC() As Double
    dValorNC = mvardValorNC
End Property

Public Property Let sFornNF(ByVal vData As String)
    mvarsFornNF = vData
End Property

Public Property Get sFornNF() As String
    sFornNF = mvarsFornNF
End Property

Public Property Let lFornNF(ByVal vData As Long)
    mvarlFornNF = vData
End Property

Public Property Get lFornNF() As Long
    lFornNF = mvarlFornNF
End Property

Public Property Let dtDataNF(ByVal vDataNF As Date)
    mvardtDataNF = vDataNF
End Property

Public Property Get dtDataNF() As Date
    dtDataNF = mvardtDataNF
End Property

Public Property Let dValorNF(ByVal vData As Double)
    mvardValorNF = vData
End Property

Public Property Get dValorNF() As Double
    dValorNF = mvardValorNF
End Property

Public Property Let lNumNF(ByVal vData As Long)
    mvarlNumNF = vData
End Property

Public Property Get lNumNF() As Long
    lNumNF = mvarlNumNF
End Property

Public Property Let lNumIntDocNF(ByVal vData As Long)
    mvarlNumIntDocNF = vData
End Property

Public Property Get lNumIntDocNF() As Long
    lNumIntDocNF = mvarlNumIntDocNF
End Property
'AUXILIAR
'##################################

Public Property Let dtDataRegistro(ByVal vData As Date)
    mvardtDataRegistro = vData
End Property

Public Property Get dtDataRegistro() As Date
    dtDataRegistro = mvardtDataRegistro
End Property

Public Property Let dHoraRegistro(ByVal vData As Double)
    mvardHoraRegistro = vData
End Property

Public Property Get dHoraRegistro() As Double
    dHoraRegistro = mvardHoraRegistro
End Property

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

Public Property Let iIndireta(ByVal vData As Integer)
    mvariIndireta = vData
End Property

Public Property Get iIndireta() As Integer
    iIndireta = mvariIndireta
End Property

Public Property Let lNumIntDocEstorno(ByVal vData As Long)
    mvarlNumIntDocEstorno = vData
End Property

Public Property Get lNumIntDocEstorno() As Long
    lNumIntDocEstorno = mvarlNumIntDocEstorno
End Property

Public Property Let lNumIntDocOCR(ByVal vData As Long)
    mvarlNumIntDocOCR = vData
End Property

Public Property Get lNumIntDocOCR() As Long
    lNumIntDocOCR = mvarlNumIntDocOCR
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
