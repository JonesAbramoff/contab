VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassParcelaPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public objInfoUsu As Object

Private mvarlNumIntDoc As Long
Private mvarlNumIntTitulo As Long
Private mvariNumParcela As Integer
Private mvariStatus As Integer
Private mvardtDataVencimento As Date
Private mvardtDataVencimentoReal As Date
Private mvardSaldo As Double
Private mvardValor As Double
Private mvariPortador As Integer
Private mvariProxSeqBaixa As Integer
Private mvariTipoCobranca As Integer
Private mvariBancoCobrador As Integer
Private mvarsNossoNumero As String
Private mvarsCodigoDeBarras As String 'Guarda o Numero Referênte ao Código da Barras
Private mvardtDataUltimaBaixa As Date
Private mvardValorOriginal As Double
Private mvariMotivoDiferenca As Integer
Private mvarsCodUsuarioLib As String
Private mvardtDataLib As Date

Public Property Let dtDataLib(ByVal vData As Date)
    mvardtDataLib = vData
End Property

Public Property Get dtDataLib() As Date
    dtDataLib = mvardtDataLib
End Property

Public Property Let sCodUsuarioLib(ByVal vData As String)
    mvarsCodUsuarioLib = vData
End Property

Public Property Get sCodUsuarioLib() As String
    sCodUsuarioLib = mvarsCodUsuarioLib
End Property

Public Property Let iMotivoDiferenca(ByVal vData As Integer)
    mvariMotivoDiferenca = vData
End Property

Public Property Get iMotivoDiferenca() As Integer
    iMotivoDiferenca = mvariMotivoDiferenca
End Property

Public Property Let dValorOriginal(ByVal vData As Double)
    mvardValorOriginal = vData
End Property

Public Property Get dValorOriginal() As Double
    dValorOriginal = mvardValorOriginal
End Property

Public Property Let dtDataUltimaBaixa(ByVal vData As Date)
    mvardtDataUltimaBaixa = vData
End Property

Public Property Get dtDataUltimaBaixa() As Date
    dtDataUltimaBaixa = mvardtDataUltimaBaixa
End Property

Public Property Let sCodigoDeBarras(ByVal vData As String)
    mvarsCodigoDeBarras = vData
End Property

Public Property Get sCodigoDeBarras() As String
    sCodigoDeBarras = mvarsCodigoDeBarras
End Property

Public Property Let sNossoNumero(ByVal vData As String)
    mvarsNossoNumero = vData
End Property

Public Property Get sNossoNumero() As String
    sNossoNumero = mvarsNossoNumero
End Property

Public Property Let iBancoCobrador(ByVal vData As Integer)
    mvariBancoCobrador = vData
End Property

Public Property Get iBancoCobrador() As Integer
    iBancoCobrador = mvariBancoCobrador
End Property

Public Property Let iTipoCobranca(ByVal vData As Integer)
    mvariTipoCobranca = vData
End Property

Public Property Get iTipoCobranca() As Integer
    iTipoCobranca = mvariTipoCobranca
End Property

Public Property Let iProxSeqBaixa(ByVal vData As Integer)
    mvariProxSeqBaixa = vData
End Property

Public Property Get iProxSeqBaixa() As Integer
    iProxSeqBaixa = mvariProxSeqBaixa
End Property

Public Property Let iPortador(ByVal vData As Integer)
    mvariPortador = vData
End Property

Public Property Get iPortador() As Integer
    iPortador = mvariPortador
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dSaldo(ByVal vData As Double)
    mvardSaldo = vData
End Property

Public Property Get dSaldo() As Double
    dSaldo = mvardSaldo
End Property

Public Property Let dtDataVencimentoReal(ByVal vData As Date)
    mvardtDataVencimentoReal = vData
End Property

Public Property Get dtDataVencimentoReal() As Date
    dtDataVencimentoReal = mvardtDataVencimentoReal
End Property

Public Property Let dtDataVencimento(ByVal vData As Date)
    mvardtDataVencimento = vData
End Property

Public Property Get dtDataVencimento() As Date
    dtDataVencimento = mvardtDataVencimento
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let iNumParcela(ByVal vData As Integer)
    mvariNumParcela = vData
End Property

Public Property Get iNumParcela() As Integer
    iNumParcela = mvariNumParcela
End Property

Public Property Let lNumIntTitulo(ByVal vData As Long)
    mvarlNumIntTitulo = vData
End Property

Public Property Get lNumIntTitulo() As Long
    lNumIntTitulo = mvarlNumIntTitulo
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property



Private Sub Class_Initialize()
    mvardtDataLib = DATA_NULA
End Sub
