VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCobrancaPorEmailSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlClienteDe As Long
Private mvarlClienteAte As Long
Private mvarlTituloDe As Long
Private mvarlTituloAte As Long
Private mvariAtrasoDe As Integer
Private mvariAtrasoAte As Integer
Private mvarsTipo As String
Private mvardSaldoDe As Double
Private mvardSaldoAte As Double
Private mvarcolParcelas As New Collection
Private mvardtDataBaixaDe As Date
Private mvardtDataBaixaAte As Date
Private mvarlModeloForcado As Long
Private mvariFaturamento As Integer
Private mvardtDataEmissaoDe As Date
Private mvardtDataEmissaoAte As Date
Private mvarobjInfoUsu As Object
Private mvarsCategoria As String
Private mvarsCategoriaDe As String
Private mvarsCategoriaAte As String
Private mvariSoComEmailValido As Integer
Private mvarsCobrador As String
Private mvariCobrador As Integer
Private mvarlNumBordero As Long
Private mvariIgnoraJaEnviados As Integer
Private mvariTipoTela As Integer

Public Property Let iTipoTela(ByVal vData As Integer)
    mvariTipoTela = vData
End Property

Public Property Get iTipoTela() As Integer
    iTipoTela = mvariTipoTela
End Property

Public Property Let iIgnoraJaEnviados(ByVal vData As Integer)
    mvariIgnoraJaEnviados = vData
End Property

Public Property Get iIgnoraJaEnviados() As Integer
    iIgnoraJaEnviados = mvariIgnoraJaEnviados
End Property

Public Property Let lNumBordero(ByVal vData As Long)
    mvarlNumBordero = vData
End Property

Public Property Get lNumBordero() As Long
    lNumBordero = mvarlNumBordero
End Property

Public Property Let iCobrador(ByVal vData As Integer)
    mvariCobrador = vData
End Property

Public Property Get iCobrador() As Integer
    iCobrador = mvariCobrador
End Property

Public Property Let sCobrador(ByVal vData As String)
    mvarsCobrador = vData
End Property

Public Property Get sCobrador() As String
    sCobrador = mvarsCobrador
End Property

Public Property Let iSoComEmailValido(ByVal vData As Integer)
    mvariSoComEmailValido = vData
End Property

Public Property Get iSoComEmailValido() As Integer
    iSoComEmailValido = mvariSoComEmailValido
End Property

Public Property Let sCategoriaAte(ByVal vData As String)
    mvarsCategoriaAte = vData
End Property

Public Property Get sCategoriaAte() As String
    sCategoriaAte = mvarsCategoriaAte
End Property

Public Property Let sCategoriaDe(ByVal vData As String)
    mvarsCategoriaDe = vData
End Property

Public Property Get sCategoriaDe() As String
    sCategoriaDe = mvarsCategoriaDe
End Property

Public Property Let sCategoria(ByVal vData As String)
    mvarsCategoria = vData
End Property

Public Property Get sCategoria() As String
    sCategoria = mvarsCategoria
End Property

Public Property Set objInfoUsu(ByVal vData As Object)
    Set mvarobjInfoUsu = vData
End Property

Public Property Get objInfoUsu() As Object
    Set objInfoUsu = mvarobjInfoUsu
End Property

Public Property Let dtDataEmissaoAte(ByVal vData As Date)
    mvardtDataEmissaoAte = vData
End Property

Public Property Get dtDataEmissaoAte() As Date
    dtDataEmissaoAte = mvardtDataEmissaoAte
End Property

Public Property Let dtDataEmissaoDe(ByVal vData As Date)
    mvardtDataEmissaoDe = vData
End Property

Public Property Get dtDataEmissaoDe() As Date
    dtDataEmissaoDe = mvardtDataEmissaoDe
End Property

Public Property Let iFaturamento(ByVal vData As Integer)
    mvariFaturamento = vData
End Property

Public Property Get iFaturamento() As Integer
    iFaturamento = mvariFaturamento
End Property

Public Property Let lModeloForcado(ByVal vData As Long)
    mvarlModeloForcado = vData
End Property

Public Property Get lModeloForcado() As Long
    lModeloForcado = mvarlModeloForcado
End Property

Public Property Let dtDataBaixaAte(ByVal vData As Date)
    mvardtDataBaixaAte = vData
End Property

Public Property Get dtDataBaixaAte() As Date
    dtDataBaixaAte = mvardtDataBaixaAte
End Property

Public Property Let dtDataBaixaDe(ByVal vData As Date)
    mvardtDataBaixaDe = vData
End Property

Public Property Get dtDataBaixaDe() As Date
    dtDataBaixaDe = mvardtDataBaixaDe
End Property

Public Property Set colParcelas(ByVal vData As Collection)
    Set mvarcolParcelas = vData
End Property

Public Property Get colParcelas() As Collection
    Set colParcelas = mvarcolParcelas
End Property

Public Property Let sTipo(ByVal vData As String)
    mvarsTipo = vData
End Property

Public Property Get sTipo() As String
    sTipo = mvarsTipo
End Property

Public Property Let iAtrasoAte(ByVal vData As Integer)
    mvariAtrasoAte = vData
End Property

Public Property Get iAtrasoAte() As Integer
    iAtrasoAte = mvariAtrasoAte
End Property

Public Property Let iAtrasoDe(ByVal vData As Integer)
    mvariAtrasoDe = vData
End Property

Public Property Get iAtrasoDe() As Integer
    iAtrasoDe = mvariAtrasoDe
End Property

Public Property Let lClienteAte(ByVal vData As Long)
    mvarlClienteAte = vData
End Property

Public Property Get lClienteAte() As Long
    lClienteAte = mvarlClienteAte
End Property

Public Property Let lClienteDe(ByVal vData As Long)
    mvarlClienteDe = vData
End Property

Public Property Get lClienteDe() As Long
    lClienteDe = mvarlClienteDe
End Property

Public Property Let lTituloAte(ByVal vData As Long)
    mvarlTituloAte = vData
End Property

Public Property Get lTituloAte() As Long
    lTituloAte = mvarlTituloAte
End Property

Public Property Let lTituloDe(ByVal vData As Long)
    mvarlTituloDe = vData
End Property

Public Property Get lTituloDe() As Long
    lTituloDe = mvarlTituloDe
End Property

Public Property Let dSaldoAte(ByVal vData As Double)
    mvardSaldoAte = vData
End Property

Public Property Get dSaldoAte() As Double
    dSaldoAte = mvardSaldoAte
End Property

Public Property Let dSaldoDe(ByVal vData As Double)
    mvardSaldoDe = vData
End Property

Public Property Get dSaldoDe() As Double
    dSaldoDe = mvardSaldoDe
End Property
