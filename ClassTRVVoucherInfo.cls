VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVVoucherInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsTipo As String
Private mvarsSerie As String
Private mvarlNumVou As Long
Private mvardtDataEmissao As Date
Private mvardtDataRecepcao As Date
Private mvarlCliente As Long
Private mvarlFornEmissor As Long
Private mvarsProduto As String
Private mvarsDestino As String
Private mvardtDataInicio As Date
Private mvardtDataTermino As Date
Private mvarsVigencia As String
Private mvarsIdioma As String
Private mvariPax As Integer
Private mvarsDestinoVou As String
Private mvariAntc As Integer
Private mvarsControle As String
Private mvarsConvenio As String
Private mvardtDataPag As Date
Private mvariCartao As Integer
Private mvariPago As Integer
Private mvarlNumFat As Long
Private mvarlCliPassageiro As Long
Private mvardtDataNasc As Date
Private mvarsSexo As String
Private mvarsTipoDoc As String
Private mvarsCartaoFid As String
Private mvarsMoeda As String
Private mvardTarifaUnitaria As Double
Private mvardCambio As Double
Private mvarsValor As String
Private mvardTarifaPerc As Double
Private mvardTarifaValorMoeda As Double
Private mvardTarifaValorReal As Double
Private mvardComissaoPerc As Double
Private mvardComissaoValorMoeda As Double
Private mvardComissaoValorReal As Double
Private mvardCartaoPerc As Double
Private mvardCartaoValorMoeda As Double
Private mvardCartaoValorReal As Double
Private mvardOverPerc As Double
Private mvardOverValorMoeda As Double
Private mvardOverValorReal As Double
Private mvardCMRPerc As Double
Private mvardCMRValorMoeda As Double
Private mvardCMRValorReal As Double
Private mvarsCia As String
Private mvarsValidade As String
Private mvarsNumeroCC As String
Private mvarsTitular As String
Private mvardValorCartao As Double
Private mvarlParcela As Long
'Private mvarlAprovacao As Long
Private mvarsAprovacao As String
Private mvarsPassageiroSobreNome As String
Private mvarsPassageiroNome As String
Private mvarsPassageiroCGC As String
Private mvarsPassageiroEndereco As String
Private mvarsPassageiroBairro As String
Private mvarsPassageiroCidade As String
Private mvarsPassageiroCEP As String
Private mvarsPassageiroUF As String
Private mvarsPassageiroEmail As String
Private mvarsPassageiroContato As String
Private mvarsPassageiroTelefone1 As String
Private mvarsPassageiroTelefone2 As String
Private mvarsGrupo As String
Private mvarsTitularCPF As String
'Auxiliar apenas
Private mvariFilialFat As Integer

Public Property Let iFilialFat(ByVal vData As Integer)
    mvariFilialFat = vData
End Property

Public Property Get iFilialFat() As Integer
    iFilialFat = mvariFilialFat
End Property
'Auxiliar apenas


Public Property Let sTitularCPF(ByVal vData As String)
    mvarsTitularCPF = vData
End Property

Public Property Get sTitularCPF() As String
    sTitularCPF = mvarsTitularCPF
End Property

Public Property Let sTipo(ByVal vData As String)
    mvarsTipo = vData
End Property

Public Property Get sTipo() As String
    sTipo = mvarsTipo
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

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let dtDataRecepcao(ByVal vData As Date)
    mvardtDataRecepcao = vData
End Property

Public Property Get dtDataRecepcao() As Date
    dtDataRecepcao = mvardtDataRecepcao
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let lFornEmissor(ByVal vData As Long)
    mvarlFornEmissor = vData
End Property

Public Property Get lFornEmissor() As Long
    lFornEmissor = mvarlFornEmissor
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sDestino(ByVal vData As String)
    mvarsDestino = vData
End Property

Public Property Get sDestino() As String
    sDestino = mvarsDestino
End Property

Public Property Let dtDataInicio(ByVal vData As Date)
    mvardtDataInicio = vData
End Property

Public Property Get dtDataInicio() As Date
    dtDataInicio = mvardtDataInicio
End Property

Public Property Let dtDataTermino(ByVal vData As Date)
    mvardtDataTermino = vData
End Property

Public Property Get dtDataTermino() As Date
    dtDataTermino = mvardtDataTermino
End Property

Public Property Let sVigencia(ByVal vData As String)
    mvarsVigencia = vData
End Property

Public Property Get sVigencia() As String
    sVigencia = mvarsVigencia
End Property

Public Property Let sIdioma(ByVal vData As String)
    mvarsIdioma = vData
End Property

Public Property Get sIdioma() As String
    sIdioma = mvarsIdioma
End Property

Public Property Let iPax(ByVal vData As Integer)
    mvariPax = vData
End Property

Public Property Get iPax() As Integer
    iPax = mvariPax
End Property

Public Property Let sDestinoVou(ByVal vData As String)
    mvarsDestinoVou = vData
End Property

Public Property Get sDestinoVou() As String
    sDestinoVou = mvarsDestinoVou
End Property

Public Property Let iAntc(ByVal vData As Integer)
    mvariAntc = vData
End Property

Public Property Get iAntc() As Integer
    iAntc = mvariAntc
End Property

Public Property Let sControle(ByVal vData As String)
    mvarsControle = vData
End Property

Public Property Get sControle() As String
    sControle = mvarsControle
End Property

Public Property Let sConvenio(ByVal vData As String)
    mvarsConvenio = vData
End Property

Public Property Get sConvenio() As String
    sConvenio = mvarsConvenio
End Property

Public Property Let dtDataPag(ByVal vData As Date)
    mvardtDataPag = vData
End Property

Public Property Get dtDataPag() As Date
    dtDataPag = mvardtDataPag
End Property

Public Property Let iCartao(ByVal vData As Integer)
    mvariCartao = vData
End Property

Public Property Get iCartao() As Integer
    iCartao = mvariCartao
End Property

Public Property Let iPago(ByVal vData As Integer)
    mvariPago = vData
End Property

Public Property Get iPago() As Integer
    iPago = mvariPago
End Property

Public Property Let lNumFat(ByVal vData As Long)
    mvarlNumFat = vData
End Property

Public Property Get lNumFat() As Long
    lNumFat = mvarlNumFat
End Property

Public Property Let lCliPassageiro(ByVal vData As Long)
    mvarlCliPassageiro = vData
End Property

Public Property Get lCliPassageiro() As Long
    lCliPassageiro = mvarlCliPassageiro
End Property

Public Property Let dtDataNasc(ByVal vData As Date)
    mvardtDataNasc = vData
End Property

Public Property Get dtDataNasc() As Date
    dtDataNasc = mvardtDataNasc
End Property

Public Property Let sSexo(ByVal vData As String)
    mvarsSexo = vData
End Property

Public Property Get sSexo() As String
    sSexo = mvarsSexo
End Property

Public Property Let sTipoDoc(ByVal vData As String)
    mvarsTipoDoc = vData
End Property

Public Property Get sTipoDoc() As String
    sTipoDoc = mvarsTipoDoc
End Property

Public Property Let sCartaoFid(ByVal vData As String)
    mvarsCartaoFid = vData
End Property

Public Property Get sCartaoFid() As String
    sCartaoFid = mvarsCartaoFid
End Property

Public Property Let sMoeda(ByVal vData As String)
    mvarsMoeda = vData
End Property

Public Property Get sMoeda() As String
    sMoeda = mvarsMoeda
End Property

Public Property Let dTarifaUnitaria(ByVal vData As Double)
    mvardTarifaUnitaria = vData
End Property

Public Property Get dTarifaUnitaria() As Double
    dTarifaUnitaria = mvardTarifaUnitaria
End Property

Public Property Let dCambio(ByVal vData As Double)
    mvardCambio = vData
End Property

Public Property Get dCambio() As Double
    dCambio = mvardCambio
End Property

Public Property Let sValor(ByVal vData As String)
    mvarsValor = vData
End Property

Public Property Get sValor() As String
    sValor = mvarsValor
End Property

Public Property Let dTarifaPerc(ByVal vData As Double)
    mvardTarifaPerc = vData
End Property

Public Property Get dTarifaPerc() As Double
    dTarifaPerc = mvardTarifaPerc
End Property

Public Property Let dTarifaValorMoeda(ByVal vData As Double)
    mvardTarifaValorMoeda = vData
End Property

Public Property Get dTarifaValorMoeda() As Double
    dTarifaValorMoeda = mvardTarifaValorMoeda
End Property

Public Property Let dTarifaValorReal(ByVal vData As Double)
    mvardTarifaValorReal = vData
End Property

Public Property Get dTarifaValorReal() As Double
    dTarifaValorReal = mvardTarifaValorReal
End Property

Public Property Let dComissaoPerc(ByVal vData As Double)
    mvardComissaoPerc = vData
End Property

Public Property Get dComissaoPerc() As Double
    dComissaoPerc = mvardComissaoPerc
End Property

Public Property Let dComissaoValorMoeda(ByVal vData As Double)
    mvardComissaoValorMoeda = vData
End Property

Public Property Get dComissaoValorMoeda() As Double
    dComissaoValorMoeda = mvardComissaoValorMoeda
End Property

Public Property Let dComissaoValorReal(ByVal vData As Double)
    mvardComissaoValorReal = vData
End Property

Public Property Get dComissaoValorReal() As Double
    dComissaoValorReal = mvardComissaoValorReal
End Property

Public Property Let dCartaoPerc(ByVal vData As Double)
    mvardCartaoPerc = vData
End Property

Public Property Get dCartaoPerc() As Double
    dCartaoPerc = mvardCartaoPerc
End Property

Public Property Let dCartaoValorMoeda(ByVal vData As Double)
    mvardCartaoValorMoeda = vData
End Property

Public Property Get dCartaoValorMoeda() As Double
    dCartaoValorMoeda = mvardCartaoValorMoeda
End Property

Public Property Let dCartaoValorReal(ByVal vData As Double)
    mvardCartaoValorReal = vData
End Property

Public Property Get dCartaoValorReal() As Double
    dCartaoValorReal = mvardCartaoValorReal
End Property

Public Property Let dOverPerc(ByVal vData As Double)
    mvardOverPerc = vData
End Property

Public Property Get dOverPerc() As Double
    dOverPerc = mvardOverPerc
End Property

Public Property Let dOverValorMoeda(ByVal vData As Double)
    mvardOverValorMoeda = vData
End Property

Public Property Get dOverValorMoeda() As Double
    dOverValorMoeda = mvardOverValorMoeda
End Property

Public Property Let dOverValorReal(ByVal vData As Double)
    mvardOverValorReal = vData
End Property

Public Property Get dOverValorReal() As Double
    dOverValorReal = mvardOverValorReal
End Property

Public Property Let dCMRPerc(ByVal vData As Double)
    mvardCMRPerc = vData
End Property

Public Property Get dCMRPerc() As Double
    dCMRPerc = mvardCMRPerc
End Property

Public Property Let dCMRValorMoeda(ByVal vData As Double)
    mvardCMRValorMoeda = vData
End Property

Public Property Get dCMRValorMoeda() As Double
    dCMRValorMoeda = mvardCMRValorMoeda
End Property

Public Property Let dCMRValorReal(ByVal vData As Double)
    mvardCMRValorReal = vData
End Property

Public Property Get dCMRValorReal() As Double
    dCMRValorReal = mvardCMRValorReal
End Property

Public Property Let sCia(ByVal vData As String)
    mvarsCia = vData
End Property

Public Property Get sCia() As String
    sCia = mvarsCia
End Property

Public Property Let sValidade(ByVal vData As String)
    mvarsValidade = vData
End Property

Public Property Get sValidade() As String
    sValidade = mvarsValidade
End Property

Public Property Let sNumeroCC(ByVal vData As String)
    mvarsNumeroCC = vData
End Property

Public Property Get sNumeroCC() As String
    sNumeroCC = mvarsNumeroCC
End Property

Public Property Let sTitular(ByVal vData As String)
    mvarsTitular = vData
End Property

Public Property Get sTitular() As String
    sTitular = mvarsTitular
End Property

Public Property Let dValorCartao(ByVal vData As Double)
    mvardValorCartao = vData
End Property

Public Property Get dValorCartao() As Double
    dValorCartao = mvardValorCartao
End Property

Public Property Let lParcela(ByVal vData As Long)
    mvarlParcela = vData
End Property

Public Property Get lParcela() As Long
    lParcela = mvarlParcela
End Property

'Public Property Let lAprovacao(ByVal vData As Long)
'    mvarlAprovacao = vData
'End Property
'
'Public Property Get lAprovacao() As Long
'    lAprovacao = mvarlAprovacao
'End Property

Public Property Let sAprovacao(ByVal vData As String)
    mvarsAprovacao = vData
End Property

Public Property Get sAprovacao() As String
    sAprovacao = mvarsAprovacao
End Property

Public Property Let sPassageiroSobreNome(ByVal vData As String)
    mvarsPassageiroSobreNome = vData
End Property

Public Property Get sPassageiroSobreNome() As String
    sPassageiroSobreNome = mvarsPassageiroSobreNome
End Property

Public Property Let sPassageiroNome(ByVal vData As String)
    mvarsPassageiroNome = vData
End Property

Public Property Get sPassageiroNome() As String
    sPassageiroNome = mvarsPassageiroNome
End Property

Public Property Let sPassageiroCGC(ByVal vData As String)
    mvarsPassageiroCGC = vData
End Property

Public Property Get sPassageiroCGC() As String
    sPassageiroCGC = mvarsPassageiroCGC
End Property

Public Property Let sPassageiroEndereco(ByVal vData As String)
    mvarsPassageiroEndereco = vData
End Property

Public Property Get sPassageiroEndereco() As String
    sPassageiroEndereco = mvarsPassageiroEndereco
End Property

Public Property Let sPassageiroBairro(ByVal vData As String)
    mvarsPassageiroBairro = vData
End Property

Public Property Get sPassageiroBairro() As String
    sPassageiroBairro = mvarsPassageiroBairro
End Property

Public Property Let sPassageiroCidade(ByVal vData As String)
    mvarsPassageiroCidade = vData
End Property

Public Property Get sPassageiroCidade() As String
    sPassageiroCidade = mvarsPassageiroCidade
End Property

Public Property Let sPassageiroCEP(ByVal vData As String)
    mvarsPassageiroCEP = vData
End Property

Public Property Get sPassageiroCEP() As String
    sPassageiroCEP = mvarsPassageiroCEP
End Property

Public Property Let sPassageiroUF(ByVal vData As String)
    mvarsPassageiroUF = vData
End Property

Public Property Get sPassageiroUF() As String
    sPassageiroUF = mvarsPassageiroUF
End Property

Public Property Let sPassageiroEmail(ByVal vData As String)
    mvarsPassageiroEmail = vData
End Property

Public Property Get sPassageiroEmail() As String
    sPassageiroEmail = mvarsPassageiroEmail
End Property

Public Property Let sPassageiroContato(ByVal vData As String)
    mvarsPassageiroContato = vData
End Property

Public Property Get sPassageiroContato() As String
    sPassageiroContato = mvarsPassageiroContato
End Property

Public Property Let sPassageiroTelefone1(ByVal vData As String)
    mvarsPassageiroTelefone1 = vData
End Property

Public Property Get sPassageiroTelefone1() As String
    sPassageiroTelefone1 = mvarsPassageiroTelefone1
End Property

Public Property Let sPassageiroTelefone2(ByVal vData As String)
    mvarsPassageiroTelefone2 = vData
End Property

Public Property Get sPassageiroTelefone2() As String
    sPassageiroTelefone2 = mvarsPassageiroTelefone2
End Property

Public Property Let sGrupo(ByVal vData As String)
    mvarsGrupo = vData
End Property

Public Property Get sGrupo() As String
    sGrupo = mvarsGrupo
End Property
