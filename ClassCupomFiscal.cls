VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCupomFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarlNumIntDoc As Long
Private mvariStatus As Integer
Private mvariECF As Integer
Private mvardtDataEmissao As Date
Private mvardHoraEmissao As Double
Private mvarlNumero As Long
Private mvarsCPFCGC As String
Private mvarlNumOrcamento As Long
Private mvardValorTotal As Double
Private mvardValorProdutos As Double
Private mvardValorDesconto As Double
Private mvardValorDesconto1 As Double
Private mvardValorAcrescimo As Double
Private mvariTabelaPreco As Integer
Private mvarsNaturezaOp As String
Private mvarlGerenteCancel As Long
Private mvarsMotivoCancel As String
Private mvarcolItens As New Collection
Private mvariVendedor As Integer
Private mvardValorTroco As Double
Private mvarlDuracao As Long
Private mvariTipo As Integer
Private mvariCodCaixa As Integer
Private mvarlCliente As Long
Private mvarsCPFCGC1 As String
Private mvarsNomeCliente As String
Private mvarlCOOCupomOrigDAV As Long 'COO do Cupom Fiscal que se originou no DAV
Private mvardtDataOrcamento As Date
Private mvarsTipoECF As String
Private mvarsMarcaECF As String
Private mvarsModeloECF As String
Private mvarsNumSerieECF As String
Private mvarlCCF As Long 'contador de cupom fiscal
Private mvarsMFAdicional As String
Private mvarlNumeroDAV As Long
Private mvarsEndereco As String
Private mvariSequencialECF As Integer
Private mvariSequencialECFOrigDAV 'Sequencial do ECF que emitiu o cupom fiscal que se origou no DAV (para uso do DAV)
Private mvariItem As Integer
Private mvariDAVImpresso As Integer
Private mvardtDataReducao As Date

'SAT
Public sSATtimeStamp As String 'Data e hora da emissÃ£o no formato: AAAAMMDDHHMMSS
Public sSATChaveAcesso As String
Public sSATQRCode As String
Public sSATArqXml As String 'nome do arquivo xml
Public sSATmensagemSEFAZ As String
Public sSATmensagem As String
Public sSATnumeroSessao As String
Public sSATEEEEE As String
Public sSATCCCC As String
Public sSATcod As String
Public sSATvalorTotalCFe As String
Public sSATCPFCNPJValue As String
Public iSATtpAmb As Integer

Public sSATCanctimeStamp As String 'Data e hora da emissÃ£o no formato: AAAAMMDDHHMMSS
Public sSATCancChaveAcesso As String
Public sSATCancQRCode As String
Public sSATCancArqXml As String 'nome do arquivo xml de cancelamento
Public sSATCancmensagemSEFAZ As String
Public sSATCancmensagem As String
Public sSATCancnumeroSessao As String
Public sSATCancEEEEE As String
Public sSATCancCCCC As String
Public sSATCanccod As String
Public sSATCancvalorTotalCFe As String
Public sSATCancCPFCNPJValue As String
Public dtSATCancDataEmissao As Date '??? em desuso
Public dSATCancHoraEmissao As Double '??? em desuso
Public dtSATCancData As Date
Public dSATCancHora As Double
Public lSATCancNumero As Long
Public lSATCancCCF As Long 'contador de cupom fiscal
Public sSATCancNumSerieECF As String

'local de entrega
Public sEndEntLogradouro As String
Public sEndEntNúmero As String
Public sEndEntComplemento As String
Public sEndEntBairro As String
Public sEndEntCidade As String
Public sEndEntUF As String
Public lEndEntIBGECidade As Long
Public sEndEntEmail As String

'do cliente
Public objEndDest As New ClassEndereco
Public iIEIsento As Integer
Public sInscricaoEstadual As String

Public sNFeChaveAcesso As String
Public sNFeArqXml As String
Public sNFeArqXmlPre As String

Public sNFecStat As String
Public sNFenProt As String
Public dtNFeData As Date
Public dNFeHora As Double
Public sNFEversao As String
Public iNFetpAmb As Integer
Public sNFCeQRCode As String
Public sNFCeMensagem As String
Public iNFeCancHomologado As Integer

Public bEditavel As Boolean 'se um envio foi começado mas nao houve autorizacao nem rejeicao entao a nfce nao pode ser editada

'cancelamento de nfce
Public sNFeCancnProt As String
Public dtNFeCancData As Date
Public dNFeCancHora As Double
Public iNFeCancPendente As Integer
Public sNFeCancArqXml As String

Public bNFCEImprimir As Boolean

Public IdExterno As String 'ex.: order id da NTK
Public NumCaixa As Integer 'do caixa que pegou o orçamento para si
Public StatusExterno As Integer 'confirmado, despachado, entregue,...

Public objTributacaoDoc As New ClassTributacaoDoc

Public objNF As New ClassNFiscal 'para nf d2 (manual) ser registrada junto com o cupom fiscal correspondente
Private mvardValorDescontoTEF As Double
Private mvardValorTrocoTEF As Double

Public Property Let dValorTrocoTEF(ByVal vData As Double)
    mvardValorTrocoTEF = vData
End Property

Public Property Get dValorTrocoTEF() As Double
    dValorTrocoTEF = mvardValorTrocoTEF
End Property

Public Property Let dValorDescontoTEF(ByVal vData As Double)
    mvardValorDescontoTEF = vData
End Property

Public Property Get dValorDescontoTEF() As Double
    dValorDescontoTEF = mvardValorDescontoTEF
End Property

Public Property Let dtDataReducao(ByVal vData As Date)
    mvardtDataReducao = vData
End Property

Public Property Get dtDataReducao() As Date
    dtDataReducao = mvardtDataReducao
End Property

Public Property Let iDAVImpresso(ByVal vData As Integer)
    mvariDAVImpresso = vData
End Property

Public Property Get iDAVImpresso() As Integer
    iDAVImpresso = mvariDAVImpresso
End Property

Public Property Let iItem(ByVal vData As Integer)
    mvariItem = vData
End Property

Public Property Get iItem() As Integer
    iItem = mvariItem
End Property

Public Property Let iSequencialECFOrigDAV(ByVal vData As Integer)
    mvariSequencialECFOrigDAV = vData
End Property

Public Property Get iSequencialECFOrigDAV() As Integer
    iSequencialECFOrigDAV = mvariSequencialECFOrigDAV
End Property

Public Property Let iSequencialECF(ByVal vData As Integer)
    mvariSequencialECF = vData
End Property

Public Property Get iSequencialECF() As Integer
    iSequencialECF = mvariSequencialECF
End Property

Public Property Let sEndereco(ByVal vData As String)
    mvarsEndereco = vData
End Property

Public Property Get sEndereco() As String
    If Len(Trim(sEndEntLogradouro)) <> 0 Then
        sEndereco = sEndEntLogradouro
        If Len(Trim(sEndEntNúmero)) <> 0 Then sEndereco = sEndereco & ", " & sEndEntNúmero
        If Len(Trim(sEndEntComplemento)) <> 0 Then sEndereco = sEndereco & ", " & sEndEntComplemento
        If Len(Trim(sEndEntBairro)) <> 0 Then sEndereco = sEndereco & " - " & sEndEntBairro
    Else
        sEndereco = mvarsEndereco
    End If
End Property

Public Property Let lNumeroDAV(ByVal vData As Long)
    mvarlNumeroDAV = vData
End Property

Public Property Get lNumeroDAV() As Long
    lNumeroDAV = mvarlNumeroDAV
End Property

Public Property Let sMFAdicional(ByVal vData As String)
    mvarsMFAdicional = vData
End Property

Public Property Get sMFAdicional() As String
    sMFAdicional = mvarsMFAdicional
End Property

Public Property Let lCCF(ByVal vData As Long)
    mvarlCCF = vData
End Property

Public Property Get lCCF() As Long
    lCCF = mvarlCCF
End Property

Public Property Let sNumSerieECF(ByVal vData As String)
    mvarsNumSerieECF = vData
End Property

Public Property Get sNumSerieECF() As String
    sNumSerieECF = mvarsNumSerieECF
End Property

Public Property Let sModeloECF(ByVal vData As String)
    mvarsModeloECF = vData
End Property

Public Property Get sModeloECF() As String
    sModeloECF = mvarsModeloECF
End Property

Public Property Let sMarcaECF(ByVal vData As String)
    mvarsMarcaECF = vData
End Property

Public Property Get sMarcaECF() As String
    sMarcaECF = mvarsMarcaECF
End Property

Public Property Let sTipoECF(ByVal vData As String)
    mvarsTipoECF = vData
End Property

Public Property Get sTipoECF() As String
    sTipoECF = mvarsTipoECF
End Property

Public Property Let dtDataOrcamento(ByVal vData As Date)
    mvardtDataOrcamento = vData
End Property

Public Property Get dtDataOrcamento() As Date
    dtDataOrcamento = mvardtDataOrcamento
End Property

Public Property Let lCOOCupomOrigDAV(ByVal vData As Long)
    mvarlCOOCupomOrigDAV = vData
End Property

Public Property Get lCOOCupomOrigDAV() As Long
    lCOOCupomOrigDAV = mvarlCOOCupomOrigDAV
End Property

Public Property Let sNomeCliente(ByVal vData As String)
    mvarsNomeCliente = vData
End Property

Public Property Get sNomeCliente() As String
    sNomeCliente = mvarsNomeCliente
End Property

Public Property Let sCPFCGC1(ByVal vData As String)
    mvarsCPFCGC1 = vData
End Property

Public Property Get sCPFCGC1() As String
    sCPFCGC1 = mvarsCPFCGC1
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let iCodCaixa(ByVal vData As Integer)
    mvariCodCaixa = vData
End Property

Public Property Get iCodCaixa() As Integer
    iCodCaixa = mvariCodCaixa
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let lDuracao(ByVal vData As Long)
    mvarlDuracao = vData
End Property

Public Property Get lDuracao() As Long
    lDuracao = mvarlDuracao
End Property

Public Property Let dValorTroco(ByVal vData As Double)
    mvardValorTroco = vData
End Property

Public Property Get dValorTroco() As Double
    dValorTroco = mvardValorTroco
End Property

Public Property Let iVendedor(ByVal vData As Integer)
    mvariVendedor = vData
End Property

Public Property Get iVendedor() As Integer
    iVendedor = mvariVendedor
End Property

Public Property Set colItens(ByVal vData As Collection)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Let sMotivoCancel(ByVal vData As String)
    mvarsMotivoCancel = vData
End Property

Public Property Get sMotivoCancel() As String
    sMotivoCancel = mvarsMotivoCancel
End Property

Public Property Let lGerenteCancel(ByVal vData As Long)
    mvarlGerenteCancel = vData
End Property

Public Property Get lGerenteCancel() As Long
    lGerenteCancel = mvarlGerenteCancel
End Property

Public Property Let sNaturezaOp(ByVal vData As String)
    mvarsNaturezaOp = vData
End Property

Public Property Get sNaturezaOp() As String
    sNaturezaOp = mvarsNaturezaOp
End Property

Public Property Let iTabelaPreco(ByVal vData As Integer)
    mvariTabelaPreco = vData
End Property

Public Property Get iTabelaPreco() As Integer
    iTabelaPreco = mvariTabelaPreco
End Property

Public Property Let dValorAcrescimo(ByVal vData As Double)
    mvardValorAcrescimo = vData
End Property

Public Property Get dValorAcrescimo() As Double
    dValorAcrescimo = mvardValorAcrescimo
End Property

Public Property Let dValorDesconto1(ByVal vData As Double)
    mvardValorDesconto1 = vData
End Property

Public Property Get dValorDesconto1() As Double
    dValorDesconto1 = mvardValorDesconto1
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property

Public Property Let dValorProdutos(ByVal vData As Double)
    mvardValorProdutos = vData
End Property

Public Property Get dValorProdutos() As Double
    dValorProdutos = mvardValorProdutos
End Property

Public Property Let dValorTotal(ByVal vData As Double)
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal() As Double
    dValorTotal = mvardValorTotal
End Property

Public Property Let lNumOrcamento(ByVal vData As Long)
    mvarlNumOrcamento = vData
End Property

Public Property Get lNumOrcamento() As Long
    lNumOrcamento = mvarlNumOrcamento
End Property

Public Property Let sCPFCGC(ByVal vData As String)
    mvarsCPFCGC = vData
End Property

Public Property Get sCPFCGC() As String
    sCPFCGC = mvarsCPFCGC
End Property

Public Property Let lNumero(ByVal vData As Long)
    mvarlNumero = vData
End Property

Public Property Get lNumero() As Long
    lNumero = mvarlNumero
End Property

Public Property Let dHoraEmissao(ByVal vData As Double)
    mvardHoraEmissao = vData
End Property

Public Property Get dHoraEmissao() As Double
    dHoraEmissao = mvardHoraEmissao
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let iECF(ByVal vData As Integer)
    mvariECF = vData
End Property

Public Property Get iECF() As Integer
    iECF = mvariECF
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property


Private Sub Class_Initialize()
    objNF.dtDataEmissao = DATA_NULA
    dtNFeCancData = DATA_NULA
    bNFCEImprimir = True
    bEditavel = True
End Sub
