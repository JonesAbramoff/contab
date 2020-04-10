VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProdutoSolicSRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsCodigo As String
Private mvariTipo As Integer
Private mvarsDescricao As String
Private mvarsNomeReduzido As String
Private mvarsModelo As String
Private mvariGerencial As Integer
Private mvariNivel As Integer
Private mvarsSubstituto1 As String
Private mvarsSubstituto2 As String
Private mvariPrazoValidade As Integer
Private mvarsCodigoBarras As String
Private mvariEtiquetasCodBarras As Integer
Private mvardPesoLiq As Double
Private mvardPesoBruto As Double
Private mvardComprimento As Double
Private mvardEspessura As Double
Private mvardLargura As Double
Private mvarsCor As String
Private mvarsObsFisica As String
Private mvariClasseUM As Integer
Private mvarsSiglaUMEstoque As String
Private mvarsSiglaUMCompra As String
Private mvarsSiglaUMVenda As String
Private mvariAtivo As Integer
Private mvariFaturamento As Integer
Private mvariCompras As Integer
Private mvariPCP As Integer
Private mvariKitBasico As Integer
Private mvariKitInt As Integer
Private mvarcolCategoriaItem As New Collection
Private mvarcolProdutoEmbalagem As New Collection
Private mvarcolProdutoTeste As New Collection
Private mvardIPIAliquota As Double
Private mvarsIPICodigo As String
Private mvarsIPICodDIPI As String
'Private mvardISSAliquota As Double
'Private mvarsISSCodigo As String
'Private mvariIRIncide As Integer
Private mvariControleEstoque As Integer
Private mvariICMSAgregaCusto As Integer
Private mvariIPIAgregaCusto As Integer
Private mvariFreteAgregaCusto As Integer
Private mvariApropriacaoCusto As Integer
Private mvarsContaContabil As String
Private mvardPercentMaisQuantCotAnt As Double
Private mvardPercentMenosQuantCotAnt As Double
Private mvariConsideraQuantCotAnt As Integer
Private mvariTemFaixaReceb As Integer
Private mvardPercentMaisReceb As Double
Private mvardPercentMenosReceb As Double
Private mvariRecebForaFaixa As Integer
Private mvariCreditoICMS As Integer
Private mvariCreditoIPI As Integer
Private mvardResiduo As Double
Private mvarsContaContabilProducao As String
Private mvariNatureza As Integer
Private mvardCustoReposicao As Double
Private mvariOrigemMercadoria As Integer
Private mvariTabelaPreco As Integer
Private mvariTempoProducao As Integer
Private mvariRastro As Integer
Private mvarlHorasMaquina As Long
Private mvardPesoEspecifico As Double
Private mvarsReferencia As String
Private mvarsFigura As String
Private mvarsSituacaoTribECF As String
Private mvarsICMSAliquota As String
Private mvariEmbalagem As Integer
Private mvardINSSPercBase As Double
Private mvariKitVendaComp As Integer
Private mvardPrecoLoja As Double
Private mvarcolCodBarras As New Collection
Private mvarsGrade As String
Public objInfoUsu As Object
Private mvardtDataLog As Date
Private mvardLimiteDesconto As Double
Private mvardDesconto As Double
Private mvardDescontoValor As Double
Private mvariUsaBalanca As Integer
Private mvarsSerieProx As String
Private mvariSerieParteNum As Integer
Private mvardtDataVenda As Date
Private mvarsUM As String
Private mvardQuantidade As Double
Private mvarsLote As String
Private mvariFilialOP As Integer
Private mvarsSolicitacao As String
Private mvariFilialEmpresa As Integer
Private mvarlCodSolic As Long

Public Property Let lCodSolic(ByVal vData As Long)
    mvarlCodSolic = vData
End Property

Public Property Get lCodSolic() As Long
    lCodSolic = mvarlCodSolic
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sSolicitacao(ByVal vData As String)
    mvarsSolicitacao = vData
End Property

Public Property Get sSolicitacao() As String
    sSolicitacao = mvarsSolicitacao
End Property

Public Property Let iFilialOP(ByVal vData As Integer)
    mvariFilialOP = vData
End Property

Public Property Get iFilialOP() As Integer
    iFilialOP = mvariFilialOP
End Property

Public Property Let sLote(ByVal vData As String)
    mvarsLote = vData
End Property

Public Property Get sLote() As String
    sLote = mvarsLote
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let dtDataVenda(ByVal vData As Date)
    mvardtDataVenda = vData
End Property

Public Property Get dtDataVenda() As Date
    dtDataVenda = mvardtDataVenda
End Property




Public Property Let iSerieParteNum(ByVal vData As Integer)
    mvariSerieParteNum = vData
End Property

Public Property Get iSerieParteNum() As Integer
    iSerieParteNum = mvariSerieParteNum
End Property

Public Property Let sSerieProx(ByVal vData As String)
    mvarsSerieProx = vData
End Property

Public Property Get sSerieProx() As String
    sSerieProx = mvarsSerieProx
End Property

Public Property Let iUsaBalanca(ByVal vData As Integer)
    mvariUsaBalanca = vData
End Property

Public Property Get iUsaBalanca() As Integer
    iUsaBalanca = mvariUsaBalanca
End Property

Public Property Set colProdutoTeste(ByVal vData As Object)
    Set mvarcolProdutoTeste = vData
End Property

Public Property Get colProdutoTeste() As Collection
    Set colProdutoTeste = mvarcolProdutoTeste
End Property

Public Property Let dDescontoValor(ByVal vData As Double)
    mvardDescontoValor = vData
End Property

Public Property Get dDescontoValor() As Double
    dDescontoValor = mvardDescontoValor
End Property

Public Property Let dDesconto(ByVal vData As Double)
    mvardDesconto = vData
End Property

Public Property Get dDesconto() As Double
    dDesconto = mvardDesconto
End Property

Public Property Let dLimiteDesconto(ByVal vData As Double)
    mvardLimiteDesconto = vData
End Property

Public Property Get dLimiteDesconto() As Double
    dLimiteDesconto = mvardLimiteDesconto
End Property

Public Property Let dtDataLog(ByVal vData As Date)
    mvardtDataLog = vData
End Property

Public Property Get dtDataLog() As Date
    dtDataLog = mvardtDataLog
End Property

Public Property Let sGrade(ByVal vData As String)
    mvarsGrade = vData
End Property

Public Property Get sGrade() As String
    sGrade = mvarsGrade
End Property

Public Property Let dPrecoLoja(ByVal vData As Double)
    mvardPrecoLoja = vData
End Property

Public Property Get dPrecoLoja() As Double
    dPrecoLoja = mvardPrecoLoja
End Property

Public Property Set colCodBarras(ByVal vData As Collection)
    Set mvarcolCodBarras = vData
End Property
Public Property Get colCodBarras() As Collection
    Set colCodBarras = mvarcolCodBarras
End Property
Public Property Let iKitVendaComp(ByVal vData As Integer)
   mvariKitVendaComp = vData
End Property

Public Property Get iKitVendaComp() As Integer
    iKitVendaComp = mvariKitVendaComp
End Property

Public Property Let dINSSPercBase(ByVal vData As Double)
    mvardINSSPercBase = vData
End Property

Public Property Get dINSSPercBase() As Double
    dINSSPercBase = mvardINSSPercBase
End Property

Public Property Let iEmbalagem(ByVal vData As Integer)
    mvariEmbalagem = vData
End Property

Public Property Get iEmbalagem() As Integer
    iEmbalagem = mvariEmbalagem
End Property

Public Property Let sICMSAliquota(ByVal vData As String)
    mvarsICMSAliquota = vData
End Property

Public Property Get sICMSAliquota() As String
    sICMSAliquota = mvarsICMSAliquota
End Property

Public Property Let sSituacaoTribECF(ByVal vData As String)
    mvarsSituacaoTribECF = vData
End Property

Public Property Get sSituacaoTribECF() As String
    sSituacaoTribECF = mvarsSituacaoTribECF
End Property

Public Property Let sFigura(ByVal vData As String)
    mvarsFigura = vData
End Property

Public Property Get sFigura() As String
    sFigura = mvarsFigura
End Property

Public Property Let sReferencia(ByVal vData As String)
    mvarsReferencia = vData
End Property

Public Property Get sReferencia() As String
    sReferencia = mvarsReferencia
End Property

Public Property Let dPesoEspecifico(ByVal vData As Double)
    mvardPesoEspecifico = vData
End Property

Public Property Get dPesoEspecifico() As Double
    dPesoEspecifico = mvardPesoEspecifico
End Property

Public Property Let lHorasMaquina(ByVal vData As Long)
    mvarlHorasMaquina = vData
End Property

Public Property Get lHorasMaquina() As Long
    lHorasMaquina = mvarlHorasMaquina
End Property

Public Property Let iRastro(ByVal vData As Integer)
    mvariRastro = vData
End Property

Public Property Get iRastro() As Integer
    iRastro = mvariRastro
End Property

Public Property Let iTempoProducao(ByVal vData As Integer)
    mvariTempoProducao = vData
End Property

Public Property Get iTempoProducao() As Integer
    iTempoProducao = mvariTempoProducao
End Property

Public Property Let iTabelaPreco(ByVal vData As Integer)
    mvariTabelaPreco = vData
End Property

Public Property Get iTabelaPreco() As Integer
    iTabelaPreco = mvariTabelaPreco
End Property

Public Property Let iOrigemMercadoria(ByVal vData As Integer)
    mvariOrigemMercadoria = vData
End Property

Public Property Get iOrigemMercadoria() As Integer
    iOrigemMercadoria = mvariOrigemMercadoria
End Property

Public Property Let dCustoReposicao(ByVal vData As Double)
    mvardCustoReposicao = vData
End Property

Public Property Get dCustoReposicao() As Double
    dCustoReposicao = mvardCustoReposicao
End Property

Public Property Let iNatureza(ByVal vData As Integer)
    mvariNatureza = vData
End Property

Public Property Get iNatureza() As Integer
    iNatureza = mvariNatureza
End Property

Public Property Let sContaContabilProducao(ByVal vData As String)
    mvarsContaContabilProducao = vData
End Property

Public Property Get sContaContabilProducao() As String
    sContaContabilProducao = mvarsContaContabilProducao
End Property

Public Property Let dResiduo(ByVal vData As Double)
    mvardResiduo = vData
End Property

Public Property Get dResiduo() As Double
    dResiduo = mvardResiduo
End Property

Public Property Let iCreditoIPI(ByVal vData As Integer)
    mvariCreditoIPI = vData
End Property

Public Property Get iCreditoIPI() As Integer
    iCreditoIPI = mvariCreditoIPI
End Property

Public Property Let iCreditoICMS(ByVal vData As Integer)
    mvariCreditoICMS = vData
End Property

Public Property Get iCreditoICMS() As Integer
    iCreditoICMS = mvariCreditoICMS
End Property

Public Property Let iRecebForaFaixa(ByVal vData As Integer)
    mvariRecebForaFaixa = vData
End Property

Public Property Get iRecebForaFaixa() As Integer
    iRecebForaFaixa = mvariRecebForaFaixa
End Property

Public Property Let dPercentMenosReceb(ByVal vData As Double)
    mvardPercentMenosReceb = vData
End Property

Public Property Get dPercentMenosReceb() As Double
    dPercentMenosReceb = mvardPercentMenosReceb
End Property

Public Property Let dPercentMaisReceb(ByVal vData As Double)
    mvardPercentMaisReceb = vData
End Property

Public Property Get dPercentMaisReceb() As Double
    dPercentMaisReceb = mvardPercentMaisReceb
End Property

Public Property Let iTemFaixaReceb(ByVal vData As Integer)
    mvariTemFaixaReceb = vData
End Property

Public Property Get iTemFaixaReceb() As Integer
    iTemFaixaReceb = mvariTemFaixaReceb
End Property

Public Property Let iConsideraQuantCotAnt(ByVal vData As Integer)
    mvariConsideraQuantCotAnt = vData
End Property

Public Property Get iConsideraQuantCotAnt() As Integer
    iConsideraQuantCotAnt = mvariConsideraQuantCotAnt
End Property

Public Property Let dPercentMenosQuantCotAnt(ByVal vData As Double)
    mvardPercentMenosQuantCotAnt = vData
End Property

Public Property Get dPercentMenosQuantCotAnt() As Double
    dPercentMenosQuantCotAnt = mvardPercentMenosQuantCotAnt
End Property

Public Property Let dPercentMaisQuantCotAnt(ByVal vData As Double)
    mvardPercentMaisQuantCotAnt = vData
End Property

Public Property Get dPercentMaisQuantCotAnt() As Double
    dPercentMaisQuantCotAnt = mvardPercentMaisQuantCotAnt
End Property

Public Property Let sContaContabil(ByVal vData As String)
    mvarsContaContabil = vData
End Property

Public Property Get sContaContabil() As String
    sContaContabil = mvarsContaContabil
End Property

Public Property Let iApropriacaoCusto(ByVal vData As Integer)
    mvariApropriacaoCusto = vData
End Property

Public Property Get iApropriacaoCusto() As Integer
    iApropriacaoCusto = mvariApropriacaoCusto
End Property

Public Property Let iFreteAgregaCusto(ByVal vData As Integer)
    mvariFreteAgregaCusto = vData
End Property

Public Property Get iFreteAgregaCusto() As Integer
    iFreteAgregaCusto = mvariFreteAgregaCusto
End Property

Public Property Let iICMSAgregaCusto(ByVal vData As Integer)
    mvariICMSAgregaCusto = vData
End Property

Public Property Get iICMSAgregaCusto() As Integer
    iICMSAgregaCusto = mvariICMSAgregaCusto
End Property

Public Property Let iIPIAgregaCusto(ByVal vData As Integer)
    mvariIPIAgregaCusto = vData
End Property

Public Property Get iIPIAgregaCusto() As Integer
    iIPIAgregaCusto = mvariIPIAgregaCusto
End Property

Public Property Let iControleEstoque(ByVal vData As Integer)
    mvariControleEstoque = vData
End Property

Public Property Get iControleEstoque() As Integer
    iControleEstoque = mvariControleEstoque
End Property

'Public Property Let iIRIncide(ByVal vData As Integer)
'    mvariIRIncide = vData
'End Property
'
'Public Property Get iIRIncide() As Integer
'    iIRIncide = mvariIRIncide
'End Property
'
'Public Property Let sISSCodigo(ByVal vData As String)
'    mvarsISSCodigo = vData
'End Property
'
'Public Property Get sISSCodigo() As String
'    sISSCodigo = mvarsISSCodigo
'End Property
'
'Public Property Let dISSAliquota(ByVal vData As Double)
'    mvardISSAliquota = vData
'End Property
'
'Public Property Get dISSAliquota() As Double
'    dISSAliquota = mvardISSAliquota
'End Property

Public Property Let sIPICodDIPI(ByVal vData As String)
    mvarsIPICodDIPI = vData
End Property

Public Property Get sIPICodDIPI() As String
    sIPICodDIPI = mvarsIPICodDIPI
End Property

Public Property Let sIPICodigo(ByVal vData As String)
    mvarsIPICodigo = vData
End Property

Public Property Get sIPICodigo() As String
    sIPICodigo = mvarsIPICodigo
End Property

Public Property Let dIPIAliquota(ByVal vData As Double)
    mvardIPIAliquota = vData
End Property

Public Property Get dIPIAliquota() As Double
    dIPIAliquota = mvardIPIAliquota
End Property

Public Property Set colCategoriaItem(ByVal vData As Object)
    Set mvarcolCategoriaItem = vData
End Property

Public Property Get colCategoriaItem() As Collection
    Set colCategoriaItem = mvarcolCategoriaItem
End Property

Public Property Set colProdutoEmbalagem(ByVal vData As Object)
    Set mvarcolProdutoEmbalagem = vData
End Property

Public Property Get colProdutoEmbalagem() As Collection
    Set colProdutoEmbalagem = mvarcolProdutoEmbalagem
End Property

Public Property Let iKitInt(ByVal vData As Integer)
    mvariKitInt = vData
End Property

Public Property Get iKitInt() As Integer
    iKitInt = mvariKitInt
End Property

Public Property Let iKitBasico(ByVal vData As Integer)
    mvariKitBasico = vData
End Property

Public Property Get iKitBasico() As Integer
    iKitBasico = mvariKitBasico
End Property

Public Property Let iPCP(ByVal vData As Integer)
    mvariPCP = vData
End Property

Public Property Get iPCP() As Integer
    iPCP = mvariPCP
End Property

Public Property Let iCompras(ByVal vData As Integer)
    mvariCompras = vData
End Property

Public Property Get iCompras() As Integer
    iCompras = mvariCompras
End Property

Public Property Let iFaturamento(ByVal vData As Integer)
    mvariFaturamento = vData
End Property

Public Property Get iFaturamento() As Integer
    iFaturamento = mvariFaturamento
End Property

''Public Property Let dPontoPedido(ByVal vData As Double)
''    mvardPontoPedido = vData
''End Property
''
''Public Property Get dPontoPedido() As Double
''    dPontoPedido = mvardPontoPedido
''End Property
''
''Public Property Let dEstoqueMaximo(ByVal vData As Double)
''    mvardEstoqueMaximo = vData
''End Property
''
''Public Property Get dEstoqueMaximo() As Double
''    dEstoqueMaximo = mvardEstoqueMaximo
''End Property
''
''Public Property Let dEstoqueSeguranca(ByVal vData As Double)
''    mvardEstoqueSeguranca = vData
''End Property
''
''Public Property Get dEstoqueSeguranca() As Double
''    dEstoqueSeguranca = mvardEstoqueSeguranca
''End Property
''
''Public Property Let iControleEstoque(ByVal vData As Integer)
''    mvariControleEstoque = vData
''End Property
''
''Public Property Get iControleEstoque() As Integer
''    iControleEstoque = mvariControleEstoque
''End Property
''
Public Property Let sObsFisica(ByVal vData As String)
    mvarsObsFisica = vData
End Property

Public Property Get sObsFisica() As String
    sObsFisica = mvarsObsFisica
End Property

Public Property Let sCor(ByVal vData As String)
    mvarsCor = vData
End Property

Public Property Get sCor() As String
    sCor = mvarsCor
End Property

Public Property Let dLargura(ByVal vData As Double)
    mvardLargura = vData
End Property

Public Property Get dLargura() As Double
    dLargura = mvardLargura
End Property

Public Property Let dEspessura(ByVal vData As Double)
    mvardEspessura = vData
End Property

Public Property Get dEspessura() As Double
    dEspessura = mvardEspessura
End Property

Public Property Let dComprimento(ByVal vData As Double)
    mvardComprimento = vData
End Property

Public Property Get dComprimento() As Double
    dComprimento = mvardComprimento
End Property

Public Property Let dPesoBruto(ByVal vData As Double)
    mvardPesoBruto = vData
End Property

Public Property Get dPesoBruto() As Double
    dPesoBruto = mvardPesoBruto
End Property

Public Property Let dPesoLiq(ByVal vData As Double)
    mvardPesoLiq = vData
End Property

Public Property Get dPesoLiq() As Double
    dPesoLiq = mvardPesoLiq
End Property

Public Property Let iEtiquetasCodBarras(ByVal vData As Integer)
    mvariEtiquetasCodBarras = vData
End Property

Public Property Get iEtiquetasCodBarras() As Integer
    iEtiquetasCodBarras = mvariEtiquetasCodBarras
End Property

Public Property Let sCodigoBarras(ByVal vData As String)
    mvarsCodigoBarras = vData
End Property

Public Property Get sCodigoBarras() As String
    sCodigoBarras = mvarsCodigoBarras
End Property

Public Property Let iPrazoValidade(ByVal vData As Integer)
    mvariPrazoValidade = vData
End Property

Public Property Get iPrazoValidade() As Integer
    iPrazoValidade = mvariPrazoValidade
End Property

Public Property Let iNivel(ByVal vData As Integer)
    mvariNivel = vData
End Property

Public Property Get iNivel() As Integer
    iNivel = mvariNivel
End Property

Public Property Let sModelo(ByVal vData As String)
    mvarsModelo = vData
End Property

Public Property Get sModelo() As String
    sModelo = mvarsModelo
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let sSubstituto2(ByVal vData As String)
    mvarsSubstituto2 = vData
End Property

Public Property Get sSubstituto2() As String
    sSubstituto2 = mvarsSubstituto2
End Property

Public Property Let sSubstituto1(ByVal vData As String)
    mvarsSubstituto1 = vData
End Property

Public Property Get sSubstituto1() As String
    sSubstituto1 = mvarsSubstituto1
End Property

Public Property Let sSiglaUMVenda(ByVal vData As String)
    mvarsSiglaUMVenda = vData
End Property

Public Property Get sSiglaUMVenda() As String
    sSiglaUMVenda = mvarsSiglaUMVenda
End Property

Public Property Let sSiglaUMCompra(ByVal vData As String)
    mvarsSiglaUMCompra = vData
End Property

Public Property Get sSiglaUMCompra() As String
    sSiglaUMCompra = mvarsSiglaUMCompra
End Property

Public Property Let sSiglaUMEstoque(ByVal vData As String)
    mvarsSiglaUMEstoque = vData
End Property

Public Property Get sSiglaUMEstoque() As String
    sSiglaUMEstoque = mvarsSiglaUMEstoque
End Property

Public Property Let iClasseUM(ByVal vData As Integer)
    mvariClasseUM = vData
End Property


Public Property Get iClasseUM() As Integer
    iClasseUM = mvariClasseUM
End Property

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property

Public Property Let iGerencial(ByVal vData As Integer)
    mvariGerencial = vData
End Property

Public Property Get iGerencial() As Integer
    iGerencial = mvariGerencial
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sCodigo(ByVal vData As String)
    mvarsCodigo = vData
End Property

Public Property Get sCodigo() As String
    sCodigo = mvarsCodigo
End Property

