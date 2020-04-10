VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassParcelaReceber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'**** ATENCAO: QUANDO INCLUIR NOVO CAMPO ATUALIZAR A SUB COPIAR ******'

Private mvarlNumIntDoc As Long
Private mvarlNumIntTitulo As Long
Private mvariNumParcela As Integer
Private mvariStatus As Integer
Private mvardtDataVencimento As Date
Private mvardtDataVencimentoReal As Date
Private mvardSaldo As Double
Private mvardValor As Double
Private mvariCobrador As Integer
Private mvariCarteiraCobranca As Integer
Private mvarsNumTitCobrador As String
Private mvarlNumIntCheque As Long
Private mvariProxSeqBaixa As Integer
Private mvariProxSeqOcorr As Integer
Private mvariDesconto1Codigo As Integer
Private mvardtDesconto1Ate As Date
Private mvardDesconto1Valor As Double
Private mvariDesconto2Codigo As Integer
Private mvardtDesconto2Ate As Date
Private mvardDesconto2Valor As Double
Private mvariDesconto3Codigo As Integer
Private mvardtDesconto3Ate As Date
Private mvardDesconto3Valor As Double
Private mvariAceite As Integer
Private mvariDescontada As Integer
Private mvarcolComissoesBaixa As New Collection

Private mvardValorOriginal As Double

'##########################################################
'INSERIDO POR WAGNER
Private mvariPrevisao As Integer
Private mvarsObservacao As String
Private mvariTipoPagto As Integer
Private mvariCodConta As Integer
Private mvardtDataCredito As Date
Private mvardtDataEmissaoCheque As Date
Private mvariBancoCheque As Integer
Private mvarsAgenciaCheque As String
Private mvarsContaCorrenteCheque As String
Private mvarlNumeroCheque As Long
Private mvardtDataDepositoCheque As Date
Private mvariAdmMeioPagto As Integer
Private mvariParcelamento As Integer
Private mvarsNumeroCartao As String
Private mvardtValidadeCartao As Date
Private mvarsAprovacaoCartao As String
Private mvardtDataTransacaoCartao As Date

Private mvardtDataProxCobr As Date
Private mvardtDataPrevReceb As Date
Private mvarobjInfoUsu As Object
Private mvardtDataUltimaBaixa As Date
Private mvardValorAux As Double
Private mvarlIdImpressaoBoleto As Long

Public Property Let lIdImpressaoBoleto(ByVal vData As Long)
    mvarlIdImpressaoBoleto = vData
End Property

Public Property Get lIdImpressaoBoleto() As Long
    lIdImpressaoBoleto = mvarlIdImpressaoBoleto
End Property

Public Property Let dValorAux(ByVal vData As Double)
    mvardValorAux = vData
End Property

Public Property Get dValorAux() As Double
    dValorAux = mvardValorAux
End Property

Public Property Let dtDataUltimaBaixa(ByVal vData As Date)
    mvardtDataUltimaBaixa = vData
End Property

Public Property Get dtDataUltimaBaixa() As Date
    dtDataUltimaBaixa = mvardtDataUltimaBaixa
End Property

Public Property Set objInfoUsu(ByVal vData As Object)
    Set mvarobjInfoUsu = vData
End Property

Public Property Get objInfoUsu() As Object
    Set objInfoUsu = mvarobjInfoUsu
End Property

Public Property Let dtDataPrevReceb(ByVal vData As Date)
    mvardtDataPrevReceb = vData
End Property

Public Property Get dtDataPrevReceb() As Date
    dtDataPrevReceb = mvardtDataPrevReceb
End Property

Public Property Let dtDataProxCobr(ByVal vData As Date)
    mvardtDataProxCobr = vData
End Property

Public Property Get dtDataProxCobr() As Date
    dtDataProxCobr = mvardtDataProxCobr
End Property

Public Property Let dtDataTransacaoCartao(ByVal vData As Date)
    mvardtDataTransacaoCartao = vData
End Property

Public Property Get dtDataTransacaoCartao() As Date
    dtDataTransacaoCartao = mvardtDataTransacaoCartao
End Property


Public Property Let sAprovacaoCartao(ByVal vData As String)
    mvarsAprovacaoCartao = vData
End Property

Public Property Get sAprovacaoCartao() As String
    sAprovacaoCartao = mvarsAprovacaoCartao
End Property

Public Property Let dtValidadeCartao(ByVal vData As Date)
    mvardtValidadeCartao = vData
End Property

Public Property Get dtValidadeCartao() As Date
    dtValidadeCartao = mvardtValidadeCartao
End Property

Public Property Let sNumeroCartao(ByVal vData As String)
    mvarsNumeroCartao = vData
End Property

Public Property Get sNumeroCartao() As String
    sNumeroCartao = mvarsNumeroCartao
End Property

Public Property Let iParcelamento(ByVal vData As Integer)
    mvariParcelamento = vData
End Property

Public Property Get iParcelamento() As Integer
    iParcelamento = mvariParcelamento
End Property

Public Property Let iAdmMeioPagto(ByVal vData As Integer)
    mvariAdmMeioPagto = vData
End Property

Public Property Get iAdmMeioPagto() As Integer
    iAdmMeioPagto = mvariAdmMeioPagto
End Property

Public Property Let dtDataDepositoCheque(ByVal vData As Date)
    mvardtDataDepositoCheque = vData
End Property

Public Property Get dtDataDepositoCheque() As Date
    dtDataDepositoCheque = mvardtDataDepositoCheque
End Property

Public Property Let lNumeroCheque(ByVal vData As Long)
    mvarlNumeroCheque = vData
End Property

Public Property Get lNumeroCheque() As Long
    lNumeroCheque = mvarlNumeroCheque
End Property

Public Property Let sContaCorrenteCheque(ByVal vData As String)
    mvarsContaCorrenteCheque = vData
End Property

Public Property Get sContaCorrenteCheque() As String
    sContaCorrenteCheque = mvarsContaCorrenteCheque
End Property

Public Property Let sAgenciaCheque(ByVal vData As String)
    mvarsAgenciaCheque = vData
End Property

Public Property Get sAgenciaCheque() As String
    sAgenciaCheque = mvarsAgenciaCheque
End Property

Public Property Let iBancoCheque(ByVal vData As Integer)
    mvariBancoCheque = vData
End Property

Public Property Get iBancoCheque() As Integer
    iBancoCheque = mvariBancoCheque
End Property

Public Property Let dtDataEmissaoCheque(ByVal vData As Date)
    mvardtDataEmissaoCheque = vData
End Property

Public Property Get dtDataEmissaoCheque() As Date
    dtDataEmissaoCheque = mvardtDataEmissaoCheque
End Property

Public Property Let dtDataCredito(ByVal vData As Date)
    mvardtDataCredito = vData
End Property

Public Property Get dtDataCredito() As Date
    dtDataCredito = mvardtDataCredito
End Property

Public Property Let iCodConta(ByVal vData As Integer)
    mvariCodConta = vData
End Property

Public Property Get iCodConta() As Integer
    iCodConta = mvariCodConta
End Property

Public Property Let iTipoPagto(ByVal vData As Integer)
    mvariTipoPagto = vData
End Property

Public Property Get iTipoPagto() As Integer
    iTipoPagto = mvariTipoPagto
End Property

Public Property Let iPrevisao(ByVal vData As Integer)
    mvariPrevisao = vData
End Property

Public Property Get iPrevisao() As Integer
    iPrevisao = mvariPrevisao
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property
'##########################################################

Public Property Let dValorOriginal(ByVal vData As Double)
    mvardValorOriginal = vData
End Property

Public Property Get dValorOriginal() As Double
    dValorOriginal = mvardValorOriginal
End Property

Public Property Set colComissoesBaixa(ByVal vData As Object)
    Set mvarcolComissoesBaixa = vData
End Property

Public Property Get colComissoesBaixa() As Collection
    Set colComissoesBaixa = mvarcolComissoesBaixa
End Property

Public Property Let iAceite(ByVal vData As Integer)
    mvariAceite = vData
End Property

Public Property Get iAceite() As Integer
    iAceite = mvariAceite
End Property

Public Property Let iDescontada(ByVal vData As Integer)
    mvariDescontada = vData
End Property

Public Property Get iDescontada() As Integer
    iDescontada = mvariDescontada
End Property

Public Property Let lNumIntCheque(ByVal vData As Long)
    mvarlNumIntCheque = vData
End Property

Public Property Get lNumIntCheque() As Long
    lNumIntCheque = mvarlNumIntCheque
End Property

Public Property Let dDesconto3Valor(ByVal vData As Double)
    mvardDesconto3Valor = vData
End Property

Public Property Get dDesconto3Valor() As Double
    dDesconto3Valor = mvardDesconto3Valor
End Property

Public Property Let dtDesconto3Ate(ByVal vData As Date)
    mvardtDesconto3Ate = vData
End Property

Public Property Get dtDesconto3Ate() As Date
    dtDesconto3Ate = mvardtDesconto3Ate
End Property

Public Property Let iDesconto3Codigo(ByVal vData As Integer)
    mvariDesconto3Codigo = vData
End Property

Public Property Get iDesconto3Codigo() As Integer
    iDesconto3Codigo = mvariDesconto3Codigo
End Property

Public Property Let dDesconto2Valor(ByVal vData As Double)
    mvardDesconto2Valor = vData
End Property

Public Property Get dDesconto2Valor() As Double
    dDesconto2Valor = mvardDesconto2Valor
End Property

Public Property Let dtDesconto2Ate(ByVal vData As Date)
    mvardtDesconto2Ate = vData
End Property

Public Property Get dtDesconto2Ate() As Date
    dtDesconto2Ate = mvardtDesconto2Ate
End Property

Public Property Let iDesconto2Codigo(ByVal vData As Integer)
    mvariDesconto2Codigo = vData
End Property

Public Property Get iDesconto2Codigo() As Integer
    iDesconto2Codigo = mvariDesconto2Codigo
End Property

Public Property Let dDesconto1Valor(ByVal vData As Double)
    mvardDesconto1Valor = vData
End Property

Public Property Get dDesconto1Valor() As Double
    dDesconto1Valor = mvardDesconto1Valor
End Property

Public Property Let dtDesconto1Ate(ByVal vData As Date)
    mvardtDesconto1Ate = vData
End Property

Public Property Get dtDesconto1Ate() As Date
    dtDesconto1Ate = mvardtDesconto1Ate
End Property

Public Property Let iDesconto1Codigo(ByVal vData As Integer)
    mvariDesconto1Codigo = vData
End Property

Public Property Get iDesconto1Codigo() As Integer
    iDesconto1Codigo = mvariDesconto1Codigo
End Property

Public Property Let sNumTitCobrador(ByVal vData As String)
    mvarsNumTitCobrador = vData
End Property

Public Property Get sNumTitCobrador() As String
    sNumTitCobrador = mvarsNumTitCobrador
End Property

Public Property Let iCobrador(ByVal vData As Integer)
    mvariCobrador = vData
End Property

Public Property Get iCobrador() As Integer
    iCobrador = mvariCobrador
End Property

Public Property Let iCarteiraCobranca(ByVal vData As Integer)
    mvariCarteiraCobranca = vData
End Property

Public Property Get iCarteiraCobranca() As Integer
    iCarteiraCobranca = mvariCarteiraCobranca
End Property

Public Property Let iProxSeqBaixa(ByVal vData As Integer)
    mvariProxSeqBaixa = vData
End Property

Public Property Get iProxSeqBaixa() As Integer
    iProxSeqBaixa = mvariProxSeqBaixa
End Property

Public Property Let iProxSeqOcorr(ByVal vData As Integer)
    mvariProxSeqOcorr = vData
End Property

Public Property Get iProxSeqOcorr() As Integer
    iProxSeqOcorr = mvariProxSeqOcorr
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

Public Sub Copiar(objParcRec As ClassParcelaReceber)
'copia os dados de objParcRec p/"Me"

Dim objComissao As ClassComissao, objComissaoAux As ClassComissao

    With objParcRec
        
        mvarlNumIntDoc = .lNumIntDoc
        mvarlNumIntTitulo = .lNumIntTitulo
        mvariNumParcela = .iNumParcela
        mvariStatus = .iStatus
        mvardtDataVencimento = .dtDataVencimento
        mvardtDataVencimentoReal = .dtDataVencimentoReal
        mvardSaldo = .dSaldo
        mvardValor = .dValor
        mvariCobrador = .iCobrador
        mvariCarteiraCobranca = .iCarteiraCobranca
        mvarsNumTitCobrador = .sNumTitCobrador
        mvarlNumIntCheque = .lNumIntCheque
        mvariProxSeqBaixa = .iProxSeqBaixa
        mvariProxSeqOcorr = .iProxSeqOcorr
        mvariDesconto1Codigo = .iDesconto1Codigo
        mvardtDesconto1Ate = .dtDesconto1Ate
        mvardDesconto1Valor = .dDesconto1Valor
        mvariDesconto2Codigo = .iDesconto2Codigo
        mvardtDesconto2Ate = .dtDesconto2Ate
        mvardDesconto2Valor = .dDesconto2Valor
        mvariDesconto3Codigo = .iDesconto3Codigo
        mvardtDesconto3Ate = .dtDesconto3Ate
        mvardDesconto3Valor = .dDesconto3Valor
        mvariAceite = .iAceite
        mvariDescontada = .iDescontada
        mvardValorOriginal = .dValorOriginal
        mvariPrevisao = .iPrevisao
        mvarsObservacao = .sObservacao
    
        mvariTipoPagto = .iTipoPagto
        mvariCodConta = .iCodConta
        mvardtDataCredito = .dtDataCredito
        mvardtDataEmissaoCheque = .dtDataEmissaoCheque
        mvariBancoCheque = .iBancoCheque
        mvarsAgenciaCheque = .sAgenciaCheque
        mvarsContaCorrenteCheque = .sContaCorrenteCheque
        mvarlNumeroCheque = .lNumeroCheque
        mvardtDataDepositoCheque = .dtDataDepositoCheque
        mvariAdmMeioPagto = .iAdmMeioPagto
        mvariParcelamento = .iParcelamento
        mvarsNumeroCartao = .sNumeroCartao
        mvardtValidadeCartao = .dtValidadeCartao
        mvarsAprovacaoCartao = .sAprovacaoCartao
        mvardtDataTransacaoCartao = .dtDataTransacaoCartao
        
        mvardtDataProxCobr = .dtDataProxCobr
        mvardtDataPrevReceb = .dtDataPrevReceb
        
        mvardtDataUltimaBaixa = .dtDataUltimaBaixa
    
    End With
    
    Set mvarcolComissoesBaixa = New Collection
    
    For Each objComissao In objParcRec.colComissoesBaixa
    
        Set objComissaoAux = New ClassComissao
        Call objComissaoAux.Copiar(objComissao)
        mvarcolComissoesBaixa.Add objComissaoAux
        
    Next

End Sub
