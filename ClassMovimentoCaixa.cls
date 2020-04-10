VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMovimentoCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumMovto As Long
Private mvariFilialEmpresa As Integer
Private mvariCaixa As Integer
Private mvarlSequencial As Long
Private mvariTipo As Integer
Private mvariAdmMeioPagto As Integer
Private mvariParcelamento As Integer
Private mvariTipoCartao As Integer
Private mvarlNumero As Long
Private mvardtDataMovimento As Date
Private mvardValor As Double
Private mvarsHistorico As String
Private mvarsFavorecido As String
Private mvarlCupomFiscal As Long
Private mvarlNumRefInterna As Long
Private mvarlMovtoTransf As Long
Private mvarlMovtoEstorno As Long
Private mvariGerente As Integer
Private mvariCodConta As Integer
Private mvarlSequencialConta As Long
Private mvardHora As Double
Private mvariCodOperador As Integer
Private mvarlTransferencia As Long
'local variable(s) to hold property value(s)
Private mvariExcluiu As Integer 'local copy
Private mvarlNumIntExt As Long
Private mvarlNumIntDocLog As Long
Private mvariQuantLog As Integer 'local copy
Private mvardValorUnitario As Double
Private mvardQuantidade As Double
Private mvarsVinculado As String
Private mvariIndiceImpChq As Integer 'local copy
Private mvardtDataInicio As Date
Private mvardtDataFim As Date
Private mvarsArquivo As String
Private mvarlNumMovtoArq As Long
Private mvardtDataPreDatado As Date
Private mvarsAutorizacao As String

Public iCodigoCFe As Integer
Public sCNPJRede As String 'para incluir no xml da nfce
Public iBandeiraCartao As Integer 'para incluir no xml da nfce

Public iCodModeloECF As Integer

Public Property Let sAutorizacao(ByVal vData As String)
    mvarsAutorizacao = vData
End Property

Public Property Get sAutorizacao() As String
    sAutorizacao = mvarsAutorizacao
End Property

Public Property Let dtDataPreDatado(ByVal vData As Date)
    mvardtDataPreDatado = vData
End Property

Public Property Get dtDataPreDatado() As Date
    dtDataPreDatado = mvardtDataPreDatado
End Property

Public Property Let lNumMovtoArq(ByVal vData As Long)
    mvarlNumMovtoArq = vData
End Property

Public Property Get lNumMovtoArq() As Long
    lNumMovtoArq = mvarlNumMovtoArq
End Property

Public Property Let sArquivo(ByVal vData As String)
    mvarsArquivo = vData
End Property

Public Property Get sArquivo() As String
    sArquivo = mvarsArquivo
End Property


Public Property Let dtDataFim(ByVal vData As Date)
    mvardtDataFim = vData
End Property

Public Property Get dtDataFim() As Date
    dtDataFim = mvardtDataFim
End Property

Public Property Let dtDataInicio(ByVal vData As Date)
    mvardtDataInicio = vData
End Property

Public Property Get dtDataInicio() As Date
    dtDataInicio = mvardtDataInicio
End Property

Public Property Let iIndiceImpChq(ByVal vData As Integer)
    mvariIndiceImpChq = vData
End Property


Public Property Get iIndiceImpChq() As Integer
    iIndiceImpChq = mvariIndiceImpChq
End Property

Public Property Let sVinculado(ByVal vData As String)
    mvarsVinculado = vData
End Property

Public Property Get sVinculado() As String
    sVinculado = mvarsVinculado
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let dValorUnitario(ByVal vData As Double)
    mvardValorUnitario = vData
End Property

Public Property Get dValorUnitario() As Double
    dValorUnitario = mvardValorUnitario
End Property

Public Property Let iQuantLog(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iQuantLog = 5
    mvariQuantLog = vData
End Property

Public Property Get iQuantLog() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iQuantLog
    iQuantLog = mvariQuantLog
End Property

Public Property Let lNumIntDocLog(ByVal vData As Long)
    mvarlNumIntDocLog = vData
End Property

Public Property Get lNumIntDocLog() As Long
    lNumIntDocLog = mvarlNumIntDocLog
End Property

Public Property Let lNumIntExt(ByVal vData As Long)
    mvarlNumIntExt = vData
End Property

Public Property Get lNumIntExt() As Long
    lNumIntExt = mvarlNumIntExt
End Property

Public Property Let iExcluiu(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iExcluiu = 5
    mvariExcluiu = vData
End Property


Public Property Get iExcluiu() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iExcluiu
    iExcluiu = mvariExcluiu
End Property


Public Property Let lTransferencia(ByVal vData As Long)
    mvarlTransferencia = vData
End Property

Public Property Get lTransferencia() As Long
    lTransferencia = mvarlTransferencia
End Property

Public Property Let iCodOperador(ByVal vData As Integer)
    mvariCodOperador = vData
End Property

Public Property Get iCodOperador() As Integer
    iCodOperador = mvariCodOperador
End Property


Public Property Let dHora(ByVal vData As Double)
    mvardHora = vData
End Property

Public Property Get dHora() As Double
    dHora = mvardHora
End Property

Public Property Let lSequencialConta(ByVal vData As Long)
Attribute lSequencialConta.VB_Description = "Sequencial do movimento na conta corrente (usados no depósito bancário)."
    mvarlSequencialConta = vData
End Property

Public Property Get lSequencialConta() As Long
    lSequencialConta = mvarlSequencialConta
End Property

Public Property Let iCodConta(ByVal vData As Integer)
Attribute iCodConta.VB_Description = "Código da Conta corrente (usada no depósito bancário)"
    mvariCodConta = vData
End Property

Public Property Get iCodConta() As Integer
    iCodConta = mvariCodConta
End Property

Public Property Let iGerente(ByVal vData As Integer)
Attribute iGerente.VB_Description = "Código do gerente que autorizou o movimento (caso o necessite)"
    mvariGerente = vData
End Property

Public Property Get iGerente() As Integer
    iGerente = mvariGerente
End Property

Public Property Let lMovtoEstorno(ByVal vData As Long)
Attribute lMovtoEstorno.VB_Description = "Número Interno do Movimento de Estorno, se houver."
    mvarlMovtoEstorno = vData
End Property

Public Property Get lMovtoEstorno() As Long
    lMovtoEstorno = mvarlMovtoEstorno
End Property

Public Property Let lMovtoTransf(ByVal vData As Long)
Attribute lMovtoTransf.VB_Description = "Número interno do movimento de transferencia (MovimentoCaixa) correspondente."
    mvarlMovtoTransf = vData
End Property

Public Property Get lMovtoTransf() As Long
    lMovtoTransf = mvarlMovtoTransf
End Property

Public Property Let lNumRefInterna(ByVal vData As Long)
Attribute lNumRefInterna.VB_Description = "identifica cheque, boleto, vale/ticket, bordero de cheque, bordero de cheques, bordero de boletos, etc..."
    mvarlNumRefInterna = vData
End Property

Public Property Get lNumRefInterna() As Long
    lNumRefInterna = mvarlNumRefInterna
End Property

Public Property Let lCupomFiscal(ByVal vData As Long)
Attribute lCupomFiscal.VB_Description = "Número Interno do Cupom Fiscal"
    mvarlCupomFiscal = vData
End Property

Public Property Get lCupomFiscal() As Long
    lCupomFiscal = mvarlCupomFiscal
End Property

Public Property Let sFavorecido(ByVal vData As String)
    mvarsFavorecido = vData
End Property

Public Property Get sFavorecido() As String
    sFavorecido = mvarsFavorecido
End Property

Public Property Let sHistorico(ByVal vData As String)
Attribute sHistorico.VB_Description = "Descritivo"
    mvarsHistorico = vData
End Property

Public Property Get sHistorico() As String
    sHistorico = mvarsHistorico
End Property

Public Property Let dValor(ByVal vData As Double)
Attribute dValor.VB_Description = "Valor do movimento. sempre positivo, independente se for crédito ou débito."
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dtDataMovimento(ByVal vData As Date)
Attribute dtDataMovimento.VB_Description = "data/hora do movimento"
    mvardtDataMovimento = vData
End Property

Public Property Get dtDataMovimento() As Date
    dtDataMovimento = mvardtDataMovimento
End Property

Public Property Let lNumero(ByVal vData As Long)
Attribute lNumero.VB_Description = "Número do cheque, bordero de pagto,sequencial se for dinherio..."
    mvarlNumero = vData
End Property

Public Property Get lNumero() As Long
    lNumero = mvarlNumero
End Property

Public Property Let iParcelamento(ByVal vData As Integer)
Attribute iParcelamento.VB_Description = "0 -> A Vista    1 -> A Prazo . Preenchido quando preenche o Valor. No caso do cheque, a Prazo significa cheque-pre."
    mvariParcelamento = vData
End Property

Public Property Get iParcelamento() As Integer
    iParcelamento = mvariParcelamento
End Property

Public Property Let iTipoCartao(ByVal vData As Integer)
    mvariTipoCartao = vData
End Property

Public Property Get iTipoCartao() As Integer
    iTipoCartao = mvariTipoCartao
End Property

Public Property Let iAdmMeioPagto(ByVal vData As Integer)
Attribute iAdmMeioPagto.VB_Description = "1:dinheiro, 2:cheque, 3:bordero,..."
    mvariAdmMeioPagto = vData
End Property

Public Property Get iAdmMeioPagto() As Integer
    iAdmMeioPagto = mvariAdmMeioPagto
End Property

Public Property Let iTipo(ByVal vData As Integer)
Attribute iTipo.VB_Description = "Tipo do moviento que está sendo executado pela caixa."
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let lSequencial(ByVal vData As Long)
Attribute lSequencial.VB_Description = "Sequencial do movimento dentro da caixa"
    mvarlSequencial = vData
End Property

Public Property Get lSequencial() As Long
    lSequencial = mvarlSequencial
End Property

Public Property Let iCaixa(ByVal vData As Integer)
Attribute iCaixa.VB_Description = "Código da caixa"
    mvariCaixa = vData
End Property

Public Property Get iCaixa() As Integer
    iCaixa = mvariCaixa
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
Attribute iFilialEmpresa.VB_Description = "Filial Empresa que representa a Loja."
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lNumMovto(ByVal vData As Long)
Attribute lNumMovto.VB_Description = "Identificador do movimento"
    mvarlNumMovto = vData
End Property

Public Property Get lNumMovto() As Long
    lNumMovto = mvarlNumMovto
End Property

Private Sub Class_Initialize()
    mvardtDataPreDatado = DATA_NULA
End Sub
