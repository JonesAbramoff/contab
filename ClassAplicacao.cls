VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassAplicacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlCodigo As Long
Private mvariFilialEmpresa As Integer
Private mvariStatus As Integer
Private mvarlNumMovto As Long
Private mvariTipoAplicacao As Integer
Private mvardSaldoAplicado As Double
Private mvardtDataResgatePrevista As Date
Private mvardValorResgatePrevisto As Double
Private mvardTaxaPrevista As Double
Private mvariProxSeqResgate As Integer
Private mvardtDataAplicacao As Date
Private mvardValorAplicado As Double
Private mvardtDataBaixa As Date

Public Property Let dtDataBaixa(ByVal vData As Date)
Attribute dtDataBaixa.VB_Description = "data em que a aplicacao foi completamente resgatada (ou DATA_NULA caso o resgate total ainda nao tenha ocorrido)"
    mvardtDataBaixa = vData
End Property

Public Property Get dtDataBaixa() As Date
    dtDataBaixa = mvardtDataBaixa
End Property

Public Property Let dValorAplicado(ByVal vData As Double)
Attribute dValorAplicado.VB_Description = "Valor da Aplicação"
    mvardValorAplicado = vData
End Property

Public Property Get dValorAplicado() As Double
    dValorAplicado = mvardValorAplicado
End Property

Public Property Let dtDataAplicacao(ByVal vData As Date)
Attribute dtDataAplicacao.VB_Description = "Data da Aplicacao"
    mvardtDataAplicacao = vData
End Property

Public Property Get dtDataAplicacao() As Date
    dtDataAplicacao = mvardtDataAplicacao
End Property

Public Property Let iProxSeqResgate(ByVal vData As Integer)
Attribute iProxSeqResgate.VB_Description = "proximo sequencial p/identificar resgate desta aplicacao"
    mvariProxSeqResgate = vData
End Property

Public Property Get iProxSeqResgate() As Integer
    iProxSeqResgate = mvariProxSeqResgate
End Property

Public Property Let dTaxaPrevista(ByVal vData As Double)
Attribute dTaxaPrevista.VB_Description = "rentabilidade nominal prevista"
    mvardTaxaPrevista = vData
End Property

Public Property Get dTaxaPrevista() As Double
    dTaxaPrevista = mvardTaxaPrevista
End Property

Public Property Let dValorResgatePrevisto(ByVal vData As Double)
Attribute dValorResgatePrevisto.VB_Description = "Valor previsto para o resgate da aplicação."
    mvardValorResgatePrevisto = vData
End Property

Public Property Get dValorResgatePrevisto() As Double
    dValorResgatePrevisto = mvardValorResgatePrevisto
End Property

Public Property Let dtDataResgatePrevista(ByVal vData As Date)
Attribute dtDataResgatePrevista.VB_Description = "Data prevista para o resgate da aplicação."
    mvardtDataResgatePrevista = vData
End Property

Public Property Get dtDataResgatePrevista() As Date
    dtDataResgatePrevista = mvardtDataResgatePrevista
End Property

Public Property Let dSaldoAplicado(ByVal vData As Double)
Attribute dSaldoAplicado.VB_Description = "valor aplicado menos valor resgatado"
    mvardSaldoAplicado = vData
End Property

Public Property Get dSaldoAplicado() As Double
    dSaldoAplicado = mvardSaldoAplicado
End Property

Public Property Let iTipoAplicacao(ByVal vData As Integer)
Attribute iTipoAplicacao.VB_Description = "identifica o saque em MovtosContaCorrente"
    mvariTipoAplicacao = vData
End Property

Public Property Get iTipoAplicacao() As Integer
    iTipoAplicacao = mvariTipoAplicacao
End Property

Public Property Let lNumMovto(ByVal vData As Long)
Attribute lNumMovto.VB_Description = "identifica o saque em MovtosContaCorrente"
    mvarlNumMovto = vData
End Property

Public Property Get lNumMovto() As Long
    lNumMovto = mvarlNumMovto
End Property

Public Property Let iStatus(ByVal vData As Integer)
Attribute iStatus.VB_Description = "0:ativa, 1:inativa (ou excluida)"
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lCodigo(ByVal vData As Long)
Attribute lCodigo.VB_Description = "Identificador da aplicação. Sequencial crescente a partir de valor armazenado em cprconfig"
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property



