VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassInfoParcPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvariSeqCheque As Integer
Private mvarsNomeRedForn As String
Private mvarsRazaoSocialForn As String
Private mvariFilialForn As Integer
Private mvarlNumTitulo As Long
Private mvariNumParcela As Integer
Private mvarlNumIntParc As Long
Private mvariTipoCobranca As Integer
Private mvarsNomeRedPortador As String
Private mvardValorJuros As Double
Private mvardValorMulta As Double
Private mvardValorDesconto As Double
Private mvardValor As Double
Private mvardValorOriginal As Double
Private mvariBancoCobrador As Integer
Private mvardtDataVencimento As Date
Private mvarlFornecedor As Long
Private mvariPortador As Integer
Private mvarsSiglaDocumento As String
Private mvariFilialEmpresa As Integer
Private mvarsContaFilForn As String 'conta contabil da filial do fornecedor
Private mvardtDataEmissao As Date
Private mvariMotivo As Integer
Private mvarlNumMovCta As Long
Private mvarlNumIntDoc As Long
Private mvarlNumIntBaixa As Long
Private mvariSequencial As Integer
Private mvariSelecionado As Integer
Private mvariAtivo As Integer
Private mvariMotivoDiferenca As Integer

Public Property Let iMotivoDiferenca(ByVal vData As Integer)
    mvariMotivoDiferenca = vData
End Property

Public Property Get iMotivoDiferenca() As Integer
    iMotivoDiferenca = mvariMotivoDiferenca
End Property

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property

Public Property Let iSelecionado(ByVal vData As Integer)
    mvariSelecionado = vData
End Property

Public Property Get iSelecionado() As Integer
    iSelecionado = mvariSelecionado
End Property

Public Property Let iSequencial(ByVal vData As Integer)
    mvariSequencial = vData
End Property

Public Property Get iSequencial() As Integer
    iSequencial = mvariSequencial
End Property

Public Property Let lNumIntBaixa(ByVal vData As Long)
    mvarlNumIntBaixa = vData
End Property

Public Property Get lNumIntBaixa() As Long
    lNumIntBaixa = mvarlNumIntBaixa
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
Attribute lNumIntDoc.VB_Description = "Número Interno do Crédito"
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumMovCta(ByVal vData As Long)
    mvarlNumMovCta = vData
End Property

Public Property Get lNumMovCta() As Long
    lNumMovCta = mvarlNumMovCta
End Property

Public Property Let iMotivo(ByVal vData As Integer)
    mvariMotivo = vData
End Property

Public Property Get iMotivo() As Integer
    iMotivo = mvariMotivo
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iPortador(ByVal vData As Integer)
    mvariPortador = vData
End Property

Public Property Get iPortador() As Integer
    iPortador = mvariPortador
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Let dtDataVencimento(ByVal vData As Date)
    mvardtDataVencimento = vData
End Property

Public Property Get dtDataVencimento() As Date
    dtDataVencimento = mvardtDataVencimento
End Property

Public Property Let iBancoCobrador(ByVal vData As Integer)
    mvariBancoCobrador = vData
End Property

Public Property Get iBancoCobrador() As Integer
    iBancoCobrador = mvariBancoCobrador
End Property

Public Property Let dValorOriginal(ByVal vData As Double)
    mvardValorOriginal = vData
End Property

Public Property Get dValorOriginal() As Double
    dValorOriginal = mvardValorOriginal
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property

Public Property Let dValorMulta(ByVal vData As Double)
    mvardValorMulta = vData
End Property

Public Property Get dValorMulta() As Double
    dValorMulta = mvardValorMulta
End Property

Public Property Let dValorJuros(ByVal vData As Double)
    mvardValorJuros = vData
End Property

Public Property Get dValorJuros() As Double
    dValorJuros = mvardValorJuros
End Property

Public Property Let sNomeRedPortador(ByVal vData As String)
    mvarsNomeRedPortador = vData
End Property

Public Property Get sNomeRedPortador() As String
    sNomeRedPortador = mvarsNomeRedPortador
End Property

Public Property Let iTipoCobranca(ByVal vData As Integer)
    mvariTipoCobranca = vData
End Property

Public Property Get iTipoCobranca() As Integer
    iTipoCobranca = mvariTipoCobranca
End Property

Public Property Let lNumIntParc(ByVal vData As Long)
    mvarlNumIntParc = vData
End Property

Public Property Get lNumIntParc() As Long
    lNumIntParc = mvarlNumIntParc
End Property

Public Property Let iNumParcela(ByVal vData As Integer)
    mvariNumParcela = vData
End Property

Public Property Get iNumParcela() As Integer
    iNumParcela = mvariNumParcela
End Property

Public Property Let lNumTitulo(ByVal vData As Long)
    mvarlNumTitulo = vData
End Property

Public Property Get lNumTitulo() As Long
    lNumTitulo = mvarlNumTitulo
End Property

Public Property Let iFilialForn(ByVal vData As Integer)
    mvariFilialForn = vData
End Property

Public Property Get iFilialForn() As Integer
    iFilialForn = mvariFilialForn
End Property

Public Property Let sRazaoSocialForn(ByVal vData As String)
    mvarsRazaoSocialForn = vData
End Property

Public Property Get sRazaoSocialForn() As String
    sRazaoSocialForn = mvarsRazaoSocialForn
End Property

Public Property Let sNomeRedForn(ByVal vData As String)
    mvarsNomeRedForn = vData
End Property

Public Property Get sNomeRedForn() As String
    sNomeRedForn = mvarsNomeRedForn
End Property

Public Property Let iSeqCheque(ByVal vData As Integer)
    mvariSeqCheque = vData
End Property

Public Property Get iSeqCheque() As Integer
    iSeqCheque = mvariSeqCheque
End Property

Public Property Let sSiglaDocumento(ByVal vData As String)
    mvarsSiglaDocumento = vData
End Property

Public Property Get sSiglaDocumento() As String
    sSiglaDocumento = mvarsSiglaDocumento
End Property

Public Property Let sContaFilForn(ByVal vData As String)
    mvarsContaFilForn = vData
End Property

Public Property Get sContaFilForn() As String
    sContaFilForn = mvarsContaFilForn
End Property



