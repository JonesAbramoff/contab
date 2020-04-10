VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassContasCorrentesInternas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsNomeReduzido As String
Private mvariCodigo As Integer
Private mvarsDescricao As Variant
Private mvariCodBanco As Integer
Private mvarsAgencia As String
Private mvarsDVAgencia As String
Private mvarsNumConta As String
Private mvarsDVNumConta As String
Private mvarsDVAgConta As String
Private mvarsContato As String
Private mvarsConvenioPagto As String
Private mvarsTelefone As String
Private mvardSaldoInicial As Double
Private mvardtDataInicial As Date
Private mvarsContaContabil As String
Private mvarlProxSeqMov As Long
Private mvarsFax As String
Private mvariNumMenorExtratoNaoConciliado As Integer
Private mvarlProxBordero As Long
Private mvariChequeBordero As Integer
Private mvariFilialEmpresa As Integer
Private mvariChequePre As Integer
Private mvariAtivo As Integer
Private mvardtDataLog As Date
Private mvardRotativo As Double
Private mvarsDirArqBordPagto As String
Private mvarsContaContabilChqPre As String
Private mvarlCNABProxSeqArqCobr As Long

Public Property Let lCNABProxSeqArqCobr(ByVal vData As Long)
    mvarlCNABProxSeqArqCobr = vData
End Property

Public Property Get lCNABProxSeqArqCobr() As Long
    lCNABProxSeqArqCobr = mvarlCNABProxSeqArqCobr
End Property

Public Property Let sContaContabilChqPre(ByVal vData As String)
    mvarsContaContabilChqPre = vData
End Property

Public Property Get sContaContabilChqPre() As String
    sContaContabilChqPre = mvarsContaContabilChqPre
End Property

Public Property Let sDirArqBordPagto(ByVal vData As String)
    mvarsDirArqBordPagto = vData
End Property

Public Property Get sDirArqBordPagto() As String
    sDirArqBordPagto = mvarsDirArqBordPagto
End Property

Public Property Let dRotativo(ByVal vData As Double)
    mvardRotativo = vData
End Property

Public Property Get dRotativo() As Double
    dRotativo = mvardRotativo
End Property


Public Property Let dtDataLog(ByVal vData As Date)
    mvardtDataLog = vData
End Property

Public Property Get dtDataLog() As Date
    dtDataLog = mvardtDataLog
End Property

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property

Public Property Let iChequePre(ByVal vData As Integer)
    mvariChequePre = vData
End Property

Public Property Get iChequePre() As Integer
    iChequePre = mvariChequePre
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iChequeBordero(ByVal vData As Integer)
Attribute iChequeBordero.VB_Description = "1:conta está sendo utilizada pela função de emissão de cheque/bordero 0:caso contrário"
    mvariChequeBordero = vData
End Property

Public Property Get iChequeBordero() As Integer
    iChequeBordero = mvariChequeBordero
End Property

Public Property Let lProxBordero(ByVal vData As Long)
Attribute lProxBordero.VB_Description = "proximo numero de bordero a ser impresso"
    mvarlProxBordero = vData
End Property

Public Property Get lProxBordero() As Long
    lProxBordero = mvarlProxBordero
End Property

Public Property Let iNumMenorExtratoNaoConciliado(ByVal vData As Integer)
Attribute iNumMenorExtratoNaoConciliado.VB_Description = "sequencial do menor extrato nao conciliado p/ser a sugestao na rotina de conciliacao de extratos bancarios, se estiver zerado nao tem sugestao."
    mvariNumMenorExtratoNaoConciliado = vData
End Property

Public Property Get iNumMenorExtratoNaoConciliado() As Integer
    iNumMenorExtratoNaoConciliado = mvariNumMenorExtratoNaoConciliado
End Property

Public Property Let sFax(ByVal vData As String)
Attribute sFax.VB_Description = "Telefone 2 do banco"
    mvarsFax = vData
End Property

Public Property Get sFax() As String
    sFax = mvarsFax
End Property

Public Property Let lProxSeqMov(ByVal vData As Long)
Attribute lProxSeqMov.VB_Description = "proximo sequencial de movimento"
    mvarlProxSeqMov = vData
End Property

Public Property Get lProxSeqMov() As Long
    lProxSeqMov = mvarlProxSeqMov
End Property

Public Property Let sContaContabil(ByVal vData As String)
Attribute sContaContabil.VB_Description = "Numero da Conta Contabil. Se estiver em branco nao será gerada contabilizacao default para esta conta"
    mvarsContaContabil = vData
End Property

Public Property Get sContaContabil() As String
    sContaContabil = mvarsContaContabil
End Property

Public Property Let dtDataInicial(ByVal vData As Date)
Attribute dtDataInicial.VB_Description = "Data do Saldo Inicial."
    mvardtDataInicial = vData
End Property

Public Property Get dtDataInicial() As Date
    dtDataInicial = mvardtDataInicial
End Property

Public Property Let dSaldoInicial(ByVal vData As Double)
Attribute dSaldoInicial.VB_Description = "Saldo Inicial da Conta."
    mvardSaldoInicial = vData
End Property

Public Property Get dSaldoInicial() As Double
    dSaldoInicial = mvardSaldoInicial
End Property

Public Property Let sTelefone(ByVal vData As String)
Attribute sTelefone.VB_Description = "Telefone 1 do banco"
    mvarsTelefone = vData
End Property

Public Property Get sTelefone() As String
    sTelefone = mvarsTelefone
End Property

Public Property Let sContato(ByVal vData As String)
Attribute sContato.VB_Description = "Nome do Contado no banco"
    mvarsContato = vData
End Property

Public Property Get sContato() As String
    sContato = mvarsContato
End Property

Public Property Let sConvenioPagto(ByVal vData As String)
    mvarsConvenioPagto = vData
End Property

Public Property Get sConvenioPagto() As String
    sConvenioPagto = mvarsConvenioPagto
End Property

Public Property Let sDVAgConta(ByVal vData As String)
Attribute sDVAgConta.VB_Description = "Digito Verificador do conjunto agencia e conta corrente. Formado por 1 caracter alfanumerico"
    mvarsDVAgConta = vData
End Property

Public Property Get sDVAgConta() As String
    sDVAgConta = mvarsDVAgConta
End Property

Public Property Let sDVNumConta(ByVal vData As String)
Attribute sDVNumConta.VB_Description = "Digito Verificador associado à conta corrented. Formado por 1 caracter alfanumerico"
    mvarsDVNumConta = vData
End Property

Public Property Get sDVNumConta() As String
    sDVNumConta = mvarsDVNumConta
End Property

Public Property Let sNumConta(ByVal vData As String)
Attribute sNumConta.VB_Description = "Número da conta corrente. É formado por até 12 digitos."
    mvarsNumConta = vData
End Property

Public Property Get sNumConta() As String
    sNumConta = mvarsNumConta
End Property

Public Property Let sDVAgencia(ByVal vData As String)
Attribute sDVAgencia.VB_Description = "Digito Verificador da agencia bancária. É formado por um caracter alfanumerico."
    mvarsDVAgencia = vData
End Property

Public Property Get sDVAgencia() As String
    sDVAgencia = mvarsDVAgencia
End Property

Public Property Let sAgencia(ByVal vData As String)
Attribute sAgencia.VB_Description = "Se estiver preenchida, Contem o codigo da agencia bancária. Este codigo é formado por até 5 digitos."
    mvarsAgencia = vData
End Property

Public Property Get sAgencia() As String
    sAgencia = mvarsAgencia
End Property

Public Property Let iCodBanco(ByVal vData As Integer)
Attribute iCodBanco.VB_Description = "Codigo do Banco ao qual a conta se refere. Se este campo estiver zerado significa que esta conta não é uma conta corrente bancária."
    mvariCodBanco = vData
End Property

Public Property Get iCodBanco() As Integer
    iCodBanco = mvariCodBanco
End Property

Public Property Let sDescricao(ByVal vData As Variant)
Attribute sDescricao.VB_Description = "Texto descritivo da conta."
    mvarsDescricao = vData
End Property

Public Property Set sDescricao(ByVal vData As Object)
    Set mvarsDescricao = vData
End Property

Public Property Get sDescricao() As Variant
    If IsObject(mvarsDescricao) Then
        Set sDescricao = mvarsDescricao
    Else
        sDescricao = mvarsDescricao
    End If
End Property

Public Property Let iCodigo(ByVal vData As Integer)
Attribute iCodigo.VB_Description = "identificador da conta corrente interna"
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
Attribute sNomeReduzido.VB_Description = "Nome Reduzido da Conta Corrente Interna"
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property



