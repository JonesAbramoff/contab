VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMovContaCorrente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumMovto As Long
Private mvariCodConta As Integer
Private mvarlSequencial As Long
Private mvariTipo As Integer
Private mvariExcluido As Integer
Private mvariTipoMeioPagto As Integer
Private mvarlNumero As Long
Private mvardtDataMovimento As Date
''Private mvardtDataDeposito As Date
Private mvardtDataContabil As Date
Private mvardValor As Double
Private mvarsHistorico As String
Private mvariPortador As Integer
Private mvariConciliado As Integer
''Private mvardSaldoNaoApropriado As Double
''Private mvarlCliente As Long
''Private mvariFilial_Cliente As Integer
''Private mvarlFornecedor As Long
''Private mvariFilial_Fornecedor As Integer
Private mvarsNumRefExterna As String
''Private mvarsBeneficiario As String
Private mvariFavorecido As Integer
Private mvarlNumRefInterna As Long
Private mvarsOrigem As String
Private mvariExercicio As Integer
Private mvariLote As Integer
Private mvariPeriodo As Integer
Private mvarlDoc As Long
Private mvariFilialEmpresa As Integer
Private mvardtDataBaixa As Date
Private mvarlNumMovtoTransf As Long
Private mvariTipoMeioPagtoLoja As Integer
'############################################
'Inserido por Wagner
Private mvarsObservacao As String
Private mvarobjProjetoInfo As Object
Private mvarsNatureza As String
Private mvarsCcl As String

Public Property Let sCcl(ByVal vData As String)
    mvarsCcl = vData
End Property

Public Property Get sCcl() As String
    sCcl = mvarsCcl
End Property

Public Property Let sNatureza(ByVal vData As String)
    mvarsNatureza = vData
End Property

Public Property Get sNatureza() As String
    sNatureza = mvarsNatureza
End Property

Public Property Set objProjetoInfo(ByVal vData As Object)
    Set mvarobjProjetoInfo = vData
End Property

Public Property Get objProjetoInfo() As Object
    Set objProjetoInfo = mvarobjProjetoInfo
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property
'############################################

Public Property Let iTipoMeioPagtoLoja(ByVal vData As Integer)
    mvariTipoMeioPagtoLoja = vData
End Property

Public Property Get iTipoMeioPagtoLoja() As Integer
    iTipoMeioPagtoLoja = mvariTipoMeioPagtoLoja
End Property

Public Property Let lNumMovtoTransf(ByVal vData As Long)
    mvarlNumMovtoTransf = vData
End Property

Public Property Get lNumMovtoTransf() As Long
    lNumMovtoTransf = mvarlNumMovtoTransf
End Property


Public Property Let dtDataBaixa(ByVal vData As Date)
    mvardtDataBaixa = vData
End Property

Public Property Get dtDataBaixa() As Date
    dtDataBaixa = mvardtDataBaixa
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lDoc(ByVal vData As Long)
Attribute lDoc.VB_Description = "Código do Documento varia de 1 a 99999"
    mvarlDoc = vData
End Property

Public Property Get lDoc() As Long
    lDoc = mvarlDoc
End Property

Public Property Let iPeriodo(ByVal vData As Integer)
Attribute iPeriodo.VB_Description = "Periodo do Lote"
    mvariPeriodo = vData
End Property

Public Property Get iPeriodo() As Integer
    iPeriodo = mvariPeriodo
End Property

Public Property Let iLote(ByVal vData As Integer)
Attribute iLote.VB_Description = "Número do Lote (de 1 a 9999)"
    mvariLote = vData
End Property

Public Property Get iLote() As Integer
    iLote = mvariLote
End Property

Public Property Let iExercicio(ByVal vData As Integer)
Attribute iExercicio.VB_Description = "Exercicio do Lote"
    mvariExercicio = vData
End Property

Public Property Get iExercicio() As Integer
    iExercicio = mvariExercicio
End Property

Public Property Let sOrigem(ByVal vData As String)
Attribute sOrigem.VB_Description = "Onde o lote foi gerado"
    mvarsOrigem = vData
End Property

Public Property Get sOrigem() As String
    sOrigem = mvarsOrigem
End Property

Public Property Let lNumRefInterna(ByVal vData As Long)
Attribute lNumRefInterna.VB_Description = "identifica bordero de pagto, bordero de retorno da cobranca,"
    mvarlNumRefInterna = vData
End Property

Public Property Get lNumRefInterna() As Long
    lNumRefInterna = mvarlNumRefInterna
End Property

Public Property Let iFavorecido(ByVal vData As Integer)
Attribute iFavorecido.VB_Description = "p/saques de tesouraria e saques referentes a aplicacoes"
    mvariFavorecido = vData
End Property

Public Property Get iFavorecido() As Integer
    iFavorecido = mvariFavorecido
End Property

'Public Property Let sBeneficiario(ByVal vData As String)
'    mvarsBeneficiario = vData
'End Property
'
'Public Property Get sBeneficiario() As String
'    sBeneficiario = mvarsBeneficiario
'End Property

Public Property Let sNumRefExterna(ByVal vData As String)
Attribute sNumRefExterna.VB_Description = "referencia externa"
    mvarsNumRefExterna = vData
End Property

Public Property Get sNumRefExterna() As String
    sNumRefExterna = mvarsNumRefExterna
End Property

'Public Property Let iFilial_Fornecedor(ByVal vData As Integer)
'    mvariFilial_Fornecedor = vData
'End Property
'
'Public Property Get iFilial_Fornecedor() As Integer
'    iFilial_Fornecedor = mvariFilial_Fornecedor
'End Property
'
'Public Property Let lFornecedor(ByVal vData As Long)
'    mvarlFornecedor = vData
'End Property
'
'Public Property Get lFornecedor() As Long
'    lFornecedor = mvarlFornecedor
'End Property
'
'Public Property Let iFilial_Cliente(ByVal vData As Integer)
'    mvariFilial_Cliente = vData
'End Property
'
'Public Property Get iFilial_Cliente() As Integer
'    iFilial_Cliente = mvariFilial_Cliente
'End Property
'
'Public Property Let lCliente(ByVal vData As Long)
'    mvarlCliente = vData
'End Property
'
'Public Property Get lCliente() As Long
'    lCliente = mvarlCliente
'End Property
'
'Public Property Let dSaldoNaoApropriado(ByVal vData As Double)
'    mvardSaldoNaoApropriado = vData
'End Property
'
'Public Property Get dSaldoNaoApropriado() As Double
'    dSaldoNaoApropriado = mvardSaldoNaoApropriado
'End Property

Public Property Let iConciliado(ByVal vData As Integer)
Attribute iConciliado.VB_Description = "0: nao conciliado 1:conciliado manualmente, 2:conciliado via CNAB. Indica tipo de conferencia com extrato bancario"
    mvariConciliado = vData
End Property

Public Property Get iConciliado() As Integer
    iConciliado = mvariConciliado
End Property

Public Property Let iPortador(ByVal vData As Integer)
    mvariPortador = vData
End Property

Public Property Get iPortador() As Integer
    iPortador = mvariPortador
End Property

Public Property Let sHistorico(ByVal vData As String)
Attribute sHistorico.VB_Description = "Descritivo"
    mvarsHistorico = vData
End Property

Public Property Get sHistorico() As String
    sHistorico = mvarsHistorico
End Property

Public Property Let dValor(ByVal vData As Double)
Attribute dValor.VB_Description = "valor do movimento"
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dtDataContabil(ByVal vData As Date)
Attribute dtDataContabil.VB_Description = "Data da Contabilizacao"
    mvardtDataContabil = vData
End Property

Public Property Get dtDataContabil() As Date
    dtDataContabil = mvardtDataContabil
End Property

'Public Property Let dtDataDeposito(ByVal vData As Date)
'    mvardtDataDeposito = vData
'End Property
'
'Public Property Get dtDataDeposito() As Date
'    dtDataDeposito = mvardtDataDeposito
'End Property

Public Property Let dtDataMovimento(ByVal vData As Date)
Attribute dtDataMovimento.VB_Description = "Data do Movimento"
    mvardtDataMovimento = vData
End Property

Public Property Get dtDataMovimento() As Date
    dtDataMovimento = mvardtDataMovimento
End Property

Public Property Let lNumero(ByVal vData As Long)
Attribute lNumero.VB_Description = "numero do cheque, bordero,..."
    mvarlNumero = vData
End Property

Public Property Get lNumero() As Long
    lNumero = mvarlNumero
End Property

Public Property Let iTipoMeioPagto(ByVal vData As Integer)
Attribute iTipoMeioPagto.VB_Description = "0:cheque, 1:bordero,..."
    mvariTipoMeioPagto = vData
End Property

Public Property Get iTipoMeioPagto() As Integer
    iTipoMeioPagto = mvariTipoMeioPagto
End Property

Public Property Let iExcluido(ByVal vData As Integer)
Attribute iExcluido.VB_Description = "1=Excluido"
    mvariExcluido = vData
End Property

Public Property Get iExcluido() As Integer
    iExcluido = mvariExcluido
End Property

Public Property Let iTipo(ByVal vData As Integer)
Attribute iTipo.VB_Description = "0:saque, 1:deposito, 2:aplicacao, 3:resgate,4:saida transferencia,5:entrada transf, 6:pagto antecipado,7:receb antecipado,8:pagto de titulo,9:receto de titulo,..."
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let lSequencial(ByVal vData As Long)
Attribute lSequencial.VB_Description = "seq do movto dentro da conta"
    mvarlSequencial = vData
End Property

Public Property Get lSequencial() As Long
    lSequencial = mvarlSequencial
End Property

Public Property Let iCodConta(ByVal vData As Integer)
Attribute iCodConta.VB_Description = "codigo da conta corrente interna"
    mvariCodConta = vData
End Property

Public Property Get iCodConta() As Integer
    iCodConta = mvariCodConta
End Property

Public Property Let lNumMovto(ByVal vData As Long)
Attribute lNumMovto.VB_Description = "Identificador do movimento"
    mvarlNumMovto = vData
End Property

Public Property Get lNumMovto() As Long
    lNumMovto = mvarlNumMovto
End Property



