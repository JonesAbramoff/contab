VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassItensPedCompraPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDocItem AS Long
Private mvarlNumIntDoc AS Long
Private mvariFilialEmpresa AS Integer
Private mvarlCodigo AS Long
Private mvarlFornecedor AS Long
Private mvariFilial AS Integer
Private mvariComprador AS Integer
Private mvardtData AS Date
Private mvardtDataEmissao AS Date
Private mvardtDataEnvio AS Date
Private mvardValorTotal AS Double
Private mvarlObservacao AS Long
Private mvarlCodigoPV AS Long
Private mvarsNomeFornecedor AS String
Private mvarsNomeFilialForn AS String
Private mvarsNomeComprador AS String
Private mvarsNomeCondPagto AS String
Private mvarsObservacaoTexto AS String
Private mvarsProduto AS String
Private mvarsDescProduto AS String
Private mvardQuantidade AS Double
Private mvardQuantRecebida AS Double
Private mvardQuantRecebimento AS Double
Private mvarsUM AS String
Private mvardPrecoUnitario AS Double
Private mvardValorDesconto AS Double
Private mvardPercentMaisReceb AS Double
Private mvardPercentMenosReceb AS Double
Private mvariRecebForaFaixa AS Integer
Private mvariStatus AS Integer
Private mvarlObsItem AS Long
Private mvarsProdutoFornecedor AS String
Private mvardLoteMinimo AS Double
Private mvarsDescricao AS String
Private mvardValorIPI AS Double
Private mvardAliquotaIPI AS Double
Private mvardAliquotaICMS AS Double
Private mvariStatusBaixa AS Integer
Private mvardtDataLimite AS Date

Public Property Let lNumIntDocItem (ByVal vData As Long)
    mvarlNumIntDocItem = vData
End Property

Public Property Get lNumIntDocItem () AS Long
    lNumIntDocItem= mvarlNumIntDocItem
End Property

Public Property Let lNumIntDoc (ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc () AS Long
    lNumIntDoc= mvarlNumIntDoc
End Property

Public Property Let iFilialEmpresa (ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa () AS Integer
    iFilialEmpresa= mvariFilialEmpresa
End Property

Public Property Let lCodigo (ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo () AS Long
    lCodigo= mvarlCodigo
End Property

Public Property Let lFornecedor (ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor () AS Long
    lFornecedor= mvarlFornecedor
End Property

Public Property Let iFilial (ByVal vData As Integer)
    mvariFilial = vData
End Property

Public Property Get iFilial () AS Integer
    iFilial= mvariFilial
End Property

Public Property Let iComprador (ByVal vData As Integer)
    mvariComprador = vData
End Property

Public Property Get iComprador () AS Integer
    iComprador= mvariComprador
End Property

Public Property Let dtData (ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData () AS Date
    dtData= mvardtData
End Property

Public Property Let dtDataEmissao (ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao () AS Date
    dtDataEmissao= mvardtDataEmissao
End Property

Public Property Let dtDataEnvio (ByVal vData As Date)
    mvardtDataEnvio = vData
End Property

Public Property Get dtDataEnvio () AS Date
    dtDataEnvio= mvardtDataEnvio
End Property

Public Property Let dValorTotal (ByVal vData As Double)
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal () AS Double
    dValorTotal= mvardValorTotal
End Property

Public Property Let lObservacao (ByVal vData As Long)
    mvarlObservacao = vData
End Property

Public Property Get lObservacao () AS Long
    lObservacao= mvarlObservacao
End Property

Public Property Let lCodigoPV (ByVal vData As Long)
    mvarlCodigoPV = vData
End Property

Public Property Get lCodigoPV () AS Long
    lCodigoPV= mvarlCodigoPV
End Property

Public Property Let sNomeFornecedor (ByVal vData As String)
    mvarsNomeFornecedor = vData
End Property

Public Property Get sNomeFornecedor () AS String
    sNomeFornecedor= mvarsNomeFornecedor
End Property

Public Property Let sNomeFilialForn (ByVal vData As String)
    mvarsNomeFilialForn = vData
End Property

Public Property Get sNomeFilialForn () AS String
    sNomeFilialForn= mvarsNomeFilialForn
End Property

Public Property Let sNomeComprador (ByVal vData As String)
    mvarsNomeComprador = vData
End Property

Public Property Get sNomeComprador () AS String
    sNomeComprador= mvarsNomeComprador
End Property

Public Property Let sNomeCondPagto (ByVal vData As String)
    mvarsNomeCondPagto = vData
End Property

Public Property Get sNomeCondPagto () AS String
    sNomeCondPagto= mvarsNomeCondPagto
End Property

Public Property Let sObservacaoTexto (ByVal vData As String)
    mvarsObservacaoTexto = vData
End Property

Public Property Get sObservacaoTexto () AS String
    sObservacaoTexto= mvarsObservacaoTexto
End Property

Public Property Let sProduto (ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto () AS String
    sProduto= mvarsProduto
End Property

Public Property Let sDescProduto (ByVal vData As String)
    mvarsDescProduto = vData
End Property

Public Property Get sDescProduto () AS String
    sDescProduto= mvarsDescProduto
End Property

Public Property Let dQuantidade (ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade () AS Double
    dQuantidade= mvardQuantidade
End Property

Public Property Let dQuantRecebida (ByVal vData As Double)
    mvardQuantRecebida = vData
End Property

Public Property Get dQuantRecebida () AS Double
    dQuantRecebida= mvardQuantRecebida
End Property

Public Property Let dQuantRecebimento (ByVal vData As Double)
    mvardQuantRecebimento = vData
End Property

Public Property Get dQuantRecebimento () AS Double
    dQuantRecebimento= mvardQuantRecebimento
End Property

Public Property Let sUM (ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM () AS String
    sUM= mvarsUM
End Property

Public Property Let dPrecoUnitario (ByVal vData As Double)
    mvardPrecoUnitario = vData
End Property

Public Property Get dPrecoUnitario () AS Double
    dPrecoUnitario= mvardPrecoUnitario
End Property

Public Property Let dValorDesconto (ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto () AS Double
    dValorDesconto= mvardValorDesconto
End Property

Public Property Let dPercentMaisReceb (ByVal vData As Double)
    mvardPercentMaisReceb = vData
End Property

Public Property Get dPercentMaisReceb () AS Double
    dPercentMaisReceb= mvardPercentMaisReceb
End Property

Public Property Let dPercentMenosReceb (ByVal vData As Double)
    mvardPercentMenosReceb = vData
End Property

Public Property Get dPercentMenosReceb () AS Double
    dPercentMenosReceb= mvardPercentMenosReceb
End Property

Public Property Let iRecebForaFaixa (ByVal vData As Integer)
    mvariRecebForaFaixa = vData
End Property

Public Property Get iRecebForaFaixa () AS Integer
    iRecebForaFaixa= mvariRecebForaFaixa
End Property

Public Property Let iStatus (ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus () AS Integer
    iStatus= mvariStatus
End Property

Public Property Let lObsItem (ByVal vData As Long)
    mvarlObsItem = vData
End Property

Public Property Get lObsItem () AS Long
    lObsItem= mvarlObsItem
End Property

Public Property Let sProdutoFornecedor (ByVal vData As String)
    mvarsProdutoFornecedor = vData
End Property

Public Property Get sProdutoFornecedor () AS String
    sProdutoFornecedor= mvarsProdutoFornecedor
End Property

Public Property Let dLoteMinimo (ByVal vData As Double)
    mvardLoteMinimo = vData
End Property

Public Property Get dLoteMinimo () AS Double
    dLoteMinimo= mvardLoteMinimo
End Property

Public Property Let sDescricao (ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao () AS String
    sDescricao= mvarsDescricao
End Property

Public Property Let dValorIPI (ByVal vData As Double)
    mvardValorIPI = vData
End Property

Public Property Get dValorIPI () AS Double
    dValorIPI= mvardValorIPI
End Property

Public Property Let dAliquotaIPI (ByVal vData As Double)
    mvardAliquotaIPI = vData
End Property

Public Property Get dAliquotaIPI () AS Double
    dAliquotaIPI= mvardAliquotaIPI
End Property

Public Property Let dAliquotaICMS (ByVal vData As Double)
    mvardAliquotaICMS = vData
End Property

Public Property Get dAliquotaICMS () AS Double
    dAliquotaICMS= mvardAliquotaICMS
End Property

Public Property Let iStatusBaixa (ByVal vData As Integer)
    mvariStatusBaixa = vData
End Property

Public Property Get iStatusBaixa () AS Integer
    iStatusBaixa= mvariStatusBaixa
End Property

Public Property Let dtDataLimite (ByVal vData As Date)
    mvardtDataLimite = vData
End Property

Public Property Get dtDataLimite () AS Date
    dtDataLimite= mvardtDataLimite
End Property
