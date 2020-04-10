VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemMovEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvardCusto As Double
Private mvardCustoAnt As Double 'custo anterior utilizado em apuração de custo de produção para guardar o custo antigo e na atualização de estoque produto atualizar pela diferença
Private mvariApropriacao As Integer
Private mvariApropriacaoProd As Integer 'apropriacao de custo do produto
Private mvarsProduto As String
Private mvariClasseUM As Integer 'classe de unidade de medida ao qual o produto está associado
Private mvarsSiglaUM As String  'sigla da UM com que o movimento foi cadastrado
Private mvarsSiglaUMEst As String ' sigla da UM em que o produto é estocado
Private mvardQuantidade As Double 'quantidade em mvarsSiglaUM
Private mvardQuantidadeEst As Double ' quantidade em unidade de medida do estoque
Private mvardQuantidadeOP As Double ' quantidade a ser adicionada/subtraida de QuantOP em mvarsSiglaUM
Private mvardQuantidadeOPEst As Double ' quantidade a ser adicionada/subtraida de QuantOP em unidade de medida do estoque
Private mvariAlmoxarifado As Integer
Private mvarsCcl As String
Private mvarlNumIntDocEst As Long
Private mvariTipoMov As Integer
Private mvarsProdutoDesc As String
Private mvarsAlmoxarifadoNomeRed As String
Private mvarsOPCodigo As String
Private mvarsProdutoOP As String
Private mvarlNumIntDocOrigem As Long
Private mvariFilialEmpresa As Integer
Private mvarlCodigo As Long
Private mvardtData As Date
Private mvarlCliente As Long
Private mvarlFornecedor As Long
Private mvarsDocOrigem As String
Private mvarsContaContabilAplic As String
Private mvarsContaContabilEst As String
Private mvariTipoNumIntDocOrigem As Integer
Private mvariControleEstoque As Integer
Private mvarcolRastreamentoMovto As Collection
Private mvarcolApropriacaoInsumo As Collection
Private mvarlHorasMaquina As Long
Private mvardtDataInicioProducao As Date
Private mvardtDataRegistro As Date
Private mvardtHora As Date
Private mvariAtualizaSoLote As Integer
Private mvariNaturezaProduto As Integer
Private mvariPossuiGrade As Integer
Private mvarcolItensRomaneioGrade As New Collection
Private mvariBenef As Integer
Private mvarsDescricao As String
Private mvariIndicePaiGrade As Integer
Private mvarlNumIntDocGrade As Long
Private mvardQuantInsumos As Double 'quantidade na unidade de medida do item da nota fiscal vinculada
Private mvariItemNF As Integer 'indice do item da nota fiscal vinculada a este movimento. Utilizado em beneficiamento nota dev.simbolica
Private mvarlRequisitante As Long
Private mvarsObservacao As String
Private mvariItem As Integer
Private mvariFilialCli As Integer
Private mvariFilialForn As Integer

Public Property Let iFilialForn(ByVal vData As Integer)
    mvariFilialForn = vData
End Property

Public Property Get iFilialForn() As Integer
    iFilialForn = mvariFilialForn
End Property

Public Property Let iFilialCli(ByVal vData As Integer)
    mvariFilialCli = vData
End Property

Public Property Get iFilialCli() As Integer
    iFilialCli = mvariFilialCli
End Property

Public Property Let iItem(ByVal vData As Integer)
    mvariItem = vData
End Property

Public Property Get iItem() As Integer
    iItem = mvariItem
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let lRequisitante(ByVal vData As Long)
    mvarlRequisitante = vData
End Property

Public Property Get lRequisitante() As Long
    lRequisitante = mvarlRequisitante
End Property

Public Property Let iItemNF(ByVal vData As Integer)
    mvariItemNF = vData
End Property

Public Property Get iItemNF() As Integer
    iItemNF = mvariItemNF
End Property

Public Property Let dQuantInsumos(ByVal vData As Double)
    mvardQuantInsumos = vData
End Property

Public Property Get dQuantInsumos() As Double
    dQuantInsumos = mvardQuantInsumos
End Property

Public Property Let lNumIntDocGrade(ByVal vData As Long)
    mvarlNumIntDocGrade = vData
End Property

Public Property Get lNumIntDocGrade() As Long
    lNumIntDocGrade = mvarlNumIntDocGrade
End Property

Public Property Let iIndicePaiGrade(ByVal vData As Integer)
    mvariIndicePaiGrade = vData
End Property

Public Property Get iIndicePaiGrade() As Integer
    iIndicePaiGrade = mvariIndicePaiGrade
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iBenef(ByVal vData As Integer)
    mvariBenef = vData
End Property

Public Property Get iBenef() As Integer
    iBenef = mvariBenef
End Property

Public Property Set colItensRomaneioGrade(ByVal vData As Collection)
    Set mvarcolItensRomaneioGrade = vData
End Property

Public Property Get colItensRomaneioGrade() As Collection
    Set colItensRomaneioGrade = mvarcolItensRomaneioGrade
End Property

Public Property Let iPossuiGrade(ByVal vData As Integer)
    mvariPossuiGrade = vData
End Property

Public Property Get iPossuiGrade() As Integer
    iPossuiGrade = mvariPossuiGrade
End Property

Public Property Let iNaturezaProduto(ByVal vData As Integer)
    mvariNaturezaProduto = vData
End Property

Public Property Get iNaturezaProduto() As Integer
    iNaturezaProduto = mvariNaturezaProduto
End Property

Public Property Let iAtualizaSoLote(ByVal vData As Integer)
    mvariAtualizaSoLote = vData
End Property

Public Property Get iAtualizaSoLote() As Integer
    iAtualizaSoLote = mvariAtualizaSoLote
End Property

Public Property Let dtHora(ByVal vData As Date)
    mvardtHora = vData
End Property

Public Property Get dtHora() As Date
    dtHora = mvardtHora
End Property

Public Property Let dtDataRegistro(ByVal vData As Date)
    mvardtDataRegistro = vData
End Property

Public Property Get dtDataRegistro() As Date
    dtDataRegistro = mvardtDataRegistro
End Property

Public Property Let dtDataInicioProducao(ByVal vData As Date)
    mvardtDataInicioProducao = vData
End Property

Public Property Get dtDataInicioProducao() As Date
    dtDataInicioProducao = mvardtDataInicioProducao
End Property

Public Property Let lHorasMaquina(ByVal vData As Long)
    mvarlHorasMaquina = vData
End Property

Public Property Get lHorasMaquina() As Long
    lHorasMaquina = mvarlHorasMaquina
End Property

Public Property Set colApropriacaoInsumo(ByVal vData As Collection)
    Set mvarcolApropriacaoInsumo = vData
End Property

Public Property Get colApropriacaoInsumo() As Collection
    Set colApropriacaoInsumo = mvarcolApropriacaoInsumo
End Property

Public Property Set colRastreamentoMovto(ByVal vData As Collection)
    Set mvarcolRastreamentoMovto = vData
End Property

Public Property Get colRastreamentoMovto() As Collection
    Set colRastreamentoMovto = mvarcolRastreamentoMovto
End Property

Public Property Let iControleEstoque(ByVal vData As Integer)
    mvariControleEstoque = vData
End Property

Public Property Get iControleEstoque() As Integer
    iControleEstoque = mvariControleEstoque
End Property

Public Property Let iTipoNumIntDocOrigem(ByVal vData As Integer)
    mvariTipoNumIntDocOrigem = vData
End Property

Public Property Get iTipoNumIntDocOrigem() As Integer
    iTipoNumIntDocOrigem = mvariTipoNumIntDocOrigem
End Property

Public Property Let sContaContabilEst(ByVal vData As String)
    mvarsContaContabilEst = vData
End Property

Public Property Get sContaContabilEst() As String
    sContaContabilEst = mvarsContaContabilEst
End Property

Public Property Let sContaContabilAplic(ByVal vData As String)
    mvarsContaContabilAplic = vData
End Property

Public Property Get sContaContabilAplic() As String
    sContaContabilAplic = mvarsContaContabilAplic
End Property

Public Property Let sDocOrigem(ByVal vData As String)
    mvarsDocOrigem = vData
End Property

Public Property Get sDocOrigem() As String
    sDocOrigem = mvarsDocOrigem
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lNumIntDocOrigem(ByVal vData As Long)
    mvarlNumIntDocOrigem = vData
End Property

Public Property Get lNumIntDocOrigem() As Long
    lNumIntDocOrigem = mvarlNumIntDocOrigem
End Property

Public Property Let sProdutoOP(ByVal vData As String)
    mvarsProdutoOP = vData
End Property

Public Property Get sProdutoOP() As String
    sProdutoOP = mvarsProdutoOP
End Property

Public Property Let sOPCodigo(ByVal vData As String)
    mvarsOPCodigo = vData
End Property

Public Property Get sOPCodigo() As String
    sOPCodigo = mvarsOPCodigo
End Property

Public Property Let sAlmoxarifadoNomeRed(ByVal vData As String)
    mvarsAlmoxarifadoNomeRed = vData
End Property

Public Property Get sAlmoxarifadoNomeRed() As String
    sAlmoxarifadoNomeRed = mvarsAlmoxarifadoNomeRed
End Property

Public Property Let sProdutoDesc(ByVal vData As String)
    mvarsProdutoDesc = vData
End Property

Public Property Get sProdutoDesc() As String
    sProdutoDesc = mvarsProdutoDesc
End Property

Public Property Let iTipoMov(ByVal vData As Integer)
    mvariTipoMov = vData
End Property

Public Property Get iTipoMov() As Integer
    iTipoMov = mvariTipoMov
End Property

Public Property Let lNumIntDocEst(ByVal vData As Long)
    mvarlNumIntDocEst = vData
End Property

Public Property Get lNumIntDocEst() As Long
    lNumIntDocEst = mvarlNumIntDocEst
End Property

Public Property Let sCcl(ByVal vData As String)
    mvarsCcl = vData
End Property

Public Property Get sCcl() As String
    sCcl = mvarsCcl
End Property

Public Property Let iAlmoxarifado(ByVal vData As Integer)
    mvariAlmoxarifado = vData
End Property

Public Property Get iAlmoxarifado() As Integer
    iAlmoxarifado = mvariAlmoxarifado
End Property

Public Property Let dQuantidadeOP(ByVal vData As Double)
    mvardQuantidadeOP = vData
End Property

Public Property Get dQuantidadeOP() As Double
    dQuantidadeOP = mvardQuantidadeOP
End Property

Public Property Let dQuantidadeOPEst(ByVal vData As Double)
    mvardQuantidadeOPEst = vData
End Property

Public Property Get dQuantidadeOPEst() As Double
    dQuantidadeOPEst = mvardQuantidadeOPEst
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let dQuantidadeEst(ByVal vData As Double)
    mvardQuantidadeEst = vData
End Property

Public Property Get dQuantidadeEst() As Double
    dQuantidadeEst = mvardQuantidadeEst
End Property

Public Property Let sSiglaUM(ByVal vData As String)
    mvarsSiglaUM = vData
End Property

Public Property Get sSiglaUM() As String
    sSiglaUM = mvarsSiglaUM
End Property

Public Property Let sSiglaUMEst(ByVal vData As String)
    mvarsSiglaUMEst = vData
End Property

Public Property Get sSiglaUMEst() As String
    sSiglaUMEst = mvarsSiglaUMEst
End Property

Public Property Let iClasseUM(ByVal vData As Integer)
    mvariClasseUM = vData
End Property

Public Property Get iClasseUM() As Integer
    iClasseUM = mvariClasseUM
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let iApropriacao(ByVal vData As Integer)
    mvariApropriacao = vData
End Property

Public Property Get iApropriacaoProd() As Integer
    iApropriacaoProd = mvariApropriacaoProd
End Property

Public Property Let iApropriacaoProd(ByVal vData As Integer)
    mvariApropriacaoProd = vData
End Property

Public Property Get iApropriacao() As Integer
    iApropriacao = mvariApropriacao
End Property

Public Property Let dCustoAnt(ByVal vData As Double)
    mvardCustoAnt = vData
End Property

Public Property Get dCustoAnt() As Double
    dCustoAnt = mvardCustoAnt
End Property

Public Property Let dCusto(ByVal vData As Double)
    mvardCusto = vData
End Property

Public Property Get dCusto() As Double
    dCusto = mvardCusto
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property





