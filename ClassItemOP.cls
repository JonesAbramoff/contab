VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public objInfoUsu As Object

Private mvarlNumIntDoc As Long
Private mvariFilialEmpresa As Integer
Private mvarsCodigo As String
Private mvariItem As Integer
Private mvarsProduto As String
Private mvariFilialPedido As Integer
Private mvarlCodPedido As Long
Private mvariAlmoxarifado As Integer
Private mvarsCcl As String
Private mvarsSiglaUM As String
Private mvardQuantidade As Double
Private mvardtDataInicioProd As Date
Private mvardtDataFimProd As Date
Private mvariPrioridade As Integer
Private mvardQuantidadeProd As Double
Private mvariSituacao As Integer
Private mvariDestinacao As Integer
Private mvarlNumIntOrigem As Long
Private mvarsSiglaUMEstoque As String
Private mvariClasseUM As Integer
Private mvariBeneficiamento As Integer
Private mvarlNumIntEquipamento As Long
Private mvarlNumIntEquipamento2 As Long
Private mvarsVersao As String
Private mvarcolItensRomaneioGrade As New Collection
Private mvariPossuiGrade As Integer 'Indica se o Item é um pai de grade
Private mvarsDescricao As String
Private mvarsProdutoPai As String 'Codigo do produto pai de um item de grade
Private mvariOrigemPedido As Integer
Private mvarlNumItemOP As Long
Private mvarsAlmoxarifadoNomeRed As String
Private mvardMetragemCons As Double
Private mvardEnfesto As Double
Private mvardRisco As Double
Private mvariTipo As Integer '0 = Ordem de Producao, 1 = Ordem de Corte

'#####################################################
'INSERIDO POR WAGNER
Private mvarlNumIntDocPai As Long
Private mvarsJustificativaBloqueio As String
'#####################################################

'------------------------------------------
'Inserido por Jorge Specian - 03/05/2005
Private mvarcolOrdemProducaoOperacoes As New Collection
Private mvariProduzLogo As Integer
Public mvarcolProdutoKitInfo As New Collection
Public mvarcolItemPV As New Collection

Public Property Set colItemPV(ByVal vData As Collection)
    Set mvarcolItemPV = vData
End Property

Public Property Get colItemPV() As Collection
    Set colItemPV = mvarcolItemPV
End Property

Public Property Set colProdutoKitInfo(ByVal vData As Collection)
    Set mvarcolProdutoKitInfo = vData
End Property

Public Property Get colProdutoKitInfo() As Collection
    Set colProdutoKitInfo = mvarcolProdutoKitInfo
End Property

Public Property Let iProduzLogo(ByVal vData As Integer)
    mvariProduzLogo = vData
End Property

Public Property Get iProduzLogo() As Integer
    iProduzLogo = mvariProduzLogo
End Property

Public Property Get colOrdemProducaoOperacoes() As Collection
    Set colOrdemProducaoOperacoes = mvarcolOrdemProducaoOperacoes
End Property

Public Property Set colOrdemProducaoOperacoes(vData As Collection)
    Set mvarcolOrdemProducaoOperacoes = vData
End Property
'-------------------------------------------

'#####################################################
'INSERIDO POR WAGNER
Public Property Let lNumIntDocPai(ByVal vData As Long)
    mvarlNumIntDocPai = vData
End Property

Public Property Get lNumIntDocPai() As Long
    lNumIntDocPai = mvarlNumIntDocPai
End Property

Public Property Let sJustificativaBloqueio(ByVal vData As String)
    mvarsJustificativaBloqueio = vData
End Property

Public Property Get sJustificativaBloqueio() As String
    sJustificativaBloqueio = mvarsJustificativaBloqueio
End Property
'#####################################################

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let dRisco(ByVal vData As Double)
    mvardRisco = vData
End Property

Public Property Get dRisco() As Double
    dRisco = mvardRisco
End Property

Public Property Let dEnfesto(ByVal vData As Double)
    mvardEnfesto = vData
End Property

Public Property Get dEnfesto() As Double
    dEnfesto = mvardEnfesto
End Property

Public Property Let dMetragemCons(ByVal vData As Double)
    mvardMetragemCons = vData
End Property

Public Property Get dMetragemCons() As Double
    dMetragemCons = mvardMetragemCons
End Property

Public Property Let sAlmoxarifadoNomeRed(ByVal vData As String)
    mvarsAlmoxarifadoNomeRed = vData
End Property

Public Property Get sAlmoxarifadoNomeRed() As String
    sAlmoxarifadoNomeRed = mvarsAlmoxarifadoNomeRed
End Property

Public Property Let lNumItemOP(ByVal vData As Long)
    mvarlNumItemOP = vData
End Property

Public Property Get lNumItemOP() As Long
    lNumItemOP = mvarlNumItemOP
End Property

Public Property Get iOrigemPedido() As Integer
    iOrigemPedido = mvariOrigemPedido
End Property

Public Property Let iOrigemPedido(ByVal vData As Integer)
    mvariOrigemPedido = vData
End Property

Public Property Let sProdutoPai(ByVal vData As String)
    mvarsProdutoPai = vData
End Property

Public Property Get sProdutoPai() As String
    sProdutoPai = mvarsProdutoPai
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Get iPossuiGrade() As Integer
    iPossuiGrade = mvariPossuiGrade
End Property

Public Property Let iPossuiGrade(ByVal vData As Integer)
    mvariPossuiGrade = vData
End Property

Public Property Set colItensRomaneioGrade(ByVal vData As Collection)
    Set mvarcolItensRomaneioGrade = vData
End Property

Public Property Get colItensRomaneioGrade() As Collection
    Set colItensRomaneioGrade = mvarcolItensRomaneioGrade
End Property

Public Property Let sVersao(ByVal vData As String)
    mvarsVersao = vData
End Property

Public Property Get sVersao() As String
    sVersao = mvarsVersao
End Property

Public Property Let lNumIntEquipamento(ByVal vData As Long)
    mvarlNumIntEquipamento = vData
End Property

Public Property Get lNumIntEquipamento() As Long
    lNumIntEquipamento = mvarlNumIntEquipamento
End Property

Public Property Let iBeneficiamento(ByVal vData As Integer)
    mvariBeneficiamento = vData
End Property

Public Property Get iBeneficiamento() As Integer
    iBeneficiamento = mvariBeneficiamento
End Property

Public Property Let iClasseUM(ByVal vData As Integer)
    mvariClasseUM = vData
End Property

Public Property Get iClasseUM() As Integer
    iClasseUM = mvariClasseUM
End Property

Public Property Let sSiglaUMEstoque(ByVal vData As String)
    mvarsSiglaUMEstoque = vData
End Property

Public Property Get sSiglaUMEstoque() As String
    sSiglaUMEstoque = mvarsSiglaUMEstoque
End Property

Public Property Let lNumIntOrigem(ByVal vData As Long)
    mvarlNumIntOrigem = vData
End Property

Public Property Get lNumIntOrigem() As Long
    lNumIntOrigem = mvarlNumIntOrigem
End Property

Public Property Let iItem(ByVal vData As Integer)
    mvariItem = vData
End Property

Public Property Get iItem() As Integer
    iItem = mvariItem
End Property

Public Property Let iDestinacao(ByVal vData As Integer)
    mvariDestinacao = vData
End Property

Public Property Get iDestinacao() As Integer
    iDestinacao = mvariDestinacao
End Property

Public Property Let iSituacao(ByVal vData As Integer)
    mvariSituacao = vData
End Property

Public Property Get iSituacao() As Integer
    iSituacao = mvariSituacao
End Property

Public Property Let dQuantidadeProd(ByVal vData As Double)
    mvardQuantidadeProd = vData
End Property

Public Property Get dQuantidadeProd() As Double
    dQuantidadeProd = mvardQuantidadeProd
End Property

Public Property Let iPrioridade(ByVal vData As Integer)
    mvariPrioridade = vData
End Property

Public Property Get iPrioridade() As Integer
    iPrioridade = mvariPrioridade
End Property

Public Property Let dtDataFimProd(ByVal vData As Date)
    mvardtDataFimProd = vData
End Property

Public Property Get dtDataFimProd() As Date
    dtDataFimProd = mvardtDataFimProd
End Property

Public Property Let dtDataInicioProd(ByVal vData As Date)
    mvardtDataInicioProd = vData
End Property

Public Property Get dtDataInicioProd() As Date
    dtDataInicioProd = mvardtDataInicioProd
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sSiglaUM(ByVal vData As String)
    mvarsSiglaUM = vData
End Property

Public Property Get sSiglaUM() As String
    sSiglaUM = mvarsSiglaUM
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

Public Property Let lCodPedido(ByVal vData As Long)
    mvarlCodPedido = vData
End Property

Public Property Get lCodPedido() As Long
    lCodPedido = mvarlCodPedido
End Property

Public Property Let iFilialPedido(ByVal vData As Integer)
    mvariFilialPedido = vData
End Property

Public Property Get iFilialPedido() As Integer
    iFilialPedido = mvariFilialPedido
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sCodigo(ByVal vData As String)
    mvarsCodigo = vData
End Property

Public Property Get sCodigo() As String
    sCodigo = mvarsCodigo
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntEquipamento2(ByVal vData As Long)
    mvarlNumIntEquipamento2 = vData
End Property

Public Property Get lNumIntEquipamento2() As Long
    lNumIntEquipamento2 = mvarlNumIntEquipamento2
End Property


