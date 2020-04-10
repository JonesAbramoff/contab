VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassIN86ItensNFEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public objTrib As ClassTributacaoDocItem

Private mvarsIndicadorMov As String
Private mvarsIndicadorMovCMP As String
Private mvariModeloDoc As Integer
Private mvarsSerie As String
Private mvarlNumeroNF As Long
Private mvardtDataEmissao As Date
Private mvariItemNF As Integer
Private mvarsProduto As String
Private mvarsDescProd As String
Private mvarsNaturezaOp As String
Private mvarsIPICodigo As String
Private mvardQuantidade As Double
Private mvarsUM As String
Private mvardPrecoUnitario As Double
Private mvardValorDesconto As Double
Private mvariIPITipo As Integer
Private mvardIPIAliquota As Double
Private mvardIPIBaseCalculo As Double
Private mvardIPIValor As Double
Private mvariOrigemMerc As Integer
Private mvariTipoTribCST As Integer
Private mvariICMSTipo As Integer
Private mvardICMSAliquota As Double
Private mvardICMSBase As Double
Private mvardICMSValor As Double
Private mvardICMSSubstValor As Double
Private mvardICMSSubstBase As Double
Private mvariTipoMovEstoque As Integer
Private mvariFilialEmpresa As Integer
Private mvariOrigem As Integer
Private mvarlFornecedor As Long
Private mvariFilialForn As Integer
Private mvarlCliente As Long
Private mvariFilialCli As Long
Private mvariColunaNoLivro As Integer
Private mvariColunaNoLivroEntrada As Integer
Private mvariColunaNoLivroSaida As Integer
Private mvariTipoTipoDocInfo As Integer 'Incluído por Luiz Nogueira em 28/01/04
Private mvarsCSTIPI As String

Public Property Let sCSTIPI(ByVal vData As String)
    mvarsCSTIPI = vData
End Property

Public Property Get sCSTIPI() As String
    sCSTIPI = mvarsCSTIPI
End Property

Public Property Let iColunaNoLivroSaida(ByVal vData As Integer)
    mvariColunaNoLivroSaida = vData
End Property

Public Property Get iColunaNoLivroSaida() As Integer
    iColunaNoLivroSaida = mvariColunaNoLivroSaida
End Property

Public Property Let iColunaNoLivroEntrada(ByVal vData As Integer)
    mvariColunaNoLivroEntrada = vData
End Property

Public Property Get iColunaNoLivroEntrada() As Integer
    iColunaNoLivroEntrada = mvariColunaNoLivroEntrada
End Property

Public Property Let iColunaNoLivro(ByVal vData As Integer)
    mvariColunaNoLivro = vData
End Property

Public Property Get iColunaNoLivro() As Integer
    iColunaNoLivro = mvariColunaNoLivro
End Property

Public Property Let iFilialCli(ByVal vData As Long)
    mvariFilialCli = vData
End Property

Public Property Get iFilialCli() As Long
    iFilialCli = mvariFilialCli
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let iFilialForn(ByVal vData As Integer)
    mvariFilialForn = vData
End Property

Public Property Get iFilialForn() As Integer
    iFilialForn = mvariFilialForn
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Let iOrigem(ByVal vData As Integer)
    mvariOrigem = vData
End Property

Public Property Get iOrigem() As Integer
    iOrigem = mvariOrigem
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iTipoMovEstoque(ByVal vData As Integer)
    mvariTipoMovEstoque = vData
End Property

Public Property Get iTipoMovEstoque() As Integer
    iTipoMovEstoque = mvariTipoMovEstoque
End Property

Public Property Let dICMSSubstBase(ByVal vData As Double)
    mvardICMSSubstBase = vData
End Property

Public Property Get dICMSSubstBase() As Double
    dICMSSubstBase = mvardICMSSubstBase
End Property

Public Property Let dICMSSubstValor(ByVal vData As Double)
    mvardICMSSubstValor = vData
End Property

Public Property Get dICMSSubstValor() As Double
    dICMSSubstValor = mvardICMSSubstValor
End Property

Public Property Let dICMSValor(ByVal vData As Double)
    mvardICMSValor = vData
End Property

Public Property Get dICMSValor() As Double
    dICMSValor = mvardICMSValor
End Property

Public Property Let dICMSBase(ByVal vData As Double)
    mvardICMSBase = vData
End Property

Public Property Get dICMSBase() As Double
    dICMSBase = mvardICMSBase
End Property

Public Property Let dICMSAliquota(ByVal vData As Double)
    mvardICMSAliquota = vData
End Property

Public Property Get dICMSAliquota() As Double
    dICMSAliquota = mvardICMSAliquota
End Property

Public Property Let iICMSTipo(ByVal vData As Integer)
    mvariICMSTipo = vData
End Property

Public Property Get iICMSTipo() As Integer
    iICMSTipo = mvariICMSTipo
End Property

Public Property Let iTipoTribCST(ByVal vData As Integer)
    mvariTipoTribCST = vData
End Property

Public Property Get iTipoTribCST() As Integer
    iTipoTribCST = mvariTipoTribCST
End Property

Public Property Let iOrigemMerc(ByVal vData As Integer)
    mvariOrigemMerc = vData
End Property

Public Property Get iOrigemMerc() As Integer
    iOrigemMerc = mvariOrigemMerc
End Property

Public Property Let dIPIValor(ByVal vData As Double)
    mvardIPIValor = vData
End Property

Public Property Get dIPIValor() As Double
    dIPIValor = mvardIPIValor
End Property

Public Property Let dIPIBaseCalculo(ByVal vData As Double)
    mvardIPIBaseCalculo = vData
End Property

Public Property Get dIPIBaseCalculo() As Double
    dIPIBaseCalculo = mvardIPIBaseCalculo
End Property

Public Property Let dIPIALiquota(ByVal vData As Double)
    mvardIPIAliquota = vData
End Property

Public Property Get dIPIALiquota() As Double
    dIPIALiquota = mvardIPIAliquota
End Property

Public Property Let iIPITipo(ByVal vData As Integer)
    mvariIPITipo = vData
End Property

Public Property Get iIPITipo() As Integer
    iIPITipo = mvariIPITipo
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property

Public Property Let dPrecoUnitario(ByVal vData As Double)
    mvardPrecoUnitario = vData
End Property

Public Property Get dPrecoUnitario() As Double
    dPrecoUnitario = mvardPrecoUnitario
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sIPICodigo(ByVal vData As String)
    mvarsIPICodigo = vData
End Property

Public Property Get sIPICodigo() As String
    sIPICodigo = mvarsIPICodigo
End Property

Public Property Let sNaturezaOp(ByVal vData As String)
    mvarsNaturezaOp = vData
End Property

Public Property Get sNaturezaOp() As String
    sNaturezaOp = mvarsNaturezaOp
End Property

Public Property Let sDescProd(ByVal vData As String)
    mvarsDescProd = vData
End Property

Public Property Get sDescProd() As String
    sDescProd = mvarsDescProd
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let iItemNF(ByVal vData As Integer)
    mvariItemNF = vData
End Property

Public Property Get iItemNF() As Integer
    iItemNF = mvariItemNF
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let lNumeroNF(ByVal vData As Long)
    mvarlNumeroNF = vData
End Property

Public Property Get lNumeroNF() As Long
    lNumeroNF = mvarlNumeroNF
End Property

Public Property Let sSerie(ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let iModeloDoc(ByVal vData As Integer)
    mvariModeloDoc = vData
End Property

Public Property Get iModeloDoc() As Integer
    iModeloDoc = mvariModeloDoc
End Property

Public Property Let sIndicadorMovCMP(ByVal vData As String)
    mvarsIndicadorMovCMP = vData
End Property

Public Property Get sIndicadorMovCMP() As String
    sIndicadorMovCMP = mvarsIndicadorMovCMP
End Property

Public Property Let sIndicadorMov(ByVal vData As String)
    mvarsIndicadorMov = vData
End Property

Public Property Get sIndicadorMov() As String
    sIndicadorMov = mvarsIndicadorMov
End Property

'Incluído por Luiz Nogueira em 28/01/04
Public Property Let iTipoTipoDocInfo(ByVal vData As Integer)
    mvariTipoTipoDocInfo = vData
End Property

'Incluído por Luiz Nogueira em 28/01/04
Public Property Get iTipoTipoDocInfo() As Integer
    iTipoTipoDocInfo = mvariTipoTipoDocInfo
End Property

