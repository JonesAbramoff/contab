VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassICMSExcecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarsEstadoDestino As String
Private mvarsCategoriaProduto As String
Private mvarsCategoriaProdutoItem As String
Private mvarsCategoriaCliente As String
Private mvarsCategoriaClienteItem As String
Private mvariTipo As Integer
Private mvariTipoSimples As Integer
Private mvardPercRedBaseCalculo As Double
Private mvardAliquota As Double
Private mvardPercMargemLucro As Double
Private mvarsFundamentacao As String
Private mvariPrioridade As Integer
Private mvarsEstadoOrigem As String
Private mvariTipoCliForn As Integer
Private mvarsCategoriaFornecedor As String
Private mvarsCategoriaFornecedorItem As String
Private mvariUsaPauta As Integer '0:Substituicao Tributaria usa margem, 1:ST usa valor de Pauta
Private mvardValorPauta As Double
Private mvardPercRedBaseCalculoSubst As Double
Private mvariGrupoOrigemMercadoria As Integer
Private mvardICMSPercFCP As Double
Private mvariTipoAplicacao As Integer
Private mvariICMSSTBaseDupla As Integer
Private mvardtICMSSTBaseDuplaIni As Date

Public scBenef As String
Public iICMSMotivo As Integer

Public Property Let dtICMSSTBaseDuplaIni(ByVal vData As Date)
    mvardtICMSSTBaseDuplaIni = vData
End Property

Public Property Get dtICMSSTBaseDuplaIni() As Date
    dtICMSSTBaseDuplaIni = mvardtICMSSTBaseDuplaIni
End Property

Public Property Let iICMSSTBaseDupla(ByVal vData As Integer)
    mvariICMSSTBaseDupla = vData
End Property

Public Property Get iICMSSTBaseDupla() As Integer
    iICMSSTBaseDupla = mvariICMSSTBaseDupla
End Property

Public Property Let iTipoAplicacao(ByVal vData As Integer)
    mvariTipoAplicacao = vData
End Property

Public Property Get iTipoAplicacao() As Integer
    iTipoAplicacao = mvariTipoAplicacao
End Property

Public Property Let dICMSPercFCP(ByVal vData As Double)
    mvardICMSPercFCP = vData
End Property

Public Property Get dICMSPercFCP() As Double
    dICMSPercFCP = mvardICMSPercFCP
End Property

Public Property Let iGrupoOrigemMercadoria(ByVal vData As Integer)
    mvariGrupoOrigemMercadoria = vData
End Property

Public Property Get iGrupoOrigemMercadoria() As Integer
    iGrupoOrigemMercadoria = mvariGrupoOrigemMercadoria
End Property

Public Property Let dPercRedBaseCalculoSubst(ByVal vData As Double)
    mvardPercRedBaseCalculoSubst = vData
End Property

Public Property Get dPercRedBaseCalculoSubst() As Double
    dPercRedBaseCalculoSubst = mvardPercRedBaseCalculoSubst
End Property

Public Property Let dValorPauta(ByVal vData As Double)
    mvardValorPauta = vData
End Property

Public Property Get dValorPauta() As Double
    dValorPauta = mvardValorPauta
End Property

Public Property Let iUsaPauta(ByVal vData As Integer)
    mvariUsaPauta = vData
End Property

Public Property Get iUsaPauta() As Integer
    iUsaPauta = mvariUsaPauta
End Property

Public Property Let sCategoriaFornecedorItem(ByVal vData As String)
    mvarsCategoriaFornecedorItem = vData
End Property

Public Property Get sCategoriaFornecedorItem() As String
    sCategoriaFornecedorItem = mvarsCategoriaFornecedorItem
End Property

Public Property Let sCategoriaFornecedor(ByVal vData As String)
    mvarsCategoriaFornecedor = vData
End Property

Public Property Get sCategoriaFornecedor() As String
    sCategoriaFornecedor = mvarsCategoriaFornecedor
End Property

Public Property Let iTipoCliForn(ByVal vData As Integer)
    mvariTipoCliForn = vData
End Property

Public Property Get iTipoCliForn() As Integer
    iTipoCliForn = mvariTipoCliForn
End Property

Public Property Let sEstadoOrigem(ByVal vData As String)
    mvarsEstadoOrigem = vData
End Property

Public Property Get sEstadoOrigem() As String
    sEstadoOrigem = mvarsEstadoOrigem
End Property

Public Property Let iPrioridade(ByVal vData As Integer)
    mvariPrioridade = vData
End Property

Public Property Get iPrioridade() As Integer
    iPrioridade = mvariPrioridade
End Property

Public Property Let sFundamentacao(ByVal vData As String)
    mvarsFundamentacao = vData
End Property

Public Property Get sFundamentacao() As String
    sFundamentacao = mvarsFundamentacao
End Property

Public Property Let dPercMargemLucro(ByVal vData As Double)
    mvardPercMargemLucro = vData
End Property

Public Property Get dPercMargemLucro() As Double
    dPercMargemLucro = mvardPercMargemLucro
End Property

Public Property Let dAliquota(ByVal vData As Double)
    mvardAliquota = vData
End Property

Public Property Get dAliquota() As Double
    dAliquota = mvardAliquota
End Property

Public Property Let dPercRedBaseCalculo(ByVal vData As Double)
    mvardPercRedBaseCalculo = vData
End Property

Public Property Get dPercRedBaseCalculo() As Double
    dPercRedBaseCalculo = mvardPercRedBaseCalculo
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let iTipoSimples(ByVal vData As Integer)
    mvariTipoSimples = vData
End Property

Public Property Get iTipoSimples() As Integer
    iTipoSimples = mvariTipoSimples
End Property

Public Property Let sCategoriaClienteItem(ByVal vData As String)
    mvarsCategoriaClienteItem = vData
End Property

Public Property Get sCategoriaClienteItem() As String
    sCategoriaClienteItem = mvarsCategoriaClienteItem
End Property

Public Property Let sCategoriaCliente(ByVal vData As String)
    mvarsCategoriaCliente = vData
End Property

Public Property Get sCategoriaCliente() As String
    sCategoriaCliente = mvarsCategoriaCliente
End Property

Public Property Let sCategoriaProdutoItem(ByVal vData As String)
    mvarsCategoriaProdutoItem = vData
End Property

Public Property Get sCategoriaProdutoItem() As String
    sCategoriaProdutoItem = mvarsCategoriaProdutoItem
End Property

Public Property Let sCategoriaProduto(ByVal vData As String)
    mvarsCategoriaProduto = vData
End Property

Public Property Get sCategoriaProduto() As String
    sCategoriaProduto = mvarsCategoriaProduto
End Property

Public Property Let sEstadoDestino(ByVal vData As String)
    mvarsEstadoDestino = vData
End Property

Public Property Get sEstadoDestino() As String
    sEstadoDestino = mvarsEstadoDestino
End Property



Private Sub Class_Initialize()
    scBenef = ""
    iICMSMotivo = 0
End Sub
