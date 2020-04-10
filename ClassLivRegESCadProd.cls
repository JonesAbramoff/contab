VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLivRegESCadProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlNumIntArqICMS As Long
Private mvarsProduto As String
Private mvardtDataInicial As Date
Private mvardtDataFinal As Date
Private mvarsCodigoNCM As String
Private mvarsDescricao As String
Private mvarsSiglaUM As String
Private mvarsSituacaoTrib As String
Private mvardAliquotaIPI As Double
Private mvardAliquotaICMS As Double
Private mvardReducaoBaseCalculoICMS As Double
Private mvardBaseCalculoICMSSubst As Double

Public Property Let dBaseCalculoICMSSubst(ByVal vData As Double)
Attribute dBaseCalculoICMSSubst.VB_Description = "Base de Cálculo ICMS de Substituição tributária"
    mvardBaseCalculoICMSSubst = vData
End Property

Public Property Get dBaseCalculoICMSSubst() As Double
    dBaseCalculoICMSSubst = mvardBaseCalculoICMSSubst
End Property

Public Property Let dReducaoBaseCalculoICMS(ByVal vData As Double)
Attribute dReducaoBaseCalculoICMS.VB_Description = "% de Redução na base de cálculo do ICMS, nas Operações internas."
    mvardReducaoBaseCalculoICMS = vData
End Property

Public Property Get dReducaoBaseCalculoICMS() As Double
    dReducaoBaseCalculoICMS = mvardReducaoBaseCalculoICMS
End Property

Public Property Let dAliquotaICMS(ByVal vData As Double)
Attribute dAliquotaICMS.VB_Description = "Aliquota ICMS do Produto"
    mvardAliquotaICMS = vData
End Property

Public Property Get dAliquotaICMS() As Double
    dAliquotaICMS = mvardAliquotaICMS
End Property

Public Property Let dAliquotaIPI(ByVal vData As Double)
Attribute dAliquotaIPI.VB_Description = "Aliquota IPI do Produto"
    mvardAliquotaIPI = vData
End Property

Public Property Get dAliquotaIPI() As Double
    dAliquotaIPI = mvardAliquotaIPI
End Property

Public Property Let sSituacaoTrib(ByVal vData As String)
Attribute sSituacaoTrib.VB_Description = "OrigemMercadoria &  iTipoTribCST (Tabela: TiposTribICMS)"
    mvarsSituacaoTrib = vData
End Property

Public Property Get sSituacaoTrib() As String
    sSituacaoTrib = mvarsSituacaoTrib
End Property

Public Property Let sSiglaUM(ByVal vData As String)
Attribute sSiglaUM.VB_Description = "Unidade de Medida de Comercialização do Produto"
    mvarsSiglaUM = vData
End Property

Public Property Get sSiglaUM() As String
    sSiglaUM = mvarsSiglaUM
End Property

Public Property Let sDescricao(ByVal vData As String)
Attribute sDescricao.VB_Description = "Descricao do Produto ou Serviço"
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sCodigoNCM(ByVal vData As String)
Attribute sCodigoNCM.VB_Description = "Codificação da Nomeclatura Comum do Mercosul (Obrigatório para IPI)"
    mvarsCodigoNCM = vData
End Property

Public Property Get sCodigoNCM() As String
    sCodigoNCM = mvarsCodigoNCM
End Property

Public Property Let dtDataFinal(ByVal vData As Date)
Attribute dtDataFinal.VB_Description = "Data Final"
    mvardtDataFinal = vData
End Property

Public Property Get dtDataFinal() As Date
    dtDataFinal = mvardtDataFinal
End Property

Public Property Let dtDataInicial(ByVal vData As Date)
Attribute dtDataInicial.VB_Description = "Data Inicial"
    mvardtDataInicial = vData
End Property

Public Property Get dtDataInicial() As Date
    dtDataInicial = mvardtDataInicial
End Property

Public Property Let sProduto(ByVal vData As String)
Attribute sProduto.VB_Description = "Código de Produto"
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let lNumIntArqICMS(ByVal vData As Long)
Attribute lNumIntArqICMS.VB_Description = "Número Interno do Arquivo de ICMS"
    mvarlNumIntArqICMS = vData
End Property

Public Property Get lNumIntArqICMS() As Long
    lNumIntArqICMS = mvarlNumIntArqICMS
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
Attribute lNumIntDoc.VB_Description = "Número Interno"
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property
