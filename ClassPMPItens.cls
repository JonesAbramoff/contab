VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPMPItens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarsProduto As String
Private mvarsVersao As String
Private mvardQuantidade As Double
Private mvarsUM As String
Private mvardtDataNecessidade As Date
Private mvarlCodGeracao As Long
Private mvarsCodOPOrigem As String
Private mvariFilialEmpresa As Integer
Private mvarobjItemOP As New ClassItemOP
Private mvarobjOP As New ClassOrdemDeProducao
Private mvariStatus As Integer
Private mvarcolPO As New Collection
Private mvarlCliente As Long
Private mvariFilialCli As Integer
Private mvarsDescErro As String
Private mvariAlterado As Integer
Private mvariPrioridade As Integer
Private mvariOrdem As Integer
Private mvarlUltimoProxPO As Long
Private mvariProduzLogo As Integer
Private mvardtDataEntregaPV As Date
Private mvardQuantidadeNecess As Double
Private mvarcolPOAux As New Collection

Public Property Set colPOAux(ByVal vData As Collection)
    Set mvarcolPOAux = vData
End Property

Public Property Get colPOAux() As Collection
    Set colPOAux = mvarcolPOAux
End Property

Public Property Let dQuantidadeNecess(ByVal vData As Double)
    mvardQuantidadeNecess = vData
End Property

Public Property Get dQuantidadeNecess() As Double
    dQuantidadeNecess = mvardQuantidadeNecess
End Property

Public Property Let dtDataEntregaPV(ByVal vData As Date)
    mvardtDataEntregaPV = vData
End Property

Public Property Get dtDataEntregaPV() As Date
    dtDataEntregaPV = mvardtDataEntregaPV
End Property

Public Property Let iProduzLogo(ByVal vData As Integer)
    mvariProduzLogo = vData
End Property

Public Property Get iProduzLogo() As Integer
    iProduzLogo = mvariProduzLogo
End Property

Public Property Let lUltimoProxPO(ByVal vData As Long)
    mvarlUltimoProxPO = vData
End Property

Public Property Get lUltimoProxPO() As Long
    lUltimoProxPO = mvarlUltimoProxPO
End Property

Public Property Let iOrdem(ByVal vData As Integer)
    mvariOrdem = vData
End Property

Public Property Get iOrdem() As Integer
    iOrdem = mvariOrdem
End Property

Public Property Let iPrioridade(ByVal vData As Integer)
    mvariPrioridade = vData
End Property

Public Property Get iPrioridade() As Integer
    iPrioridade = mvariPrioridade
End Property

Public Property Let iAlterado(ByVal vData As Integer)
    mvariAlterado = vData
End Property

Public Property Get iAlterado() As Integer
    iAlterado = mvariAlterado
End Property

Public Property Let sDescErro(ByVal vData As String)
    mvarsDescErro = vData
End Property

Public Property Get sDescErro() As String
    sDescErro = mvarsDescErro
End Property

Public Property Let iFilialCli(ByVal vData As Integer)
    mvariFilialCli = vData
End Property

Public Property Get iFilialCli() As Integer
    iFilialCli = mvariFilialCli
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Set colPO(ByVal vData As Collection)
    Set mvarcolPO = vData
End Property

Public Property Get colPO() As Collection
    Set colPO = mvarcolPO
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Set objOP(ByVal vData As ClassOrdemDeProducao)
    Set mvarobjOP = vData
End Property

Public Property Get objOP() As ClassOrdemDeProducao
    Set objOP = mvarobjOP
End Property

Public Property Set objItemOP(ByVal vData As ClassItemOP)
    Set mvarobjItemOP = vData
End Property

Public Property Get objItemOP() As ClassItemOP
    Set objItemOP = mvarobjItemOP
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sCodOPOrigem(ByVal vData As String)
    mvarsCodOPOrigem = vData
End Property

Public Property Get sCodOPOrigem() As String
    sCodOPOrigem = mvarsCodOPOrigem
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sVersao(ByVal vData As String)
    mvarsVersao = vData
End Property

Public Property Get sVersao() As String
    sVersao = mvarsVersao
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let dtDataNecessidade(ByVal vData As Date)
    mvardtDataNecessidade = vData
End Property

Public Property Get dtDataNecessidade() As Date
    dtDataNecessidade = mvardtDataNecessidade
End Property

Public Property Let lCodGeracao(ByVal vData As Long)
    mvarlCodGeracao = vData
End Property

Public Property Get lCodGeracao() As Long
    lCodGeracao = mvarlCodGeracao
End Property
