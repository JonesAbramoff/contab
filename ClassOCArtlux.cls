VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassOCArtlux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvariFilialEmpresa As Integer
Private mvarsProduto As String
Private mvardQuantidade As Double
Private mvarsUsuCorte As String
Private mvardtDataIniCorte As Date
Private mvardHoraIniCorte As Double
Private mvardtDataFimCorte As Date
Private mvardHoraFimCorte As Double
Private mvarsUsuForro As String
Private mvardtDataIniForro As Date
Private mvardHoraIniForro As Double
Private mvardtDataFimForro As Date
Private mvardHoraFimForro As Double
Private mvardQuantidadeProd As Double
Private mvarsUsuMontagem As String
Private mvardtDataIniMontagem As Date
Private mvardtDataFimMontagem As Date
Private mvariExibir As Integer
Private mvariPrioridade As Integer
Private mvardQuantidadePreProd As Double
Private mvarcolItens As New Collection
Private mvardQuantidadeEst As Double
Private mvardQuantidadePV As Double
Private mvardEstoqueSeguranca As Double
Private mvarsProdutoDesc As String
Private mvariManual As Integer
Private mvarsUsuManual As String
Private mvardtDataManual As Date
Private mvarsTipoCouro As String
Private mvardQuantidadeNecess1 As Double
Private mvardQuantidadeNecess2 As Double
Private mvarsGrupo As String

Public Property Let sGrupo(ByVal vData As String)
    mvarsGrupo = vData
End Property

Public Property Get sGrupo() As String
    sGrupo = mvarsGrupo
End Property

Public Property Let dQuantidadeNecess1(ByVal vData As Double)
    mvardQuantidadeNecess1 = vData
End Property

Public Property Get dQuantidadeNecess1() As Double
    dQuantidadeNecess1 = mvardQuantidadeNecess1
End Property

Public Property Let dQuantidadeNecess2(ByVal vData As Double)
    mvardQuantidadeNecess2 = vData
End Property

Public Property Get dQuantidadeNecess2() As Double
    dQuantidadeNecess2 = mvardQuantidadeNecess2
End Property

Public Property Let sTipoCouro(ByVal vData As String)
    mvarsTipoCouro = vData
End Property

Public Property Get sTipoCouro() As String
    sTipoCouro = mvarsTipoCouro
End Property

Public Property Let dtDataManual(ByVal vData As Date)
    mvardtDataManual = vData
End Property

Public Property Get dtDataManual() As Date
    dtDataManual = mvardtDataManual
End Property

Public Property Let sUsuManual(ByVal vData As String)
    mvarsUsuManual = vData
End Property

Public Property Get sUsuManual() As String
    sUsuManual = mvarsUsuManual
End Property

Public Property Let iManual(ByVal vData As Integer)
    mvariManual = vData
End Property

Public Property Get iManual() As Integer
    iManual = mvariManual
End Property

Public Property Let sProdutoDesc(ByVal vData As String)
    mvarsProdutoDesc = vData
End Property

Public Property Get sProdutoDesc() As String
    sProdutoDesc = mvarsProdutoDesc
End Property

Public Property Let dEstoqueSeguranca(ByVal vData As Double)
    mvardEstoqueSeguranca = vData
End Property

Public Property Get dEstoqueSeguranca() As Double
    dEstoqueSeguranca = mvardEstoqueSeguranca
End Property

Public Property Let dQuantidadePV(ByVal vData As Double)
    mvardQuantidadePV = vData
End Property

Public Property Get dQuantidadePV() As Double
    dQuantidadePV = mvardQuantidadePV
End Property

Public Property Let dQuantidadeEst(ByVal vData As Double)
    mvardQuantidadeEst = vData
End Property

Public Property Get dQuantidadeEst() As Double
    dQuantidadeEst = mvardQuantidadeEst
End Property

Public Property Set colItens(ByVal vData As Object)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Let dQuantidadePreProd(ByVal vData As Double)
    mvardQuantidadePreProd = vData
End Property

Public Property Get dQuantidadePreProd() As Double
    dQuantidadePreProd = mvardQuantidadePreProd
End Property

Public Property Let iPrioridade(ByVal vData As Integer)
    mvariPrioridade = vData
End Property

Public Property Get iPrioridade() As Integer
    iPrioridade = mvariPrioridade
End Property

Public Property Let dtDataIniMontagem(ByVal vData As Date)
    mvardtDataIniMontagem = vData
End Property

Public Property Get dtDataIniMontagem() As Date
    dtDataIniMontagem = mvardtDataIniMontagem
End Property

Public Property Let dtDataFimMontagem(ByVal vData As Date)
    mvardtDataFimMontagem = vData
End Property

Public Property Get dtDataFimMontagem() As Date
    dtDataFimMontagem = mvardtDataFimMontagem
End Property

Public Property Let iExibir(ByVal vData As Integer)
    mvariExibir = vData
End Property

Public Property Get iExibir() As Integer
    iExibir = mvariExibir
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sUsuCorte(ByVal vData As String)
    mvarsUsuCorte = vData
End Property

Public Property Get sUsuCorte() As String
    sUsuCorte = mvarsUsuCorte
End Property

Public Property Let dtDataIniCorte(ByVal vData As Date)
    mvardtDataIniCorte = vData
End Property

Public Property Get dtDataIniCorte() As Date
    dtDataIniCorte = mvardtDataIniCorte
End Property

Public Property Let dHoraIniCorte(ByVal vData As Double)
    mvardHoraIniCorte = vData
End Property

Public Property Get dHoraIniCorte() As Double
    dHoraIniCorte = mvardHoraIniCorte
End Property

Public Property Let dtDataFimCorte(ByVal vData As Date)
    mvardtDataFimCorte = vData
End Property

Public Property Get dtDataFimCorte() As Date
    dtDataFimCorte = mvardtDataFimCorte
End Property

Public Property Let dHoraFimCorte(ByVal vData As Double)
    mvardHoraFimCorte = vData
End Property

Public Property Get dHoraFimCorte() As Double
    dHoraFimCorte = mvardHoraFimCorte
End Property

Public Property Let sUsuForro(ByVal vData As String)
    mvarsUsuForro = vData
End Property

Public Property Get sUsuForro() As String
    sUsuForro = mvarsUsuForro
End Property

Public Property Let dtDataIniForro(ByVal vData As Date)
    mvardtDataIniForro = vData
End Property

Public Property Get dtDataIniForro() As Date
    dtDataIniForro = mvardtDataIniForro
End Property

Public Property Let dHoraIniForro(ByVal vData As Double)
    mvardHoraIniForro = vData
End Property

Public Property Get dHoraIniForro() As Double
    dHoraIniForro = mvardHoraIniForro
End Property

Public Property Let dtDataFimForro(ByVal vData As Date)
    mvardtDataFimForro = vData
End Property

Public Property Get dtDataFimForro() As Date
    dtDataFimForro = mvardtDataFimForro
End Property

Public Property Let dHoraFimForro(ByVal vData As Double)
    mvardHoraFimForro = vData
End Property

Public Property Get dHoraFimForro() As Double
    dHoraFimForro = mvardHoraFimForro
End Property

Public Property Let dQuantidadeProd(ByVal vData As Double)
    mvardQuantidadeProd = vData
End Property

Public Property Get dQuantidadeProd() As Double
    dQuantidadeProd = mvardQuantidadeProd
End Property

Public Property Let sUsuMontagem(ByVal vData As String)
    mvarsUsuMontagem = vData
End Property

Public Property Get sUsuMontagem() As String
    sUsuMontagem = mvarsUsuMontagem
End Property

Public Sub Copiar(ByVal objOC As ClassOCArtlux)

Dim objProd As ClassOCProdArtlux
Dim objProdAux As ClassOCProdArtlux

    mvarlNumIntDoc = objOC.lNumIntDoc
    mvariFilialEmpresa = objOC.iFilialEmpresa
    mvarsProduto = objOC.sProduto
    mvardQuantidade = objOC.dQuantidade
    mvarsUsuCorte = objOC.sUsuCorte
    mvardtDataIniCorte = objOC.dtDataIniCorte
    mvardHoraIniCorte = objOC.dHoraIniCorte
    mvardtDataFimCorte = objOC.dtDataFimCorte
    mvardHoraFimCorte = objOC.dHoraFimCorte
    mvarsUsuForro = objOC.sUsuForro
    mvardtDataIniForro = objOC.dtDataIniForro
    mvardHoraIniForro = objOC.dHoraIniForro
    mvardtDataFimForro = objOC.dtDataFimForro
    mvardHoraFimForro = objOC.dHoraFimForro
    mvardQuantidadeProd = objOC.dQuantidadeProd
    mvarsUsuMontagem = objOC.sUsuMontagem
    mvardtDataIniMontagem = objOC.dtDataIniMontagem
    mvardtDataFimMontagem = objOC.dtDataFimMontagem
    mvariExibir = objOC.iExibir
    mvariPrioridade = objOC.iPrioridade
    mvardQuantidadePreProd = objOC.dQuantidadePreProd
    mvardQuantidadeEst = objOC.dQuantidadeEst
    mvardQuantidadePV = objOC.dQuantidadePV
    mvardEstoqueSeguranca = objOC.dEstoqueSeguranca
    mvarsProdutoDesc = objOC.sProdutoDesc

    For Each objProd In objOC.colItens
        Set objProdAux = New ClassOCProdArtlux
        Call objProdAux.Copiar(objProd)
        mvarcolItens.Add objProdAux
    Next

End Sub

Private Sub Class_Initialize()
    'Sumiu a etapa de corte do forro, já virá como se estivesse concluida
    mvarsUsuForro = gsUsuario
    mvardtDataIniForro = Date
    mvardHoraIniForro = CDbl(Time)
    mvardtDataFimForro = Date
    mvardHoraFimForro = CDbl(Time)
    
    mvardtDataIniCorte = DATA_NULA
    mvardtDataFimCorte = DATA_NULA
    
    mvardtDataIniMontagem = Date
    mvardtDataFimMontagem = Date
End Sub
