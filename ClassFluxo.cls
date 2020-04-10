VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFluxo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsFluxo As String
Private mvarsDescricao As String
Private mvardtDataBase As Date
Private mvardtDataFinal As Date
Private mvardtDataDadosReais As Date
Private mvarlFluxoId As Long
Private mvariFilialEmpresa As Integer
Private mvarlNumMovCta As Long  'contém o numero da ultima movimentação de conta corrente que foi encontrada durante a criação do fluxo. Impedirá que na obtenção de dados reais sejam levados em consideração movimentos que já existiam na criação do fluxo.
Private mvarlNumIntBaixaPag As Long 'contém o numero da ultima baixa de pagamento que foi encontrada durante a criação do fluxo. Impedirá que na obtenção de dados reais sejam levados em consideração baixas que já existiam na criação do fluxo.
Private mvarlNumIntBaixaRec As Long 'contém o numero da ultima baixa de recebimento que foi encontrada durante a criação do fluxo. Impedirá que na obtenção de dados reais sejam levados em consideração baixas que já existiam na criação do fluxo.
Private mvardtData As Date
Private mvarobjTela As Object

Public Property Get objTela() As Object
    Set objTela = mvarobjTela
End Property

Public Property Set objTela(ByVal vData As Object)
    Set mvarobjTela = vData
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let lNumIntBaixaRec(ByVal vData As Long)
    mvarlNumIntBaixaRec = vData
End Property

Public Property Get lNumIntBaixaRec() As Long
    lNumIntBaixaRec = mvarlNumIntBaixaRec
End Property

Public Property Let lNumIntBaixaPag(ByVal vData As Long)
    mvarlNumIntBaixaPag = vData
End Property

Public Property Get lNumIntBaixaPag() As Long
    lNumIntBaixaPag = mvarlNumIntBaixaPag
End Property

Public Property Let lNumMovCta(ByVal vData As Long)
    mvarlNumMovCta = vData
End Property

Public Property Get lNumMovCta() As Long
    lNumMovCta = mvarlNumMovCta
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
        iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lFluxoId(ByVal vData As Long)
Attribute lFluxoId.VB_Description = "identificador do fluxo de caixa"
    mvarlFluxoId = vData
End Property

Public Property Get lFluxoId() As Long
    lFluxoId = mvarlFluxoId
End Property

Public Property Let dtDataDadosReais(ByVal vData As Date)
Attribute dtDataDadosReais.VB_Description = "os valores que realmente ocorreram foram atualizados até esta data."
    mvardtDataDadosReais = vData
End Property

Public Property Get dtDataDadosReais() As Date
    dtDataDadosReais = mvardtDataDadosReais
End Property

Public Property Let dtDataFinal(ByVal vData As Date)
Attribute dtDataFinal.VB_Description = "data final do fluxo de caixa"
    mvardtDataFinal = vData
End Property

Public Property Get dtDataFinal() As Date
    dtDataFinal = mvardtDataFinal
End Property

Public Property Let dtDataBase(ByVal vData As Date)
Attribute dtDataBase.VB_Description = "data inicial do fluxo de caixa"
    mvardtDataBase = vData
End Property

Public Property Get dtDataBase() As Date
    dtDataBase = mvardtDataBase
End Property

Public Property Let sDescricao(ByVal vData As String)
Attribute sDescricao.VB_Description = "descrição do fluxo de caixa"
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sFluxo(ByVal vData As String)
Attribute sFluxo.VB_Description = "Nome do fluxo de caixa"
    mvarsFluxo = vData
End Property

Public Property Get sFluxo() As String
    sFluxo = mvarsFluxo
End Property



