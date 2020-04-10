VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPRJFluxoFinanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvardtData As String
Private mvarsDescricao As String
Private mvarobjPRJCR As New ClassPRJCR
Private mvarobjPRJInfo As New ClassProjetoInfo
Private mvardDespesaPrev As Double
Private mvardDespesaReal As Double
Private mvardDespesaPlan As Double
Private mvardReceitaPrev As Double
Private mvardReceitaReal As Double
Private mvardReceitaPlan As Double
Private mvarobjProjeto As New ClassProjetos
Private mvarobjEtapa As New ClassPRJEtapas
Private mvarobjOrigem As Object
Private mvarsDocOrigem As String
Private mvarsNomeTela As String
Private mvardCustoProdPrevInfo As Double
Private mvardCustoProdPrevCalc As Double
Private mvardCustoProdRealInfo As Double
Private mvardCustoProdRealCalc As Double
Private mvarsObs As String
Private mvarsNatureza As String
Private mvarsDescNatureza As String

Public Property Let sDescNatureza(ByVal vData As String)
    mvarsDescNatureza = vData
End Property

Public Property Get sDescNatureza() As String
    sDescNatureza = mvarsDescNatureza
End Property

Public Property Let sNatureza(ByVal vData As String)
    mvarsNatureza = vData
End Property

Public Property Get sNatureza() As String
    sNatureza = mvarsNatureza
End Property

Public Property Let sObs(ByVal vData As String)
    mvarsObs = vData
End Property

Public Property Get sObs() As String
    sObs = mvarsObs
End Property

Public Property Let dCustoProdRealCalc(ByVal vData As Double)
    mvardCustoProdRealCalc = vData
End Property

Public Property Get dCustoProdRealCalc() As Double
    dCustoProdRealCalc = mvardCustoProdRealCalc
End Property

Public Property Let dCustoProdRealInfo(ByVal vData As Double)
    mvardCustoProdRealInfo = vData
End Property

Public Property Get dCustoProdRealInfo() As Double
    dCustoProdRealInfo = mvardCustoProdRealInfo
End Property
Public Property Let dCustoProdPrevCalc(ByVal vData As Double)
    mvardCustoProdPrevCalc = vData
End Property

Public Property Get dCustoProdPrevCalc() As Double
    dCustoProdPrevCalc = mvardCustoProdPrevCalc
End Property

Public Property Let dCustoProdPrevInfo(ByVal vData As Double)
    mvardCustoProdPrevInfo = vData
End Property

Public Property Get dCustoProdPrevInfo() As Double
    dCustoProdPrevInfo = mvardCustoProdPrevInfo
End Property

Public Property Let sNomeTela(ByVal vData As String)
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property

Public Property Let sDocOrigem(ByVal vData As String)
    mvarsDocOrigem = vData
End Property

Public Property Get sDocOrigem() As String
    sDocOrigem = mvarsDocOrigem
End Property

Public Property Set objOrigem(ByVal vData As Object)
    Set mvarobjOrigem = vData
End Property

Public Property Get objOrigem() As Object
    Set objOrigem = mvarobjOrigem
End Property

Public Property Set objEtapa(ByVal vData As ClassPRJEtapas)
    Set mvarobjEtapa = vData
End Property

Public Property Get objEtapa() As ClassPRJEtapas
    Set objEtapa = mvarobjEtapa
End Property

Public Property Set objProjeto(ByVal vData As ClassProjetos)
    Set mvarobjProjeto = vData
End Property

Public Property Get objProjeto() As ClassProjetos
    Set objProjeto = mvarobjProjeto
End Property

Public Property Let dReceitaPlan(ByVal vData As Double)
    mvardReceitaPlan = vData
End Property

Public Property Get dReceitaPlan() As Double
    dReceitaPlan = mvardReceitaPlan
End Property

Public Property Let dReceitaReal(ByVal vData As Double)
    mvardReceitaReal = vData
End Property

Public Property Get dReceitaReal() As Double
    dReceitaReal = mvardReceitaReal
End Property

Public Property Let dReceitaPrev(ByVal vData As Double)
    mvardReceitaPrev = vData
End Property

Public Property Get dReceitaPrev() As Double
    dReceitaPrev = mvardReceitaPrev
End Property
Public Property Let dDespesaPlan(ByVal vData As Double)
    mvardDespesaPlan = vData
End Property

Public Property Get dDespesaPlan() As Double
    dDespesaPlan = mvardDespesaPlan
End Property

Public Property Let dDespesaReal(ByVal vData As Double)
    mvardDespesaReal = vData
End Property

Public Property Get dDespesaReal() As Double
    dDespesaReal = mvardDespesaReal
End Property

Public Property Let dDespesaPrev(ByVal vData As Double)
    mvardDespesaPrev = vData
End Property

Public Property Get dDespesaPrev() As Double
    dDespesaPrev = mvardDespesaPrev
End Property

Public Property Let dtData(ByVal vData As String)
    mvardtData = vData
End Property

Public Property Get dtData() As String
    dtData = mvardtData
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Set objPRJCR(ByVal vData As ClassPRJCR)
    Set mvarobjPRJCR = vData
End Property

Public Property Get objPRJCR() As ClassPRJCR
    Set objPRJCR = mvarobjPRJCR
End Property

Public Property Set objPRJInfo(ByVal vData As ClassProjetoInfo)
    Set mvarobjPRJInfo = vData
End Property

Public Property Get objPRJInfo() As ClassProjetoInfo
    Set objPRJInfo = mvarobjPRJInfo
End Property
