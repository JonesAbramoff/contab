VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMargContr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlCodCliente As Long
Private mvarlNumIntRel As Long
Private mvariCodFilial As Integer
Private mvariFilialFaturamento As Integer
Private mvarsProduto As String
Private mvardQuantidade As Double
Private mvarcolDVVLin As New Collection
Private mvarcolDVVLinCol As New Collection
Private mvarcolPlanMargContrCol As New Collection
Private mvarcolPlanMargContrLin As New Collection
Private mvarcolPlanMargContrLinCol As New Collection
Private mvariCodVendedor As Integer
Private mvariTabelaPreco As Integer

Public iColunaCli As Integer
Public iColunaPraticado As Integer
Public iColunaSimulacao As Integer
Public iLinhaPercMarg As Integer
Public iLinhaPercLucro As Integer
Public iLinhaValorMarg As Integer
Public iLinhaPrazo As Integer

Public objTelaComissoes As Object
Public colComissoes As Collection
Public dPrecoComissoes As Double
Public dtDataCF As Date
Public sDVVTotal1 As String
Public sDVVTotal2 As String
Public sDVVTotal3 As String
Public dTaxaDescPadrao As Double, dTaxaValFut As Double, dDiasValFut As Double

Public dPrecoPraticadoInformado As Double

Public iUltFilialEmpresa As Integer
Public lUltCliente As Long

Public sUltNomeRedCli As String
Public sUnidadeMed As String

'Valores armazenados para a última pesquisa
Public dUltPrecoCli As Double
Public dUltPrecoPraticado As Double
Public dUltPercMargPrat As Double
Public dUltPercLucroPrat As Double
Public dUltValorMargPrat As Double

Public Property Let iTabelaPreco(ByVal vData As Integer)
    mvariTabelaPreco = vData
End Property

Public Property Get iTabelaPreco() As Integer
    iTabelaPreco = mvariTabelaPreco
End Property

Public Property Let iCodVendedor(ByVal vData As Integer)
    mvariCodVendedor = vData
End Property

Public Property Get iCodVendedor() As Integer
    iCodVendedor = mvariCodVendedor
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let iFilialFaturamento(ByVal vData As Integer)
    mvariFilialFaturamento = vData
End Property

Public Property Get iFilialFaturamento() As Integer
    iFilialFaturamento = mvariFilialFaturamento
End Property

Public Property Let iCodFilial(ByVal vData As Integer)
    mvariCodFilial = vData
End Property

Public Property Get iCodFilial() As Integer
    iCodFilial = mvariCodFilial
End Property

Public Property Let lNumIntRel(ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel() As Long
    lNumIntRel = mvarlNumIntRel
End Property

Public Property Let lCodCliente(ByVal vData As Long)
    mvarlCodCliente = vData
End Property

Public Property Get lCodCliente() As Long
    lCodCliente = mvarlCodCliente
End Property

Public Property Set colDVVLin(ByVal vData As Object)
    Set mvarcolDVVLin = vData
End Property

Public Property Get colDVVLin() As Collection
    Set colDVVLin = mvarcolDVVLin
End Property

Public Property Set colDVVLinCol(ByVal vData As Object)
    Set mvarcolDVVLinCol = vData
End Property

Public Property Get colDVVLinCol() As Collection
    Set colDVVLinCol = mvarcolDVVLinCol
End Property

Public Property Set colPlanMargContrCol(ByVal vData As Object)
    Set mvarcolPlanMargContrCol = vData
End Property

Public Property Get colPlanMargContrCol() As Collection
    Set colPlanMargContrCol = mvarcolPlanMargContrCol
End Property

Public Property Set colPlanMargContrLin(ByVal vData As Object)
    Set mvarcolPlanMargContrLin = vData
End Property

Public Property Get colPlanMargContrLin() As Collection
    Set colPlanMargContrLin = mvarcolPlanMargContrLin
End Property

Public Property Set colPlanMargContrLinCol(ByVal vData As Object)
    Set mvarcolPlanMargContrLinCol = vData
End Property

Public Property Get colPlanMargContrLinCol() As Collection
    Set colPlanMargContrLinCol = mvarcolPlanMargContrLinCol
End Property

Public Property Get IndAnalise(ByVal iLinha As Integer, ByVal iColuna As Integer) As Integer
    IndAnalise = ((iLinha - 1) * colPlanMargContrCol.Count) + iColuna
End Property

Public Property Get IndDVV(ByVal iLinha As Integer, ByVal iColuna As Integer) As Integer
    IndDVV = ((iLinha - 1) * (MAX_NUM_FORMULAS_DVV)) + iColuna
End Property

