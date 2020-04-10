VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTelaGrafico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsNomeTela As String
Private mvarsTextoImpressao As String
Private mvarcolItens As New Collection
Private mvarcolBotoes As New Collection
Private mvarobjTela As Object
Private mvariTamanhoDia As Integer
Private mvariModal As Integer
Private mvariAtualizaRetornoClick As Integer
Private mvarsNomeFuncAtualiza As String
Private mvariZOOM As Integer
Private mvariAlturaDia As Integer
Private mvardtDataInicio As Date
Private mvariNumDias As Integer
Private mvarcolParametros As New Collection
Private mvariAlteraQtdLinhasPara As Integer
Private mvariExibirGridAux As Integer
Private mvarsNomeFuncTrataGrid As String
Private mvariCadaEtapaUmaLinha As Integer
Private mvarcolParametrosTrataGrid As New Collection
Private mvarsNomeArqFigura As String
Private mvariNumFiguras As Integer

Public Property Let iNumFiguras(ByVal vData As Integer)
    mvariNumFiguras = vData
End Property

Public Property Get iNumFiguras() As Integer
    iNumFiguras = mvariNumFiguras
End Property

Public Property Let sNomeArqFigura(ByVal vData As String)
    mvarsNomeArqFigura = vData
End Property

Public Property Get sNomeArqFigura() As String
    sNomeArqFigura = mvarsNomeArqFigura
End Property

Public Property Set colParametrosTrataGrid(ByVal vData As Collection)
    Set mvarcolParametrosTrataGrid = vData
End Property

Public Property Get colParametrosTrataGrid() As Collection
    Set colParametrosTrataGrid = mvarcolParametrosTrataGrid
End Property

Public Property Let iCadaEtapaUmaLinha(ByVal vData As Integer)
    mvariCadaEtapaUmaLinha = vData
End Property

Public Property Get iCadaEtapaUmaLinha() As Integer
    iCadaEtapaUmaLinha = mvariCadaEtapaUmaLinha
End Property

Public Property Let sNomeFuncTrataGrid(ByVal vData As String)
    mvarsNomeFuncTrataGrid = vData
End Property

Public Property Get sNomeFuncTrataGrid() As String
    sNomeFuncTrataGrid = mvarsNomeFuncTrataGrid
End Property

Public Property Let iExibirGridAux(ByVal vData As Integer)
    mvariExibirGridAux = vData
End Property

Public Property Get iExibirGridAux() As Integer
    iExibirGridAux = mvariExibirGridAux
End Property

Public Property Let iAlteraQtdLinhasPara(ByVal vData As Integer)
    mvariAlteraQtdLinhasPara = vData
End Property

Public Property Get iAlteraQtdLinhasPara() As Integer
    iAlteraQtdLinhasPara = mvariAlteraQtdLinhasPara
End Property

Public Property Set colParametros(ByVal vData As Collection)
    Set mvarcolParametros = vData
End Property

Public Property Get colParametros() As Collection
    Set colParametros = mvarcolParametros
End Property

Public Property Let iNumDias(ByVal vData As Integer)
    mvariNumDias = vData
End Property

Public Property Get iNumDias() As Integer
    iNumDias = mvariNumDias
End Property

Public Property Let dtDataInicio(ByVal vData As Date)
    mvardtDataInicio = vData
End Property

Public Property Get dtDataInicio() As Date
    dtDataInicio = mvardtDataInicio
End Property

Public Property Let iAlturaDia(ByVal vData As Integer)
    mvariAlturaDia = vData
End Property

Public Property Get iAlturaDia() As Integer
    iAlturaDia = mvariAlturaDia
End Property

Public Property Let iZOOM(ByVal vData As Integer)
    mvariZOOM = vData
End Property

Public Property Get iZOOM() As Integer
    iZOOM = mvariZOOM
End Property

Public Property Let sNomeFuncAtualiza(ByVal vData As String)
    mvarsNomeFuncAtualiza = vData
End Property

Public Property Get sNomeFuncAtualiza() As String
    sNomeFuncAtualiza = mvarsNomeFuncAtualiza
End Property

Public Property Let iAtualizaRetornoClick(ByVal vData As Integer)
    mvariAtualizaRetornoClick = vData
End Property

Public Property Get iAtualizaRetornoClick() As Integer
    iAtualizaRetornoClick = mvariAtualizaRetornoClick
End Property

Public Property Let iModal(ByVal vData As Integer)
    mvariModal = vData
End Property

Public Property Get iModal() As Integer
    iModal = mvariModal
End Property

Public Property Let iTamanhoDia(ByVal vData As Integer)
    mvariTamanhoDia = vData
End Property

Public Property Get iTamanhoDia() As Integer
    iTamanhoDia = mvariTamanhoDia
End Property

Public Property Set colItens(ByVal vData As Collection)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Set colBotoes(ByVal vData As Collection)
    Set mvarcolBotoes = vData
End Property

Public Property Get colBotoes() As Collection
    Set colBotoes = mvarcolBotoes
End Property

Public Property Set objTela(ByVal vData As Object)
    Set mvarobjTela = vData
End Property

Public Property Get objTela() As Object
    Set objTela = mvarobjTela
End Property

Public Property Let sNomeTela(ByVal vData As String)
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property

Public Property Let sTextoImpressao(ByVal vData As String)
    mvarsTextoImpressao = vData
End Property

Public Property Get sTextoImpressao() As String
    sTextoImpressao = mvarsTextoImpressao
End Property

Private Sub Class_Initialize()
    mvardtDataInicio = DATA_NULA
End Sub
