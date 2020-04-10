VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPlanilhaExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarsNomePlanilha As String
Private mvarsNomeGrafico As String
Private mvarsTituloGrafico As String
Private mvarcolColunas As New Collection
Private mvarlTipoGrafico As Long
Private mvarlPosicaoLabelsX As Long
Private mvarlOrientacaoLabelsX As Long
Private mvarlDadosLabelOrientacao As Long
Private mvarvPlotLinhaColuna As Variant
Private mvariEixosGrafico As Integer
Private mvarcolCabecalhoRodape As Collection
Private mvarlPosicaoLegenda As Long
Private mvarsFonteNome As String
Private mvardFonteTamanho As Double
Private mvariModoImpressao As Integer
Private mvarsNomeDiretorioArq As String
Private mvariModoGravacao As Integer

Public Property Let iModoGravacao(ByVal vData As Integer)
    mvariModoGravacao = vData
End Property

Public Property Get iModoGravacao() As Integer
    iModoGravacao = mvariModoGravacao
End Property

Public Property Let sNomeDiretorioArq(ByVal vData As String)
    mvarsNomeDiretorioArq = vData
End Property

Public Property Get sNomeDiretorioArq() As String
    sNomeDiretorioArq = mvarsNomeDiretorioArq
End Property

Public Property Let iModoImpressao(ByVal vData As Integer)
    mvariModoImpressao = vData
End Property

Public Property Get iModoImpressao() As Integer
    iModoImpressao = mvariModoImpressao
End Property

Public Property Let sFonteNome(ByVal vData As String)
    mvarsFonteNome = vData
End Property

Public Property Get sFonteNome() As String
    sFonteNome = mvarsFonteNome
End Property

Public Property Let dFonteTamanho(ByVal vData As Double)
    mvardFonteTamanho = vData
End Property

Public Property Get dFonteTamanho() As Double
    dFonteTamanho = mvardFonteTamanho
End Property

Public Property Let lPosicaoLegenda(ByVal vData As Long)
    mvarlPosicaoLegenda = vData
End Property

Public Property Get lPosicaoLegenda() As Long
    lPosicaoLegenda = mvarlPosicaoLegenda
End Property

Public Property Set colCabecalhoRodape(ByVal vData As Collection)
    Set mvarcolCabecalhoRodape = vData
End Property

Public Property Get colCabecalhoRodape() As Collection
    Set colCabecalhoRodape = mvarcolCabecalhoRodape
End Property

Public Property Let iEixosGrafico(ByVal vData As Integer)
    mvariEixosGrafico = vData
End Property

Public Property Get iEixosGrafico() As Integer
    iEixosGrafico = mvariEixosGrafico
End Property

Public Property Let vPlotLinhaColuna(ByVal vData As Variant)
    mvarvPlotLinhaColuna = vData
End Property

Public Property Set vPlotLinhaColuna(ByVal vData As Variant)
    Set mvarvPlotLinhaColuna = vData
End Property

Public Property Get vPlotLinhaColuna() As Variant
    If IsObject(mvarvPlotLinhaColuna) Then
        Set vPlotLinhaColuna = mvarvPlotLinhaColuna
    Else
        vPlotLinhaColuna = mvarvPlotLinhaColuna
    End If
End Property

Public Property Let lDadosLabelOrientacao(ByVal vData As Long)
    mvarlDadosLabelOrientacao = vData
End Property

Public Property Get lDadosLabelOrientacao() As Long
    lDadosLabelOrientacao = mvarlDadosLabelOrientacao
End Property

Public Property Let lLabelsXOrientacao(ByVal vData As Long)
    mvarlOrientacaoLabelsX = vData
End Property

Public Property Get lLabelsXOrientacao() As Long
    lLabelsXOrientacao = mvarlOrientacaoLabelsX
End Property

Public Property Let lLabelsXPosicao(ByVal vData As Long)
    mvarlPosicaoLabelsX = vData
End Property

Public Property Get lLabelsXPosicao() As Long
    lLabelsXPosicao = mvarlPosicaoLabelsX
End Property

Public Property Let lTipoGrafico(ByVal vData As Long)
    mvarlTipoGrafico = vData
End Property

Public Property Get lTipoGrafico() As Long
    lTipoGrafico = mvarlTipoGrafico
End Property

Public Property Set colColunas(ByVal vData As Collection)
    Set mvarcolColunas = vData
End Property

Public Property Get colColunas() As Collection
    Set colColunas = mvarcolColunas
End Property

Public Property Let sTituloGrafico(ByVal vData As String)
    mvarsTituloGrafico = vData
End Property

Public Property Get sTituloGrafico() As String
    sTituloGrafico = mvarsTituloGrafico
End Property

Public Property Let sNomeGrafico(ByVal vData As String)
    mvarsNomeGrafico = vData
End Property

Public Property Get sNomeGrafico() As String
    sNomeGrafico = mvarsNomeGrafico
End Property

Public Property Let sNomePlanilha(ByVal vData As String)
    mvarsNomePlanilha = vData
End Property

Public Property Get sNomePlanilha() As String
    sNomePlanilha = mvarsNomePlanilha
End Property

