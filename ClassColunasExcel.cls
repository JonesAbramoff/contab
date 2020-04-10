VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassColunasExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarcolCelulas As New Collection
Private mvariParticipaGrafico As Integer
Private mvarlDataLabels As Long
Private mvarlDataLabelsOrientacao As Long
Private mvarlTipoGraficoColuna As Long
Private mvardLarguraColuna As Double

Public Property Let dLarguraColuna(ByVal vData As Double)
    mvardLarguraColuna = vData
End Property

Public Property Get dLarguraColuna() As Double
    dLarguraColuna = mvardLarguraColuna
End Property

Public Property Let lDataLabelsOrientacao(ByVal vData As Long)
    mvarlDataLabelsOrientacao = vData
End Property

Public Property Get lDataLabelsOrientacao() As Long
    lDataLabelsOrientacao = mvarlDataLabelsOrientacao
End Property

Public Property Let lDataLabels(ByVal vData As Long)
    mvarlDataLabels = vData
End Property

Public Property Get lDataLabels() As Long
    lDataLabels = mvarlDataLabels
End Property

Public Property Let lTipoGraficoColuna(ByVal vData As Long)
    mvarlTipoGraficoColuna = vData
End Property

Public Property Get lTipoGraficoColuna() As Long
    lTipoGraficoColuna = mvarlTipoGraficoColuna
End Property

Public Property Let iParticipaGrafico(ByVal vData As Integer)
    mvariParticipaGrafico = vData
End Property

Public Property Get iParticipaGrafico() As Integer
    iParticipaGrafico = mvariParticipaGrafico
End Property

Public Property Set colCelulas(ByVal vData As Collection)
    Set mvarcolCelulas = vData
End Property

Public Property Get colCelulas() As Collection
    Set colCelulas = mvarcolCelulas
End Property

