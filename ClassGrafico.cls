VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGrafico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarChartType As Integer
Private mvarTitleText As String
Private mvarFootNote As String
Private mvarcolcolItensGrafico As New Collection

Public Property Set colcolItensGrafico(ByVal vData As Collection)
    Set mvarcolcolItensGrafico = vData
End Property

Public Property Get colcolItensGrafico() As Collection
    Set colcolItensGrafico = mvarcolcolItensGrafico
End Property

Public Property Let FootNote(ByVal vData As String)
    mvarFootNote = vData
End Property

Public Property Get FootNote() As String
    FootNote = mvarFootNote
End Property

Public Property Let TitleText(ByVal vData As String)
Attribute TitleText.VB_Description = "Título do gráfico"
    mvarTitleText = vData
End Property

Public Property Get TitleText() As String
    TitleText = mvarTitleText
End Property

Public Property Let ChartType(ByVal vData As Integer)
Attribute ChartType.VB_Description = "O tipo (aparência) que o gráfico deve possuir.\r\nEx.: Torta, Filas em 3D, ..."
    mvarChartType = vData
End Property

Public Property Get ChartType() As Integer
    ChartType = mvarChartType
End Property
