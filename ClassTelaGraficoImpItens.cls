VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTelaGraficoImpItens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsText As String
Private mvarobjPicture As Object
Private mvarsFontName As String
Private mvarlLeft As Long
Private mvarlTop As Long
Private mvarlHeight As Long
Private mvarlWidth As Long
Private mvariBorderStyle As Integer
Private mvarlBackColor As Long
Private mvarlFontSize As Long
Private mvarlForeColor As Long
Private mvariTipo As Integer
Private mvariLegenda As Integer
Private mvarsDescricao As String

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iLegenda(ByVal vData As Integer)
    mvariLegenda = vData
End Property

Public Property Get iLegenda() As Integer
    iLegenda = mvariLegenda
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let lForeColor(ByVal vData As Long)
    mvarlForeColor = vData
End Property

Public Property Get lForeColor() As Long
    lForeColor = mvarlForeColor
End Property

Public Property Let lFontSize(ByVal vData As Long)
    mvarlFontSize = vData
End Property

Public Property Get lFontSize() As Long
    lFontSize = mvarlFontSize
End Property

Public Property Let lBackColor(ByVal vData As Long)
    mvarlBackColor = vData
End Property

Public Property Get lBackColor() As Long
    lBackColor = mvarlBackColor
End Property

Public Property Let iBorderStyle(ByVal vData As Integer)
    mvariBorderStyle = vData
End Property

Public Property Get iBorderStyle() As Integer
    iBorderStyle = mvariBorderStyle
End Property

Public Property Let lWidth(ByVal vData As Long)
    mvarlWidth = vData
End Property

Public Property Get lWidth() As Long
    lWidth = mvarlWidth
End Property

Public Property Let lHeight(ByVal vData As Long)
    mvarlHeight = vData
End Property

Public Property Get lHeight() As Long
    lHeight = mvarlHeight
End Property

Public Property Let lTop(ByVal vData As Long)
    mvarlTop = vData
End Property

Public Property Get lTop() As Long
    lTop = mvarlTop
End Property

Public Property Let lLeft(ByVal vData As Long)
    mvarlLeft = vData
End Property

Public Property Get lLeft() As Long
    lLeft = mvarlLeft
End Property

Public Property Let sText(ByVal vData As String)
    mvarsText = vData
End Property

Public Property Get sText() As String
    sText = mvarsText
End Property

Public Property Let objPicture(ByVal vData As Object)
    mvarobjPicture = vData
End Property

Public Property Get objPicture() As Object
    objPicture = mvarobjPicture
End Property

Public Property Let sFontName(ByVal vData As String)
    mvarsFontName = vData
End Property

Public Property Get sFontName() As String
    sFontName = mvarsFontName
End Property

