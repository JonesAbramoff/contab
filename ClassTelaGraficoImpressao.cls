VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTelaGraficoImpressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarcolItens As New Collection
Private mvarsNome As String
Private mvarsTexto As String
Private mvarsTexto2 As String
Private mvarsFontName As String
Private mvarlBackColor As Long
Private mvarlFontSize As Long
Private mvarlForeColor As Long
Private mvarlFontSizeNome As Long
Private mvarlFontSizeTexto As Long
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

Public Property Let lFontSizeTexto(ByVal vData As Long)
    mvarlFontSizeTexto = vData
End Property

Public Property Get lFontSizeTexto() As Long
    lFontSizeTexto = mvarlFontSizeTexto
End Property

Public Property Let lFontSizeNome(ByVal vData As Long)
    mvarlFontSizeNome = vData
End Property

Public Property Get lFontSizeNome() As Long
    lFontSizeNome = mvarlFontSizeNome
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

Public Property Let sTexto(ByVal vData As String)
    mvarsTexto = vData
End Property

Public Property Get sTexto() As String
    sTexto = mvarsTexto
End Property

Public Property Let sTexto2(ByVal vData As String)
    mvarsTexto2 = vData
End Property

Public Property Get sTexto2() As String
    sTexto2 = mvarsTexto2
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Set colItens(ByVal vData As Collection)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Let sFontName(ByVal vData As String)
    mvarsFontName = vData
End Property

Public Property Get sFontName() As String
    sFontName = mvarsFontName
End Property
