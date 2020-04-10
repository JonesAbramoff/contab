VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGradeLinCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarcolGradeCategorias As New Collection
Private mvariColunas As Integer
Private mvariLinhas As Integer

Public Property Let iLinhas(ByVal vData As Integer)
    mvariLinhas = vData
End Property

Public Property Get iLinhas() As Integer
    iLinhas = mvariLinhas
End Property

Public Property Let iColunas(ByVal vData As Integer)
    mvariColunas = vData
End Property

Public Property Get iColunas() As Integer
    iColunas = mvariColunas
End Property

Public Property Set colGradeCategorias(ByVal vData As Object)
    Set mvarcolGradeCategorias = vData
End Property

Public Property Get colGradeCategorias() As Collection
    Set colGradeCategorias = mvarcolGradeCategorias
End Property
