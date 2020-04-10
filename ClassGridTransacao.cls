VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassGridTransacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsNomeGrid As String
Private mvariNumLinhas As Integer

Public Property Let iNumLinhas(ByVal vData As Integer)
Attribute iNumLinhas.VB_Description = "Numero de Linhas do grid que estão preenchidas"
    mvariNumLinhas = vData
End Property

Public Property Get iNumLinhas() As Integer
    iNumLinhas = mvariNumLinhas
End Property

Public Property Let sNomeGrid(ByVal vData As String)
Attribute sNomeGrid.VB_Description = "Nome do Grid"
    mvarsNomeGrid = vData
End Property

Public Property Get sNomeGrid() As String
    sNomeGrid = mvarsNomeGrid
End Property



