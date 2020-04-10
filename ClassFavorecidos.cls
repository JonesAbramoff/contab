VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFavorecidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvarsNome As String
Private mvariInativo As Integer

Public Property Let iInativo(ByVal vData As Integer)
    mvariInativo = vData
End Property

Public Property Get iInativo() As Integer
    iInativo = mvariInativo
End Property

Public Property Let sNome(ByVal vData As String)
Attribute sNome.VB_Description = "Nome do Favorecido"
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let iCodigo(ByVal vData As Integer)
Attribute iCodigo.VB_Description = "Identificador do Favorecido"
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property



