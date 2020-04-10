VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassChaveFilial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodFilial As Integer
Private mvarlCodPai As Long
Private mvarsNome As String

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let lCodPai(ByVal vData As Long)
    mvarlCodPai = vData
End Property

Public Property Get lCodPai() As Long
    lCodPai = mvarlCodPai
End Property

Public Property Let iCodFilial(ByVal vData As Integer)
    mvariCodFilial = vData
End Property

Public Property Get iCodFilial() As Integer
    iCodFilial = mvariCodFilial
End Property



