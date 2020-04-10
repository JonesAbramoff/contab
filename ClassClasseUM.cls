VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassClasseUM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariClasse As Integer
Private mvarsDescricao As String
Private mvarsSiglaUMBase As String

Public Property Let sSiglaUMBase(ByVal vData As String)
    mvarsSiglaUMBase = vData
End Property

Public Property Get sSiglaUMBase() As String
    sSiglaUMBase = mvarsSiglaUMBase
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iClasse(ByVal vData As Integer)
    mvariClasse = vData
End Property

Public Property Get iClasse() As Integer
    iClasse = mvariClasse
End Property



