VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdmRelSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Classe para passagem de parametros para a tela RelSelecionar"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsCodRel As String
Private mvariCancela As Integer

Public Property Let iCancela(ByVal vData As Integer)
    mvariCancela = vData
    Exit Property
End Property

Public Property Get iCancela() As Integer
    iCancela = mvariCancela
    Exit Property
End Property

Public Property Let sCodRel(ByVal vData As String)
    mvarsCodRel = vData
    Exit Property
End Property

Public Property Get sCodRel() As String
    sCodRel = mvarsCodRel
    Exit Property
End Property





