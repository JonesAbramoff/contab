VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFormulaFuncao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsFuncaoCombo As String
Private mvarsFuncaoDesc As String

Public Property Let sFuncaoDesc(ByVal vData As String)
Attribute sFuncaoDesc.VB_Description = "Texto Descritivo da Funcao"
    mvarsFuncaoDesc = vData
End Property

Public Property Get sFuncaoDesc() As String
    sFuncaoDesc = mvarsFuncaoDesc
End Property

Public Property Let sFuncaoCombo(ByVal vData As String)
Attribute sFuncaoCombo.VB_Description = "Nome da Funcao como vai aparecer na combobox"
    mvarsFuncaoCombo = vData
End Property

Public Property Get sFuncaoCombo() As String
    sFuncaoCombo = mvarsFuncaoCombo
End Property



