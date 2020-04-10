VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdmTelaIndice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Classe TelaIndice."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarNomeTela As String
Private mvarIndice As Integer
Private mvarNomeExterno As String

Public Property Let sNomeExterno(ByVal vData As String)
    mvarNomeExterno = vData
End Property

Public Property Get sNomeExterno() As String
    sNomeExterno = mvarNomeExterno
End Property

Public Property Let iIndice(ByVal vData As Integer)
    mvarIndice = vData
End Property

Public Property Get iIndice() As Integer
    iIndice = mvarIndice
End Property

Public Property Let sNomeTela(ByVal vData As String)
    mvarNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarNomeTela
End Property


