VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdmBrowseIndiceSegmentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsNomeTela As String
Private mvariIndice As Integer
Private mvariPosicaoCampo As Integer
Private mvarsNomeCampo As String

Public Property Let sNomeCampo(ByVal vData As String)
    mvarsNomeCampo = vData
End Property

Public Property Get sNomeCampo() As String
    sNomeCampo = mvarsNomeCampo
End Property

Public Property Let iPosicaoCampo(ByVal vData As Integer)
    mvariPosicaoCampo = vData
End Property

Public Property Get iPosicaoCampo() As Integer
    iPosicaoCampo = mvariPosicaoCampo
End Property

Public Property Let iIndice(ByVal vData As Integer)
    mvariIndice = vData
End Property

Public Property Get iIndice() As Integer
    iIndice = mvariIndice
End Property

Public Property Let sNomeTela(ByVal vData As String)
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property

Private Sub Class_Initialize()

    mvarsNomeTela = String(STRING_NOME_TELA, 0)
    mvarsNomeCampo = String(STRING_NOME_CAMPO, 0)

End Sub

