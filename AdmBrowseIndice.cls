VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdmBrowseIndice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsNomeTela As String
Private mvariIndice As Integer
Private mvarsNomeIndice As String
Private mvarsOrdenacaoSQL As String
Private mvarsSelecaoSQL As String

Public Property Let sSelecaoSQL(ByVal vData As String)
    mvarsSelecaoSQL = vData
End Property

Public Property Get sSelecaoSQL() As String
    sSelecaoSQL = mvarsSelecaoSQL
End Property

Public Property Let sOrdenacaoSQL(ByVal vData As String)
    mvarsOrdenacaoSQL = vData
End Property

Public Property Get sOrdenacaoSQL() As String
    sOrdenacaoSQL = mvarsOrdenacaoSQL
End Property

Public Property Let sNomeIndice(ByVal vData As String)
    mvarsNomeIndice = vData
End Property

Public Property Get sNomeIndice() As String
    sNomeIndice = mvarsNomeIndice
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
    mvarsNomeIndice = String(STRING_NOME_INDICE, 0)
    mvarsOrdenacaoSQL = String(STRING_ORDENACAO_SQL, 0)
    mvarsSelecaoSQL = String(STRING_SELECAO_SQL, 0)

End Sub
