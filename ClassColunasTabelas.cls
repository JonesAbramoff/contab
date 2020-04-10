VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassColunasTabelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarsArquivo As String
Private mvarsAtributoClasse As String
Private mvarsArquivoTipo As String
Private mvarsColuna As String
Private mvarsColunaTipo As String
Private mvarlColunaTamanho As Long
Private mvarlTamanhoTela As Long
Private mvarlColunaPrecisao As Long
Private mvariOrdem As Integer
Private mvariBrowse As Integer
Private mvariIndice As Integer
Private mvariClasse As Integer
Private mvariChave As Integer
Private mvarsDescricao As String
Private mvariSubTipo As Integer

Public Property Let iSubTipo(ByVal vData As Integer)
    mvariSubTipo = vData
End Property

Public Property Get iSubTipo() As Integer
    iSubTipo = mvariSubTipo
End Property

Public Property Let iChave(ByVal vData As Integer)
    mvariChave = vData
End Property

Public Property Get iChave() As Integer
    iChave = mvariChave
End Property

Public Property Let iClasse(ByVal vData As Integer)
    mvariClasse = vData
End Property

Public Property Get iClasse() As Integer
    iClasse = mvariClasse
End Property

Public Property Let iBrowse(ByVal vData As Integer)
    mvariBrowse = vData
End Property

Public Property Get iBrowse() As Integer
    iBrowse = mvariBrowse
End Property

Public Property Let iIndice(ByVal vData As Integer)
    mvariIndice = vData
End Property

Public Property Get iIndice() As Integer
    iIndice = mvariIndice
End Property

Public Property Let iOrdem(ByVal vData As Integer)
    mvariOrdem = vData
End Property

Public Property Get iOrdem() As Integer
    iOrdem = mvariOrdem
End Property

Public Property Let lTamanhoTela(ByVal vData As Long)
    mvarlTamanhoTela = vData
End Property

Public Property Get lTamanhoTela() As Long
    lTamanhoTela = mvarlTamanhoTela
End Property

Public Property Let lColunaPrecisao(ByVal vData As Long)
    mvarlColunaPrecisao = vData
End Property

Public Property Get lColunaPrecisao() As Long
    lColunaPrecisao = mvarlColunaPrecisao
End Property

Public Property Let lColunaTamanho(ByVal vData As Long)
    mvarlColunaTamanho = vData
End Property

Public Property Get lColunaTamanho() As Long
    lColunaTamanho = mvarlColunaTamanho
End Property

Public Property Let sColunaTipo(ByVal vData As String)
    mvarsColunaTipo = vData
End Property

Public Property Get sColunaTipo() As String
    sColunaTipo = mvarsColunaTipo
End Property

Public Property Let sColuna(ByVal vData As String)
    mvarsColuna = vData
End Property

Public Property Get sColuna() As String
    sColuna = mvarsColuna
End Property

Public Property Let sArquivoTipo(ByVal vData As String)
    mvarsArquivoTipo = vData
End Property

Public Property Get sArquivoTipo() As String
    sArquivoTipo = mvarsArquivoTipo
End Property

Public Property Let sArquivo(ByVal vData As String)
    mvarsArquivo = vData
End Property

Public Property Get sArquivo() As String
    sArquivo = mvarsArquivo
End Property

Public Property Let sAtributoClasse(ByVal vData As String)
    mvarsAtributoClasse = vData
End Property

Public Property Get sAtributoClasse() As String
    sAtributoClasse = mvarsAtributoClasse
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property



