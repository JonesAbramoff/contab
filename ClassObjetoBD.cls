VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassObjetoBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarsClasseObjeto As String
Private mvarsNomeArquivo As String
Private mvariTipo As Integer
Private mvarsSelecaoSQL As String
Private mvariAvisaSobrePosicao As Integer
Private mvarsNomeObjetoMSG As String

Public Property Let sNomeObjetoMSG(ByVal vData As String)
    mvarsNomeObjetoMSG = vData
End Property

Public Property Get sNomeObjetoMSG() As String
    sNomeObjetoMSG = mvarsNomeObjetoMSG
End Property

Public Property Let iAvisaSobrePosicao(ByVal vData As Integer)
    mvariAvisaSobrePosicao = vData
End Property

Public Property Get iAvisaSobrePosicao() As Integer
    iAvisaSobrePosicao = mvariAvisaSobrePosicao
End Property

Public Property Let sSelecaoSQL(ByVal vData As String)
    mvarsSelecaoSQL = vData
End Property

Public Property Get sSelecaoSQL() As String
    sSelecaoSQL = mvarsSelecaoSQL
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let sNomeArquivo(ByVal vData As String)
    mvarsNomeArquivo = vData
End Property

Public Property Get sNomeArquivo() As String
    sNomeArquivo = mvarsNomeArquivo
End Property

Public Property Let sClasseObjeto(ByVal vData As String)
    mvarsClasseObjeto = vData
End Property

Public Property Get sClasseObjeto() As String
    sClasseObjeto = mvarsClasseObjeto
End Property



