VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmBrowseUsuarioOrdenacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsNomeTela As String
Private mvarsCodUsuario As String
Private mvariIndice As Integer
Private mvarsSelecaoSQL1 As String
Private mvarsSelecaoSQL1Usuario As String
Private mvarsNomeIndice As String

Public Property Let sNomeIndice(ByVal vData As String)
    mvarsNomeIndice = vData
End Property

Public Property Get sNomeIndice() As String
    sNomeIndice = mvarsNomeIndice
End Property

Public Property Let sSelecaoSQL1Usuario(ByVal vData As String)
    mvarsSelecaoSQL1Usuario = vData
End Property

Public Property Get sSelecaoSQL1Usuario() As String
    sSelecaoSQL1Usuario = mvarsSelecaoSQL1Usuario
End Property

Public Property Let sSelecaoSQL1(ByVal vData As String)
    mvarsSelecaoSQL1 = vData
End Property

Public Property Get sSelecaoSQL1() As String
    sSelecaoSQL1 = mvarsSelecaoSQL1
End Property

Public Property Let iIndice(ByVal vData As Integer)
    mvariIndice = vData
End Property

Public Property Get iIndice() As Integer
    iIndice = mvariIndice
End Property

Public Property Let sCodUsuario(ByVal vData As String)
    mvarsCodUsuario = vData
End Property

Public Property Get sCodUsuario() As String
    sCodUsuario = mvarsCodUsuario
End Property

Public Property Let sNomeTela(ByVal vData As String)
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property

Private Sub Class_Initialize()

    mvarsCodUsuario = String(STRING_USUARIO, 0)
    mvarsNomeTela = String(STRING_NOME_TELA, 0)

End Sub
