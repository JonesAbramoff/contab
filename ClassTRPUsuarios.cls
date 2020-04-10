VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRPUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvariTipoUsuario As Integer
Private mvarlCodigo As Long
Private mvarsLogin As String
Private mvarsSenha As String
Private mvariAlteraSenhaProxLog As Integer
Private mvarsGrupoAcesso As String
Private mvarsNome As String

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let iTipoUsuario(ByVal vData As Integer)
    mvariTipoUsuario = vData
End Property

Public Property Get iTipoUsuario() As Integer
    iTipoUsuario = mvariTipoUsuario
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let sLogin(ByVal vData As String)
    mvarsLogin = vData
End Property

Public Property Get sLogin() As String
    sLogin = mvarsLogin
End Property

Public Property Let sSenha(ByVal vData As String)
    mvarsSenha = vData
End Property

Public Property Get sSenha() As String
    sSenha = mvarsSenha
End Property

Public Property Let iAlteraSenhaProxLog(ByVal vData As Integer)
    mvariAlteraSenhaProxLog = vData
End Property

Public Property Get iAlteraSenhaProxLog() As Integer
    iAlteraSenhaProxLog = mvariAlteraSenhaProxLog
End Property

Public Property Let sGrupoAcesso(ByVal vData As String)
    mvarsGrupoAcesso = vData
End Property

Public Property Get sGrupoAcesso() As String
    sGrupoAcesso = mvarsGrupoAcesso
End Property
