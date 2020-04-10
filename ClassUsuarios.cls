VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Tabela de Usuarios do Dic. Dados"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsCodUsuario As String
Private mvarsCodGrupo As String
Private mvarsNome As String
Private mvarsNomeReduzido As String
Private mvarsSenha As String
Private mvardtDataValidade As Date
Private mvariAtivo As Integer
Private mvariWorkFlowAtivo As Integer
Private mvarsEmail As String

Public Property Let sEmail(ByVal vData As String)
    mvarsEmail = vData
End Property

Public Property Get sEmail() As String
    sEmail = mvarsEmail
End Property

Public Property Let iWorkFlowAtivo(ByVal vData As Integer)
    mvariWorkFlowAtivo = vData
End Property

Public Property Get iWorkFlowAtivo() As Integer
    iWorkFlowAtivo = mvariWorkFlowAtivo
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property

Public Property Let dtDataValidade(ByVal vData As Date)
    mvardtDataValidade = vData
End Property

Public Property Get dtDataValidade() As Date
    dtDataValidade = mvardtDataValidade
End Property

Public Property Let sSenha(ByVal vData As String)
    mvarsSenha = vData
End Property

Public Property Get sSenha() As String
    sSenha = mvarsSenha
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let sCodUsuario(ByVal vData As String)
    mvarsCodUsuario = vData
End Property

Public Property Get sCodUsuario() As String
    sCodUsuario = mvarsCodUsuario
End Property

Public Property Let sCodGrupo(ByVal vData As String)
    mvarsCodGrupo = vData
End Property

Public Property Get sCodGrupo() As String
    sCodGrupo = mvarsCodGrupo
End Property



