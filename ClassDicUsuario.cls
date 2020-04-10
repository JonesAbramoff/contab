VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDicUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Classe Usuario"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsCodUsuario As String
Private mvarsCodGrupo As String 'local copy
Private mvarsNome As String 'local copy
Private mvarsNomeReduzido As String 'local copy
Private mvarsSenha As String 'local copy
Private mvardtDataValidade As Date 'local copy
Private mvariAtivo As Integer 'local copy
Private mvarsNomeLogin As String
Private mvarsComputador As String
Private mvariLogado As Integer
Private mvarsEmail As String

Public Property Let sEmail(ByVal vData As String)
    mvarsEmail = vData
End Property

Public Property Get sEmail() As String
    sEmail = mvarsEmail
End Property

Public Property Let iLogado(ByVal vData As Integer)
    mvariLogado = vData
End Property

Public Property Get iLogado() As Integer
    iLogado = mvariLogado
End Property

Public Property Let sComputador(ByVal vData As String)
    mvarsComputador = vData
End Property

Public Property Get sComputador() As String
    sComputador = mvarsComputador
End Property

Public Property Let sNomeLogin(ByVal vData As String)
    mvarsNomeLogin = vData
End Property

Public Property Get sNomeLogin() As String
    sNomeLogin = mvarsNomeLogin
End Property

Public Property Let iAtivo(ByVal vData As Integer)
Attribute iAtivo.VB_Description = "1-> Usuário Ativo  0->Usuário Inativo"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iAtivo = 5
    mvariAtivo = vData
End Property


Public Property Get iAtivo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iAtivo
    iAtivo = mvariAtivo
End Property



Public Property Let dtDataValidade(ByVal vData As Date)
Attribute dtDataValidade.VB_Description = "Data de Validade do Usuário"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataValidade = 5
    mvardtDataValidade = vData
End Property


Public Property Get dtDataValidade() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataValidade
    dtDataValidade = mvardtDataValidade
End Property



Public Property Let sSenha(ByVal vData As String)
Attribute sSenha.VB_Description = "Senha do Usuário"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSenha = 5
    mvarsSenha = vData
End Property


Public Property Get sSenha() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSenha
    sSenha = mvarsSenha
End Property



Public Property Let sNome(ByVal vData As String)
Attribute sNome.VB_Description = "Nome do Usuário"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNome = 5
    mvarsNome = vData
End Property


Public Property Get sNome() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNome
    sNome = mvarsNome
End Property

Public Property Let sCodGrupo(ByVal vData As String)
Attribute sCodGrupo.VB_Description = "Código do Usuário"
    mvarsCodGrupo = vData
End Property

Public Property Get sCodGrupo() As String
    sCodGrupo = mvarsCodGrupo
End Property

Public Property Let sCodUsuario(ByVal vData As String)
    mvarsCodUsuario = vData
End Property

Public Property Get sCodUsuario() As String
    sCodUsuario = mvarsCodUsuario
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property



