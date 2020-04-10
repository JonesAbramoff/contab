VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDicEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlCodigo As Long
Private mvarsNome As String
Private mvarsNomeReduzido As String
Private mvarsStringConexao As String
Private mvariInativa As Integer
Private mvarcolModulos As New Collection 'local copy
Public Property Set colModulos(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colModulos = Form1
    Set mvarcolModulos = vData
End Property

Public Property Get colModulos() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colModulos
    Set colModulos = mvarcolModulos
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let iInativa(ByVal vData As Integer)
    mvariInativa = vData
End Property

Public Property Get iInativa() As Integer
    iInativa = mvariInativa
End Property

Public Property Let sStringConexao(ByVal vData As String)
Attribute sStringConexao.VB_Description = "String de Conexão da Empresa ao Banco de Dados"
    mvarsStringConexao = vData
End Property

Public Property Get sStringConexao() As String
    sStringConexao = mvarsStringConexao
End Property

Public Property Let sNome(ByVal vData As String)
Attribute sNome.VB_Description = "Nome da Empresa"
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let lCodigo(ByVal vData As Long)
Attribute lCodigo.VB_Description = "codigo da empresa"
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property



