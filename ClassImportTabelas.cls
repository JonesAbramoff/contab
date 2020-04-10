VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassImportTabelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlCodigo As Long
Private mvarsTabela As String
Private mvarsDescricao As String
Private mvarsFuncaoGrava As String
Private mvarsFuncaoValida As String
Private mvarcolCampos As New Collection
Private mvarcolcolDados As New Collection

Public Property Set colcolDados(ByVal vData As Object)
    Set mvarcolcolDados = vData
End Property

Public Property Get colcolDados() As Collection
    Set colcolDados = mvarcolcolDados
End Property

Public Property Set colCampos(ByVal vData As Object)
    Set mvarcolCampos = vData
End Property

Public Property Get colCampos() As Collection
    Set colCampos = mvarcolCampos
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let sTabela(ByVal vData As String)
    mvarsTabela = vData
End Property

Public Property Get sTabela() As String
    sTabela = mvarsTabela
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sFuncaoGrava(ByVal vData As String)
    mvarsFuncaoGrava = vData
End Property

Public Property Get sFuncaoGrava() As String
    sFuncaoGrava = mvarsFuncaoGrava
End Property

Public Property Let sFuncaoValida(ByVal vData As String)
    mvarsFuncaoValida = vData
End Property

Public Property Get sFuncaoValida() As String
    sFuncaoValida = mvarsFuncaoValida
End Property
