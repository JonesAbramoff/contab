VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassConfigOutros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsTabela As String
Private mvarsCodigo As String
Private mvarsModuloExibicao As String
Private mvarsNomeObj As String
Private mvarsNomeProperty As String
Private mvarsDescricaoGrid As String
Private mvarsTipoControle As String
Private mvarlCodVlrValidos As Long

'AUXILIARES
Private mvariFilialEmpresa As Long
Private mvarsConteudo As String
Private mvarcolValores As New Collection
Private mvarsConteudoNovo As String

Public Property Let sConteudoNovo(ByVal vData As String)
    mvarsConteudoNovo = vData
End Property

Public Property Get sConteudoNovo() As String
    sConteudoNovo = mvarsConteudoNovo
End Property

Public Property Set colValores(ByVal vData As Collection)
    Set mvarcolValores = vData
End Property

Public Property Get colValores() As Collection
    Set colValores = mvarcolValores
End Property

Public Property Let sConteudo(ByVal vData As String)
    mvarsConteudo = vData
End Property

Public Property Get sConteudo() As String
    sConteudo = mvarsConteudo
End Property

Public Property Let iFilialEmpresa(ByVal vData As Long)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Long
    iFilialEmpresa = mvariFilialEmpresa
End Property
'AUXILIARES

Public Property Let sTabela(ByVal vData As String)
    mvarsTabela = vData
End Property

Public Property Get sTabela() As String
    sTabela = mvarsTabela
End Property

Public Property Let sCodigo(ByVal vData As String)
    mvarsCodigo = vData
End Property

Public Property Get sCodigo() As String
    sCodigo = mvarsCodigo
End Property

Public Property Let sModuloExibicao(ByVal vData As String)
    mvarsModuloExibicao = vData
End Property

Public Property Get sModuloExibicao() As String
    sModuloExibicao = mvarsModuloExibicao
End Property

Public Property Let sNomeObj(ByVal vData As String)
    mvarsNomeObj = vData
End Property

Public Property Get sNomeObj() As String
    sNomeObj = mvarsNomeObj
End Property

Public Property Let sNomeProperty(ByVal vData As String)
    mvarsNomeProperty = vData
End Property

Public Property Get sNomeProperty() As String
    sNomeProperty = mvarsNomeProperty
End Property

Public Property Let sDescricaoGrid(ByVal vData As String)
    mvarsDescricaoGrid = vData
End Property

Public Property Get sDescricaoGrid() As String
    sDescricaoGrid = mvarsDescricaoGrid
End Property

Public Property Let sTipoControle(ByVal vData As String)
    mvarsTipoControle = vData
End Property

Public Property Get sTipoControle() As String
    sTipoControle = mvarsTipoControle
End Property

Public Property Let lCodVlrValidos(ByVal vData As Long)
    mvarlCodVlrValidos = vData
End Property

Public Property Get lCodVlrValidos() As Long
    lCodVlrValidos = mvarlCodVlrValidos
End Property
