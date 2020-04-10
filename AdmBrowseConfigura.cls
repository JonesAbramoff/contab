VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmBrowseConfigura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsNomeTela As String
Private mvarcolBrowseUsuarioCampo As Collection
Private mvariTelaOK As Integer
Private mvarsSelecaoSQL1 As String
Private mvarsSelecaoSQL1Usuario As String
Private mvariPesquisa As Integer 'indica que a tela de configuracao foi chamada atraves do botao pesquisa
Private mvariAlteradoOrdenacao As Integer 'indica que a ordenação foi alterada pelo usuario
Private mvarcolBrowseIndiceUsuario As Collection 'guarda a coleção de indices criados pelo usuario
Private mvarcolBrowseIndice As Collection 'guarda a coleção de indices criados pelo sistema
Private mvarobjBrowseExcel As New AdmBrowseExcel

Public Property Set objBrowseExcel(ByVal vData As Object)
    Set mvarobjBrowseExcel = vData
End Property

Public Property Get objBrowseExcel() As AdmBrowseExcel
    Set objBrowseExcel = mvarobjBrowseExcel
End Property

Public Property Set colBrowseIndice(ByVal vData As Object)
    Set mvarcolBrowseIndice = vData
End Property

Public Property Get colBrowseIndice() As Collection
    Set colBrowseIndice = mvarcolBrowseIndice
End Property

Public Property Set colBrowseIndiceUsuario(ByVal vData As Object)
    Set mvarcolBrowseIndiceUsuario = vData
End Property

Public Property Get colBrowseIndiceUsuario() As Collection
    Set colBrowseIndiceUsuario = mvarcolBrowseIndiceUsuario
End Property

Public Property Let iAlteradoOrdenacao(ByVal vData As Integer)
    mvariAlteradoOrdenacao = vData
End Property

Public Property Get iAlteradoOrdenacao() As Integer
    iAlteradoOrdenacao = mvariAlteradoOrdenacao
End Property

Public Property Let iPesquisa(ByVal vData As Integer)
    mvariPesquisa = vData
End Property

Public Property Get iPesquisa() As Integer
    iPesquisa = mvariPesquisa
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

Public Property Let iTelaOK(ByVal vData As Integer)
    mvariTelaOK = vData
End Property

Public Property Get iTelaOK() As Integer
    iTelaOK = mvariTelaOK
End Property

Public Property Set colBrowseUsuarioCampo(ByVal vData As Object)
Attribute colBrowseUsuarioCampo.VB_Description = "Coleção dos campos selecionados pelo usuário para serem exibidos no browse"
    Set mvarcolBrowseUsuarioCampo = vData
End Property

Public Property Get colBrowseUsuarioCampo() As Collection
    Set colBrowseUsuarioCampo = mvarcolBrowseUsuarioCampo
End Property

Public Property Let sNomeTela(ByVal vData As String)
Attribute sNomeTela.VB_Description = "Nome de uma tela"
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property



