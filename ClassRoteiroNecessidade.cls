VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRoteiroNecessidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

Private mvarsProdutoRaiz As String
Private mvarsVersao As String
Private mvardQuantidade As Double
Private mvarsUM As String
Private mvarcolMaoDeObra As New Collection
Private mvarcolMP As New Collection
Private mvarcolInsumosMaquina As New Collection
Private mvariFilialEmpresa As Integer
Private mvarcolMaquinas As New Collection

Public Property Set colMaquinas(ByVal vData As Collection)
    Set mvarcolMaquinas = vData
End Property

Public Property Get colMaquinas() As Collection
    Set colMaquinas = mvarcolMaquinas
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Set colInsumosMaquina(ByVal vData As Collection)
    Set mvarcolInsumosMaquina = vData
End Property

Public Property Get colInsumosMaquina() As Collection
    Set colInsumosMaquina = mvarcolInsumosMaquina
End Property

Public Property Set colMP(ByVal vData As Collection)
    Set mvarcolMP = vData
End Property

Public Property Get colMP() As Collection
    Set colMP = mvarcolMP
End Property

Public Property Set colMaoDeObra(ByVal vData As Collection)
    Set mvarcolMaoDeObra = vData
End Property

Public Property Get colMaoDeObra() As Collection
    Set colMaoDeObra = mvarcolMaoDeObra
End Property

Public Property Let sProdutoRaiz(ByVal vData As String)
    mvarsProdutoRaiz = vData
End Property

Public Property Get sProdutoRaiz() As String
    sProdutoRaiz = mvarsProdutoRaiz
End Property

Public Property Let sVersao(ByVal vData As String)
    mvarsVersao = vData
End Property

Public Property Get sVersao() As String
    sVersao = mvarsVersao
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

