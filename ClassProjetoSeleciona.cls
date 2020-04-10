VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProjetoSeleciona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlClienteInicial As Long
Private mvarlClienteFinal As Long
Private mvarlProjetoInicial As Long
Private mvarlProjetoFinal As Long
Private mvarsProdutoInicial As String
Private mvarsProdutoFinal As String
Private mvardtDataIniInicial As Date
Private mvardtDataIniFinal As Date
Private mvardtDataTerminoInicial As Date
Private mvardtDataTerminoFinal As Date
Private mvariDestino As Integer
Private mvarcolProjetos As New Collection

Public Property Let iDestino(ByVal vData As Integer)
    mvariDestino = vData
End Property

Public Property Get iDestino() As Integer
    iDestino = mvariDestino
End Property

Public Property Let dtDataIniFinal(ByVal vData As Date)
    mvardtDataIniFinal = vData
End Property

Public Property Get dtDataIniFinal() As Date
    dtDataIniFinal = mvardtDataIniFinal
End Property

Public Property Let dtDataIniInicial(ByVal vData As Date)
    mvardtDataIniInicial = vData
End Property

Public Property Get dtDataIniInicial() As Date
    dtDataIniInicial = mvardtDataIniInicial
End Property

Public Property Let dtDataTerminoFinal(ByVal vData As Date)
    mvardtDataTerminoFinal = vData
End Property

Public Property Get dtDataTerminoFinal() As Date
    dtDataTerminoFinal = mvardtDataTerminoFinal
End Property

Public Property Let dtDataTerminoInicial(ByVal vData As Date)
    mvardtDataTerminoInicial = vData
End Property

Public Property Get dtDataTerminoInicial() As Date
    dtDataTerminoInicial = mvardtDataTerminoInicial
End Property

Public Property Let sProdutoFinal(ByVal vData As String)
    mvarsProdutoFinal = vData
End Property

Public Property Get sProdutoFinal() As String
    sProdutoFinal = mvarsProdutoFinal
End Property

Public Property Let sProdutoInicial(ByVal vData As String)
    mvarsProdutoInicial = vData
End Property

Public Property Get sProdutoInicial() As String
    sProdutoInicial = mvarsProdutoInicial
End Property

Public Property Let lProjetoFinal(ByVal vData As Long)
    mvarlProjetoFinal = vData
End Property

Public Property Get lProjetoFinal() As Long
    lProjetoFinal = mvarlProjetoFinal
End Property

Public Property Let lProjetoInicial(ByVal vData As Long)
    mvarlProjetoInicial = vData
End Property

Public Property Get lProjetoInicial() As Long
    lProjetoInicial = mvarlProjetoInicial
End Property

Public Property Let lClienteFinal(ByVal vData As Long)
    mvarlClienteFinal = vData
End Property

Public Property Get lClienteFinal() As Long
    lClienteFinal = mvarlClienteFinal
End Property

Public Property Let lClienteInicial(ByVal vData As Long)
    mvarlClienteInicial = vData
End Property

Public Property Get lClienteInicial() As Long
    lClienteInicial = mvarlClienteInicial
End Property

Public Property Set colProjetos(ByVal vData As Collection)
    Set mvarcolProjetos = vData
End Property

Public Property Get colProjetos() As Collection
    Set colProjetos = mvarcolProjetos
End Property

