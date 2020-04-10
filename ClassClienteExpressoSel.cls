VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassClienteExpressoSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsUsuCobrador As String
Private mvarsCategoria As String
Private mvarsCategoriaDe As String
Private mvarsCategoriaAte As String
Private mvarcolCidades As New Collection
Private mvarcolClientes As New Collection
Private mvarsUsuRespCallCenter As String
Private mvariRegiao As Integer
Private mvariVendedor As Integer
Private mvariCodTransportadora As Integer
Private mvarlClienteDe As Long
Private mvarlClienteAte As Long
Private mvarcolFiliais As New Collection
Private mvarcolBairros As New Collection
Private mvarcolUFs As New Collection
Private mvariTodasUFs As Integer
Private mvarcolTipoCli As New Collection
Private mvariTodosTipoCli As Integer

Public Property Let iTodosTipoCli(ByVal vData As Integer)
    mvariTodosTipoCli = vData
End Property

Public Property Get iTodosTipoCli() As Integer
    iTodosTipoCli = mvariTodosTipoCli
End Property

Public Property Get colTipoCli() As Collection
    Set colTipoCli = mvarcolTipoCli
End Property

Public Property Set colTipoCli(ByVal vData As Collection)
    Set mvarcolTipoCli = vData
End Property

Public Property Let iTodasUFs(ByVal vData As Integer)
    mvariTodasUFs = vData
End Property

Public Property Get iTodasUFs() As Integer
    iTodasUFs = mvariTodasUFs
End Property

Public Property Get colUFs() As Collection
    Set colUFs = mvarcolUFs
End Property

Public Property Set colUFs(ByVal vData As Collection)
    Set mvarcolUFs = vData
End Property

Public Property Get colBairros() As Collection
    Set colBairros = mvarcolBairros
End Property

Public Property Set colBairros(ByVal vData As Collection)
    Set mvarcolBairros = vData
End Property

Public Property Get colFiliais() As Collection
    Set colFiliais = mvarcolFiliais
End Property

Public Property Set colFiliais(ByVal vData As Collection)
    Set mvarcolFiliais = vData
End Property

Public Property Let lClienteAte(ByVal vData As Long)
    mvarlClienteAte = vData
End Property

Public Property Get lClienteAte() As Long
    lClienteAte = mvarlClienteAte
End Property

Public Property Let lClienteDe(ByVal vData As Long)
    mvarlClienteDe = vData
End Property

Public Property Get lClienteDe() As Long
    lClienteDe = mvarlClienteDe
End Property

Public Property Let iCodTransportadora(ByVal vData As Integer)
    mvariCodTransportadora = vData
End Property

Public Property Get iCodTransportadora() As Integer
    iCodTransportadora = mvariCodTransportadora
End Property

Public Property Let iVendedor(ByVal vData As Integer)
    mvariVendedor = vData
End Property

Public Property Get iVendedor() As Integer
    iVendedor = mvariVendedor
End Property

Public Property Let iRegiao(ByVal vData As Integer)
    mvariRegiao = vData
End Property

Public Property Get iRegiao() As Integer
    iRegiao = mvariRegiao
End Property

Public Property Let sUsuRespCallCenter(ByVal vData As String)
    mvarsUsuRespCallCenter = vData
End Property

Public Property Get sUsuRespCallCenter() As String
    sUsuRespCallCenter = mvarsUsuRespCallCenter
End Property

Public Property Get colClientes() As Collection
    Set colClientes = mvarcolClientes
End Property

Public Property Set colClientes(ByVal vData As Collection)
    Set mvarcolClientes = vData
End Property

Public Property Get colCidades() As Collection
    Set colCidades = mvarcolCidades
End Property

Public Property Set colCidades(ByVal vData As Collection)
    Set mvarcolCidades = vData
End Property

Public Property Let sCategoriaAte(ByVal vData As String)
    mvarsCategoriaAte = vData
End Property

Public Property Get sCategoriaAte() As String
    sCategoriaAte = mvarsCategoriaAte
End Property

Public Property Let sCategoriaDe(ByVal vData As String)
    mvarsCategoriaDe = vData
End Property

Public Property Get sCategoriaDe() As String
    sCategoriaDe = mvarsCategoriaDe
End Property

Public Property Let sCategoria(ByVal vData As String)
    mvarsCategoria = vData
End Property

Public Property Get sCategoria() As String
    sCategoria = mvarsCategoria
End Property

Public Property Let sUsuCobrador(ByVal vData As String)
    mvarsUsuCobrador = vData
End Property

Public Property Get sUsuCobrador() As String
    sUsuCobrador = mvarsUsuCobrador
End Property


