VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarobjCliente As New ClassCliente
Private mvarobjProduto As New ClassProduto
Private mvarobjFilialCliente As New ClassFilialCliente
Private mvarobjFilialEmpresa As New AdmFiliais
Private mvarobjEnderecoCli As New ClassEndereco
Private mvarobjEnderecoEmp As New ClassEndereco
Private mvarobjTipoDocInfo As New ClassTipoDocInfo
Private mvarobjTipoTribICMS As New ClassTipoTribICMS
Private mvarobjTipoTribIPI As New ClassTipoTribIPI
Private mvarobjTipoDeTributacaoMovto As New ClassTipoDeTributacaoMovto

Public Property Set objTipoTribICMS(ByVal vData As ClassTipoTribICMS)
    Set mvarobjTipoTribICMS = vData
End Property

Public Property Get objTipoTribICMS() As ClassTipoTribICMS
    Set objTipoTribICMS = mvarobjTipoTribICMS
End Property

Public Property Set objTipoTribIPI(ByVal vData As ClassTipoTribIPI)
    Set mvarobjTipoTribIPI = vData
End Property

Public Property Get objTipoTribIPI() As ClassTipoTribIPI
    Set objTipoTribIPI = mvarobjTipoTribIPI
End Property

Public Property Set objTipoDeTributacaoMovto(ByVal vData As ClassTipoDeTributacaoMovto)
    Set mvarobjTipoDeTributacaoMovto = vData
End Property

Public Property Get objTipoDeTributacaoMovto() As ClassTipoDeTributacaoMovto
    Set objTipoDeTributacaoMovto = mvarobjTipoDeTributacaoMovto
End Property

Public Property Set objTipoDocInfo(ByVal vData As ClassTipoDocInfo)
    Set mvarobjTipoDocInfo = vData
End Property

Public Property Get objTipoDocInfo() As ClassTipoDocInfo
    Set objTipoDocInfo = mvarobjTipoDocInfo
End Property

Public Property Set objEnderecoEmp(ByVal vData As ClassEndereco)
    Set mvarobjEnderecoEmp = vData
End Property

Public Property Get objEnderecoEmp() As ClassEndereco
    Set objEnderecoEmp = mvarobjEnderecoEmp
End Property

Public Property Set objEnderecoCli(ByVal vData As ClassEndereco)
    Set mvarobjEnderecoCli = vData
End Property

Public Property Get objEnderecoCli() As ClassEndereco
    Set objEnderecoCli = mvarobjEnderecoCli
End Property

Public Property Set objFilialEmpresa(ByVal vData As AdmFiliais)
    Set mvarobjFilialEmpresa = vData
End Property

Public Property Get objFilialEmpresa() As AdmFiliais
    Set objFilialEmpresa = mvarobjFilialEmpresa
End Property

Public Property Set objFilialCliente(ByVal vData As ClassFilialCliente)
    Set mvarobjFilialCliente = vData
End Property

Public Property Get objFilialCliente() As ClassFilialCliente
    Set objFilialCliente = mvarobjFilialCliente
End Property

Public Property Set objProduto(ByVal vData As ClassProduto)
    Set mvarobjProduto = vData
End Property

Public Property Get objProduto() As ClassProduto
    Set objProduto = mvarobjProduto
End Property

Public Property Set objCliente(ByVal vData As ClassCliente)
    Set mvarobjCliente = vData
End Property

Public Property Get objCliente() As ClassCliente
    Set objCliente = mvarobjCliente
End Property



