VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCobrador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvariInativo As Integer
Private mvarsNomeReduzido As String
Private mvarsNome As String
Private mvarlEndereco As Long
Private mvariCodBanco As Integer
Private mvariCodCCI As Integer
Private mvarlCNABProxSeqArqCobr As Long
Private mvarcolCarteira As New Collection
Private mvariCobrancaEletronica As Integer
Private mvariFilialEmpresa As Integer
Private mvarlFornecedor As Long
Private mvariFilial As Integer

Public Property Let iFilial(ByVal vData As Integer)
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iCobrancaEletronica(ByVal vData As Integer)
    mvariCobrancaEletronica = vData
End Property

Public Property Get iCobrancaEletronica() As Integer
    iCobrancaEletronica = mvariCobrancaEletronica
End Property

Public Property Get colCarteira() As Collection
    Set colCarteira = mvarcolCarteira
End Property

Public Property Let lCNABProxSeqArqCobr(ByVal vData As Long)
Attribute lCNABProxSeqArqCobr.VB_Description = "prox sequencial p/identificar arquivo de remessa/retorno (sem pular e sem repetir)"
    mvarlCNABProxSeqArqCobr = vData
End Property

Public Property Get lCNABProxSeqArqCobr() As Long
    lCNABProxSeqArqCobr = mvarlCNABProxSeqArqCobr
End Property

Public Property Let iCodCCI(ByVal vData As Integer)
    mvariCodCCI = vData
End Property

Public Property Get iCodCCI() As Integer
    iCodCCI = mvariCodCCI
End Property

Public Property Let iCodBanco(ByVal vData As Integer)
    mvariCodBanco = vData
End Property

Public Property Get iCodBanco() As Integer
    iCodBanco = mvariCodBanco
End Property

Public Property Let lEndereco(ByVal vData As Long)
    mvarlEndereco = vData
End Property

Public Property Get lEndereco() As Long
    lEndereco = mvarlEndereco
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let iInativo(ByVal vData As Integer)
    mvariInativo = vData
End Property

Public Property Get iInativo() As Integer
    iInativo = mvariInativo
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

Private Sub Class_Initialize()
    Set mvarcolCarteira = New Collection
End Sub
