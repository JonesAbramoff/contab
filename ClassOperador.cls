VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassOperador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvariCodigo As Integer
Private mvarsSenha As String
Private mvariDesconto As Integer
Private mvariLimiteDesconto As Integer
Private mvarsNome As String
Private mvariCodigoVendedor As Integer
Private mvariGerente As Integer
Private mvariAtivo As Integer
Private mvardtDataLog As Date

Public Property Let dtDataLog(ByVal vData As Date)
    mvardtDataLog = vData
End Property

Public Property Get dtDataLog() As Date
    dtDataLog = mvardtDataLog
End Property

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property

Public Property Let iGerente(ByVal vData As Integer)
    mvariGerente = vData
End Property

Public Property Get iGerente() As Integer
    iGerente = mvariGerente
End Property

Public Property Let iCodigoVendedor(ByVal vData As Integer)
    mvariCodigoVendedor = vData
End Property

Public Property Get iCodigoVendedor() As Integer
    iCodigoVendedor = mvariCodigoVendedor
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let iLimiteDesconto(ByVal vData As Integer)
    mvariLimiteDesconto = vData
End Property

Public Property Get iLimiteDesconto() As Integer
    iLimiteDesconto = mvariLimiteDesconto
End Property

Public Property Let iDesconto(ByVal vData As Integer)
    mvariDesconto = vData
End Property

Public Property Get iDesconto() As Integer
    iDesconto = mvariDesconto
End Property

Public Property Let sSenha(ByVal vData As String)
    mvarsSenha = vData
End Property

Public Property Get sSenha() As String
    sSenha = mvarsSenha
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property



