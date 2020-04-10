VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRede"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvarsNome As String
Private mvariFilialEmpresa As Integer
Private mvarlCliente As Long
Private mvariFilialCli As Integer
Private mvariAtivo As Integer
Private mvardtDataLog As Date
Private mvarlClienteBack As Long
Private mvariFilialCliBack As Integer
Private mvariTEF As Integer
Private mvariBandeira As Integer
Private mvarsEstabelecimento As String
Private mvariCodConta As Integer
Private mvarsDirImportacaoExtrato As String
Public sCGC As String

Public Property Let sDirImportacaoExtrato(ByVal vData As String)
    mvarsDirImportacaoExtrato = vData
End Property

Public Property Get sDirImportacaoExtrato() As String
    sDirImportacaoExtrato = mvarsDirImportacaoExtrato
End Property

Public Property Let iCodConta(ByVal vData As Integer)
    mvariCodConta = vData
End Property

Public Property Get iCodConta() As Integer
    iCodConta = mvariCodConta
End Property

Public Property Let sEstabelecimento(ByVal vData As String)
    mvarsEstabelecimento = vData
End Property

Public Property Get sEstabelecimento() As String
    sEstabelecimento = mvarsEstabelecimento
End Property

Public Property Let iBandeira(ByVal vData As Integer)
    mvariBandeira = vData
End Property

Public Property Get iBandeira() As Integer
    iBandeira = mvariBandeira
End Property

Public Property Let iTEF(ByVal vData As Integer)
    mvariTEF = vData
End Property

Public Property Get iTEF() As Integer
    iTEF = mvariTEF
End Property

Public Property Let iFilialCliBack(ByVal vData As Integer)
    mvariFilialCliBack = vData
End Property

Public Property Get iFilialCliBack() As Integer
    iFilialCliBack = mvariFilialCliBack
End Property

Public Property Let lClienteBack(ByVal vData As Long)
    mvarlClienteBack = vData
End Property

Public Property Get lClienteBack() As Long
    lClienteBack = mvarlClienteBack
End Property

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

Public Property Let iFilialCli(ByVal vData As Integer)
    mvariFilialCli = vData
End Property

Public Property Get iFilialCli() As Integer
    iFilialCli = mvariFilialCli
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property
