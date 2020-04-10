VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTransfECF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsFTPURL As String
Private mvarsFTPUserName As String
Private mvarsFTPPassword As String
Private mvarsFTPDiretorio As String
Private mvarlIntervaloTrans As Long
Private mvarsDirDadosECF As String
Private mvariCaixa As Integer
Private mvarsDirDadosCCC As String

Public Property Let sDirDadosCCC(ByVal vData As String)
    mvarsDirDadosCCC = vData
End Property

Public Property Get sDirDadosCCC() As String
    sDirDadosCCC = mvarsDirDadosCCC
End Property

Public Property Let iCaixa(ByVal vData As Integer)
    mvariCaixa = vData
End Property

Public Property Get iCaixa() As Integer
    iCaixa = mvariCaixa
End Property

Public Property Let sDirDadosECF(ByVal vData As String)
    mvarsDirDadosECF = vData
End Property

Public Property Get sDirDadosECF() As String
    sDirDadosECF = mvarsDirDadosECF
End Property

Public Property Let lIntervaloTrans(ByVal vData As Long)
    mvarlIntervaloTrans = vData
End Property

Public Property Get lIntervaloTrans() As Long
    lIntervaloTrans = mvarlIntervaloTrans
End Property

Public Property Let sFTPDiretorio(ByVal vData As String)
    mvarsFTPDiretorio = vData
End Property

Public Property Get sFTPDiretorio() As String
    sFTPDiretorio = mvarsFTPDiretorio
End Property

Public Property Let sFTPPassword(ByVal vData As String)
    mvarsFTPPassword = vData
End Property

Public Property Get sFTPPassword() As String
    sFTPPassword = mvarsFTPPassword
End Property

Public Property Let sFTPUserName(ByVal vData As String)
    mvarsFTPUserName = vData
End Property

Public Property Get sFTPUserName() As String
    sFTPUserName = mvarsFTPUserName
End Property

Public Property Let sFTPURL(ByVal vData As String)
    mvarsFTPURL = vData
End Property

Public Property Get sFTPURL() As String
    sFTPURL = mvarsFTPURL
End Property


