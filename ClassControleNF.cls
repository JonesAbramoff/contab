VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassControleNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlNumFormSistema As Long
Private mvarlNumFormDigitado As Long
Private mvarlNumNFiscalSistema As Long
Private mvarlNumNFiscalDigitado As Long
Private mvardtDataRegistro As Date
Private mvardHoraRegistro As Double
Private mvarsObservacao As String
Private mvariFilialEmpresa As Integer
Private mvarsUsuario As String
Private mvarsSerie As String

Public Property Let sSerie(ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let dHoraRegistro(ByVal vData As Double)
    mvardHoraRegistro = vData
End Property

Public Property Get dHoraRegistro() As Double
    dHoraRegistro = mvardHoraRegistro
End Property

Public Property Let dtDataRegistro(ByVal vData As Date)
    mvardtDataRegistro = vData
End Property

Public Property Get dtDataRegistro() As Date
    dtDataRegistro = mvardtDataRegistro
End Property

Public Property Let lNumNFiscalDigitado(ByVal vData As Long)
    mvarlNumNFiscalDigitado = vData
End Property

Public Property Get lNumNFiscalDigitado() As Long
    lNumNFiscalDigitado = mvarlNumNFiscalDigitado
End Property

Public Property Let lNumNFiscalSistema(ByVal vData As Long)
    mvarlNumNFiscalSistema = vData
End Property

Public Property Get lNumNFiscalSistema() As Long
    lNumNFiscalSistema = mvarlNumNFiscalSistema
End Property

Public Property Let lNumFormDigitado(ByVal vData As Long)
    mvarlNumFormDigitado = vData
End Property

Public Property Get lNumFormDigitado() As Long
    lNumFormDigitado = mvarlNumFormDigitado
End Property

Public Property Let lNumFormSistema(ByVal vData As Long)
    mvarlNumFormSistema = vData
End Property

Public Property Get lNumFormSistema() As Long
    lNumFormSistema = mvarlNumFormSistema
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

