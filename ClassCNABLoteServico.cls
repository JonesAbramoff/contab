VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCNABLoteServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsLote As String
Private mvarsTipoOperacao As String
Private mvarsTipoServico As String
Private mvarsFormaLcto As String
Private mvarsAgencia As String
Private mvarsDVAgencia As String
Private mvarsNumCtaCorrente As String
Private mvarsDVCtaCorrente As String

Public Property Let sAgencia(ByVal vData As String)
    mvarsAgencia = vData
End Property

Public Property Get sAgencia() As String
    sAgencia = mvarsAgencia
End Property

Public Property Let sDVAgencia(ByVal vData As String)
    mvarsDVAgencia = vData
End Property

Public Property Get sDVAgencia() As String
    sDVAgencia = mvarsDVAgencia
End Property

Public Property Let sNumCtaCorrente(ByVal vData As String)
    mvarsNumCtaCorrente = vData
End Property

Public Property Get sNumCtaCorrente() As String
    sNumCtaCorrente = mvarsNumCtaCorrente
End Property

Public Property Let sDVCtaCorrente(ByVal vData As String)
    mvarsDVCtaCorrente = vData
End Property

Public Property Get sDVCtaCorrente() As String
    sDVCtaCorrente = mvarsDVCtaCorrente
End Property

Public Property Let sFormaLcto(ByVal vData As String)
    mvarsFormaLcto = vData
End Property

Public Property Get sFormaLcto() As String
    sFormaLcto = mvarsFormaLcto
End Property

Public Property Let sLote(ByVal vData As String)
    mvarsLote = vData
End Property

Public Property Get sLote() As String
    sLote = mvarsLote
End Property

Public Property Let sTipoOperacao(ByVal vData As String)
    mvarsTipoOperacao = vData
End Property

Public Property Get sTipoOperacao() As String
    sTipoOperacao = mvarsTipoOperacao
End Property

Public Property Let sTipoServico(ByVal vData As String)
    mvarsTipoServico = vData
End Property

Public Property Get sTipoServico() As String
    sTipoServico = mvarsTipoServico
End Property


