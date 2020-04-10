VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSpedFiscalReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariIndice As Integer
Private mvarsCampo As String
Private mvarsValor As String

Public Property Let iIndice(ByVal vData As Integer)
    mvariIndice = vData
End Property

Public Property Get iIndice() As Integer
    iIndice = mvariIndice
End Property

Public Property Let sCampo(ByVal vData As String)
    mvarsCampo = vData
End Property

Public Property Get sCampo() As String
    sCampo = mvarsCampo
End Property

Public Property Let sValor(ByVal vData As String)
    mvarsValor = vData
End Property

Public Property Get sValor() As String
    sValor = mvarsValor
End Property



