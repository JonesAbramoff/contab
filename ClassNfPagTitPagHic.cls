VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNfPagTitPagHic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mvarlSituacao As Long
Private mvarlHistorico As Long
Private mvarlSubConta As Long
Private mvarsObservacao As String
Private mvariChequeRetido As Integer

Public Property Let iChequeRetido(ByVal vData As Integer)
    mvariChequeRetido = vData
End Property

Public Property Get iChequeRetido() As Integer
    iChequeRetido = mvariChequeRetido
End Property

Public Property Let lSituacao(ByVal vData As Long)
    mvarlSituacao = vData
End Property

Public Property Get lSituacao() As Long
    lSituacao = mvarlSituacao
End Property

Public Property Let lHistorico(ByVal vData As Long)
    mvarlHistorico = vData
End Property

Public Property Get lHistorico() As Long
    lHistorico = mvarlHistorico
End Property

Public Property Let lSubConta(ByVal vData As Long)
    mvarlSubConta = vData
End Property

Public Property Get lSubConta() As Long
    lSubConta = mvarlSubConta
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

