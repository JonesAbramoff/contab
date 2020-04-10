VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFiscalMgz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsContrato As String
Private mvarsProcesso As String
Private mvarsIDCorrespondencia As String

Public Property Let sIDCorrespondencia(ByVal vData As String)
    mvarsIDCorrespondencia = vData
End Property

Public Property Get sIDCorrespondencia() As String
    sIDCorrespondencia = mvarsIDCorrespondencia
End Property

Public Property Let sProcesso(ByVal vData As String)
    mvarsProcesso = vData
End Property

Public Property Get sProcesso() As String
    sProcesso = mvarsProcesso
End Property

Public Property Let sContrato(ByVal vData As String)
    mvarsContrato = vData
End Property

Public Property Get sContrato() As String
    sContrato = mvarsContrato
End Property
