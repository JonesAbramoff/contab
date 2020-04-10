VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassOpcoesTelasCampos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsNomeTela As String
Private mvarsNomeControle As String
Private mvariExecutaValidate As Integer

Public Property Let sNomeTela(ByVal vData As String)
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property

Public Property Let sNomeControle(ByVal vData As String)
    mvarsNomeControle = vData
End Property

Public Property Get sNomeControle() As String
    sNomeControle = mvarsNomeControle
End Property

Public Property Let iExecutaValidate(ByVal vData As Integer)
    mvariExecutaValidate = vData
End Property

Public Property Get iExecutaValidate() As Integer
    iExecutaValidate = mvariExecutaValidate
End Property
