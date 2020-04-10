VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassOrdenaColecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarvColuna(1 To 500) As Variant
Private mvarsColuna(1 To 500) As String
Private mvariIndice As Integer

Public Property Let iIndice(ByVal vData As Integer)
    mvariIndice = vData
End Property

Public Property Get iIndice() As Integer
    iIndice = mvariIndice
End Property

Public Property Let sColuna(ByVal iIndice As Integer, ByVal vData As String)
    mvarsColuna(iIndice) = vData
End Property

Public Property Get sColuna(ByVal iIndice As Integer) As String
    sColuna = mvarsColuna(iIndice)
End Property

Public Property Let vColuna(ByVal iIndice As Integer, ByVal vData As Variant)
    mvarvColuna(iIndice) = vData
End Property

Public Property Get vColuna(ByVal iIndice As Integer) As Variant
    vColuna = mvarvColuna(iIndice)
End Property
