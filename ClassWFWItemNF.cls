VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassWFWItemNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsProduto As String
Private mvariItemNF As Integer
Private mvardAliquota As Double
Private mvardAliquotaPC As Double

Public Property Let dAliquotaPC(ByVal vData As Double)
    mvardAliquotaPC = vData
End Property

Public Property Get dAliquotaPC() As Double
    dAliquotaPC = mvardAliquotaPC
End Property

Public Property Let dAliquota(ByVal vData As Double)
    mvardAliquota = vData
End Property

Public Property Get dAliquota() As Double
    dAliquota = mvardAliquota
End Property

Public Property Let iItemNF(ByVal vData As Integer)
    mvariItemNF = vData
End Property

Public Property Get iItemNF() As Integer
    iItemNF = mvariItemNF
End Property


Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

