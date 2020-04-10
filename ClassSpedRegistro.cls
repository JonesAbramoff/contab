VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSpedRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlCount As Long
Private mvariCount As Integer
Private mvarsTipo As String

Public Property Let sTipo(ByVal vData As String)
    mvarsTipo = vData
End Property

Public Property Get sTipo() As String
    sTipo = mvarsTipo
End Property

Public Property Let lCount(ByVal vData As Long)
    mvarlCount = vData
End Property

Public Property Get lCount() As Long
    lCount = mvarlCount
End Property

Public Property Let iCount(ByVal vData As Integer)
    mvariCount = vData
End Property

Public Property Get iCount() As Integer
    iCount = mvariCount
End Property


