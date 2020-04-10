VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSpedC460"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsCodObs As String
Private mvarsTxt As String

Public Property Let sTxt(ByVal vData As String)
    mvarsTxt = vData
End Property

Public Property Get sTxt() As String
    sTxt = mvarsTxt
End Property

Public Property Let sCodObs(ByVal vData As String)
    mvarsCodObs = vData
End Property

Public Property Get sCodObs() As String
    sCodObs = mvarsCodObs
End Property

