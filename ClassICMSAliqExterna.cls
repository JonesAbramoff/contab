VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassICMSAliqExterna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsSiglaOrig As String
Private mvarsSiglaDest As String
Private mvardAliquota As Double
Private mvardAliquotaFCP As Double

Public Property Let dAliquotaFCP(ByVal vData As Double)
    mvardAliquotaFCP = vData
End Property

Public Property Get dAliquotaFCP() As Double
    dAliquotaFCP = mvardAliquotaFCP
End Property

Public Property Let dAliquota(ByVal vData As Double)
    mvardAliquota = vData
End Property

Public Property Get dAliquota() As Double
    dAliquota = mvardAliquota
End Property

Public Property Let sSiglaDest(ByVal vData As String)
    mvarsSiglaDest = vData
End Property

Public Property Get sSiglaDest() As String
    sSiglaDest = mvarsSiglaDest
End Property

Public Property Let sSiglaOrig(ByVal vData As String)
    mvarsSiglaOrig = vData
End Property

Public Property Get sSiglaOrig() As String
    sSiglaOrig = mvarsSiglaOrig
End Property



