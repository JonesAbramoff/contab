VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFsPagTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarcolTiposDoc As New Collection

Public Property Set colTiposDoc(ByVal vData As Collection)
    Set mvarcolTiposDoc = vData
End Property

Public Property Get colTiposDoc() As Collection
    Set colTiposDoc = mvarcolTiposDoc
End Property
