VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEnderecoBol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsRadio As String
Private mvarsSkype As String

Public Property Let sSkype(ByVal vData As String)
    mvarsSkype = vData
End Property

Public Property Get sSkype() As String
    sSkype = mvarsSkype
End Property

Public Property Let sRadio(ByVal vData As String)
    mvarsRadio = vData
End Property

Public Property Get sRadio() As String
    sRadio = mvarsRadio
End Property
