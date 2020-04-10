VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarobjCT As Object 'local copy
Private mvarsNomeTela As String 'local copy
Public Property Let sNomeTela(ByVal vData As String)
    mvarsNomeTela = vData
End Property


Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property



Public Property Set objCT(ByVal vData As Object)
    Set mvarobjCT = vData
End Property


Public Property Get objCT() As Object
    Set objCT = mvarobjCT
End Property



