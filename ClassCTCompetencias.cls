VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCTCompetencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDocCT As Long 'local copy
Private mvarlNumIntDocCompet As Long 'local copy
Public Property Let lNumIntDocCompet(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntDocCompet = 5
    mvarlNumIntDocCompet = vData
End Property


Public Property Get lNumIntDocCompet() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntDocCompet
    lNumIntDocCompet = mvarlNumIntDocCompet
End Property



Public Property Let lNumIntDocCT(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntDocCT = 5
    mvarlNumIntDocCT = vData
End Property


Public Property Get lNumIntDocCT() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntDocCT
    lNumIntDocCT = mvarlNumIntDocCT
End Property



