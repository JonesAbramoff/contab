VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassVersao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Private mvarsCodigo As String 'local copy
Private mvardtData As Date 'local copy
Private mvariVigente As Integer 'local copy
Public Property Let iVigente(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iVigente = 5
    mvariVigente = vData
End Property


Public Property Get iVigente() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iVigente
    iVigente = mvariVigente
End Property



Public Property Let dtData(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtData = 5
    mvardtData = vData
End Property


Public Property Get dtData() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtData
    dtData = mvardtData
End Property



Public Property Let sCodigo(ByVal vData As String)
Attribute sCodigo.VB_Description = "Código da versão."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCodigo = 5
    mvarsCodigo = vData
End Property


Public Property Get sCodigo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCodigo
    sCodigo = mvarsCodigo
End Property



