VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlNumIntDoc As Long 'local copy
Private mvariOperacao As Integer 'local copy
Private mvarsLog As String 'local copy
Private mvardtData As Date 'local copy
Private mvardHora As Double 'local copy
Private mvarcolRegistro As New Collection 'local copy
Private mvariQuantLog As Integer 'local copy
Private mvariContinuacao As Integer 'local copy

Public Property Let iContinuacao(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iContinuacao = 5
    mvariContinuacao = vData
End Property


Public Property Get iContinuacao() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iContinuacao
    iContinuacao = mvariContinuacao
End Property


Public Property Let iQuantLog(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iQuantLog = 5
    mvariQuantLog = vData
End Property


Public Property Get iQuantLog() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iQuantLog
    iQuantLog = mvariQuantLog
End Property

Public Property Set colRegistro(ByVal vData As Collection)
    Set mvarcolRegistro = vData
End Property

Public Property Get colRegistro() As Collection
    Set colRegistro = mvarcolRegistro
End Property

Public Property Let dHora(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dHora = 5
    mvardHora = vData
End Property


Public Property Get dHora() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dHora
    dHora = mvardHora
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



Public Property Let sLog(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sLog = 5
    mvarsLog = vData
End Property


Public Property Get sLog() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sLog
    sLog = mvarsLog
End Property



Public Property Let iOperacao(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iOperacao = 5
    mvariOperacao = vData
End Property


Public Property Get iOperacao() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iOperacao
    iOperacao = mvariOperacao
End Property



Public Property Let lNumIntDoc(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntDoc = 5
    mvarlNumIntDoc = vData
End Property


Public Property Get lNumIntDoc() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntDoc
    lNumIntDoc = mvarlNumIntDoc
End Property



