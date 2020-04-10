VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFechamentoBoletos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvariNumIntDoc As Integer
Private mvarsNumIntDoc As String
Private mvarsLote As String
Private mvardtLote As Date
Private mvarcolBoletos As New Collection
Public Property Set colBoletos(ByVal vData As Collection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colBoletos = Form1
    Set mvarcolBoletos = vData
End Property


Public Property Get colBoletos() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colBoletos
    Set colBoletos = mvarcolBoletos
End Property



Public Property Let dtLote(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtLote = 5
    mvardtLote = vData
End Property


Public Property Get dtLote() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtLote
    dtLote = mvardtLote
End Property



Public Property Let sLote(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sLote = 5
    mvarsLote = vData
End Property


Public Property Get sLote() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sLote
    sLote = mvarsLote
End Property



Public Property Let sNumIntDoc(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNumIntDoc = 5
    mvarsNumIntDoc = vData
End Property


Public Property Get sNumIntDoc() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNumIntDoc
    sNumIntDoc = mvarsNumIntDoc
End Property



Public Property Let iNumIntDoc(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iNumIntDoc = 5
    mvariNumIntDoc = vData
End Property


Public Property Get iNumIntDoc() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iNumIntDoc
    iNumIntDoc = mvariNumIntDoc
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



