VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemOPCarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlNumIntItemOP As Long 'local copy
Private mvarsProdutoBase As String 'local copy
Private mvardKgProdBase As Double 'local copy
Private mvardKgCarga As Double 'local copy
Private mvardQtdCarga As Double 'local copy
Private mvardKgKit As Double 'local copy
Private mvarcolInsumos As New Collection 'local copy
'local variable(s) to hold property value(s)
Private mvariItem As Integer 'local copy
Public Property Let iItem(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iItem = 5
    mvariItem = vData
End Property


Public Property Get iItem() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iItem
    iItem = mvariItem
End Property



Public Property Set colInsumos(ByVal vData As Collection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colInsumos = Form1
    Set mvarcolInsumos = vData
End Property


Public Property Get colInsumos() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colInsumos
    Set colInsumos = mvarcolInsumos
End Property

Public Property Let dKgKit(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dKgKit = 5
    mvardKgKit = vData
End Property

Public Property Get dKgKit() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dKgKit
    dKgKit = mvardKgKit
End Property

Public Property Let dQtdCarga(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQtdCarga = 5
    mvardQtdCarga = vData
End Property

Public Property Get dQtdCarga() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQtdCarga
    dQtdCarga = mvardQtdCarga
End Property



Public Property Let dKgCarga(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dKgCarga = 5
    mvardKgCarga = vData
End Property


Public Property Get dKgCarga() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dKgCarga
    dKgCarga = mvardKgCarga
End Property



Public Property Let dKgProdBase(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dKgProdBase = 5
    mvardKgProdBase = vData
End Property


Public Property Get dKgProdBase() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dKgProdBase
    dKgProdBase = mvardKgProdBase
End Property



Public Property Let sProdutoBase(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sProdutoBase = 5
    mvarsProdutoBase = vData
End Property


Public Property Get sProdutoBase() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sProdutoBase
    sProdutoBase = mvarsProdutoBase
End Property



Public Property Let lNumIntItemOP(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntItemOP = 5
    mvarlNumIntItemOP = vData
End Property


Public Property Get lNumIntItemOP() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntItemOP
    lNumIntItemOP = mvarlNumIntItemOP
End Property



