VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTecladoProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvarsDescricao As String
Private mvariPadrao As Integer
Private mvariFilialEmpresa As Integer
Private mvariTeclado As Integer
Private mvarcolTecladoProdutoItem As New Collection

Public Property Set colTecladoProdutoItem(ByVal vData As Collection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colTecladoProduto = Form1
    Set mvarcolTecladoProdutoItem = vData
End Property

Public Property Get colTecladoProdutoItem() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colTecladoProduto
    Set colTecladoProdutoItem = mvarcolTecladoProdutoItem
End Property


Public Property Let iFilialEmpresa(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFilialEmpresa = 5
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilialEmpresa
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iPadrao(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iPadrao = 5
    mvariPadrao = vData
End Property

Public Property Get iPadrao() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iPadrao
    iPadrao = mvariPadrao
End Property

Public Property Let sDescricao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sDescricao = 5
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sDescricao
    sDescricao = mvarsDescricao
End Property

Public Property Let iCodigo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCodigo = 5
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCodigo
    iCodigo = mvariCodigo
End Property

Public Property Let iTeclado(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.i = 5
    mvariTeclado = vData
End Property

Public Property Get iTeclado() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.i
    iTeclado = mvariTeclado
End Property

