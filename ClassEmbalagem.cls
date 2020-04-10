VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEmbalagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvariCodigo As Integer 'local copy
Private mvarsDescricao As String 'local copy
Private mvarsSigla As String 'local copy
Private mvardCapacidade As Double 'local copy
Private mvardPeso As Double 'local copy
Private mvarsProduto As String
Private mvarcolEmbExpedicao As New Collection

Public Property Set colEmbExpedicao(ByVal vData As Object)
    Set mvarcolEmbExpedicao = vData
End Property

Public Property Get colEmbExpedicao() As Collection
    Set colEmbExpedicao = mvarcolEmbExpedicao
End Property

Public Property Let dPeso(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPeso = 5
    mvardPeso = vData
End Property


Public Property Get dPeso() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPeso
    dPeso = mvardPeso
End Property



Public Property Let dCapacidade(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dCapacidade = 5
    mvardCapacidade = vData
End Property


Public Property Get dCapacidade() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dCapacidade
    dCapacidade = mvardCapacidade
End Property



Public Property Let sSigla(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSigla = 5
    mvarsSigla = vData
End Property


Public Property Get sSigla() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSigla
    sSigla = mvarsSigla
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
'Syntax: X.iCodigo = 5
    mvariCodigo = vData
End Property


Public Property Get iCodigo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iCodigo
    iCodigo = mvariCodigo
End Property

Public Property Let sProduto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sDescricao = 5
    mvarsProduto = vData
End Property


Public Property Get sProduto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sDescricao
    sProduto = mvarsProduto
End Property


