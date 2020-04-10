VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPrevVendaPrevConsumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvardQtdEstoqueAux As Double 'local copy
Private mvarsUMEstoque As String 'local copy
Private mvarsProduto As String 'local copy
Private mvariMes As Integer 'local copy
Private mvariAno As Integer 'local copy
Private mvariFilialEmpresa As Integer 'local copy
Private mvarsCodigoPrevVenda As String 'local copy
Private mvardQuantPrevVenda As Double 'local copy
Private mvardQuantPrevInsumo As Double 'local copy
Private mvardQuantEstoque As Double 'local copy
Private mvardtData As Date 'local copy
Private mvariAlterado As Integer 'local copy
Private mvarsVersao As String 'local copy

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



Public Property Let dQuantEstoque(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantEstoque = 5
    mvardQuantEstoque = vData
End Property


Public Property Get dQuantEstoque() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantEstoque
    dQuantEstoque = mvardQuantEstoque
End Property



Public Property Let dQuantPrevInsumo(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantPrevInsumo = 5
    mvardQuantPrevInsumo = vData
End Property


Public Property Get dQuantPrevInsumo() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantPrevInsumo
    dQuantPrevInsumo = mvardQuantPrevInsumo
End Property



Public Property Let dQuantPrevVenda(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantPrevVenda = 5
    mvardQuantPrevVenda = vData
End Property


Public Property Get dQuantPrevVenda() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantPrevVenda
    dQuantPrevVenda = mvardQuantPrevVenda
End Property



Public Property Let sCodigoPrevVenda(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCodigoPrevVenda = 5
    mvarsCodigoPrevVenda = vData
End Property


Public Property Get sCodigoPrevVenda() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCodigoPrevVenda
    sCodigoPrevVenda = mvarsCodigoPrevVenda
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

Public Property Let iAno(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iAno = 5
    mvariAno = vData
End Property


Public Property Get iAno() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iAno
    iAno = mvariAno
End Property



Public Property Let iMes(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iMes = 5
    mvariMes = vData
End Property


Public Property Get iMes() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iMes
    iMes = mvariMes
End Property



Public Property Let sProduto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCodProduto = 5
    mvarsProduto = vData
End Property


Public Property Get sProduto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCodProduto
    sProduto = mvarsProduto
End Property

Public Property Let dQtdEstoqueAux(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQtdEstoque = 5
    mvardQtdEstoqueAux = vData
End Property


Public Property Get dQtdEstoqueAux() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQtdEstoque
    dQtdEstoqueAux = mvardQtdEstoqueAux
End Property


Public Property Let sUMEstoque(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUMEstoque = 5
    mvarsUMEstoque = vData
End Property


Public Property Get sUMEstoque() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUMEstoque
    sUMEstoque = mvarsUMEstoque
End Property


Public Property Let iAlterado(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFilialEmpresa = 5
    mvariAlterado = vData
End Property


Public Property Get iAlterado() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilialEmpresa
    iAlterado = mvariAlterado
End Property

Public Property Let sVersao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sVersao = 5
    mvarsVersao = vData
End Property

Public Property Get sVersao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sVersao
    sVersao = mvarsVersao
End Property

