VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassParcelaOV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvardDesconto1Valor As Double 'local copy
Private mvardDesconto2Valor As Double 'local copy
Private mvardDesconto3Valor As Double 'local copy
Private mvardtDataVencimento As Date 'local copy
Private mvardtDesconto1Ate As Date 'local copy
Private mvardtDesconto2Ate As Date 'local copy
Private mvardtDesconto3Ate As Date 'local copy
Private mvardValor As Double 'local copy
Private mvariDesconto2Codigo As Integer 'local copy
Private mvariDesconto1Codigo As Integer 'local copy
Private mvariDesconto3Codigo As Integer 'local copy
Private mvariNumParcela As Integer 'local copy
Public Property Let iNumParcela(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iNumParcela = 5
    mvariNumParcela = vData
End Property


Public Property Get iNumParcela() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iNumParcela
    iNumParcela = mvariNumParcela
End Property



Public Property Let iDesconto3Codigo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iDesconto3Codigo = 5
    mvariDesconto3Codigo = vData
End Property


Public Property Get iDesconto3Codigo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iDesconto3Codigo
    iDesconto3Codigo = mvariDesconto3Codigo
End Property



Public Property Let iDesconto1Codigo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iDesconto1Codigo = 5
    mvariDesconto1Codigo = vData
End Property


Public Property Get iDesconto1Codigo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iDesconto1Codigo
    iDesconto1Codigo = mvariDesconto1Codigo
End Property



Public Property Let iDesconto2Codigo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iDesconto2Codigo = 5
    mvariDesconto2Codigo = vData
End Property


Public Property Get iDesconto2Codigo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iDesconto2Codigo
    iDesconto2Codigo = mvariDesconto2Codigo
End Property



Public Property Let dValor(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValor = 5
    mvardValor = vData
End Property


Public Property Get dValor() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValor
    dValor = mvardValor
End Property



Public Property Let dtDesconto3Ate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDesconto3Ate = 5
    mvardtDesconto3Ate = vData
End Property


Public Property Get dtDesconto3Ate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDesconto3Ate
    dtDesconto3Ate = mvardtDesconto3Ate
End Property



Public Property Let dtDesconto2Ate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDesconto2Ate = 5
    mvardtDesconto2Ate = vData
End Property


Public Property Get dtDesconto2Ate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDesconto2Ate
    dtDesconto2Ate = mvardtDesconto2Ate
End Property



Public Property Let dtDesconto1Ate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDesconto1Ate = 5
    mvardtDesconto1Ate = vData
End Property


Public Property Get dtDesconto1Ate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDesconto1Ate
    dtDesconto1Ate = mvardtDesconto1Ate
End Property



Public Property Let dtDataVencimento(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataVencimento = 5
    mvardtDataVencimento = vData
End Property


Public Property Get dtDataVencimento() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataVencimento
    dtDataVencimento = mvardtDataVencimento
End Property



Public Property Let dDesconto3Valor(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dDesconto3Valor = 5
    mvardDesconto3Valor = vData
End Property


Public Property Get dDesconto3Valor() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dDesconto3Valor
    dDesconto3Valor = mvardDesconto3Valor
End Property



Public Property Let dDesconto2Valor(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dDesconto2Valor = 5
    mvardDesconto2Valor = vData
End Property


Public Property Get dDesconto2Valor() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dDesconto2Valor
    dDesconto2Valor = mvardDesconto2Valor
End Property



Public Property Let dDesconto1Valor(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dDesconto1Valor = 5
    mvardDesconto1Valor = vData
End Property


Public Property Get dDesconto1Valor() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dDesconto1Valor
    dDesconto1Valor = mvardDesconto1Valor
End Property



