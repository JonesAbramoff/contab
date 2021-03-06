VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProjetoMaoDeObra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlNumIntDoc As Long 'local copy
Private mvarlNumIntDocCusteioItem As Long 'local copy
Private mvariSeq As Integer 'local copy
Private mvariCodMO As Integer 'local copy
Private mvarsUMedida As String 'local copy
Private mvardQuantidade As Double 'local copy
Private mvardCustoUnitarioCalculado As Double 'local copy
Private mvardCustoUnitarioInformado As Double 'local copy
Private mvarsObservacao As String 'local copy
Public Property Let sObservacao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sObservacao = 5
    mvarsObservacao = vData
End Property


Public Property Get sObservacao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sObservacao
    sObservacao = mvarsObservacao
End Property



Public Property Let dCustoUnitarioInformado(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dCustoUnitarioInformado = 5
    mvardCustoUnitarioInformado = vData
End Property


Public Property Get dCustoUnitarioInformado() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dCustoUnitarioInformado
    dCustoUnitarioInformado = mvardCustoUnitarioInformado
End Property



Public Property Let dCustoUnitarioCalculado(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dCustoUnitarioCalculado = 5
    mvardCustoUnitarioCalculado = vData
End Property


Public Property Get dCustoUnitarioCalculado() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dCustoUnitarioCalculado
    dCustoUnitarioCalculado = mvardCustoUnitarioCalculado
End Property



Public Property Let dQuantidade(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantidade = 5
    mvardQuantidade = vData
End Property


Public Property Get dQuantidade() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantidade
    dQuantidade = mvardQuantidade
End Property



Public Property Let sUMedida(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUMedida = 5
    mvarsUMedida = vData
End Property


Public Property Get sUMedida() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUMedida
    sUMedida = mvarsUMedida
End Property



Public Property Let iCodMO(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iCodMO = 5
    mvariCodMO = vData
End Property


Public Property Get iCodMO() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iCodMO
    iCodMO = mvariCodMO
End Property



Public Property Let iSeq(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iSeq = 5
    mvariSeq = vData
End Property


Public Property Get iSeq() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iSeq
    iSeq = mvariSeq
End Property



Public Property Let lNumIntDocCusteioItem(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntDocCusteioItem = 5
    mvarlNumIntDocCusteioItem = vData
End Property


Public Property Get lNumIntDocCusteioItem() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntDocCusteioItem
    lNumIntDocCusteioItem = mvarlNumIntDocCusteioItem
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




