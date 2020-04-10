VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProjetoCusteioItens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlNumIntDoc As Long 'local copy
Private mvarlNumIntDocCusteio As Long 'local copy
Private mvarlNumIntDocProjetoItem As Long 'local copy
Private mvardCustoTotalInsumosKit As Double 'local copy
Private mvardCustoTotalInsumosMaq As Double 'local copy
Private mvardCustoTotalMaoDeObra As Double 'local copy
Private mvardPrecoTotalItem As Double 'local copy
Private mvarcolProjetoInsumosKit As New Collection
Private mvarcolProjetoInsumosMaquina As New Collection
Private mvarcolProjetoMaoDeObra As New Collection

Public Property Set colProjetoInsumosKit(ByVal vData As Collection)
    Set mvarcolProjetoInsumosKit = vData
End Property

Public Property Get colProjetoInsumosKit() As Collection
    Set colProjetoInsumosKit = mvarcolProjetoInsumosKit
End Property

Public Property Set colProjetoInsumosMaquina(ByVal vData As Collection)
    Set mvarcolProjetoInsumosMaquina = vData
End Property

Public Property Get colProjetoInsumosMaquina() As Collection
    Set colProjetoInsumosMaquina = mvarcolProjetoInsumosMaquina
End Property

Public Property Set colProjetoMaoDeObra(ByVal vData As Collection)
    Set mvarcolProjetoMaoDeObra = vData
End Property

Public Property Get colProjetoMaoDeObra() As Collection
    Set colProjetoMaoDeObra = mvarcolProjetoMaoDeObra
End Property

Public Property Let lNumIntDocProjetoItem(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntDocProjetoItem = 5
    mvarlNumIntDocProjetoItem = vData
End Property

Public Property Get lNumIntDocProjetoItem() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntDocProjetoItem
    lNumIntDocProjetoItem = mvarlNumIntDocProjetoItem
End Property

Public Property Let dPrecoTotalItem(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPrecoTotalItem = 5
    mvardPrecoTotalItem = vData
End Property

Public Property Get dPrecoTotalItem() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPrecoTotalItem
    dPrecoTotalItem = mvardPrecoTotalItem
End Property

Public Property Let dCustoTotalInsumosKit(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dCustoTotalInsumosKit = 5
    mvardCustoTotalInsumosKit = vData
End Property

Public Property Get dCustoTotalInsumosKit() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dCustoTotalInsumosKit
    dCustoTotalInsumosKit = mvardCustoTotalInsumosKit
End Property

Public Property Let dCustoTotalInsumosMaq(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dCustoTotalInsumosMaq = 5
    mvardCustoTotalInsumosMaq = vData
End Property

Public Property Get dCustoTotalInsumosMaq() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dCustoTotalInsumosMaq
    dCustoTotalInsumosMaq = mvardCustoTotalInsumosMaq
End Property

Public Property Let dCustoTotalMaoDeObra(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dCustoTotalMaoDeObra = 5
    mvardCustoTotalMaoDeObra = vData
End Property

Public Property Get dCustoTotalMaoDeObra() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dCustoTotalMaoDeObra
    dCustoTotalMaoDeObra = mvardCustoTotalMaoDeObra
End Property

Public Property Let lNumIntDocCusteio(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntDocCusteio = 5
    mvarlNumIntDocCusteio = vData
End Property

Public Property Get lNumIntDocCusteio() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntDocCusteio
    lNumIntDocCusteio = mvarlNumIntDocCusteio
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


