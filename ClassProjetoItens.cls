VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProjetoItens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlNumIntDoc As Long 'local copy
Private mvarlNumIntDocProj As Long 'local copy
Private mvariSeq As Integer 'local copy
Private mvarsProduto As String 'local copy
Private mvarsVersao As String 'local copy
Private mvarsUMedida As String 'local copy
Private mvardQuantidade As Double 'local copy
Private mvardtDataMaxTermino As Date
Private mvardtDataInicioPrev As Date
Private mvardtDataTerminoPrev As Date
Private mvariDestino As Integer
Private mvardCustoTotalItem As Double
Private mvardPrecoTotalItem As Double
Private mvarlNumIntDocCusteio As Long

Public Property Let lNumIntDocCusteio(ByVal vData As Long)
    mvarlNumIntDocCusteio = vData
End Property

Public Property Get lNumIntDocCusteio() As Long
    lNumIntDocCusteio = mvarlNumIntDocCusteio
End Property

Public Property Let dtDataMaxTermino(ByVal vData As Date)
    mvardtDataMaxTermino = vData
End Property

Public Property Get dtDataMaxTermino() As Date
    dtDataMaxTermino = mvardtDataMaxTermino
End Property

Public Property Let dtDataInicioPrev(ByVal vData As Date)
    mvardtDataInicioPrev = vData
End Property

Public Property Get dtDataInicioPrev() As Date
    dtDataInicioPrev = mvardtDataInicioPrev
End Property

Public Property Let dtDataTerminoPrev(ByVal vData As Date)
    mvardtDataTerminoPrev = vData
End Property

Public Property Get dtDataTerminoPrev() As Date
    dtDataTerminoPrev = mvardtDataTerminoPrev
End Property

Public Property Let iDestino(ByVal vData As Integer)
    mvariDestino = vData
End Property

Public Property Get iDestino() As Integer
    iDestino = mvariDestino
End Property

Public Property Let dCustoTotalItem(ByVal vData As Double)
    mvardCustoTotalItem = vData
End Property

Public Property Get dCustoTotalItem() As Double
    dCustoTotalItem = mvardCustoTotalItem
End Property

Public Property Let dPrecoTotalItem(ByVal vData As Double)
    mvardPrecoTotalItem = vData
End Property

Public Property Get dPrecoTotalItem() As Double
    dPrecoTotalItem = mvardPrecoTotalItem
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

Public Property Let sProduto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sProduto = 5
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sProduto
    sProduto = mvarsProduto
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

Public Property Let lNumIntDocProj(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntDocProj = 5
    mvarlNumIntDocProj = vData
End Property

Public Property Get lNumIntDocProj() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntDocProj
    lNumIntDocProj = mvarlNumIntDocProj
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


