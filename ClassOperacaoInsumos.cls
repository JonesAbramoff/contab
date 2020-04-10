VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassOperacaoInsumos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlNumIntDocOper As Long 'local copy
Private mvarsProduto As String 'local copy
Private mvardQuantidade As Double 'local copy
Private mvarsUMProduto As String 'local copy
Private mvariComposicao As Integer 'local copy
Private mvardPercentualPerda As Double 'local copy
Private mvardCustoStandard As Double 'local copy
Private mvarsVersaoKitComp As String 'local copy
Public Property Let sVersaoKitComp(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sVersaoKitComp = 5
    mvarsVersaoKitComp = vData
End Property


Public Property Get sVersaoKitComp() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sVersaoKitComp
    sVersaoKitComp = mvarsVersaoKitComp
End Property



Public Property Let dCustoStandard(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dCustoStandard = 5
    mvardCustoStandard = vData
End Property


Public Property Get dCustoStandard() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dCustoStandard
    dCustoStandard = mvardCustoStandard
End Property



Public Property Let dPercentualPerda(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPercentualPerda = 5
    mvardPercentualPerda = vData
End Property


Public Property Get dPercentualPerda() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPercentualPerda
    dPercentualPerda = mvardPercentualPerda
End Property



Public Property Let iComposicao(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iComposicao = 5
    mvariComposicao = vData
End Property


Public Property Get iComposicao() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iComposicao
    iComposicao = mvariComposicao
End Property



Public Property Let sUMProduto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUMProduto = 5
    mvarsUMProduto = vData
End Property


Public Property Get sUMProduto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUMProduto
    sUMProduto = mvarsUMProduto
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



Public Property Let lNumIntDocOper(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntDocOper = 5
    mvarlNumIntDocOper = vData
End Property


Public Property Get lNumIntDocOper() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntDocOper
    lNumIntDocOper = mvarlNumIntDocOper
End Property



