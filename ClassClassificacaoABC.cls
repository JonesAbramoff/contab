VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassClassificacaoABC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colProdutoDemanda"
Option Explicit

'local variable(s) to hold property value(s)
Private mvariFilialEmpresa As Integer 'local copy
Private mvarlNumInt As Long 'local copy
Private mvarsCodigo As String 'local copy
Private mvarsDescricao As String 'local copy
Private mvardtData As Date 'local copy
Private mvariMesInicial As Integer 'local copy
Private mvariAnoInicial As Integer 'local copy
Private mvariMesFinal As Integer 'local copy
Private mvariFaixaA As Integer 'local copy
Private mvariFaixaB As Integer 'local copy
Private mvariTipoProduto As Integer 'local copy
Private mvardDemandaTotal As Double 'local copy
Private mvariAtualizaProdutosFilial As Integer 'local copy
Private mvariAnoFinal As Integer 'local copy
Public Property Let iAnoFinal(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iAnoFinal = 5
    mvariAnoFinal = vData
End Property


Public Property Get iAnoFinal() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iAnoFinal
    iAnoFinal = mvariAnoFinal
End Property



Public Property Let iAtualizaProdutosFilial(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iAtualizaProdutosFilial = 5
    mvariAtualizaProdutosFilial = vData
End Property


Public Property Get iAtualizaProdutosFilial() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iAtualizaProdutosFilial
    iAtualizaProdutosFilial = mvariAtualizaProdutosFilial
End Property




Public Property Let dDemandaTotal(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dDemandaTotal = 5
    mvardDemandaTotal = vData
End Property


Public Property Get dDemandaTotal() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dDemandaTotal
    dDemandaTotal = mvardDemandaTotal
End Property



Public Property Let iTipoProduto(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTipoProduto = 5
    mvariTipoProduto = vData
End Property


Public Property Get iTipoProduto() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTipoProduto
    iTipoProduto = mvariTipoProduto
End Property



Public Property Let iFaixaB(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFaixaB = 5
    mvariFaixaB = vData
End Property


Public Property Get iFaixaB() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFaixaB
    iFaixaB = mvariFaixaB
End Property



Public Property Let iFaixaA(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFaixaA = 5
    mvariFaixaA = vData
End Property


Public Property Get iFaixaA() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFaixaA
    iFaixaA = mvariFaixaA
End Property



Public Property Let iMesFinal(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iMesFinal = 5
    mvariMesFinal = vData
End Property


Public Property Get iMesFinal() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iMesFinal
    iMesFinal = mvariMesFinal
End Property



Public Property Let iAnoInicial(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iAnoInicial = 5
    mvariAnoInicial = vData
End Property


Public Property Get iAnoInicial() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iAnoInicial
    iAnoInicial = mvariAnoInicial
End Property



Public Property Let iMesInicial(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iMesInicial = 5
    mvariMesInicial = vData
End Property


Public Property Get iMesInicial() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iMesInicial
    iMesInicial = mvariMesInicial
End Property



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



Public Property Let sCodigo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCodigo = 5
    mvarsCodigo = vData
End Property


Public Property Get sCodigo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCodigo
    sCodigo = mvarsCodigo
End Property



Public Property Let lNumInt(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumInt = 5
    mvarlNumInt = vData
End Property


Public Property Get lNumInt() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumInt
    lNumInt = mvarlNumInt
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



