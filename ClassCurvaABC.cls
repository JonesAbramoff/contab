VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassCurvaABC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ColPontos"
Option Explicit

Private mvarcolPontos As New ColPontos 'local copy
Private mvarlPasso As Long 'local copy
Private mvarobjPontoClasseA As New ClassPonto 'local copy
Private mvarobjPontoClasseB As New ClassPonto 'local copy
Private mvarlItensA As Long 'local copy
Private mvarlItensB As Long 'local copy
Private mvarlItensC As Long 'local copy
Private mvardDemandaTotal As Double 'local copy
Private mvarlClassifABC As Long 'local copy
Private mvariFaixaA As Integer 'local copy
Private mvariFaixaB As Integer 'local copy
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



Public Property Let lClassifABC(ByVal vData As Long)
Attribute lClassifABC.VB_Description = "Número interno da Classificação ABC"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lClassifABC = 5
    mvarlClassifABC = vData
End Property


Public Property Get lClassifABC() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lClassifABC
    lClassifABC = mvarlClassifABC
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



Public Property Let lItensC(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lItensC = 5
    mvarlItensC = vData
End Property


Public Property Get lItensC() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lItensC
    lItensC = mvarlItensC
End Property



Public Property Let lItensB(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lItensB = 5
    mvarlItensB = vData
End Property


Public Property Get lItensB() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lItensB
    lItensB = mvarlItensB
End Property



Public Property Let lItensA(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lItensA = 5
    mvarlItensA = vData
End Property


Public Property Get lItensA() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lItensA
    lItensA = mvarlItensA
End Property



Public Property Set objPontoClasseB(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.objPontoClasseB = Form1
    Set mvarobjPontoClasseB = vData
End Property


Public Property Get objPontoClasseB() As ClassPonto
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.objPontoClasseB
    Set objPontoClasseB = mvarobjPontoClasseB
End Property



Public Property Set objPontoClasseA(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.objPontoClasseA = Form1
    Set mvarobjPontoClasseA = vData
End Property


Public Property Get objPontoClasseA() As ClassPonto
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.objPontoClasseA
    Set objPontoClasseA = mvarobjPontoClasseA
End Property






Public Property Let lPasso(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lPasso = 5
    mvarlPasso = vData
End Property


Public Property Get lPasso() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lPasso
    lPasso = mvarlPasso
End Property



Public Property Set ColPontos(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colPontos = Form1
    Set mvarcolPontos = vData
End Property


Public Property Get ColPontos() As ColPontos
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colPontos
    Set ColPontos = mvarcolPontos
End Property
