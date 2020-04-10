VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBoletoParcela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlNumIntBoleto As Long
Private mvariSequencial As Integer
Private mvarsNSU As String
Private mvardValor As Double
Private mvardtDataVencimento As Date
Private mvariStatus As Integer
Private mvardValorRecebido As Double
Public Property Let dValorRecebido(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValorRecebido = 5
    mvardValorRecebido = vData
End Property


Public Property Get dValorRecebido() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValorRecebido
    dValorRecebido = mvardValorRecebido
End Property



Public Property Let iStatus(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iStatus = 5
    mvariStatus = vData
End Property


Public Property Get iStatus() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iStatus
    iStatus = mvariStatus
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



Public Property Let sNSU(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNSU = 5
    mvarsNSU = vData
End Property


Public Property Get sNSU() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNSU
    sNSU = mvarsNSU
End Property



Public Property Let iSequencial(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iSequencial = 5
    mvariSequencial = vData
End Property


Public Property Get iSequencial() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iSequencial
    iSequencial = mvariSequencial
End Property



Public Property Let lNumIntBoleto(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntBoleto = 5
    mvarlNumIntBoleto = vData
End Property


Public Property Get lNumIntBoleto() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntBoleto
    lNumIntBoleto = mvarlNumIntBoleto
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



