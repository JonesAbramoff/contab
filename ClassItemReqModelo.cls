VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemReqModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlNumIntDoc As Long 'local copy
Private mvarsProduto As String 'local copy
Private mvarsDescProduto As String 'local copy
Private mvardQuantidade As Double 'local copy
Private mvarsUM As String 'local copy
Private mvarsCcl As String 'local copy
Private mvariAlmoxarifado As Integer 'local copy
Private mvarsContaContabil As String 'local copy
Private mvariCreditaICMS As Integer 'local copy
Private mvariCreditaIPI As Integer 'local copy
Private mvarsObservacao As String 'local copy
Private mvarlFornecedor As Long 'local copy
Private mvariFilial As Integer 'local copy
Private mvariExclusivo As Integer 'local copy
Private mvarlObservacao As Long 'local copy
Private mvariTipoTributacao As Integer 'local copy
Public Property Let iTipoTributacao(ByVal vData As Integer)
    mvariTipoTributacao = vData
End Property


Public Property Get iTipoTributacao() As Integer
    iTipoTributacao = mvariTipoTributacao
End Property


Public Property Let lObservacao(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lObservacao = 5
    mvarlObservacao = vData
End Property


Public Property Get lObservacao() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lObservacao
    lObservacao = mvarlObservacao
End Property



Public Property Let iExclusivo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iExclusivo = 5
    mvariExclusivo = vData
End Property


Public Property Get iExclusivo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iExclusivo
    iExclusivo = mvariExclusivo
End Property



Public Property Let iFilial(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFilial = 5
    mvariFilial = vData
End Property


Public Property Get iFilial() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilial
    iFilial = mvariFilial
End Property



Public Property Let lFornecedor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lFornecedor = 5
    mvarlFornecedor = vData
End Property


Public Property Get lFornecedor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lFornecedor
    lFornecedor = mvarlFornecedor
End Property



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



Public Property Let iCreditaIPI(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iCreditaIPI = 5
    mvariCreditaIPI = vData
End Property


Public Property Get iCreditaIPI() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iCreditaIPI
    iCreditaIPI = mvariCreditaIPI
End Property



Public Property Let iCreditaICMS(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iCreditaICMS = 5
    mvariCreditaICMS = vData
End Property


Public Property Get iCreditaICMS() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iCreditaICMS
    iCreditaICMS = mvariCreditaICMS
End Property



Public Property Let sContaContabil(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sContaContabil = 5
    mvarsContaContabil = vData
End Property


Public Property Get sContaContabil() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sContaContabil
    sContaContabil = mvarsContaContabil
End Property



Public Property Let iAlmoxarifado(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iAlmoxarifado = 5
    mvariAlmoxarifado = vData
End Property


Public Property Get iAlmoxarifado() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iAlmoxarifado
    iAlmoxarifado = mvariAlmoxarifado
End Property



Public Property Let sCcl(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCcl = 5
    mvarsCcl = vData
End Property


Public Property Get sCcl() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCcl
    sCcl = mvarsCcl
End Property



Public Property Let sUM(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUM = 5
    mvarsUM = vData
End Property


Public Property Get sUM() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUM
    sUM = mvarsUM
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



Public Property Let sDescProduto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sDescProduto = 5
    mvarsDescProduto = vData
End Property


Public Property Get sDescProduto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sDescProduto
    sDescProduto = mvarsDescProduto
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



