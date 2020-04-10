VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBaixaPedCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvariSoResiduais As Integer 'local copy
Private mvardtDataDe As Date 'local copy
Private mvardtDataAte As Date 'local copy
Private mvarsOrdenacao As String 'local copy
Private mvarcolPedCompras As New Collection 'local copy
Private mvardtDataEnvioAte As Date 'local copy
Private mvardtDataEnvioDe As Date 'local copy
Private mvarlPedCompraAte As Long 'local copy
Private mvarlPedCompraDe As Long 'local copy
Private mvarlFornecedorAte As Long 'local copy
Private mvarlFornecedorDe As Long 'local copy
Public Property Let lFornecedorDe(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lFornecedorDe = 5
    mvarlFornecedorDe = vData
End Property


Public Property Get lFornecedorDe() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lFornecedorDe
    lFornecedorDe = mvarlFornecedorDe
End Property



Public Property Let lFornecedorAte(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lFornecedorAte = 5
    mvarlFornecedorAte = vData
End Property


Public Property Get lFornecedorAte() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lFornecedorAte
    lFornecedorAte = mvarlFornecedorAte
End Property



Public Property Let lPedCompraDe(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lPedCompraDe = 5
    mvarlPedCompraDe = vData
End Property


Public Property Get lPedCompraDe() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lPedCompraDe
    lPedCompraDe = mvarlPedCompraDe
End Property



Public Property Let lPedCompraAte(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lPedCompraAte = 5
    mvarlPedCompraAte = vData
End Property


Public Property Get lPedCompraAte() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lPedCompraAte
    lPedCompraAte = mvarlPedCompraAte
End Property



Public Property Let dtDataEnvioDe(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataEnvioDe = 5
    mvardtDataEnvioDe = vData
End Property


Public Property Get dtDataEnvioDe() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataEnvioDe
    dtDataEnvioDe = mvardtDataEnvioDe
End Property



Public Property Let dtDataEnvioAte(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataEnvioAte = 5
    mvardtDataEnvioAte = vData
End Property


Public Property Get dtDataEnvioAte() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataEnvioAte
    dtDataEnvioAte = mvardtDataEnvioAte
End Property



Public Property Set colPedCompras(ByVal vData As Object)
Attribute colPedCompras.VB_Description = "coleção "
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colPedCompras = Form1
    Set mvarcolPedCompras = vData
End Property


Public Property Get colPedCompras() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colPedCompras
    Set colPedCompras = mvarcolPedCompras
End Property



Public Property Let sOrdenacao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sOrdenacao = 5
    mvarsOrdenacao = vData
End Property


Public Property Get sOrdenacao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sOrdenacao
    sOrdenacao = mvarsOrdenacao
End Property



Public Property Let dtDataAte(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataAte = 5
    mvardtDataAte = vData
End Property


Public Property Get dtDataAte() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataAte
    dtDataAte = mvardtDataAte
End Property



Public Property Let dtDataDe(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataDe = 5
    mvardtDataDe = vData
End Property


Public Property Get dtDataDe() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataDe
    dtDataDe = mvardtDataDe
End Property



Public Property Let iSoResiduais(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iSoResiduais = 5
    mvariSoResiduais = vData
End Property


Public Property Get iSoResiduais() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iSoResiduais
    iSoResiduais = mvariSoResiduais
End Property



