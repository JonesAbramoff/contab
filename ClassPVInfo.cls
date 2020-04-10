VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPVInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlCliente As Long 'local copy
Private mvarsClienteNomeReduzido As String 'local copy
Private mvariFilialCliente As Integer 'local copy
Private mvariMarcada As Integer 'local copy
Private mvardtEmissao As Date 'local copy
Private mvardValorTotal As Double 'local copy
Private mvarlNumIntDoc As Long 'local copy
Private mvarlCodPedido As Long 'local copy
Private mvardtEntrega As Date 'local copy
Public Property Let dtEntrega(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtEntrega = 5
    mvardtEntrega = vData
End Property


Public Property Get dtEntrega() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtEntrega
    dtEntrega = mvardtEntrega
End Property



Public Property Let lCodPedido(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lCodPedido = 5
    mvarlCodPedido = vData
End Property


Public Property Get lCodPedido() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lCodPedido
    lCodPedido = mvarlCodPedido
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



Public Property Let dValorTotal(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValorTotal = 5
    mvardValorTotal = vData
End Property


Public Property Get dValorTotal() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValorTotal
    dValorTotal = mvardValorTotal
End Property



Public Property Let dtEmissao(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtEmissao = 5
    mvardtEmissao = vData
End Property


Public Property Get dtEmissao() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtEmissao
    dtEmissao = mvardtEmissao
End Property



Public Property Let iMarcada(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iMarcada = 5
    mvariMarcada = vData
End Property


Public Property Get iMarcada() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iMarcada
    iMarcada = mvariMarcada
End Property



Public Property Let iFilialCliente(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFilialCliente = 5
    mvariFilialCliente = vData
End Property


Public Property Get iFilialCliente() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilialCliente
    iFilialCliente = mvariFilialCliente
End Property



Public Property Let sClienteNomeReduzido(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sClienteNomeReduzido = 5
    mvarsClienteNomeReduzido = vData
End Property


Public Property Get sClienteNomeReduzido() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sClienteNomeReduzido
    sClienteNomeReduzido = mvarsClienteNomeReduzido
End Property



Public Property Let lCliente(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lCliente = 5
    mvarlCliente = vData
End Property


Public Property Get lCliente() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lCliente
    lCliente = mvarlCliente
End Property



