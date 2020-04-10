VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBaixaPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvariTodosOsPedidos As Integer 'local copy
Private mvarlPedidosDe As Long 'local copy
Private mvarlPedidosAte As Long 'local copy
Private mvarlClientesDe As Long 'local copy
Private mvarlClientesAte As Long 'local copy
Private mvardtEmissaoDe As Date 'local copy
Private mvardtEmissaoAte As Date 'local copy
Private mvardtEntregaDe As Date 'local copy
Private mvardtEntregaAte As Date 'local copy
Private mvarsOrdenacao As String 'local copy
Private mvarcolPVInfo As New Collection 'local copy
Public Property Set colPVInfo(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colPVInfo = Form1
    Set mvarcolPVInfo = vData
End Property


Public Property Get colPVInfo() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colPVInfo
    Set colPVInfo = mvarcolPVInfo
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



Public Property Let dtEntregaAte(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtEntregaAte = 5
    mvardtEntregaAte = vData
End Property


Public Property Get dtEntregaAte() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtEntregaAte
    dtEntregaAte = mvardtEntregaAte
End Property



Public Property Let dtEntregaDe(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtEntregaDe = 5
    mvardtEntregaDe = vData
End Property


Public Property Get dtEntregaDe() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtEntregaDe
    dtEntregaDe = mvardtEntregaDe
End Property



Public Property Let dtEmissaoAte(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtEmissaoAte = 5
    mvardtEmissaoAte = vData
End Property


Public Property Get dtEmissaoAte() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtEmissaoAte
    dtEmissaoAte = mvardtEmissaoAte
End Property



Public Property Let dtEmissaoDe(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtEmissaoDe = 5
    mvardtEmissaoDe = vData
End Property


Public Property Get dtEmissaoDe() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtEmissaoDe
    dtEmissaoDe = mvardtEmissaoDe
End Property



Public Property Let lClientesAte(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lClientesAte = 5
    mvarlClientesAte = vData
End Property


Public Property Get lClientesAte() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lClientesAte
    lClientesAte = mvarlClientesAte
End Property



Public Property Let lClientesDe(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lClientesDe = 5
    mvarlClientesDe = vData
End Property


Public Property Get lClientesDe() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lClientesDe
    lClientesDe = mvarlClientesDe
End Property



Public Property Let lPedidosAte(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lPedidosAte = 5
    mvarlPedidosAte = vData
End Property


Public Property Get lPedidosAte() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lPedidosAte
    lPedidosAte = mvarlPedidosAte
End Property



Public Property Let lPedidosDe(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lPedidosDe = 5
    mvarlPedidosDe = vData
End Property


Public Property Get lPedidosDe() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lPedidosDe
    lPedidosDe = mvarlPedidosDe
End Property



Public Property Let iTodosOsPedidos(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTodosOsPedidos = 5
    mvariTodosOsPedidos = vData
End Property


Public Property Get iTodosOsPedidos() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTodosOsPedidos
    iTodosOsPedidos = mvariTodosOsPedidos
End Property



