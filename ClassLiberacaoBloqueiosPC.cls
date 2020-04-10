VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassLiberacaoBloqueiosPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarcolTipoBloqueio As New Collection 'local copy
Private mvarlPedComprasDe As Long 'local copy
Private mvarlPedComprasAte As Long 'local copy
Private mvardtBloqueioDe As Date 'local copy
Private mvardtBloqueioAte As Date 'local copy
Private mvarsOrdenacao As String 'local copy
Private mvarcolBloqueioPC As New Collection 'local copy
Public Property Set colBloqueioPC(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colBloqueioPC = Form1
    Set mvarcolBloqueioPC = vData
End Property


Public Property Get colBloqueioPC() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colBloqueioPC
    Set colBloqueioPC = mvarcolBloqueioPC
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



Public Property Let dtBloqueioAte(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtBloqueioAte = 5
    mvardtBloqueioAte = vData
End Property


Public Property Get dtBloqueioAte() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtBloqueioAte
    dtBloqueioAte = mvardtBloqueioAte
End Property



Public Property Let dtBloqueioDe(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtBloqueioDe = 5
    mvardtBloqueioDe = vData
End Property


Public Property Get dtBloqueioDe() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtBloqueioDe
    dtBloqueioDe = mvardtBloqueioDe
End Property



Public Property Let lPedComprasAte(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lPedComprasAte = 5
    mvarlPedComprasAte = vData
End Property


Public Property Get lPedComprasAte() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lPedComprasAte
    lPedComprasAte = mvarlPedComprasAte
End Property



Public Property Let lPedComprasDe(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lPedComprasDe = 5
    mvarlPedComprasDe = vData
End Property


Public Property Get lPedComprasDe() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lPedComprasDe
    lPedComprasDe = mvarlPedComprasDe
End Property



Public Property Set colTipoBloqueio(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colTipoBloqueio = Form1
    Set mvarcolTipoBloqueio = vData
End Property


Public Property Get colTipoBloqueio() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colTipoBloqueio
    Set colTipoBloqueio = mvarcolTipoBloqueio
End Property



