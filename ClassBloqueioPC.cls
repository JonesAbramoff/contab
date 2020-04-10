VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassBloqueioPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Private mvariFilialEmpresa As Integer 'local copy
Private mvarlPedCompras As Long 'local copy
Private mvariSequencial As Integer 'local copy
Private mvariTipoBloqueio As Integer 'local copy
Private mvarsCodUsuario As String 'local copy
Private mvarsResponsavel As String 'local copy
Private mvardtData As Date 'local copy
Private mvarsCodUsuarioLib As String 'local copy
Private mvardtDataLib As Date 'local copy
'local variable(s) to hold property value(s)
Private mvarsNomeReduzidoFornecedor As String 'local copy
Private mvardtDataEmissao As Date 'local copy
Private mvardValorPedido As Double 'local copy
Private mvarsNomeReduzidoTipoBloqueio As String 'local copy
Private mvariMarcado As Integer 'local copy

Public Property Let iMarcado(ByVal vData As Integer)
'Syntax: X.iMarcado = 5
    mvariMarcado = vData
End Property


Public Property Get iMarcado() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iMarcado
    iMarcado = mvariMarcado
End Property



Public Property Let sNomeReduzidoTipoBloqueio(ByVal vData As String)

'Syntax: X.sNomeReduzidoTipoBloqueio = 5
    mvarsNomeReduzidoTipoBloqueio = vData
End Property


Public Property Get sNomeReduzidoTipoBloqueio() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNomeReduzidoTipoBloqueio
    sNomeReduzidoTipoBloqueio = mvarsNomeReduzidoTipoBloqueio
End Property



Public Property Let dValorPedido(ByVal vData As Double)

'Syntax: X.dValorPedido = 5
    mvardValorPedido = vData
End Property


Public Property Get dValorPedido() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValorPedido
    dValorPedido = mvardValorPedido
End Property



Public Property Let dtDataEmissao(ByVal vData As Date)

'Syntax: X.dtDataEmissao = 5
    mvardtDataEmissao = vData
End Property


Public Property Get dtDataEmissao() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataEmissao
    dtDataEmissao = mvardtDataEmissao
End Property



Public Property Let sNomeReduzidoFornecedor(ByVal vData As String)

'Syntax: X.sNomeReduzidoFornecedor = 5
    mvarsNomeReduzidoFornecedor = vData
End Property


Public Property Get sNomeReduzidoFornecedor() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNomeReduzidoFornecedor
    sNomeReduzidoFornecedor = mvarsNomeReduzidoFornecedor
End Property



Public Property Let dtDataLib(ByVal vData As Date)

'Syntax: X.dtDataLib = 5
    mvardtDataLib = vData
End Property


Public Property Get dtDataLib() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataLib
    dtDataLib = mvardtDataLib
End Property



Public Property Let sCodUsuarioLib(ByVal vData As String)

'Syntax: X.sCodUsuarioLib = 5
    mvarsCodUsuarioLib = vData
End Property


Public Property Get sCodUsuarioLib() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCodUsuarioLib
    sCodUsuarioLib = mvarsCodUsuarioLib
End Property



Public Property Let dtData(ByVal vData As Date)

'Syntax: X.dtData = 5
    mvardtData = vData
End Property


Public Property Get dtData() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtData
    dtData = mvardtData
End Property



Public Property Let sResponsavel(ByVal vData As String)

'Syntax: X.sResponsavel = 5
    mvarsResponsavel = vData
End Property


Public Property Get sResponsavel() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sResponsavel
    sResponsavel = mvarsResponsavel
End Property



Public Property Let sCodUsuario(ByVal vData As String)

'Syntax: X.sCodUsuario = 5
    mvarsCodUsuario = vData
End Property


Public Property Get sCodUsuario() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCodUsuario
    sCodUsuario = mvarsCodUsuario
End Property



Public Property Let iTipoBloqueio(ByVal vData As Integer)

'Syntax: X.iTipoBloqueio = 5
    mvariTipoBloqueio = vData
End Property


Public Property Get iTipoBloqueio() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTipoBloqueio
    iTipoBloqueio = mvariTipoBloqueio
End Property



Public Property Let iSequencial(ByVal vData As Integer)

'Syntax: X.iSequencial = 5
    mvariSequencial = vData
End Property


Public Property Get iSequencial() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iSequencial
    iSequencial = mvariSequencial
End Property



Public Property Let lPedCompras(ByVal vData As Long)

'Syntax: X.lPedCompras = 5
    mvarlPedCompras = vData
End Property


Public Property Get lPedCompras() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lPedCompras
    lPedCompras = mvarlPedCompras
End Property



Public Property Let iFilialEmpresa(ByVal vData As Integer)

'Syntax: X.iFilialEmpresa = 5
    mvariFilialEmpresa = vData
End Property


Public Property Get iFilialEmpresa() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilialEmpresa
    iFilialEmpresa = mvariFilialEmpresa
End Property



