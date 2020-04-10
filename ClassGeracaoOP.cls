VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassGeracaoOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilialEmpresa As Integer

'campos p/a selecao inicial dos pedidos
Private mvariTodos As Integer
Private mvarlPVDe As Long
Private mvarlPVAte As Long
Private mvarlClientesDe As Long
Private mvarlClientesAte As Long
Private mvardtEntregaDe As Date
Private mvardtEntregaAte As Date
Private mvarsProdDe As String
Private mvarsProdAte As String
Private mvariIncluiPVcomOP As Integer

Private mvarsOrdenacao As String 'SQL p/ordenar os itens no frame 2 da tela
Private mvarsCodigoOP As String
Private mvardtDataOP As Date
Private mvariNumItensSelecionados As Integer 'na colecao abaixo
Private mvarcolItens As New Collection 'colecao de ClassGeracaoOPItem

Public Property Set colItens(ByVal vData As Object)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Let iNumItensSelecionados(ByVal vData As Integer)
    mvariNumItensSelecionados = vData
End Property

Public Property Get iNumItensSelecionados() As Integer
    iNumItensSelecionados = mvariNumItensSelecionados
End Property

Public Property Let dtDataOP(ByVal vData As Date)
    mvardtDataOP = vData
End Property

Public Property Get dtDataOP() As Date
    dtDataOP = mvardtDataOP
End Property

Public Property Let sCodigoOP(ByVal vData As String)
    mvarsCodigoOP = vData
End Property

Public Property Get sCodigoOP() As String
    sCodigoOP = mvarsCodigoOP
End Property

Public Property Let sOrdenacao(ByVal vData As String)
    mvarsOrdenacao = vData
End Property

Public Property Get sOrdenacao() As String
    sOrdenacao = mvarsOrdenacao
End Property

Public Property Let iIncluiPVcomOP(ByVal vData As Integer)
    mvariIncluiPVcomOP = vData
End Property

Public Property Get iIncluiPVcomOP() As Integer
    iIncluiPVcomOP = mvariIncluiPVcomOP
End Property

Public Property Let sProdAte(ByVal vData As String)
    mvarsProdAte = vData
End Property

Public Property Get sProdAte() As String
    sProdAte = mvarsProdAte
End Property

Public Property Let sProdDe(ByVal vData As String)
    mvarsProdDe = vData
End Property

Public Property Get sProdDe() As String
    sProdDe = mvarsProdDe
End Property

Public Property Let dtEntregaAte(ByVal vData As Date)
    mvardtEntregaAte = vData
End Property

Public Property Get dtEntregaAte() As Date
    dtEntregaAte = mvardtEntregaAte
End Property

Public Property Let dtEntregaDe(ByVal vData As Date)
    mvardtEntregaDe = vData
End Property

Public Property Get dtEntregaDe() As Date
    dtEntregaDe = mvardtEntregaDe
End Property

Public Property Let lClientesAte(ByVal vData As Long)
    mvarlClientesAte = vData
End Property

Public Property Get lClientesAte() As Long
    lClientesAte = mvarlClientesAte
End Property

Public Property Let lClientesDe(ByVal vData As Long)
    mvarlClientesDe = vData
End Property

Public Property Get lClientesDe() As Long
    lClientesDe = mvarlClientesDe
End Property

Public Property Let lPVAte(ByVal vData As Long)
    mvarlPVAte = vData
End Property

Public Property Get lPVAte() As Long
    lPVAte = mvarlPVAte
End Property

Public Property Let lPVDe(ByVal vData As Long)
    mvarlPVDe = vData
End Property

Public Property Get lPVDe() As Long
    lPVDe = mvarlPVDe
End Property

Public Property Let iTodos(ByVal vData As Integer)
    mvariTodos = vData
End Property

Public Property Get iTodos() As Integer
    iTodos = mvariTodos
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property



