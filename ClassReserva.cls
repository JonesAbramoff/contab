VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassReserva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvariFilialEmpresa As Integer
Private mvarlCodigo As Long
Private mvarsProduto As String
Private mvariAlmoxarifado As Integer
Private mvariTipoDoc As Integer
Private mvarlDocOrigem As Long
Private mvarlNumIntOrigem As Long
Private mvardQuantidade As Double
Private mvardtDataReserva As Date
Private mvardtDataValidade As Date
Private mvarsCodUsuario As String
Private mvarsResponsavel As String
'local variable(s) to hold property value(s)
Private mvarsAlmoxarifado As String 'local copy
Public Property Let sAlmoxarifado(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sAlmoxarifado = 5
    mvarsAlmoxarifado = vData
End Property


Public Property Get sAlmoxarifado() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sAlmoxarifado
    sAlmoxarifado = mvarsAlmoxarifado
End Property




Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sResponsavel(ByVal vData As String)
    mvarsResponsavel = vData
End Property

Public Property Get sResponsavel() As String
    sResponsavel = mvarsResponsavel
End Property

Public Property Let sCodUsuario(ByVal vData As String)
    mvarsCodUsuario = vData
End Property

Public Property Get sCodUsuario() As String
    sCodUsuario = mvarsCodUsuario
End Property

Public Property Let dtDataValidade(ByVal vData As Date)
    mvardtDataValidade = vData
End Property

Public Property Get dtDataValidade() As Date
    dtDataValidade = mvardtDataValidade
End Property

Public Property Let dtDataReserva(ByVal vData As Date)
    mvardtDataReserva = vData
End Property

Public Property Get dtDataReserva() As Date
    dtDataReserva = mvardtDataReserva
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let lNumIntOrigem(ByVal vData As Long)
    mvarlNumIntOrigem = vData
End Property

Public Property Get lNumIntOrigem() As Long
    lNumIntOrigem = mvarlNumIntOrigem
End Property

Public Property Let lDocOrigem(ByVal vData As Long)
    mvarlDocOrigem = vData
End Property

Public Property Get lDocOrigem() As Long
    lDocOrigem = mvarlDocOrigem
End Property

Public Property Let iTipoDoc(ByVal vData As Integer)
    mvariTipoDoc = vData
End Property

Public Property Get iTipoDoc() As Integer
    iTipoDoc = mvariTipoDoc
End Property

Public Property Let iAlmoxarifado(ByVal vData As Integer)
    mvariAlmoxarifado = vData
End Property

Public Property Get iAlmoxarifado() As Integer
    iAlmoxarifado = mvariAlmoxarifado
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property



