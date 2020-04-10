VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassEmpenho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarlCodigo As Long
Private mvarsCodigoOP As String
Private mvariItemOP As Integer
Private mvarsProduto As String
Private mvardQuantidade As Double
Private mvardQuantidadeRequisitada As Double
Private mvarsUnidadeMed As String
Private mvarlNumIntDocItemOP As Long
Private mvariAlmoxarifado As Integer
Private mvardtData As Date

Public Property Let sCodigoOP(ByVal vData As String)
    mvarsCodigoOP = vData
End Property

Public Property Get sCodigoOP() As String
    sCodigoOP = mvarsCodigoOP
End Property

Public Property Let iItemOP(ByVal vData As Integer)
    mvariItemOP = vData
End Property

Public Property Get iItemOP() As Integer
    iItemOP = mvariItemOP
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let iAlmoxarifado(ByVal vData As Integer)
    mvariAlmoxarifado = vData
End Property

Public Property Get iAlmoxarifado() As Integer
    iAlmoxarifado = mvariAlmoxarifado
End Property

Public Property Let lNumIntDocItemOP(ByVal vData As Long)
    mvarlNumIntDocItemOP = vData
End Property

Public Property Get lNumIntDocItemOP() As Long
    lNumIntDocItemOP = mvarlNumIntDocItemOP
End Property

Public Property Let sUnidadeMed(ByVal vData As String)
    mvarsUnidadeMed = vData
End Property

Public Property Get sUnidadeMed() As String
    sUnidadeMed = mvarsUnidadeMed
End Property

Public Property Let dQuantidadeRequisitada(ByVal vData As Double)
    mvardQuantidadeRequisitada = vData
End Property

Public Property Get dQuantidadeRequisitada() As Double
    dQuantidadeRequisitada = mvardQuantidadeRequisitada
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
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



