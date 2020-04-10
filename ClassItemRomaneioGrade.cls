VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemRomaneioGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsProduto As String
Private mvarsUMEstoque As String
Private mvarcolCategoria As New Collection
Private mvardQuantidade As Double
Private mvardQuantReservada As Double
Private mvardQuantCancelada As Double
Private mvardQuantFaturada As Double
Private mvardQuantAFaturar As Double
Private mvardQuantOP As Double
Private mvardQuantSC As Double
Private mvarsDescricao As String
Private mvarcolLocalizacao As New Collection
Private mvarlNumIntDoc As Long
Private mvardQuantPV As Double
Private mvarlNumIntItemPV As Long
Private mvariAlmoxarifado As Integer
Private mvarsAlmoxarifado As String
Private mvariControleEstoque As Integer
Private mvarsVersao As String
Private mvariFilialOP As Integer
Private mvarsLote As String
Private mvarsCodOP As String
Private mvarsProdOP As String
Private mvarlHorasMaquina As Long
Private mvardPreco As Double
Private mvarsTextoGrade As String

Public Property Let sTextoGrade(ByVal vData As String)
    mvarsTextoGrade = vData
End Property

Public Property Get sTextoGrade() As String
    sTextoGrade = mvarsTextoGrade
End Property

Public Property Let dPreco(ByVal vData As Double)
    mvardPreco = vData
End Property

Public Property Get dPreco() As Double
    dPreco = mvardPreco
End Property

Public Property Let lHorasMaquina(ByVal vData As Long)
    mvarlHorasMaquina = vData
End Property

Public Property Get lHorasMaquina() As Long
    lHorasMaquina = mvarlHorasMaquina
End Property

Public Property Let sProdOP(ByVal vData As String)
    mvarsProdOP = vData
End Property

Public Property Get sProdOP() As String
    sProdOP = mvarsProdOP
End Property

Public Property Let sCodOP(ByVal vData As String)
    mvarsCodOP = vData
End Property

Public Property Get sCodOP() As String
    sCodOP = mvarsCodOP
End Property

Public Property Let sLote(ByVal vData As String)
    mvarsLote = vData
End Property

Public Property Get sLote() As String
    sLote = mvarsLote
End Property

Public Property Let iFilialOP(ByVal vData As Integer)
    mvariFilialOP = vData
End Property

Public Property Get iFilialOP() As Integer
    iFilialOP = mvariFilialOP
End Property

Public Property Let sVersao(ByVal vData As String)
    mvarsVersao = vData
End Property

Public Property Get sVersao() As String
    sVersao = mvarsVersao
End Property

Public Property Let iControleEstoque(ByVal vData As Integer)
    mvariControleEstoque = vData
End Property

Public Property Get iControleEstoque() As Integer
    iControleEstoque = mvariControleEstoque
End Property

Public Property Let sAlmoxarifado(ByVal vData As String)
    mvarsAlmoxarifado = vData
End Property

Public Property Get sAlmoxarifado() As String
    sAlmoxarifado = mvarsAlmoxarifado
End Property
Public Property Let iAlmoxarifado(ByVal vData As Integer)
    mvariAlmoxarifado = vData
End Property

Public Property Get iAlmoxarifado() As Integer
    iAlmoxarifado = mvariAlmoxarifado
End Property

Public Property Let lNumIntItemPV(ByVal vData As Long)
    mvarlNumIntItemPV = vData
End Property

Public Property Get lNumIntItemPV() As Long
    lNumIntItemPV = mvarlNumIntItemPV
End Property

Public Property Let dQuantPV(ByVal vData As Double)
    mvardQuantPV = vData
End Property

Public Property Get dQuantPV() As Double
    dQuantPV = mvardQuantPV
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let dQuantCancelada(ByVal vData As Double)
    mvardQuantCancelada = vData
End Property

Public Property Get dQuantCancelada() As Double
    dQuantCancelada = mvardQuantCancelada
End Property
Public Property Let dQuantReservada(ByVal vData As Double)
    mvardQuantReservada = vData
End Property

Public Property Get dQuantReservada() As Double
    dQuantReservada = mvardQuantReservada
End Property
Public Property Let dQuantAFaturar(ByVal vData As Double)
    mvardQuantAFaturar = vData
End Property

Public Property Get dQuantAFaturar() As Double
    dQuantAFaturar = mvardQuantAFaturar
End Property
Public Property Let dQuantFaturada(ByVal vData As Double)
    mvardQuantFaturada = vData
End Property

Public Property Get dQuantFaturada() As Double
    dQuantFaturada = mvardQuantFaturada
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let dQuantOP(ByVal vData As Double)
    mvardQuantOP = vData
End Property

Public Property Get dQuantOP() As Double
    dQuantOP = mvardQuantOP
End Property

Public Property Let dQuantSC(ByVal vData As Double)
    mvardQuantSC = vData
End Property

Public Property Get dQuantSC() As Double
    dQuantSC = mvardQuantSC
End Property

Public Property Set colCategoria(ByVal vData As Collection)
    Set mvarcolCategoria = vData
End Property

Public Property Get colCategoria() As Collection
    Set colCategoria = mvarcolCategoria
End Property

Public Property Let sProduto(ByVal vData As String)
Attribute sProduto.VB_Description = "Produto correspondente a um item da Grade"
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Set colLocalizacao(ByVal vData As Collection)
    Set mvarcolLocalizacao = vData
End Property

Public Property Get colLocalizacao() As Collection
    Set colLocalizacao = mvarcolLocalizacao
End Property

Public Property Let sUMEstoque(ByVal vData As String)
    mvarsUMEstoque = vData
End Property

Public Property Get sUMEstoque() As String
    sUMEstoque = mvarsUMEstoque
End Property

