VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelTRVDesviosVendSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntRel As Long
Private mvardtDataGer As Date
Private mvardHoraGer As Double
Private mvarsUsuGer As String
Private mvariAno As Integer
Private mvariMes As Integer
Private mvariDesvios As Integer
Private mvardPercDesvMes As Double
Private mvardPercDesvAno As Double
Private mvardMinVendVlr As Double
Private mvariMinVendQtd As Integer
Private mvariTrazerCliNComp As Integer
Private mvariValorBase As Integer
Private mvarcolItens As New Collection
Private mvariVendedor As Integer
Private mvarsRespCallCenter As String
Private mvarsResponsavel As String
Private mvariRegiao As Integer

Public Property Let iRegiao(ByVal vData As Integer)
    mvariRegiao = vData
End Property

Public Property Get iRegiao() As Integer
    iRegiao = mvariRegiao
End Property

Public Property Let sResponsavel(ByVal vData As String)
    mvarsResponsavel = vData
End Property

Public Property Get sResponsavel() As String
    sResponsavel = mvarsResponsavel
End Property

Public Property Let sRespCallCenter(ByVal vData As String)
    mvarsRespCallCenter = vData
End Property

Public Property Get sRespCallCenter() As String
    sRespCallCenter = mvarsRespCallCenter
End Property

Public Property Let iVendedor(ByVal vData As Integer)
    mvariVendedor = vData
End Property

Public Property Get iVendedor() As Integer
    iVendedor = mvariVendedor
End Property

Public Property Set colItens(ByVal vData As Collection)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Let lNumIntRel(ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel() As Long
    lNumIntRel = mvarlNumIntRel
End Property

Public Property Let dtDataGer(ByVal vData As Date)
    mvardtDataGer = vData
End Property

Public Property Get dtDataGer() As Date
    dtDataGer = mvardtDataGer
End Property

Public Property Let dHoraGer(ByVal vData As Double)
    mvardHoraGer = vData
End Property

Public Property Get dHoraGer() As Double
    dHoraGer = mvardHoraGer
End Property

Public Property Let sUsuGer(ByVal vData As String)
    mvarsUsuGer = vData
End Property

Public Property Get sUsuGer() As String
    sUsuGer = mvarsUsuGer
End Property

Public Property Let iAno(ByVal vData As Integer)
    mvariAno = vData
End Property

Public Property Get iAno() As Integer
    iAno = mvariAno
End Property

Public Property Let iMes(ByVal vData As Integer)
    mvariMes = vData
End Property

Public Property Get iMes() As Integer
    iMes = mvariMes
End Property

Public Property Let iDesvios(ByVal vData As Integer)
    mvariDesvios = vData
End Property

Public Property Get iDesvios() As Integer
    iDesvios = mvariDesvios
End Property

Public Property Let dPercDesvMes(ByVal vData As Double)
    mvardPercDesvMes = vData
End Property

Public Property Get dPercDesvMes() As Double
    dPercDesvMes = mvardPercDesvMes
End Property

Public Property Let dPercDesvAno(ByVal vData As Double)
    mvardPercDesvAno = vData
End Property

Public Property Get dPercDesvAno() As Double
    dPercDesvAno = mvardPercDesvAno
End Property

Public Property Let dMinVendVlr(ByVal vData As Double)
    mvardMinVendVlr = vData
End Property

Public Property Get dMinVendVlr() As Double
    dMinVendVlr = mvardMinVendVlr
End Property

Public Property Let iMinVendQtd(ByVal vData As Integer)
    mvariMinVendQtd = vData
End Property

Public Property Get iMinVendQtd() As Integer
    iMinVendQtd = mvariMinVendQtd
End Property

Public Property Let iTrazerCliNComp(ByVal vData As Integer)
    mvariTrazerCliNComp = vData
End Property

Public Property Get iTrazerCliNComp() As Integer
    iTrazerCliNComp = mvariTrazerCliNComp
End Property

Public Property Let iValorBase(ByVal vData As Integer)
    mvariValorBase = vData
End Property

Public Property Get iValorBase() As Integer
    iValorBase = mvariValorBase
End Property
