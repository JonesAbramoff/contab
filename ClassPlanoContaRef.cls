VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPlanoContaRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlCodigoModelo As Long
Private mvarsConta As String
Private mvarsDescricao As String
Private mvarsOrientacao As String
Private mvariTipo As Integer
Private mvarcolConfig As New Collection
Private mvarsContaImp As String
Private mvariTipoImp As Integer
Private mvardtValidadeDe As Date
Private mvardtValidadeAte As Date
Private mvarsConta0 As String
Private mvariNatureza As Integer
Private mvarsContaPai As String

Public Property Let sContaPai(ByVal vData As String)
    mvarsContaPai = vData
End Property

Public Property Get sContaPai() As String
    sContaPai = mvarsContaPai
End Property

Public Property Let iNatureza(ByVal vData As Integer)
    mvariNatureza = vData
End Property

Public Property Get iNatureza() As Integer
    iNatureza = mvariNatureza
End Property

Public Property Let sConta0(ByVal vData As String)
    mvarsConta0 = vData
End Property

Public Property Get sConta0() As String
    sConta0 = mvarsConta0
End Property

Public Property Let dtValidadeAte(ByVal vData As Date)
    mvardtValidadeAte = vData
End Property

Public Property Get dtValidadeAte() As Date
    dtValidadeAte = mvardtValidadeAte
End Property

Public Property Let dtValidadeDe(ByVal vData As Date)
    mvardtValidadeDe = vData
End Property

Public Property Get dtValidadeDe() As Date
    dtValidadeDe = mvardtValidadeDe
End Property

Public Property Let iTipoImp(ByVal vData As Integer)
    mvariTipoImp = vData
End Property

Public Property Get iTipoImp() As Integer
    iTipoImp = mvariTipoImp
End Property

Public Property Let sContaImp(ByVal vData As String)
    mvarsContaImp = vData
End Property

Public Property Get sContaImp() As String
    sContaImp = mvarsContaImp
End Property

Public Property Set colConfig(ByVal vData As Object)
    Set mvarcolConfig = vData
End Property

Public Property Get colConfig() As Collection
    Set colConfig = mvarcolConfig
End Property

Public Property Let lCodigoModelo(ByVal vData As Long)
    mvarlCodigoModelo = vData
End Property

Public Property Get lCodigoModelo() As Long
    lCodigoModelo = mvarlCodigoModelo
End Property

Public Property Let sConta(ByVal vData As String)
    mvarsConta = vData
End Property

Public Property Get sConta() As String
    sConta = mvarsConta
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sOrientacao(ByVal vData As String)
    mvarsOrientacao = vData
End Property

Public Property Get sOrientacao() As String
    sOrientacao = mvarsOrientacao
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property
