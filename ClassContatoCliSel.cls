VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassContatoCliSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvardtDataContDe As Date
Private mvardtDataContAte As Date
Private mvardtDataPContDe As Date
Private mvardtDataPContAte As Date
Private mvardtDataPCompDe As Date
Private mvardtDataPCompAte As Date
Private mvardtDataUCompDe As Date
Private mvardtDataUCompAte As Date
Private mvardtDataNCompDe As Date
Private mvardtDataNCompAte As Date
Private mvarsRespCallCenter As String
Private mvarsCategoria As String
Private mvarsCategoriaDe As String
Private mvarsCategoriaAte As String
Private mvariSoComEmailValido As Integer
Private mvarlModeloForcado As Long
Private mvarcolTiposNaoConsiderar As New Collection
Private mvariSoAtivos As Integer
Private mvariMarca As Integer
Private mvariIgnoraJaEnviados As Integer
Private mvarlNumIntDocModelo As Long

Public Property Let lNumIntDocModelo(ByVal vData As Long)
    mvarlNumIntDocModelo = vData
End Property

Public Property Get lNumIntDocModelo() As Long
    lNumIntDocModelo = mvarlNumIntDocModelo
End Property

Public Property Let iIgnoraJaEnviados(ByVal vData As Integer)
    mvariIgnoraJaEnviados = vData
End Property

Public Property Get iIgnoraJaEnviados() As Integer
    iIgnoraJaEnviados = mvariIgnoraJaEnviados
End Property

Public Property Let iMarca(ByVal vData As Integer)
    mvariMarca = vData
End Property

Public Property Get iMarca() As Integer
    iMarca = mvariMarca
End Property

Public Property Let iSoAtivos(ByVal vData As Integer)
    mvariSoAtivos = vData
End Property

Public Property Get iSoAtivos() As Integer
    iSoAtivos = mvariSoAtivos
End Property

Public Property Set colTiposNaoConsiderar(ByVal vData As Collection)
    Set mvarcolTiposNaoConsiderar = vData
End Property

Public Property Get colTiposNaoConsiderar() As Collection
    Set colTiposNaoConsiderar = mvarcolTiposNaoConsiderar
End Property

Public Property Let lModeloForcado(ByVal vData As Long)
    mvarlModeloForcado = vData
End Property

Public Property Get lModeloForcado() As Long
    lModeloForcado = mvarlModeloForcado
End Property

Public Property Let iSoComEmailValido(ByVal vData As Integer)
    mvariSoComEmailValido = vData
End Property

Public Property Get iSoComEmailValido() As Integer
    iSoComEmailValido = mvariSoComEmailValido
End Property

Public Property Let sCategoriaAte(ByVal vData As String)
    mvarsCategoriaAte = vData
End Property

Public Property Get sCategoriaAte() As String
    sCategoriaAte = mvarsCategoriaAte
End Property

Public Property Let sCategoriaDe(ByVal vData As String)
    mvarsCategoriaDe = vData
End Property

Public Property Get sCategoriaDe() As String
    sCategoriaDe = mvarsCategoriaDe
End Property

Public Property Let sCategoria(ByVal vData As String)
    mvarsCategoria = vData
End Property

Public Property Get sCategoria() As String
    sCategoria = mvarsCategoria
End Property

Public Property Let sRespCallCenter(ByVal vData As String)
    mvarsRespCallCenter = vData
End Property

Public Property Get sRespCallCenter() As String
    sRespCallCenter = mvarsRespCallCenter
End Property

Public Property Let dtDataUCompAte(ByVal vData As Date)
    mvardtDataUCompAte = vData
End Property

Public Property Get dtDataUCompAte() As Date
    dtDataUCompAte = mvardtDataUCompAte
End Property

Public Property Let dtDataUCompDe(ByVal vData As Date)
    mvardtDataUCompDe = vData
End Property

Public Property Get dtDataUCompDe() As Date
    dtDataUCompDe = mvardtDataUCompDe
End Property

Public Property Let dtDataPCompAte(ByVal vData As Date)
    mvardtDataPCompAte = vData
End Property

Public Property Get dtDataPCompAte() As Date
    dtDataPCompAte = mvardtDataPCompAte
End Property

Public Property Let dtDataPCompDe(ByVal vData As Date)
    mvardtDataPCompDe = vData
End Property

Public Property Get dtDataPCompDe() As Date
    dtDataPCompDe = mvardtDataPCompDe
End Property

Public Property Let dtDataContAte(ByVal vData As Date)
    mvardtDataContAte = vData
End Property

Public Property Get dtDataContAte() As Date
    dtDataContAte = mvardtDataContAte
End Property

Public Property Let dtDataContDe(ByVal vData As Date)
    mvardtDataContDe = vData
End Property

Public Property Get dtDataContDe() As Date
    dtDataContDe = mvardtDataContDe
End Property

Public Property Let dtDataNCompAte(ByVal vData As Date)
    mvardtDataNCompAte = vData
End Property

Public Property Get dtDataNCompAte() As Date
    dtDataNCompAte = mvardtDataNCompAte
End Property

Public Property Let dtDataNCompDe(ByVal vData As Date)
    mvardtDataNCompDe = vData
End Property

Public Property Get dtDataNCompDe() As Date
    dtDataNCompDe = mvardtDataNCompDe
End Property

Public Property Let dtDataPContAte(ByVal vData As Date)
    mvardtDataPContAte = vData
End Property

Public Property Get dtDataPContAte() As Date
    dtDataPContAte = mvardtDataPContAte
End Property

Public Property Let dtDataPContDe(ByVal vData As Date)
    mvardtDataPContDe = vData
End Property

Public Property Get dtDataPContDe() As Date
    dtDataPContDe = mvardtDataPContDe
End Property
