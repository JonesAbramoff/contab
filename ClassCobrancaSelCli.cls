VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCobrancaSelCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvardtDataProxDe As Date
Private mvardtDataProxAte As Date
Private mvardtDataVencDe As Date
Private mvardtDataVencAte As Date
Private mvardtDataPrevDe As Date
Private mvardtDataPrevAte As Date
Private mvarsCobrador As String
Private mvariTitulosBaixados As Integer
Private mvarsCategoria As String
Private mvarsCategoriaDe As String
Private mvarsCategoriaAte As String
Private mvariVendedor As Integer
Private mvariTipoVend As Integer

Public Property Let iTipoVend(ByVal vData As Integer)
    mvariTipoVend = vData
End Property

Public Property Get iTipoVend() As Integer
    iTipoVend = mvariTipoVend
End Property

Public Property Let iVendedor(ByVal vData As Integer)
    mvariVendedor = vData
End Property

Public Property Get iVendedor() As Integer
    iVendedor = mvariVendedor
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

Public Property Let iTitulosBaixados(ByVal vData As Integer)
    mvariTitulosBaixados = vData
End Property

Public Property Get iTitulosBaixados() As Integer
    iTitulosBaixados = mvariTitulosBaixados
End Property

Public Property Let sCobrador(ByVal vData As String)
    mvarsCobrador = vData
End Property

Public Property Get sCobrador() As String
    sCobrador = mvarsCobrador
End Property

Public Property Let dtDataPrevAte(ByVal vData As Date)
    mvardtDataPrevAte = vData
End Property

Public Property Get dtDataPrevAte() As Date
    dtDataPrevAte = mvardtDataPrevAte
End Property

Public Property Let dtDataPrevDe(ByVal vData As Date)
    mvardtDataPrevDe = vData
End Property

Public Property Get dtDataPrevDe() As Date
    dtDataPrevDe = mvardtDataPrevDe
End Property

Public Property Let dtDataVencAte(ByVal vData As Date)
    mvardtDataVencAte = vData
End Property

Public Property Get dtDataVencAte() As Date
    dtDataVencAte = mvardtDataVencAte
End Property

Public Property Let dtDataVencDe(ByVal vData As Date)
    mvardtDataVencDe = vData
End Property

Public Property Get dtDataVencDe() As Date
    dtDataVencDe = mvardtDataVencDe
End Property

Public Property Let dtDataProxAte(ByVal vData As Date)
    mvardtDataProxAte = vData
End Property

Public Property Get dtDataProxAte() As Date
    dtDataProxAte = mvardtDataProxAte
End Property

Public Property Let dtDataProxDe(ByVal vData As Date)
    mvardtDataProxDe = vData
End Property

Public Property Get dtDataProxDe() As Date
    dtDataProxDe = mvardtDataProxDe
End Property
