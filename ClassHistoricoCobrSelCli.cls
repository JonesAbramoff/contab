VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassHistoricoCobrSelCli"
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
Private mvariTitulosBaixados As Integer
Private mvarlCliente As Long
Private mvariFilial As Integer
Private mvariContato As Integer

Public Property Let iContato(ByVal vData As Integer)
    mvariContato = vData
End Property

Public Property Get iContato() As Integer
    iContato = mvariContato
End Property

Public Property Let iFilial(ByVal vData As Integer)
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let iTitulosBaixados(ByVal vData As Integer)
    mvariTitulosBaixados = vData
End Property

Public Property Get iTitulosBaixados() As Integer
    iTitulosBaixados = mvariTitulosBaixados
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
