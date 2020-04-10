VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVFatCondPagto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarobjCondPagtos As New ClassCondicaoPagto
Private mvardtDataVouDe As Date
Private mvardtDataVouAte As Date
Private mvardtDataEmissao As Date
Private mvardtDataVencimento As Date

Public Property Let dtDataVencimento(ByVal vData As Date)
    mvardtDataVencimento = vData
End Property

Public Property Get dtDataVencimento() As Date
    dtDataVencimento = mvardtDataVencimento
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let dtDataVouDe(ByVal vData As Date)
    mvardtDataVouDe = vData
End Property

Public Property Get dtDataVouDe() As Date
    dtDataVouDe = mvardtDataVouDe
End Property

Public Property Let dtDataVouAte(ByVal vData As Date)
    mvardtDataVouAte = vData
End Property

Public Property Get dtDataVouAte() As Date
    dtDataVouAte = mvardtDataVouAte
End Property

Public Property Set objCondPagtos(ByVal vData As ClassCondicaoPagto)
    Set mvarobjCondPagtos = vData
End Property

Public Property Get objCondPagtos() As ClassCondicaoPagto
    Set objCondPagtos = mvarobjCondPagtos
End Property




