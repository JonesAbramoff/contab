VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPVAndamentoSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarcolPVs As New Collection
Private mvariSoAbertos As Integer
Private mvarlPedido As Long
Private mvarlCliente As Long
Private mvardtEmissaoDe As Date
Private mvardtEmissaoAte As Date
Private mvardtEntregaDe As Date
Private mvardtEntregaAte As Date
Private mvarcolcolNFs As New Collection
Private mvarcolcolOPs As New Collection

Public Property Set colcolOPs(ByVal vData As Collection)
    Set mvarcolcolOPs = vData
End Property

Public Property Get colcolOPs() As Collection
    Set colcolOPs = mvarcolcolOPs
End Property

Public Property Set colcolNFs(ByVal vData As Collection)
    Set mvarcolcolNFs = vData
End Property

Public Property Get colcolNFs() As Collection
    Set colcolNFs = mvarcolcolNFs
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

Public Property Let dtEmissaoAte(ByVal vData As Date)
    mvardtEmissaoAte = vData
End Property

Public Property Get dtEmissaoAte() As Date
    dtEmissaoAte = mvardtEmissaoAte
End Property

Public Property Let dtEmissaoDe(ByVal vData As Date)
    mvardtEmissaoDe = vData
End Property

Public Property Get dtEmissaoDe() As Date
    dtEmissaoDe = mvardtEmissaoDe
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let lPedido(ByVal vData As Long)
    mvarlPedido = vData
End Property

Public Property Get lPedido() As Long
    lPedido = mvarlPedido
End Property

Public Property Let iSoAbertos(ByVal vData As Integer)
    mvariSoAbertos = vData
End Property

Public Property Get iSoAbertos() As Integer
    iSoAbertos = mvariSoAbertos
End Property

Public Property Set colPVs(ByVal vData As Collection)
    Set mvarcolPVs = vData
End Property

Public Property Get colPVs() As Collection
    Set colPVs = mvarcolPVs
End Property
