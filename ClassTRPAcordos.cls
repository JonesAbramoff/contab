VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRPAcordos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumero As Long
Private mvarsContrato As String
Private mvarlCliente As Long
Private mvariFilialCliente As Integer
Private mvardtValidadeDe As Date
Private mvardtValidadeAte As Date
Private mvarsObservacao As String
Private mvarsDescricao As String
Private mvarcolTRPAcordoComiss As New Collection

Public Property Set colTRPAcordoComiss(ByVal vData As Collection)
    Set mvarcolTRPAcordoComiss = vData
End Property

Public Property Get colTRPAcordoComiss() As Collection
    Set colTRPAcordoComiss = mvarcolTRPAcordoComiss
End Property


Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumero(ByVal vData As Long)
    mvarlNumero = vData
End Property

Public Property Get lNumero() As Long
    lNumero = mvarlNumero
End Property

Public Property Let sContrato(ByVal vData As String)
    mvarsContrato = vData
End Property

Public Property Get sContrato() As String
    sContrato = mvarsContrato
End Property

Public Property Let iFilialCliente(ByVal vData As Integer)
    mvariFilialCliente = vData
End Property

Public Property Get iFilialCliente() As Integer
    iFilialCliente = mvariFilialCliente
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let dtValidadeDe(ByVal vData As Date)
    mvardtValidadeDe = vData
End Property

Public Property Get dtValidadeDe() As Date
    dtValidadeDe = mvardtValidadeDe
End Property

Public Property Let dtValidadeAte(ByVal vData As Date)
    mvardtValidadeAte = vData
End Property

Public Property Get dtValidadeAte() As Date
    dtValidadeAte = mvardtValidadeAte
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property
