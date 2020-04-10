VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemOPOperacoesMaquinas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumIntDocOper As Long
Private mvarlNumIntDocMaq As Long
Private mvardtData As Date
Private mvardHoras As Double
Private mvarcolMO As New Collection
Private mvariQuantidade As Integer

Public Property Let iQuantidade(ByVal vData As Integer)
    mvariQuantidade = vData
End Property

Public Property Get iQuantidade() As Integer
    iQuantidade = mvariQuantidade
End Property

Public Property Set colMO(ByVal vData As Collection)
    Set mvarcolMO = vData
End Property

Public Property Get colMO() As Collection
    Set colMO = mvarcolMO
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntDocOper(ByVal vData As Long)
    mvarlNumIntDocOper = vData
End Property

Public Property Get lNumIntDocOper() As Long
    lNumIntDocOper = mvarlNumIntDocOper
End Property

Public Property Let lNumIntDocMaq(ByVal vData As Long)
    mvarlNumIntDocMaq = vData
End Property

Public Property Get lNumIntDocMaq() As Long
    lNumIntDocMaq = mvarlNumIntDocMaq
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let dHoras(ByVal vData As Double)
    mvardHoras = vData
End Property

Public Property Get dHoras() As Double
    dHoras = mvardHoras
End Property
