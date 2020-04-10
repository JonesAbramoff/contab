VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassTRVClienteRepProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlCliente AS Long
Private mvariSeq AS Integer
Private mvarlRepresentante AS Long
Private mvarsProduto AS String
Private mvardPercComis AS Double

Public Property Let lCliente (ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente () AS Long
    lCliente= mvarlCliente
End Property

Public Property Let iSeq (ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq () AS Integer
    iSeq= mvariSeq
End Property

Public Property Let lRepresentante (ByVal vData As Long)
    mvarlRepresentante = vData
End Property

Public Property Get lRepresentante () AS Long
    lRepresentante= mvarlRepresentante
End Property

Public Property Let sProduto (ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto () AS String
    sProduto= mvarsProduto
End Property

Public Property Let dPercComis (ByVal vData As Double)
    mvardPercComis = vData
End Property

Public Property Get dPercComis () AS Double
    dPercComis= mvardPercComis
End Property
