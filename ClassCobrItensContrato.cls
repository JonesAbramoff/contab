VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCobrItensContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntItensContrato As Long
Private mvarlMedicao As Long
Private mvardtDataUltCobranca As Date
Private mvarlNumIntDocCobranca As Long
Private mvarlNumIntItemNF As Long
Private mvardtDataRefIni As Date
Private mvardtDataRefFim As Date
Private mvarcolMedicoes As New Collection

Public Property Set colMedicoes(ByVal vData As Object)
    Set mvarcolMedicoes = vData
End Property

Public Property Get colMedicoes() As Collection
    Set colMedicoes = mvarcolMedicoes
End Property

Public Property Let dtDataRefIni(ByVal vData As Date)
    mvardtDataRefIni = vData
End Property

Public Property Get dtDataRefIni() As Date
     dtDataRefIni = mvardtDataRefIni
End Property

Public Property Let dtDataRefFim(ByVal vData As Date)
    mvardtDataRefFim = vData
End Property

Public Property Get dtDataRefFim() As Date
     dtDataRefFim = mvardtDataRefFim
End Property

Public Property Let lNumIntItemNF(ByVal vData As Long)
    mvarlNumIntItemNF = vData
End Property

Public Property Get lNumIntItemNF() As Long
    lNumIntItemNF = mvarlNumIntItemNF
End Property

Public Property Let lNumIntDocCobranca(ByVal vData As Long)
    mvarlNumIntDocCobranca = vData
End Property

Public Property Get lNumIntDocCobranca() As Long
    lNumIntDocCobranca = mvarlNumIntDocCobranca
End Property

Public Property Let dtDataUltCobranca(ByVal vData As Date)
    mvardtDataUltCobranca = vData
End Property

Public Property Get dtDataUltCobranca() As Date
    dtDataUltCobranca = mvardtDataUltCobranca
End Property

Public Property Let lMedicao(ByVal vData As Long)
    mvarlMedicao = vData
End Property

Public Property Get lMedicao() As Long
    lMedicao = mvarlMedicao
End Property

Public Property Let lNumIntItensContrato(ByVal vData As Long)
    mvarlNumIntItensContrato = vData
End Property

Public Property Get lNumIntItensContrato() As Long
    lNumIntItensContrato = mvarlNumIntItensContrato
End Property



