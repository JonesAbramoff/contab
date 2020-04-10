VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemOPOperacoesMO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumIntDocItemOPMaq As Long
Private mvariTipoMO As Integer
Private mvardHoras As Double
Private mvariQuantidade As Integer

Public Property Let iQuantidade(ByVal vData As Integer)
    mvariQuantidade = vData
End Property

Public Property Get iQuantidade() As Integer
    iQuantidade = mvariQuantidade
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntDocItemOPMaq(ByVal vData As Long)
    mvarlNumIntDocItemOPMaq = vData
End Property

Public Property Get lNumIntDocItemOPMaq() As Long
    lNumIntDocItemOPMaq = mvarlNumIntDocItemOPMaq
End Property

Public Property Let iTipoMO(ByVal vData As Integer)
    mvariTipoMO = vData
End Property

Public Property Get iTipoMO() As Integer
    iTipoMO = mvariTipoMO
End Property

Public Property Let dHoras(ByVal vData As Double)
    mvardHoras = vData
End Property

Public Property Get dHoras() As Double
    dHoras = mvardHoras
End Property
