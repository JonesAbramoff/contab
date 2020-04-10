VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPOMaquinas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumIntDocPO As Long
Private mvarlNumIntDocMaq As Long
Private mvariQuantidade As Integer
Private mvardtData As Date
Private mvardHorasMaquina As Double
Private mvarobjTaxaProducao As New ClassTaxaDeProducao
Private mvarobjMaquina As New ClassMaquinas
Private mvarlNumIntDocTxProd As Long

Public Property Let lNumIntDocTxProd(ByVal vData As Long)
    mvarlNumIntDocTxProd = vData
End Property

Public Property Get lNumIntDocTxProd() As Long
    lNumIntDocTxProd = mvarlNumIntDocTxProd
End Property

Public Property Set objMaquina(ByVal vData As ClassMaquinas)
    Set mvarobjMaquina = vData
End Property

Public Property Get objMaquina() As ClassMaquinas
    Set objMaquina = mvarobjMaquina
End Property

Public Property Set objTaxaProducao(ByVal vData As ClassTaxaDeProducao)
    Set mvarobjTaxaProducao = vData
End Property

Public Property Get objTaxaProducao() As ClassTaxaDeProducao
    Set objTaxaProducao = mvarobjTaxaProducao
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntDocPO(ByVal vData As Long)
    mvarlNumIntDocPO = vData
End Property

Public Property Get lNumIntDocPO() As Long
    lNumIntDocPO = mvarlNumIntDocPO
End Property

Public Property Let lNumIntDocMaq(ByVal vData As Long)
    mvarlNumIntDocMaq = vData
End Property

Public Property Get lNumIntDocMaq() As Long
    lNumIntDocMaq = mvarlNumIntDocMaq
End Property

Public Property Let iQuantidade(ByVal vData As Integer)
    mvariQuantidade = vData
End Property

Public Property Get iQuantidade() As Integer
    iQuantidade = mvariQuantidade
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let dHorasMaquina(ByVal vData As Double)
    mvardHorasMaquina = vData
End Property

Public Property Get dHorasMaquina() As Double
    dHorasMaquina = mvardHorasMaquina
End Property
