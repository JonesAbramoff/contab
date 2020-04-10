VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPRJContratoEtapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumIntDocContrato As Long
Private mvarlNumIntDocEtapa As Long
Private mvarlNumIntDocEtapaItemProd As Long
Private mvardCustoInformado As Double
Private mvardPreco As Double
Private mvariSelecionado As Integer
Private mvariImprimir As Integer
Private mvarsObservacao As String
Private mvarsDescricao As String

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let iImprimir(ByVal vData As Integer)
    mvariImprimir = vData
End Property

Public Property Get iImprimir() As Integer
    iImprimir = mvariImprimir
End Property

Public Property Let iSelecionado(ByVal vData As Integer)
    mvariSelecionado = vData
End Property

Public Property Get iSelecionado() As Integer
    iSelecionado = mvariSelecionado
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntDocContrato(ByVal vData As Long)
    mvarlNumIntDocContrato = vData
End Property

Public Property Get lNumIntDocContrato() As Long
    lNumIntDocContrato = mvarlNumIntDocContrato
End Property

Public Property Let lNumIntDocEtapa(ByVal vData As Long)
    mvarlNumIntDocEtapa = vData
End Property

Public Property Get lNumIntDocEtapa() As Long
    lNumIntDocEtapa = mvarlNumIntDocEtapa
End Property

Public Property Let lNumIntDocEtapaItemProd(ByVal vData As Long)
    mvarlNumIntDocEtapaItemProd = vData
End Property

Public Property Get lNumIntDocEtapaItemProd() As Long
    lNumIntDocEtapaItemProd = mvarlNumIntDocEtapaItemProd
End Property

Public Property Let dCustoInformado(ByVal vData As Double)
    mvardCustoInformado = vData
End Property

Public Property Get dCustoInformado() As Double
    dCustoInformado = mvardCustoInformado
End Property

Public Property Let dPreco(ByVal vData As Double)
    mvardPreco = vData
End Property

Public Property Get dPreco() As Double
    dPreco = mvardPreco
End Property
