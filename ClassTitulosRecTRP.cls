VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTitulosRecTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDocTitRec As Long
Private mvardValorTarifa As Double
Private mvardValorDeducoes As Double
Private mvardValorComissao As Double
Private mvardValorBruto As Double
Private mvarlClienteVou As Long
Private mvarlCliPassageiro As Long
Private mvariNFAoPassageiro As Integer
Private mvarsNomeCliente As String
Private mvariUmaNFPorFatura As Integer

Public Property Let iUmaNFPorFatura(ByVal vData As Integer)
    mvariUmaNFPorFatura = vData
End Property

Public Property Get iUmaNFPorFatura() As Integer
    iUmaNFPorFatura = mvariUmaNFPorFatura
End Property

Public Property Let sNomeCliente(ByVal vData As String)
    mvarsNomeCliente = vData
End Property

Public Property Get sNomeCliente() As String
    sNomeCliente = mvarsNomeCliente
End Property

Public Property Let iNFAoPassageiro(ByVal vData As Integer)
    mvariNFAoPassageiro = vData
End Property

Public Property Get iNFAoPassageiro() As Integer
    iNFAoPassageiro = mvariNFAoPassageiro
End Property

Public Property Let lCliPassageiro(ByVal vData As Long)
    mvarlCliPassageiro = vData
End Property

Public Property Get lCliPassageiro() As Long
    lCliPassageiro = mvarlCliPassageiro
End Property

Public Property Let lClienteVou(ByVal vData As Long)
    mvarlClienteVou = vData
End Property

Public Property Get lClienteVou() As Long
    lClienteVou = mvarlClienteVou
End Property

Public Property Let lNumIntDocTitRec(ByVal vData As Long)
    mvarlNumIntDocTitRec = vData
End Property

Public Property Get lNumIntDocTitRec() As Long
    lNumIntDocTitRec = mvarlNumIntDocTitRec
End Property

Public Property Let dValorTarifa(ByVal vData As Double)
    mvardValorTarifa = vData
End Property

Public Property Get dValorTarifa() As Double
    dValorTarifa = mvardValorTarifa
End Property

Public Property Let dValorDeducoes(ByVal vData As Double)
    mvardValorDeducoes = vData
End Property

Public Property Get dValorDeducoes() As Double
    dValorDeducoes = mvardValorDeducoes
End Property

Public Property Let dValorComissao(ByVal vData As Double)
    mvardValorComissao = vData
End Property

Public Property Get dValorComissao() As Double
    dValorComissao = mvardValorComissao
End Property

Public Property Let dValorBruto(ByVal vData As Double)
    mvardValorBruto = vData
End Property

Public Property Get dValorBruto() As Double
    dValorBruto = mvardValorBruto
End Property
