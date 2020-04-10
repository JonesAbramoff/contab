VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBorderoBoletoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarlNumIntDoc As Long
Private mvariAdmMeioPagto As Integer
Private mvardValor As Double
Private mvariParcelamento As Integer
Private mvarlNumBordero As Long
Private mvardSaldo As Double
Private mvarlNumIntDocCPR As Long
Private mvarsNomeParcelamento As String
Private mvarsNomeAdmMeioPagto As String
Private mvardValorN As Double
Private mvarobjMovcx As New ClassMovimentoCaixa 'guarda o movimento de caixa associado no caso do log do caixa central para o backoffice
Private mvardtDataPreDatado As Date

Public Property Let dtDataPreDatado(ByVal vData As Date)
    mvardtDataPreDatado = vData
End Property

Public Property Get dtDataPreDatado() As Date
    dtDataPreDatado = mvardtDataPreDatado
End Property

Public Property Get objMovcx() As Object
    Set objMovcx = mvarobjMovcx
End Property

Public Property Set objMovcx(ByVal vData As Object)
    Set mvarobjMovcx = vData
End Property

Public Property Let dValorN(ByVal vData As Double)
    mvardValorN = vData
End Property

Public Property Get dValorN() As Double
    dValorN = mvardValorN
End Property

Public Property Let sNomeAdmMeioPagto(ByVal vData As String)
    mvarsNomeAdmMeioPagto = vData
End Property

Public Property Get sNomeAdmMeioPagto() As String
    sNomeAdmMeioPagto = mvarsNomeAdmMeioPagto
End Property

Public Property Let sNomeParcelamento(ByVal vData As String)
    mvarsNomeParcelamento = vData
End Property

Public Property Get sNomeParcelamento() As String
    sNomeParcelamento = mvarsNomeParcelamento
End Property

Public Property Let lNumIntDocCPR(ByVal vData As Long)
    mvarlNumIntDocCPR = vData
End Property

Public Property Get lNumIntDocCPR() As Long
    lNumIntDocCPR = mvarlNumIntDocCPR
End Property

Public Property Let dSaldo(ByVal vData As Double)
    mvardSaldo = vData
End Property

Public Property Get dSaldo() As Double
    dSaldo = mvardSaldo
End Property

Public Property Let lNumBordero(ByVal vData As Long)
    mvarlNumBordero = vData
End Property

Public Property Get lNumBordero() As Long
    lNumBordero = mvarlNumBordero
End Property

Public Property Let iParcelamento(ByVal vData As Integer)
    mvariParcelamento = vData
End Property

Public Property Get iParcelamento() As Integer
    iParcelamento = mvariParcelamento
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let iAdmMeioPagto(ByVal vData As Integer)
    mvariAdmMeioPagto = vData
End Property

Public Property Get iAdmMeioPagto() As Integer
    iAdmMeioPagto = mvariAdmMeioPagto
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property


Private Sub Class_Initialize()
    mvardtDataPreDatado = DATA_NULA
End Sub
