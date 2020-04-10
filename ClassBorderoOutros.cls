VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBorderoOutros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarlNumBordero As Long
Private mvarsAdmMeioPagto As String
Private mvariParcelamento As Integer
Private mvardtDataEnvio As Date
Private mvardtDataImpressao As Date
Private mvardtDataBackoffice As Date
Private mvardValor As Double
Private mvardValorN As Double
Private mvarlNumIntDocCPR As Long
Private mvariAdmMeioPagto As Integer
Private mvarsNomeParcelamento As String
Private mvarlNumMovtoCx As Long
Private mvariExibeTela As Integer

Public Property Let iExibeTela(ByVal vData As Integer)
    mvariExibeTela = vData
End Property

Public Property Get iExibeTela() As Integer
    iExibeTela = mvariExibeTela
End Property

Public Property Let lNumMovtoCx(ByVal vData As Long)
    mvarlNumMovtoCx = vData
End Property

Public Property Get lNumMovtoCx() As Long
    lNumMovtoCx = mvarlNumMovtoCx
End Property

Public Property Let sNomeParcelamento(ByVal vData As String)
    mvarsNomeParcelamento = vData
End Property

Public Property Get sNomeParcelamento() As String
    sNomeParcelamento = mvarsNomeParcelamento
End Property

Public Property Let iAdmMeioPagto(ByVal vData As Integer)
    mvariAdmMeioPagto = vData
End Property

Public Property Get iAdmMeioPagto() As Integer
    iAdmMeioPagto = mvariAdmMeioPagto
End Property

Public Property Let lNumIntDocCPR(ByVal vData As Long)
    mvarlNumIntDocCPR = vData
End Property

Public Property Get lNumIntDocCPR() As Long
    lNumIntDocCPR = mvarlNumIntDocCPR
End Property

Public Property Let dValorN(ByVal vData As Double)
    mvardValorN = vData
End Property

Public Property Get dValorN() As Double
    dValorN = mvardValorN
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dtDataBackoffice(ByVal vData As Date)
    mvardtDataBackoffice = vData
End Property

Public Property Get dtDataBackoffice() As Date
    dtDataBackoffice = mvardtDataBackoffice
End Property

Public Property Let dtDataImpressao(ByVal vData As Date)
    mvardtDataImpressao = vData
End Property

Public Property Get dtDataImpressao() As Date
    dtDataImpressao = mvardtDataImpressao
End Property

Public Property Let dtDataEnvio(ByVal vData As Date)
    mvardtDataEnvio = vData
End Property

Public Property Get dtDataEnvio() As Date
    dtDataEnvio = mvardtDataEnvio
End Property

Public Property Let iParcelamento(ByVal vData As Integer)
    mvariParcelamento = vData
End Property

Public Property Get iParcelamento() As Integer
    iParcelamento = mvariParcelamento
End Property

Public Property Let sAdmMeioPagto(ByVal vData As String)
    mvarsAdmMeioPagto = vData
End Property

Public Property Get sAdmMeioPagto() As String
    sAdmMeioPagto = mvarsAdmMeioPagto
End Property

Public Property Let lNumBordero(ByVal vData As Long)
    mvarlNumBordero = vData
End Property

Public Property Get lNumBordero() As Long
    lNumBordero = mvarlNumBordero
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

