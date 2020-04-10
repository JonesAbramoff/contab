VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCompServItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlNumIntDocOrigem As Long
Private mvariCodItemServico As Integer
Private mvardtDataPrev As Date
Private mvardtHoraPrev As Date
Private mvardtDataInicio As Date
Private mvardtHoraInicio As Date
Private mvardtDataFim As Date
Private mvardtHoraFim As Date
Private mvarsDocIntTipo As String
Private mvariDocIntTipo As Integer
Private mvarsDocIntNumero As String
Private mvardtDocIntDataEmissao As Date
Private mvarsDocExtTipo As String
Private mvariDocExtTipo As Integer
Private mvarsDocExtNumero As String
Private mvardtDocExtDataEmissao As Date
Private mvardtDocExtDataRec As Date
Private mvardtDocExtHoraRec As Date
Private mvarsPlacaCaminhao As String
Private mvarsPlacaCarreta As String
Private mvarsMotorista As String
Private mvarsObservacao As String
Private mvariOrdem As Integer

Public Property Let iOrdem(ByVal vData As Integer)
    mvariOrdem = vData
End Property

Public Property Get iOrdem() As Integer
    iOrdem = mvariOrdem
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let sMotorista(ByVal vData As String)
    mvarsMotorista = vData
End Property

Public Property Get sMotorista() As String
    sMotorista = mvarsMotorista
End Property

Public Property Let sPlacaCarreta(ByVal vData As String)
    mvarsPlacaCarreta = vData
End Property

Public Property Get sPlacaCarreta() As String
    sPlacaCarreta = mvarsPlacaCarreta
End Property

Public Property Let sPlacaCaminhao(ByVal vData As String)
    mvarsPlacaCaminhao = vData
End Property

Public Property Get sPlacaCaminhao() As String
    sPlacaCaminhao = mvarsPlacaCaminhao
End Property

Public Property Let dtDocExtHoraRec(ByVal vData As Date)
    mvardtDocExtHoraRec = vData
End Property

Public Property Get dtDocExtHoraRec() As Date
    dtDocExtHoraRec = mvardtDocExtHoraRec
End Property

Public Property Let dtDocExtDataRec(ByVal vData As Date)
    mvardtDocExtDataRec = vData
End Property

Public Property Get dtDocExtDataRec() As Date
    dtDocExtDataRec = mvardtDocExtDataRec
End Property

Public Property Let dtDocExtDataEmissao(ByVal vData As Date)
    mvardtDocExtDataEmissao = vData
End Property

Public Property Get dtDocExtDataEmissao() As Date
    dtDocExtDataEmissao = mvardtDocExtDataEmissao
End Property

Public Property Let sDocExtNumero(ByVal vData As String)
    mvarsDocExtNumero = vData
End Property

Public Property Get sDocExtNumero() As String
    sDocExtNumero = mvarsDocExtNumero
End Property

Public Property Let iDocExtTipo(ByVal vData As Integer)
    mvariDocExtTipo = vData
End Property

Public Property Get iDocExtTipo() As Integer
    iDocExtTipo = mvariDocExtTipo
End Property

Public Property Let sDocExtTipo(ByVal vData As String)
    mvarsDocExtTipo = vData
End Property

Public Property Get sDocExtTipo() As String
    sDocExtTipo = mvarsDocExtTipo
End Property

Public Property Let dtDocIntDataEmissao(ByVal vData As Date)
    mvardtDocIntDataEmissao = vData
End Property

Public Property Get dtDocIntDataEmissao() As Date
    dtDocIntDataEmissao = mvardtDocIntDataEmissao
End Property

Public Property Let sDocIntNumero(ByVal vData As String)
    mvarsDocIntNumero = vData
End Property

Public Property Get sDocIntNumero() As String
    sDocIntNumero = mvarsDocIntNumero
End Property

Public Property Let iDocIntTipo(ByVal vData As Integer)
    mvariDocIntTipo = vData
End Property

Public Property Get iDocIntTipo() As Integer
    iDocIntTipo = mvariDocIntTipo
End Property

Public Property Let sDocIntTipo(ByVal vData As String)
    mvarsDocIntTipo = vData
End Property

Public Property Get sDocIntTipo() As String
    sDocIntTipo = mvarsDocIntTipo
End Property

Public Property Let dtHoraFim(ByVal vData As Date)
    mvardtHoraFim = vData
End Property

Public Property Get dtHoraFim() As Date
    dtHoraFim = mvardtHoraFim
End Property

Public Property Let dtDataFim(ByVal vData As Date)
    mvardtDataFim = vData
End Property

Public Property Get dtDataFim() As Date
    dtDataFim = mvardtDataFim
End Property

Public Property Let dtHoraInicio(ByVal vData As Date)
    mvardtHoraInicio = vData
End Property

Public Property Get dtHoraInicio() As Date
    dtHoraInicio = mvardtHoraInicio
End Property

Public Property Let dtDataInicio(ByVal vData As Date)
    mvardtDataInicio = vData
End Property

Public Property Get dtDataInicio() As Date
    dtDataInicio = mvardtDataInicio
End Property

Public Property Let dtHoraPrev(ByVal vData As Date)
    mvardtHoraPrev = vData
End Property

Public Property Get dtHoraPrev() As Date
    dtHoraPrev = mvardtHoraPrev
End Property

Public Property Let dtDataPrev(ByVal vData As Date)
    mvardtDataPrev = vData
End Property

Public Property Get dtDataPrev() As Date
    dtDataPrev = mvardtDataPrev
End Property

Public Property Let iCodItemServico(ByVal vData As Integer)
    mvariCodItemServico = vData
End Property

Public Property Get iCodItemServico() As Integer
    iCodItemServico = mvariCodItemServico
End Property

Public Property Let lNumIntDocOrigem(ByVal vData As Long)
    mvarlNumIntDocOrigem = vData
End Property

Public Property Get lNumIntDocOrigem() As Long
    lNumIntDocOrigem = mvarlNumIntDocOrigem
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property


