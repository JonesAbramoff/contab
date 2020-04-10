VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassReqComprasInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'USADA na tela de BAIXA de REQUISIÇÕES

'local variable(s) to hold property value(s)
Private mvarlCodRequisicao As Long 'local copy
Private mvarsCCL As String 'local copy
Private mvarlRequisitante As Long 'local copy
Private mvarsNomeRequisitante As String 'local copy
Private mvardtData As Date 'local copy
Private mvardtDataLimite As Date 'local copy
Private mvardMinPercRecItens As Double 'local copy
Private mvariBaixa As Integer 'local copy
Private mvarlNumIntReq As Long 'local copy
Private mvariFilialEmpresa As Integer 'local copy
Private mvariTipoTributacao As Integer 'local copy
Private mvarsCodigoOP As String 'local copy

Public Property Let sCodigoOP(ByVal vData As String)
    mvarsCodigoOP = vData
End Property

Public Property Get sCodigoOP() As String
    sCodigoOP = mvarsCodigoOP
End Property


Public Property Let iTipoTributacao(ByVal vData As Integer)
    mvariTipoTributacao = vData
End Property


Public Property Get iTipoTributacao() As Integer
    iTipoTributacao = mvariTipoTributacao
End Property


Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property


Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property



Public Property Let lNumIntReq(ByVal vData As Long)
    mvarlNumIntReq = vData
End Property


Public Property Get lNumIntReq() As Long
    lNumIntReq = mvarlNumIntReq
End Property



Public Property Let iBaixa(ByVal vData As Integer)
    mvariBaixa = vData
End Property


Public Property Get iBaixa() As Integer
    iBaixa = mvariBaixa
End Property



Public Property Let dMinPercRecItens(ByVal vData As Double)
    mvardMinPercRecItens = vData
End Property


Public Property Get dMinPercRecItens() As Double
    dMinPercRecItens = mvardMinPercRecItens
End Property



Public Property Let dtDataLimite(ByVal vData As Date)
    mvardtDataLimite = vData
End Property


Public Property Get dtDataLimite() As Date
    dtDataLimite = mvardtDataLimite
End Property



Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property


Public Property Get dtData() As Date
    dtData = mvardtData
End Property



Public Property Let sNomeRequisitante(ByVal vData As String)
    mvarsNomeRequisitante = vData
End Property


Public Property Get sNomeRequisitante() As String
    sNomeRequisitante = mvarsNomeRequisitante
End Property



Public Property Let lRequisitante(ByVal vData As Long)
    mvarlRequisitante = vData
End Property


Public Property Get lRequisitante() As Long
    lRequisitante = mvarlRequisitante
End Property



Public Property Let sCcl(ByVal vData As String)
    mvarsCCL = vData
End Property


Public Property Get sCcl() As String
    sCcl = mvarsCCL
End Property



Public Property Let lCodRequisicao(ByVal vData As Long)
    mvarlCodRequisicao = vData
End Property


Public Property Get lCodRequisicao() As Long
    lCodRequisicao = mvarlCodRequisicao
End Property



