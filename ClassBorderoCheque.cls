VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBorderoCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarlNumBordero As Long
Private mvariCodNossaConta As Integer 'se o codigo nossa conta não estiver preenchido ==> bordero de transf. backoffice.
Private mvardtDataImpressao As Date
Private mvardtDataEnvio As Date
Private mvardtDataBackoffice As Date
Private mvarcolCheque As New Collection
Private mvarcolChequeN As New Collection
Private mvardValorNEspec As Double
Private mvardValorEspec As Double
Private mvarlNumMovtoCx As Long

Public Property Let lNumMovtoCx(ByVal vData As Long)
    mvarlNumMovtoCx = vData
End Property

Public Property Get lNumMovtoCx() As Long
    lNumMovtoCx = mvarlNumMovtoCx
End Property

Public Property Let dValorEspec(ByVal vData As Double)
    mvardValorEspec = vData
End Property

Public Property Get dValorEspec() As Double
    dValorEspec = mvardValorEspec
End Property

Public Property Let dValorNEspec(ByVal vData As Double)
    mvardValorNEspec = vData
End Property

Public Property Get dValorNEspec() As Double
    dValorNEspec = mvardValorNEspec
End Property

Public Property Set colChequeN(ByVal vData As Collection)
    Set mvarcolChequeN = vData
End Property

Public Property Get colChequeN() As Collection
    Set colChequeN = mvarcolChequeN
End Property

Public Property Set colCheque(ByVal vData As Collection)
    Set mvarcolCheque = vData
End Property

Public Property Get colCheque() As Collection
    Set colCheque = mvarcolCheque
End Property

Public Property Let dtDataBackoffice(ByVal vData As Date)
    mvardtDataBackoffice = vData
End Property

Public Property Get dtDataBackoffice() As Date
    dtDataBackoffice = mvardtDataBackoffice
End Property

Public Property Let dtDataEnvio(ByVal vData As Date)
    mvardtDataEnvio = vData
End Property

Public Property Get dtDataEnvio() As Date
    dtDataEnvio = mvardtDataEnvio
End Property

Public Property Let dtDataImpressao(ByVal vData As Date)
    mvardtDataImpressao = vData
End Property

Public Property Get dtDataImpressao() As Date
    dtDataImpressao = mvardtDataImpressao
End Property

Public Property Let iCodNossaConta(ByVal vData As Integer)
    mvariCodNossaConta = vData
End Property

Public Property Get iCodNossaConta() As Integer
    iCodNossaConta = mvariCodNossaConta
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



