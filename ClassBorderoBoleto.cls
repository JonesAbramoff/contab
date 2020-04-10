VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBorderoBoleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvariCodigoRede As Integer
Private mvarlNumBordero As Long
Private mvarsNumero As String
Private mvarsAdmMeioPagto As String
Private mvarsNomeRede As String
Private mvardtDataImpressao As Date
Private mvardtDataEnvio As Date
Private mvardtDataBackoffice As Date
Private mvarcolBorderoBoletoItem As New Collection
Private mvardValorEnviar As Double
Private mvardValorEnviarNCC As Double
Private mvariExibeTela As Integer
Private mvardValorEnviarNCD As Double

Public Property Let dValorEnviarNCD(ByVal vData As Double)
    mvardValorEnviarNCD = vData
End Property

Public Property Get dValorEnviarNCD() As Double
    dValorEnviarNCD = mvardValorEnviarNCD
End Property

Public Property Let sNomeRede(ByVal vData As String)
    mvarsNomeRede = vData
End Property

Public Property Get sNomeRede() As String
    sNomeRede = mvarsNomeRede
End Property

Public Property Let iExibeTela(ByVal vData As Integer)
    mvariExibeTela = vData
End Property

Public Property Get iExibeTela() As Integer
    iExibeTela = mvariExibeTela
End Property

Public Property Let dValorEnviarNCC(ByVal vData As Double)
    mvardValorEnviarNCC = vData
End Property

Public Property Get dValorEnviarNCC() As Double
    dValorEnviarNCC = mvardValorEnviarNCC
End Property

Public Property Let dValorEnviar(ByVal vData As Double)
    mvardValorEnviar = vData
End Property

Public Property Get dValorEnviar() As Double
    dValorEnviar = mvardValorEnviar
End Property

Public Property Set colBorderoBoletoItem(ByVal vData As Collection)
    Set mvarcolBorderoBoletoItem = vData
End Property

Public Property Get colBorderoBoletoItem() As Collection
    Set colBorderoBoletoItem = mvarcolBorderoBoletoItem
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

Public Property Let sAdmMeioPagto(ByVal vData As String)
    mvarsAdmMeioPagto = vData
End Property

Public Property Get sAdmMeioPagto() As String
    sAdmMeioPagto = mvarsAdmMeioPagto
End Property

Public Property Let sNumero(ByVal vData As String)
    mvarsNumero = vData
End Property

Public Property Get sNumero() As String
    sNumero = mvarsNumero
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

Public Property Let iCodigoRede(ByVal vData As Integer)
    mvariCodigoRede = vData
End Property

Public Property Get iCodigoRede() As Integer
    iCodigoRede = mvariCodigoRede
End Property



