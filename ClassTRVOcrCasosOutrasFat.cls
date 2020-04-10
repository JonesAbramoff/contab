VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVOcrCasosOutrasFat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDocOcrCaso As Long
Private mvariSeq As Integer
Private mvariConsiderar As Integer
Private mvardtDataRecepcao As Date
Private mvardtDataFatura As Date
Private mvarsNumero As String
Private mvardValorUS As Double
Private mvardValorRS As Double
Private mvarsDescricao As String

Public Property Let lNumIntDocOcrCaso(ByVal vData As Long)
    mvarlNumIntDocOcrCaso = vData
End Property

Public Property Get lNumIntDocOcrCaso() As Long
    lNumIntDocOcrCaso = mvarlNumIntDocOcrCaso
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let iConsiderar(ByVal vData As Integer)
    mvariConsiderar = vData
End Property

Public Property Get iConsiderar() As Integer
    iConsiderar = mvariConsiderar
End Property

Public Property Let dtDataRecepcao(ByVal vData As Date)
    mvardtDataRecepcao = vData
End Property

Public Property Get dtDataRecepcao() As Date
    dtDataRecepcao = mvardtDataRecepcao
End Property

Public Property Let dtDataFatura(ByVal vData As Date)
    mvardtDataFatura = vData
End Property

Public Property Get dtDataFatura() As Date
    dtDataFatura = mvardtDataFatura
End Property

Public Property Let sNumero(ByVal vData As String)
    mvarsNumero = vData
End Property

Public Property Get sNumero() As String
    sNumero = mvarsNumero
End Property

Public Property Let dValorUS(ByVal vData As Double)
    mvardValorUS = vData
End Property

Public Property Get dValorUS() As Double
    dValorUS = mvardValorUS
End Property

Public Property Let dValorRS(ByVal vData As Double)
    mvardValorRS = vData
End Property

Public Property Get dValorRS() As Double
    dValorRS = mvardValorRS
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property
