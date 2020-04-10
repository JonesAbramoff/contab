VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPRJRecebPagto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumIntDocPRJ As Long
Private mvariTipo As Integer
Private mvarlNumero As Long
Private mvardValor As Double
Private mvarlCliForn As Long
Private mvariFilial As Integer
Private mvarlNumIntDocProposta As Long
Private mvariIncluiCFF As Integer
Private mvariFilialEmpresa As Integer
Private mvarcolRegras As New Collection
Private mvarlNumIntDocContrato As Long

Public Property Let lNumIntDocContrato(ByVal vData As Long)
    mvarlNumIntDocContrato = vData
End Property

Public Property Get lNumIntDocContrato() As Long
    lNumIntDocContrato = mvarlNumIntDocContrato
End Property

Public Property Set colRegras(ByVal vData As Collection)
    Set mvarcolRegras = vData
End Property

Public Property Get colRegras() As Collection
    Set colRegras = mvarcolRegras
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntDocPRJ(ByVal vData As Long)
    mvarlNumIntDocPRJ = vData
End Property

Public Property Get lNumIntDocPRJ() As Long
    lNumIntDocPRJ = mvarlNumIntDocPRJ
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let lNumero(ByVal vData As Long)
    mvarlNumero = vData
End Property

Public Property Get lNumero() As Long
    lNumero = mvarlNumero
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let lCliForn(ByVal vData As Long)
    mvarlCliForn = vData
End Property

Public Property Get lCliForn() As Long
    lCliForn = mvarlCliForn
End Property

Public Property Let iFilial(ByVal vData As Integer)
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property

Public Property Let lNumIntDocProposta(ByVal vData As Long)
    mvarlNumIntDocProposta = vData
End Property

Public Property Get lNumIntDocProposta() As Long
    lNumIntDocProposta = mvarlNumIntDocProposta
End Property

Public Property Let iIncluiCFF(ByVal vData As Integer)
    mvariIncluiCFF = vData
End Property

Public Property Get iIncluiCFF() As Integer
    iIncluiCFF = mvariIncluiCFF
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property
