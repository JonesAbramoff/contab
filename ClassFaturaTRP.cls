VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFaturaTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvariTipoNumIntDoc As Integer
Private mvarsMotivo As String
Private mvarsUsuarioGerFat As String
Private mvardtDataGerFat As Date
Private mvardHoraGerFat As Double
Private mvarlClienteVou As Long
Private mvardValorAporte As Double
Private mvardPercAporte As Double
Private mvardtDataPagto As Date
Private mvardtDataRegPagto As Date
Private mvardHoraRegPagto As Double
Private mvarsUsuarioCanc As String
Private mvardtDataCanc As Date
Private mvardHoraCanc As Double
Private mvarlNumFat As Long

Public Property Let lNumFat(ByVal vData As Long)
    mvarlNumFat = vData
End Property

Public Property Get lNumFat() As Long
    lNumFat = mvarlNumFat
End Property

Public Property Let sUsuarioCanc(ByVal vData As String)
    mvarsUsuarioCanc = vData
End Property

Public Property Get sUsuarioCanc() As String
    sUsuarioCanc = mvarsUsuarioCanc
End Property

Public Property Let dtDataCanc(ByVal vData As Date)
    mvardtDataCanc = vData
End Property

Public Property Get dtDataCanc() As Date
    dtDataCanc = mvardtDataCanc
End Property

Public Property Let dHoraCanc(ByVal vData As Double)
    mvardHoraCanc = vData
End Property

Public Property Get dHoraCanc() As Double
    dHoraCanc = mvardHoraCanc
End Property

Public Property Let dHoraRegPagto(ByVal vData As Double)
    mvardHoraRegPagto = vData
End Property

Public Property Get dHoraRegPagto() As Double
    dHoraRegPagto = mvardHoraRegPagto
End Property

Public Property Let dtDataRegPagto(ByVal vData As Date)
    mvardtDataRegPagto = vData
End Property

Public Property Get dtDataRegPagto() As Date
    dtDataRegPagto = mvardtDataRegPagto
End Property

Public Property Let dtDataPagto(ByVal vData As Date)
    mvardtDataPagto = vData
End Property

Public Property Get dtDataPagto() As Date
    dtDataPagto = mvardtDataPagto
End Property

Public Property Let dPercAporte(ByVal vData As Double)
    mvardPercAporte = vData
End Property

Public Property Get dPercAporte() As Double
    dPercAporte = mvardPercAporte
End Property

Public Property Let dValorAporte(ByVal vData As Double)
    mvardValorAporte = vData
End Property

Public Property Get dValorAporte() As Double
    dValorAporte = mvardValorAporte
End Property

Public Property Let lClienteVou(ByVal vData As Long)
    mvarlClienteVou = vData
End Property

Public Property Get lClienteVou() As Long
    lClienteVou = mvarlClienteVou
End Property

Public Property Let sUsuarioGerFat(ByVal vData As String)
    mvarsUsuarioGerFat = vData
End Property

Public Property Get sUsuarioGerFat() As String
    sUsuarioGerFat = mvarsUsuarioGerFat
End Property

Public Property Let dtDataGerFat(ByVal vData As Date)
    mvardtDataGerFat = vData
End Property

Public Property Get dtDataGerFat() As Date
    dtDataGerFat = mvardtDataGerFat
End Property

Public Property Let dHoraGerFat(ByVal vData As Double)
    mvardHoraGerFat = vData
End Property

Public Property Get dHoraGerFat() As Double
    dHoraGerFat = mvardHoraGerFat
End Property

Public Property Let sMotivo(ByVal vData As String)
    mvarsMotivo = vData
End Property

Public Property Get sMotivo() As String
    sMotivo = mvarsMotivo
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iTipoNumIntDoc(ByVal vData As Integer)
    mvariTipoNumIntDoc = vData
End Property

Public Property Get iTipoNumIntDoc() As Integer
    iTipoNumIntDoc = mvariTipoNumIntDoc
End Property

Private Sub Class_Initialize()

    mvardtDataGerFat = DATA_NULA
    mvardtDataPagto = DATA_NULA
    mvardtDataRegPagto = DATA_NULA
    mvardtDataCanc = DATA_NULA

End Sub
