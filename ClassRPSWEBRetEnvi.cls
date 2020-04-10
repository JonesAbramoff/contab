VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassRPSWEBRetEnvi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsCorrecao AS String
Private mvarsMsg AS String
Private mvarsCodMsg AS String
Private mvardtData AS Date
Private mvardHora AS Double
Private mvarsProtocolo AS String
Private mvarsLote AS String
Private mvariFilialEmpresa AS Integer
Private mvardtDataRecebimento AS Date
Private mvardHoraRecebimento AS Double
Private mvarlNumIntDoc AS Long
Private mvariAmbiente AS Integer

Public Property Let sCorrecao (ByVal vData As String)
    mvarsCorrecao = vData
End Property

Public Property Get sCorrecao () AS String
    sCorrecao= mvarsCorrecao
End Property

Public Property Let sMsg (ByVal vData As String)
    mvarsMsg = vData
End Property

Public Property Get sMsg () AS String
    sMsg= mvarsMsg
End Property

Public Property Let sCodMsg (ByVal vData As String)
    mvarsCodMsg = vData
End Property

Public Property Get sCodMsg () AS String
    sCodMsg= mvarsCodMsg
End Property

Public Property Let dtData (ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData () AS Date
    dtData= mvardtData
End Property

Public Property Let dHora (ByVal vData As Double)
    mvardHora = vData
End Property

Public Property Get dHora () AS Double
    dHora= mvardHora
End Property

Public Property Let sProtocolo (ByVal vData As String)
    mvarsProtocolo = vData
End Property

Public Property Get sProtocolo () AS String
    sProtocolo= mvarsProtocolo
End Property

Public Property Let sLote (ByVal vData As String)
    mvarsLote = vData
End Property

Public Property Get sLote () AS String
    sLote= mvarsLote
End Property

Public Property Let iFilialEmpresa (ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa () AS Integer
    iFilialEmpresa= mvariFilialEmpresa
End Property

Public Property Let dtDataRecebimento (ByVal vData As Date)
    mvardtDataRecebimento = vData
End Property

Public Property Get dtDataRecebimento () AS Date
    dtDataRecebimento= mvardtDataRecebimento
End Property

Public Property Let dHoraRecebimento (ByVal vData As Double)
    mvardHoraRecebimento = vData
End Property

Public Property Get dHoraRecebimento () AS Double
    dHoraRecebimento= mvardHoraRecebimento
End Property

Public Property Let lNumIntDoc (ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc () AS Long
    lNumIntDoc= mvarlNumIntDoc
End Property

Public Property Let iAmbiente (ByVal vData As Integer)
    mvariAmbiente = vData
End Property

Public Property Get iAmbiente () AS Integer
    iAmbiente= mvariAmbiente
End Property
