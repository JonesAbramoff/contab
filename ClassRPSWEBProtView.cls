VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassRPSWEBProtView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsLote AS String
Private mvariFilialEmpresa AS Integer
Private mvariTipoRPS AS Integer
Private mvarsSerieRPS AS String
Private mvarsNumeroRPS AS String
Private mvariAmbiente AS Integer
Private mvardtData AS Date
Private mvardHora AS Double
Private mvarsProtocolo AS String
Private mvarsCodigoVerificacao AS String
Private mvardtCompetencia AS Date
Private mvardtDataEmissaoRPS AS Date
Private mvarsRazaoSocialCli AS String
Private mvarsNomeReduzidoCli AS String
Private mvarsRazaoSocialForn AS String
Private mvarsNomeReduzidoForn AS String
Private mvarsId AS String

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

Public Property Let iTipoRPS (ByVal vData As Integer)
    mvariTipoRPS = vData
End Property

Public Property Get iTipoRPS () AS Integer
    iTipoRPS= mvariTipoRPS
End Property

Public Property Let sSerieRPS (ByVal vData As String)
    mvarsSerieRPS = vData
End Property

Public Property Get sSerieRPS () AS String
    sSerieRPS= mvarsSerieRPS
End Property

Public Property Let sNumeroRPS (ByVal vData As String)
    mvarsNumeroRPS = vData
End Property

Public Property Get sNumeroRPS () AS String
    sNumeroRPS= mvarsNumeroRPS
End Property

Public Property Let iAmbiente (ByVal vData As Integer)
    mvariAmbiente = vData
End Property

Public Property Get iAmbiente () AS Integer
    iAmbiente= mvariAmbiente
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

Public Property Let sCodigoVerificacao (ByVal vData As String)
    mvarsCodigoVerificacao = vData
End Property

Public Property Get sCodigoVerificacao () AS String
    sCodigoVerificacao= mvarsCodigoVerificacao
End Property

Public Property Let dtCompetencia (ByVal vData As Date)
    mvardtCompetencia = vData
End Property

Public Property Get dtCompetencia () AS Date
    dtCompetencia= mvardtCompetencia
End Property

Public Property Let dtDataEmissaoRPS (ByVal vData As Date)
    mvardtDataEmissaoRPS = vData
End Property

Public Property Get dtDataEmissaoRPS () AS Date
    dtDataEmissaoRPS= mvardtDataEmissaoRPS
End Property

Public Property Let sRazaoSocialCli (ByVal vData As String)
    mvarsRazaoSocialCli = vData
End Property

Public Property Get sRazaoSocialCli () AS String
    sRazaoSocialCli= mvarsRazaoSocialCli
End Property

Public Property Let sNomeReduzidoCli (ByVal vData As String)
    mvarsNomeReduzidoCli = vData
End Property

Public Property Get sNomeReduzidoCli () AS String
    sNomeReduzidoCli= mvarsNomeReduzidoCli
End Property

Public Property Let sRazaoSocialForn (ByVal vData As String)
    mvarsRazaoSocialForn = vData
End Property

Public Property Get sRazaoSocialForn () AS String
    sRazaoSocialForn= mvarsRazaoSocialForn
End Property

Public Property Let sNomeReduzidoForn (ByVal vData As String)
    mvarsNomeReduzidoForn = vData
End Property

Public Property Get sNomeReduzidoForn () AS String
    sNomeReduzidoForn= mvarsNomeReduzidoForn
End Property

Public Property Let sId (ByVal vData As String)
    mvarsId = vData
End Property

Public Property Get sId () AS String
    sId= mvarsId
End Property
