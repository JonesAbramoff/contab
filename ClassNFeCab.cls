VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassNFeCab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc AS Long
Private mvarsNomeArquivo AS String
Private mvardtDataImportacao AS Date
Private mvardHoraImportacao AS Double
Private mvarsUsuario AS String
Private mvarsVersao AS String
Private mvarlInscricaoMunicipal AS Long
Private mvardtDataInicio AS Date
Private mvardtDataFim AS Date

Public Property Let lNumIntDoc (ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc () AS Long
    lNumIntDoc= mvarlNumIntDoc
End Property

Public Property Let sNomeArquivo (ByVal vData As String)
    mvarsNomeArquivo = vData
End Property

Public Property Get sNomeArquivo () AS String
    sNomeArquivo= mvarsNomeArquivo
End Property

Public Property Let dtDataImportacao (ByVal vData As Date)
    mvardtDataImportacao = vData
End Property

Public Property Get dtDataImportacao () AS Date
    dtDataImportacao= mvardtDataImportacao
End Property

Public Property Let dHoraImportacao (ByVal vData As Double)
    mvardHoraImportacao = vData
End Property

Public Property Get dHoraImportacao () AS Double
    dHoraImportacao= mvardHoraImportacao
End Property

Public Property Let sUsuario (ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario () AS String
    sUsuario= mvarsUsuario
End Property

Public Property Let sVersao (ByVal vData As String)
    mvarsVersao = vData
End Property

Public Property Get sVersao () AS String
    sVersao= mvarsVersao
End Property

Public Property Let lInscricaoMunicipal (ByVal vData As Long)
    mvarlInscricaoMunicipal = vData
End Property

Public Property Get lInscricaoMunicipal () AS Long
    lInscricaoMunicipal= mvarlInscricaoMunicipal
End Property

Public Property Let dtDataInicio (ByVal vData As Date)
    mvardtDataInicio = vData
End Property

Public Property Get dtDataInicio () AS Date
    dtDataInicio= mvardtDataInicio
End Property

Public Property Let dtDataFim (ByVal vData As Date)
    mvardtDataFim = vData
End Property

Public Property Get dtDataFim () AS Date
    dtDataFim= mvardtDataFim
End Property
