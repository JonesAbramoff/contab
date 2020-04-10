VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassRetPagto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc AS Long
Private mvarsNomeArq AS String
Private mvardtDataImport AS Date
Private mvardHoraImport AS Double
Private mvariBanco AS Integer
Private mvarsAgencia AS String
Private mvarsConta AS String
Private mvardtDataGeracao AS Date
Private mvardHoraGeracao AS Double
Private mvarlSeqArquivo AS Long
Private mvarsNomeEmpresa AS String
Private mvarsNomeBanco AS String

Public Property Let lNumIntDoc (ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc () AS Long
    lNumIntDoc= mvarlNumIntDoc
End Property

Public Property Let sNomeArq (ByVal vData As String)
    mvarsNomeArq = vData
End Property

Public Property Get sNomeArq () AS String
    sNomeArq= mvarsNomeArq
End Property

Public Property Let dtDataImport (ByVal vData As Date)
    mvardtDataImport = vData
End Property

Public Property Get dtDataImport () AS Date
    dtDataImport= mvardtDataImport
End Property

Public Property Let dHoraImport (ByVal vData As Double)
    mvardHoraImport = vData
End Property

Public Property Get dHoraImport () AS Double
    dHoraImport= mvardHoraImport
End Property

Public Property Let iBanco (ByVal vData As Integer)
    mvariBanco = vData
End Property

Public Property Get iBanco () AS Integer
    iBanco= mvariBanco
End Property

Public Property Let sAgencia (ByVal vData As String)
    mvarsAgencia = vData
End Property

Public Property Get sAgencia () AS String
    sAgencia= mvarsAgencia
End Property

Public Property Let sConta (ByVal vData As String)
    mvarsConta = vData
End Property

Public Property Get sConta () AS String
    sConta= mvarsConta
End Property

Public Property Let dtDataGeracao (ByVal vData As Date)
    mvardtDataGeracao = vData
End Property

Public Property Get dtDataGeracao () AS Date
    dtDataGeracao= mvardtDataGeracao
End Property

Public Property Let dHoraGeracao (ByVal vData As Double)
    mvardHoraGeracao = vData
End Property

Public Property Get dHoraGeracao () AS Double
    dHoraGeracao= mvardHoraGeracao
End Property

Public Property Let lSeqArquivo (ByVal vData As Long)
    mvarlSeqArquivo = vData
End Property

Public Property Get lSeqArquivo () AS Long
    lSeqArquivo= mvarlSeqArquivo
End Property

Public Property Let sNomeEmpresa (ByVal vData As String)
    mvarsNomeEmpresa = vData
End Property

Public Property Get sNomeEmpresa () AS String
    sNomeEmpresa= mvarsNomeEmpresa
End Property

Public Property Let sNomeBanco (ByVal vData As String)
    mvarsNomeBanco = vData
End Property

Public Property Get sNomeBanco () AS String
    sNomeBanco= mvarsNomeBanco
End Property
