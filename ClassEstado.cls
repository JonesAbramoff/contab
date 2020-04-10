VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEstado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsSigla As String
Private mvarsNome As String
Private mvardICMSAliquotaInterna As Double
Private mvardICMSAliquotaExportacao As Double
Private mvardICMSAliquotaImportacao As Double
Private mvariCodIBGE As Integer
Private mvardICMSPercFCP As Double

Private mvardICMSAliquotaInternaAnt As Double
Private mvardICMSAliquotaImportacaoAnt As Double
Private mvardICMSPercFCPAnt As Double

Private mvardtDataIniAliqInternaAtual As Date
Private mvardtDataIniAliqImportacaoAtual As Date
Private mvardtDataIniAliqFCPAtual As Date

Public Property Let dtDataIniAliqFCPAtual(ByVal vData As Date)
    mvardtDataIniAliqFCPAtual = vData
End Property

Public Property Get dtDataIniAliqFCPAtual() As Date
    dtDataIniAliqFCPAtual = mvardtDataIniAliqFCPAtual
End Property

Public Property Let dtDataIniAliqImportacaoAtual(ByVal vData As Date)
    mvardtDataIniAliqImportacaoAtual = vData
End Property

Public Property Get dtDataIniAliqImportacaoAtual() As Date
    dtDataIniAliqImportacaoAtual = mvardtDataIniAliqImportacaoAtual
End Property

Public Property Let dtDataIniAliqInternaAtual(ByVal vData As Date)
    mvardtDataIniAliqInternaAtual = vData
End Property

Public Property Get dtDataIniAliqInternaAtual() As Date
    dtDataIniAliqInternaAtual = mvardtDataIniAliqInternaAtual
End Property

Public Property Let dICMSAliquotaInternaAnt(ByVal vData As Double)
    mvardICMSAliquotaInternaAnt = vData
End Property

Public Property Get dICMSAliquotaInternaAnt() As Double
    dICMSAliquotaInternaAnt = mvardICMSAliquotaInternaAnt
End Property

Public Property Let dICMSAliquotaImportacaoAnt(ByVal vData As Double)
    mvardICMSAliquotaImportacaoAnt = vData
End Property

Public Property Get dICMSAliquotaImportacaoAnt() As Double
    dICMSAliquotaImportacaoAnt = mvardICMSAliquotaImportacaoAnt
End Property

Public Property Let dICMSPercFCPAnt(ByVal vData As Double)
    mvardICMSPercFCPAnt = vData
End Property

Public Property Get dICMSPercFCPAnt() As Double
    dICMSPercFCPAnt = mvardICMSPercFCPAnt
End Property

Public Property Let dICMSPercFCP(ByVal vData As Double)
    mvardICMSPercFCP = vData
End Property

Public Property Get dICMSPercFCP() As Double
    dICMSPercFCP = mvardICMSPercFCP
End Property

Public Property Let iCodIBGE(ByVal vData As Integer)
    mvariCodIBGE = vData
End Property

Public Property Get iCodIBGE() As Integer
    iCodIBGE = mvariCodIBGE
End Property

Public Property Let dICMSAliquotaImportacao(ByVal vData As Double)
    mvardICMSAliquotaImportacao = vData
End Property

Public Property Get dICMSAliquotaImportacao() As Double
    dICMSAliquotaImportacao = mvardICMSAliquotaImportacao
End Property

Public Property Let dICMSAliquotaExportacao(ByVal vData As Double)
    mvardICMSAliquotaExportacao = vData
End Property

Public Property Get dICMSAliquotaExportacao() As Double
    dICMSAliquotaExportacao = mvardICMSAliquotaExportacao
End Property

Public Property Let dICMSAliquotaInterna(ByVal vData As Double)
    mvardICMSAliquotaInterna = vData
End Property

Public Property Get dICMSAliquotaInterna() As Double
    dICMSAliquotaInterna = mvardICMSAliquotaInterna
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let sSigla(ByVal vData As String)
    mvarsSigla = vData
End Property

Public Property Get sSigla() As String
    sSigla = mvarsSigla
End Property

Private Sub Class_Initialize()
    mvardtDataIniAliqInternaAtual = DATA_NULA
    mvardtDataIniAliqFCPAtual = DATA_NULA
    mvardtDataIniAliqImportacaoAtual = DATA_NULA
End Sub
