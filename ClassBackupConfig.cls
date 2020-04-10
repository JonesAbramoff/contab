VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBackupConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlCodigo As Long
Private mvarsDescricao As String
Private mvariHabilitado As Integer
Private mvardtDataInicio As Date
Private mvardHora As Double
Private mvariRepetirDias As Integer
Private mvarsDiretorio As String
Private mvariIncluirDataNomeArq As Integer
Private mvardtDataUltBkp As Date
Private mvardtDataProxBkp As Date
Private mvariCompactar As Integer
Private mvariTransfFTP As Integer
Private mvarsFTPURL As String
Private mvarsFTPUsu As String
Private mvarsFTPSenha As String
Private mvarsFTPDir As String
Private mvarsDirDownload As String
Private mvarbApenasBkpeRestoreEmpTeste As Boolean

Public Property Let bApenasBkpeRestoreEmpTeste(ByVal vData As Boolean)
    mvarbApenasBkpeRestoreEmpTeste = vData
End Property

Public Property Get bApenasBkpeRestoreEmpTeste() As Boolean
    bApenasBkpeRestoreEmpTeste = mvarbApenasBkpeRestoreEmpTeste
End Property

Public Property Let sDirDownload(ByVal vData As String)
    mvarsDirDownload = vData
End Property

Public Property Get sDirDownload() As String
    sDirDownload = mvarsDirDownload
End Property

Public Property Let sFTPDir(ByVal vData As String)
    mvarsFTPDir = vData
End Property

Public Property Get sFTPDir() As String
    sFTPDir = mvarsFTPDir
End Property

Public Property Let sFTPSenha(ByVal vData As String)
    mvarsFTPSenha = vData
End Property

Public Property Get sFTPSenha() As String
    sFTPSenha = mvarsFTPSenha
End Property

Public Property Let sFTPUsu(ByVal vData As String)
    mvarsFTPUsu = vData
End Property

Public Property Get sFTPUsu() As String
    sFTPUsu = mvarsFTPUsu
End Property

Public Property Let sFTPURL(ByVal vData As String)
    mvarsFTPURL = vData
End Property

Public Property Get sFTPURL() As String
    sFTPURL = mvarsFTPURL
End Property

Public Property Let iTransfFTP(ByVal vData As Integer)
    mvariTransfFTP = vData
End Property

Public Property Get iTransfFTP() As Integer
    iTransfFTP = mvariTransfFTP
End Property

Public Property Let iCompactar(ByVal vData As Integer)
    mvariCompactar = vData
End Property

Public Property Get iCompactar() As Integer
    iCompactar = mvariCompactar
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iHabilitado(ByVal vData As Integer)
    mvariHabilitado = vData
End Property

Public Property Get iHabilitado() As Integer
    iHabilitado = mvariHabilitado
End Property

Public Property Let dtDataInicio(ByVal vData As Date)
    mvardtDataInicio = vData
End Property

Public Property Get dtDataInicio() As Date
    dtDataInicio = mvardtDataInicio
End Property

Public Property Let dHora(ByVal vData As Double)
    mvardHora = vData
End Property

Public Property Get dHora() As Double
    dHora = mvardHora
End Property

Public Property Let iRepetirDias(ByVal vData As Integer)
    mvariRepetirDias = vData
End Property

Public Property Get iRepetirDias() As Integer
    iRepetirDias = mvariRepetirDias
End Property

Public Property Let sDiretorio(ByVal vData As String)
    mvarsDiretorio = vData
End Property

Public Property Get sDiretorio() As String
    sDiretorio = mvarsDiretorio
End Property

Public Property Let iIncluirDataNomeArq(ByVal vData As Integer)
    mvariIncluirDataNomeArq = vData
End Property

Public Property Get iIncluirDataNomeArq() As Integer
    iIncluirDataNomeArq = mvariIncluirDataNomeArq
End Property

Public Property Let dtDataUltBkp(ByVal vData As Date)
    mvardtDataUltBkp = vData
End Property

Public Property Get dtDataUltBkp() As Date
    dtDataUltBkp = mvardtDataUltBkp
End Property

Public Property Let dtDataProxBkp(ByVal vData As Date)
    mvardtDataProxBkp = vData
End Property

Public Property Get dtDataProxBkp() As Date
    dtDataProxBkp = mvardtDataProxBkp
End Property

Private Sub Class_Initialize()
    mvarbApenasBkpeRestoreEmpTeste = False
End Sub
