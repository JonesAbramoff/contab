VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRPSCab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarsNomeArquivo As String
Private mvardtDataGeracao As Date
Private mvardHoraGeracao As Double
Private mvarsUsuario As String
Private mvarsVersao As String
Private mvarlInscricaoMunicipal As Long
Private mvardtDataInicio As Date
Private mvardtDataFim As Date
Private mvariAtualizaDadosCliEnd As Integer
Private mvariFilialEmpresa As Integer
Private mvarlRPSDe As Long

Public Property Let lRPSDe(ByVal vData As Long)
    mvarlRPSDe = vData
End Property

Public Property Get lRPSDe() As Long
    lRPSDe = mvarlRPSDe
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iAtualizaDadosCliEnd(ByVal vData As Integer)
    mvariAtualizaDadosCliEnd = vData
End Property

Public Property Get iAtualizaDadosCliEnd() As Integer
    iAtualizaDadosCliEnd = mvariAtualizaDadosCliEnd
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sNomeArquivo(ByVal vData As String)
    mvarsNomeArquivo = vData
End Property

Public Property Get sNomeArquivo() As String
    sNomeArquivo = mvarsNomeArquivo
End Property

Public Property Let dtDataGeracao(ByVal vData As Date)
    mvardtDataGeracao = vData
End Property

Public Property Get dtDataGeracao() As Date
    dtDataGeracao = mvardtDataGeracao
End Property

Public Property Let dHoraGeracao(ByVal vData As Double)
    mvardHoraGeracao = vData
End Property

Public Property Get dHoraGeracao() As Double
    dHoraGeracao = mvardHoraGeracao
End Property

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

Public Property Let sVersao(ByVal vData As String)
    mvarsVersao = vData
End Property

Public Property Get sVersao() As String
    sVersao = mvarsVersao
End Property

Public Property Let lInscricaoMunicipal(ByVal vData As Long)
    mvarlInscricaoMunicipal = vData
End Property

Public Property Get lInscricaoMunicipal() As Long
    lInscricaoMunicipal = mvarlInscricaoMunicipal
End Property

Public Property Let dtDataInicio(ByVal vData As Date)
    mvardtDataInicio = vData
End Property

Public Property Get dtDataInicio() As Date
    dtDataInicio = mvardtDataInicio
End Property

Public Property Let dtDataFim(ByVal vData As Date)
    mvardtDataFim = vData
End Property

Public Property Get dtDataFim() As Date
    dtDataFim = mvardtDataFim
End Property
