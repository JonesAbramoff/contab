VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassClientesAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlCliente As Long
Private mvariTipoAssociado As Integer
Private mvarlMatriculaPrinc As Long
Private mvariStatusAssociado As Integer
Private mvarlMatriculaSec As Long
Private mvarlMatriculaAF As Long
Private mvarlMatriculaFRG As Long
Private mvariEmpresa1 As Integer
Private mvariEmpresa2 As Integer
Private mvarsSexo As String
Private mvarsRGOrgaoEmissor As String
Private mvardtDataExpedicaoRG As Date
Private mvardtDataNascimento As Date
Private mvardtDataInscricao As Date
Private mvarsLocalTrabalho As String
Private mvardtDataAdmissaoFurnas As Date
Private mvarsCargo As String
Private mvardtDataAposINSS As Date
Private mvardtDataAposFRG As Date
Private mvariTipoApos As Integer
Private mvardtDataConBenf As Date
Private mvarsNoBenef As String
Private mvardtDataFalecimento As Date
Private mvarsContrSoc As String
Private mvarsObservacao1 As String
Private mvarsObservacao2 As String
Private mvariBenemerito As Integer
Private mvariFundador As Integer
Private mvardtDataUltAtualizacao As Date
Private mvariMatriculaPrincDV As Integer

Public Property Let iMatriculaPrincDV(ByVal vData As Integer)
    mvariMatriculaPrincDV = vData
End Property

Public Property Get iMatriculaPrincDV() As Integer
    iMatriculaPrincDV = mvariMatriculaPrincDV
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let iTipoAssociado(ByVal vData As Integer)
    mvariTipoAssociado = vData
End Property

Public Property Get iTipoAssociado() As Integer
    iTipoAssociado = mvariTipoAssociado
End Property

Public Property Let lMatriculaPrinc(ByVal vData As Long)
    mvarlMatriculaPrinc = vData
End Property

Public Property Get lMatriculaPrinc() As Long
    lMatriculaPrinc = mvarlMatriculaPrinc
End Property

Public Property Let iStatusAssociado(ByVal vData As Integer)
    mvariStatusAssociado = vData
End Property

Public Property Get iStatusAssociado() As Integer
    iStatusAssociado = mvariStatusAssociado
End Property

Public Property Let lMatriculaSec(ByVal vData As Long)
    mvarlMatriculaSec = vData
End Property

Public Property Get lMatriculaSec() As Long
    lMatriculaSec = mvarlMatriculaSec
End Property

Public Property Let lMatriculaAF(ByVal vData As Long)
    mvarlMatriculaAF = vData
End Property

Public Property Get lMatriculaAF() As Long
    lMatriculaAF = mvarlMatriculaAF
End Property

Public Property Let lMatriculaFRG(ByVal vData As Long)
    mvarlMatriculaFRG = vData
End Property

Public Property Get lMatriculaFRG() As Long
    lMatriculaFRG = mvarlMatriculaFRG
End Property

Public Property Let iEmpresa1(ByVal vData As Integer)
    mvariEmpresa1 = vData
End Property

Public Property Get iEmpresa1() As Integer
    iEmpresa1 = mvariEmpresa1
End Property

Public Property Let iEmpresa2(ByVal vData As Integer)
    mvariEmpresa2 = vData
End Property

Public Property Get iEmpresa2() As Integer
    iEmpresa2 = mvariEmpresa2
End Property

Public Property Let sSexo(ByVal vData As String)
    mvarsSexo = vData
End Property

Public Property Get sSexo() As String
    sSexo = mvarsSexo
End Property

Public Property Let sRGOrgaoEmissor(ByVal vData As String)
    mvarsRGOrgaoEmissor = vData
End Property

Public Property Get sRGOrgaoEmissor() As String
    sRGOrgaoEmissor = mvarsRGOrgaoEmissor
End Property

Public Property Let dtDataExpedicaoRG(ByVal vData As Date)
    mvardtDataExpedicaoRG = vData
End Property

Public Property Get dtDataExpedicaoRG() As Date
    dtDataExpedicaoRG = mvardtDataExpedicaoRG
End Property

Public Property Let dtDataNascimento(ByVal vData As Date)
    mvardtDataNascimento = vData
End Property

Public Property Get dtDataNascimento() As Date
    dtDataNascimento = mvardtDataNascimento
End Property

Public Property Let dtDataInscricao(ByVal vData As Date)
    mvardtDataInscricao = vData
End Property

Public Property Get dtDataInscricao() As Date
    dtDataInscricao = mvardtDataInscricao
End Property

Public Property Let sLocalTrabalho(ByVal vData As String)
    mvarsLocalTrabalho = vData
End Property

Public Property Get sLocalTrabalho() As String
    sLocalTrabalho = mvarsLocalTrabalho
End Property

Public Property Let dtDataAdmissaoFurnas(ByVal vData As Date)
    mvardtDataAdmissaoFurnas = vData
End Property

Public Property Get dtDataAdmissaoFurnas() As Date
    dtDataAdmissaoFurnas = mvardtDataAdmissaoFurnas
End Property

Public Property Let sCargo(ByVal vData As String)
    mvarsCargo = vData
End Property

Public Property Get sCargo() As String
    sCargo = mvarsCargo
End Property

Public Property Let dtDataAposINSS(ByVal vData As Date)
    mvardtDataAposINSS = vData
End Property

Public Property Get dtDataAposINSS() As Date
    dtDataAposINSS = mvardtDataAposINSS
End Property

Public Property Let dtDataAposFRG(ByVal vData As Date)
    mvardtDataAposFRG = vData
End Property

Public Property Get dtDataAposFRG() As Date
    dtDataAposFRG = mvardtDataAposFRG
End Property

Public Property Let iTipoApos(ByVal vData As Integer)
    mvariTipoApos = vData
End Property

Public Property Get iTipoApos() As Integer
    iTipoApos = mvariTipoApos
End Property

Public Property Let dtDataConBenf(ByVal vData As Date)
    mvardtDataConBenf = vData
End Property

Public Property Get dtDataConBenf() As Date
    dtDataConBenf = mvardtDataConBenf
End Property

Public Property Let sNoBenef(ByVal vData As String)
    mvarsNoBenef = vData
End Property

Public Property Get sNoBenef() As String
    sNoBenef = mvarsNoBenef
End Property

Public Property Let dtDataFalecimento(ByVal vData As Date)
    mvardtDataFalecimento = vData
End Property

Public Property Get dtDataFalecimento() As Date
    dtDataFalecimento = mvardtDataFalecimento
End Property

Public Property Let sContrSoc(ByVal vData As String)
    mvarsContrSoc = vData
End Property

Public Property Get sContrSoc() As String
    sContrSoc = mvarsContrSoc
End Property

Public Property Let sObservacao1(ByVal vData As String)
    mvarsObservacao1 = vData
End Property

Public Property Get sObservacao1() As String
    sObservacao1 = mvarsObservacao1
End Property

Public Property Let sObservacao2(ByVal vData As String)
    mvarsObservacao2 = vData
End Property

Public Property Get sObservacao2() As String
    sObservacao2 = mvarsObservacao2
End Property

Public Property Let iBenemerito(ByVal vData As Integer)
    mvariBenemerito = vData
End Property

Public Property Get iBenemerito() As Integer
    iBenemerito = mvariBenemerito
End Property

Public Property Let iFundador(ByVal vData As Integer)
    mvariFundador = vData
End Property

Public Property Get iFundador() As Integer
    iFundador = mvariFundador
End Property

Public Property Let dtDataUltAtualizacao(ByVal vData As Date)
    mvardtDataUltAtualizacao = vData
End Property

Public Property Get dtDataUltAtualizacao() As Date
    dtDataUltAtualizacao = mvardtDataUltAtualizacao
End Property

Private Sub Class_Initialize()
    mvardtDataAdmissaoFurnas = DATA_NULA
    mvardtDataAposFRG = DATA_NULA
    mvardtDataAposINSS = DATA_NULA
    mvardtDataConBenf = DATA_NULA
    mvardtDataExpedicaoRG = DATA_NULA
    mvardtDataFalecimento = DATA_NULA
    mvardtDataInscricao = DATA_NULA
    mvardtDataNascimento = DATA_NULA
    mvardtDataUltAtualizacao = DATA_NULA
End Sub
