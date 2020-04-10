VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassInfoArqICMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvariFilialEmpresa As Integer
Private mvardtDataInicial As Date
Private mvardtDataFinal As Date
Private mvarsCGC As String
Private mvarsInscricaoEstadual As String
Private mvarsNome As String
Private mvarsMunicipio As String
Private mvarsUF As String
Private mvarsLogradouro As String
Private mvarlNumero As Long
Private mvarsComplemento As String
Private mvarsBairro As String
Private mvarsCEP As String
Private mvarsContato As String
Private mvarsTelContato As String
Private mvarsNomeArquivo As String
Private mvarbIntegral As Boolean
Private mvardtReg74DataInicial As Date
Private mvardtReg74DataFinal As Date

Private mvarlTotalReg50 As Long
Private mvarlTotalReg51 As Long
Private mvarlTotalReg53 As Long
Private mvarlTotalReg54 As Long
Private mvarlTotalReg60 As Long
Private mvarlTotalReg61 As Long
Private mvarlTotalReg70 As Long
Private mvarlTotalReg71 As Long
Private mvarlTotalReg74 As Long
Private mvarlTotalReg75 As Long
Private mvarlTotalReg As Long
Private mvarlNumIntReg As Long
Private mvarlNumIntGeracao As Long
Private mvarsChaveOrd As String
Private mvarsBuffer As String
Private mvarlNumIntInfoArqICMS As Long
Private mvariGeraReg51 As Integer
Private mvariGeraReg54 As Integer
Private mvarsUFDestino As String

Private mvariGeraSaidas As Integer
Private mvariGeraEntradas As Integer
Private mvarsFinalidade As String

Public Property Let sFinalidade(ByVal vData As String)
    mvarsFinalidade = vData
End Property

Public Property Get sFinalidade() As String
    sFinalidade = mvarsFinalidade
End Property

Public Property Let iGeraEntradas(ByVal vData As Integer)
    mvariGeraEntradas = vData
End Property

Public Property Get iGeraEntradas() As Integer
    iGeraEntradas = mvariGeraEntradas
End Property

Public Property Let iGeraSaidas(ByVal vData As Integer)
    mvariGeraSaidas = vData
End Property

Public Property Get iGeraSaidas() As Integer
    iGeraSaidas = mvariGeraSaidas
End Property


Public Property Let iGeraReg51(ByVal vData As Integer)
    mvariGeraReg51 = vData
End Property

Public Property Get iGeraReg51() As Integer
    iGeraReg51 = mvariGeraReg51
End Property

Public Property Let lTotalReg74(ByVal vData As Long)
    mvarlTotalReg74 = vData
End Property

Public Property Get lTotalReg74() As Long
    lTotalReg74 = mvarlTotalReg74
End Property

Public Property Let dtReg74DataFinal(ByVal vData As Date)
    mvardtReg74DataFinal = vData
End Property

Public Property Get dtReg74DataFinal() As Date
    dtReg74DataFinal = mvardtReg74DataFinal
End Property

Public Property Let dtReg74DataInicial(ByVal vData As Date)
    mvardtReg74DataInicial = vData
End Property

Public Property Get dtReg74DataInicial() As Date
    dtReg74DataInicial = mvardtReg74DataInicial
End Property

Public Property Let bIntegral(ByVal vData As Boolean)
    mvarbIntegral = vData
End Property

Public Property Get bIntegral() As Boolean
    bIntegral = mvarbIntegral
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sUFDestino(ByVal vData As String)
    mvarsUFDestino = vData
End Property

Public Property Get sUFDestino() As String
    sUFDestino = mvarsUFDestino
End Property

Public Property Let iGeraReg54(ByVal vData As Integer)
    mvariGeraReg54 = vData
End Property

Public Property Get iGeraReg54() As Integer
    iGeraReg54 = mvariGeraReg54
End Property

Public Property Let lNumIntInfoArqICMS(ByVal vData As Long)
    mvarlNumIntInfoArqICMS = vData
End Property

Public Property Get lNumIntInfoArqICMS() As Long
    lNumIntInfoArqICMS = mvarlNumIntInfoArqICMS
End Property
Public Property Let sBuffer(ByVal vData As String)
    mvarsBuffer = vData
End Property

Public Property Get sBuffer() As String
    sBuffer = mvarsBuffer
End Property

Public Property Let sChaveOrd(ByVal vData As String)
    mvarsChaveOrd = vData
End Property

Public Property Get sChaveOrd() As String
    sChaveOrd = mvarsChaveOrd
End Property

Public Property Let lNumIntGeracao(ByVal vData As Long)
    mvarlNumIntGeracao = vData
End Property

Public Property Get lNumIntGeracao() As Long
    lNumIntGeracao = mvarlNumIntGeracao
End Property

Public Property Let lNumIntReg(ByVal vData As Long)
    mvarlNumIntReg = vData
End Property

Public Property Get lNumIntReg() As Long
    lNumIntReg = mvarlNumIntReg
End Property

Public Property Let lTotalReg(ByVal vData As Long)
    mvarlTotalReg = vData
End Property

Public Property Get lTotalReg() As Long
    lTotalReg = mvarlTotalReg
End Property

Public Property Let lTotalReg61(ByVal vData As Long)
    mvarlTotalReg61 = vData
End Property

Public Property Get lTotalReg61() As Long
    lTotalReg61 = mvarlTotalReg61
End Property

Public Property Let lTotalReg60(ByVal vData As Long)
    mvarlTotalReg60 = vData
End Property

Public Property Get lTotalReg60() As Long
    lTotalReg60 = mvarlTotalReg60
End Property

Public Property Let lTotalReg75(ByVal vData As Long)
    mvarlTotalReg75 = vData
End Property

Public Property Get lTotalReg75() As Long
    lTotalReg75 = mvarlTotalReg75
End Property

Public Property Let lTotalReg70(ByVal vData As Long)
    mvarlTotalReg70 = vData
End Property

Public Property Get lTotalReg70() As Long
    lTotalReg70 = mvarlTotalReg70
End Property

Public Property Let lTotalReg71(ByVal vData As Long)
    mvarlTotalReg71 = vData
End Property

Public Property Get lTotalReg71() As Long
    lTotalReg71 = mvarlTotalReg71
End Property

Public Property Let lTotalReg54(ByVal vData As Long)
    mvarlTotalReg54 = vData
End Property

Public Property Get lTotalReg54() As Long
    lTotalReg54 = mvarlTotalReg54
End Property

Public Property Let lTotalReg53(ByVal vData As Long)
    mvarlTotalReg53 = vData
End Property

Public Property Get lTotalReg53() As Long
    lTotalReg53 = mvarlTotalReg53
End Property

Public Property Let lTotalReg51(ByVal vData As Long)
    mvarlTotalReg51 = vData
End Property

Public Property Get lTotalReg51() As Long
    lTotalReg51 = mvarlTotalReg51
End Property

Public Property Let lTotalReg50(ByVal vData As Long)
    mvarlTotalReg50 = vData
End Property

Public Property Get lTotalReg50() As Long
    lTotalReg50 = mvarlTotalReg50
End Property

Public Property Let sNomeArquivo(ByVal vData As String)
Attribute sNomeArquivo.VB_Description = "Nome do Arquivo que foi gerado"
    mvarsNomeArquivo = vData
End Property

Public Property Get sNomeArquivo() As String
    sNomeArquivo = mvarsNomeArquivo
End Property

Public Property Let sTelContato(ByVal vData As String)
Attribute sTelContato.VB_Description = "Telefone de Contato na Filial Empresa"
    mvarsTelContato = vData
End Property

Public Property Get sTelContato() As String
    sTelContato = mvarsTelContato
End Property

Public Property Let sContato(ByVal vData As String)
Attribute sContato.VB_Description = "Contato na Filial Empresa"
    mvarsContato = vData
End Property

Public Property Get sContato() As String
    sContato = mvarsContato
End Property

Public Property Let sCEP(ByVal vData As String)
Attribute sCEP.VB_Description = "CEP da Filial Empresa"
    mvarsCEP = vData
End Property

Public Property Get sCEP() As String
    sCEP = mvarsCEP
End Property

Public Property Let sBairro(ByVal vData As String)
Attribute sBairro.VB_Description = "Bairro onde se localiza a Filial Empresa"
    mvarsBairro = vData
End Property

Public Property Get sBairro() As String
    sBairro = mvarsBairro
End Property

Public Property Let sComplemento(ByVal vData As String)
Attribute sComplemento.VB_Description = "Complemento do Endereço da Filial Empresa"
    mvarsComplemento = vData
End Property

Public Property Get sComplemento() As String
    sComplemento = mvarsComplemento
End Property

Public Property Let lNumero(ByVal vData As Long)
Attribute lNumero.VB_Description = "Numero da Filial Empresa (Endereço)"
    mvarlNumero = vData
End Property

Public Property Get lNumero() As Long
    lNumero = mvarlNumero
End Property

Public Property Let sLogradouro(ByVal vData As String)
Attribute sLogradouro.VB_Description = "Endereço da Filial Empresa"
    mvarsLogradouro = vData
End Property

Public Property Get sLogradouro() As String
    sLogradouro = mvarsLogradouro
End Property

Public Property Let sUF(ByVal vData As String)
Attribute sUF.VB_Description = "UF da Filial Empresa"
    mvarsUF = vData
End Property

Public Property Get sUF() As String
    sUF = mvarsUF
End Property

Public Property Let sMunicipio(ByVal vData As String)
Attribute sMunicipio.VB_Description = "Municipio da Filial Empresa"
    mvarsMunicipio = vData
End Property

Public Property Get sMunicipio() As String
    sMunicipio = mvarsMunicipio
End Property

Public Property Let sNome(ByVal vData As String)
Attribute sNome.VB_Description = "Razão Social da Filial Empresa"
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let sInscricaoEstadual(ByVal vData As String)
Attribute sInscricaoEstadual.VB_Description = "Inscrição Estadual da Filial Empresa"
    mvarsInscricaoEstadual = vData
End Property

Public Property Get sInscricaoEstadual() As String
    sInscricaoEstadual = mvarsInscricaoEstadual
End Property

Public Property Let sCgc(ByVal vData As String)
Attribute sCgc.VB_Description = "CGC da Filial Empresa"
    mvarsCGC = vData
End Property

Public Property Get sCgc() As String
    sCgc = mvarsCGC
End Property

Public Property Let dtDataFinal(ByVal vData As Date)
Attribute dtDataFinal.VB_Description = "Data Final  utilizada para Filtro no Arquivo de ICMS"
    mvardtDataFinal = vData
End Property

Public Property Get dtDataFinal() As Date
    dtDataFinal = mvardtDataFinal
End Property

Public Property Let dtDataInicial(ByVal vData As Date)
Attribute dtDataInicial.VB_Description = "Data Inicial utilizada para Filtro no Arquivo de ICMS"
    mvardtDataInicial = vData
End Property

Public Property Get dtDataInicial() As Date
    dtDataInicial = mvardtDataInicial
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
Attribute lNumIntDoc.VB_Description = "Número Interno do Arquivo de ICMS"
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property
