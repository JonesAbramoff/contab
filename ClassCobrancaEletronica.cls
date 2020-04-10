VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCobrancaEletronica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvardtDataEmissaoArq As Date
Private mvariCobrador As Integer
Private mvariNumBorderoIni As Integer
Private mvariNumBorderoFim As Integer
Private mvarcolOcorrRemParcRec As New Collection
Private mvarlCNABProxSeqArqCobr As Long
Private mvarsCodEmpresaBanco As String
Private mvarlSequencialRegistro As Long
Private mvarobjCobrador As New ClassCobrador
Private mvarobjTelaAtualizacao As Object
Private mvarcolBorderos As New Collection
Private mvarsCedente As String
Private mvarsCGCCedente As String
Private mvarsCGCSacadorAvalista As String
Private mvarsSacadorAvalista As String
Private mvarlNumTitulo As Long
Private mvariEspecieTitulo As Integer
Private mvardtDataEmissaoTitulo As Date
Private mvarsCGCSacado As String
Private mvarsNomeSacado As String
Private mvarlEnderecoSacado As Long
Private mvarlEnderecoCobrancaSacado As Long
Private mvarlCliente As Long
Private mvariFilialCli As Integer
Private mvarobjCCI As New ClassContasCorrentesInternas
Private mvarlSeqNomeArq As Long
Private mvarcolCarteirasCNABInfo As New Collection
Private mvarobjCobradorCNABInfo As New ClassCNABInfo
Private mvarlQuantTitulos As Long
Private mvarsDiretorio As String
'local variable(s) to hold property value(s)
Private mvarsNomeArquivoRetorno As String 'local copy
Private mvardPercMulta As Double
Private mvardPercJurosDiario As Double
Private mvariFilialEmpTitulo As Integer
Private mvariRegerarArquivo As Integer
Private mvariCarteira As Integer
Private mvarsVersaoLayout As String

Public Property Let sVersaoLayout(ByVal vData As String)
    mvarsVersaoLayout = vData
End Property

Public Property Get sVersaoLayout() As String
    sVersaoLayout = mvarsVersaoLayout
End Property

Public Property Let iCarteira(ByVal vData As Integer)
    mvariCarteira = vData
End Property

Public Property Get iCarteira() As Integer
    iCarteira = mvariCarteira
End Property

Public Property Let iRegerarArquivo(ByVal vData As Integer)
    mvariRegerarArquivo = vData
End Property

Public Property Get iRegerarArquivo() As Integer
    iRegerarArquivo = mvariRegerarArquivo
End Property

Public Property Let iFilialEmpTitulo(ByVal vData As Integer)
    mvariFilialEmpTitulo = vData
End Property

Public Property Get iFilialEmpTitulo() As Integer
    iFilialEmpTitulo = mvariFilialEmpTitulo
End Property

Public Property Let dPercJurosDiario(ByVal vData As Double)
    mvardPercJurosDiario = vData
End Property

Public Property Get dPercJurosDiario() As Double
    dPercJurosDiario = mvardPercJurosDiario
End Property

Public Property Let iNumBorderoFim(ByVal vData As Integer)
    mvariNumBorderoFim = vData
End Property

Public Property Get iNumBorderoFim() As Integer
    iNumBorderoFim = mvariNumBorderoFim
End Property

Public Property Let iNumBorderoIni(ByVal vData As Integer)
    mvariNumBorderoIni = vData
End Property

Public Property Get iNumBorderoIni() As Integer
    iNumBorderoIni = mvariNumBorderoIni
End Property

Public Property Let dPercMulta(ByVal vData As Double)
    mvardPercMulta = vData
End Property

Public Property Get dPercMulta() As Double
    dPercMulta = mvardPercMulta
End Property

Public Property Let sNomeArquivoRetorno(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNomeArquivoRetorno = 5
    mvarsNomeArquivoRetorno = vData
End Property


Public Property Get sNomeArquivoRetorno() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNomeArquivoRetorno
    sNomeArquivoRetorno = mvarsNomeArquivoRetorno
End Property




Public Property Let sDiretorio(ByVal vData As String)
    mvarsDiretorio = vData
End Property

Public Property Get sDiretorio() As String
    sDiretorio = mvarsDiretorio
End Property

Public Property Let lQuantTitulos(ByVal vData As Long)
    mvarlQuantTitulos = vData
End Property

Public Property Get lQuantTitulos() As Long
    lQuantTitulos = mvarlQuantTitulos
End Property

Public Property Set objCobradorCNABInfo(ByVal vData As ClassCNABInfo)
    Set mvarobjCobradorCNABInfo = vData
End Property

Public Property Get objCobradorCNABInfo() As ClassCNABInfo
    Set objCobradorCNABInfo = mvarobjCobradorCNABInfo
End Property

Public Property Set colCarteirasCNABInfo(ByVal vData As Collection)
    Set mvarcolCarteirasCNABInfo = vData
End Property

Public Property Get colCarteirasCNABInfo() As Collection
    Set colCarteirasCNABInfo = mvarcolCarteirasCNABInfo
End Property

Public Property Let lSeqNomeArq(ByVal vData As Long)
    mvarlSeqNomeArq = vData
End Property

Public Property Get lSeqNomeArq() As Long
    lSeqNomeArq = mvarlSeqNomeArq
End Property

Public Property Set objCCI(ByVal vData As ClassContasCorrentesInternas)
    Set mvarobjCCI = vData
End Property

Public Property Get objCCI() As ClassContasCorrentesInternas
    Set objCCI = mvarobjCCI
End Property

Public Property Let iFilialCli(ByVal vData As Integer)
    mvariFilialCli = vData
End Property

Public Property Get iFilialCli() As Integer
    iFilialCli = mvariFilialCli
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let lEnderecoCobrancaSacado(ByVal vData As Long)
    mvarlEnderecoCobrancaSacado = vData
End Property

Public Property Get lEnderecoCobrancaSacado() As Long
    lEnderecoCobrancaSacado = mvarlEnderecoCobrancaSacado
End Property

Public Property Let lEnderecoSacado(ByVal vData As Long)
    mvarlEnderecoSacado = vData
End Property

Public Property Get lEnderecoSacado() As Long
    lEnderecoSacado = mvarlEnderecoSacado
End Property

Public Property Let sNomeSacado(ByVal vData As String)
    mvarsNomeSacado = vData
End Property

Public Property Get sNomeSacado() As String
    sNomeSacado = mvarsNomeSacado
End Property

Public Property Let sCGCSacado(ByVal vData As String)
Attribute sCGCSacado.VB_Description = "O CGC ou CPF da Filial do Cliente"
    mvarsCGCSacado = vData
End Property

Public Property Get sCGCSacado() As String
    sCGCSacado = mvarsCGCSacado
End Property

Public Property Let dtDataEmissaoTitulo(ByVal vData As Date)
Attribute dtDataEmissaoTitulo.VB_Description = "Data de Emissao do Título"
    mvardtDataEmissaoTitulo = vData
End Property

Public Property Get dtDataEmissaoTitulo() As Date
    dtDataEmissaoTitulo = mvardtDataEmissaoTitulo
End Property

Public Property Let iEspecieTitulo(ByVal vData As Integer)
Attribute iEspecieTitulo.VB_Description = "Espécie do Título a Receber"
    mvariEspecieTitulo = vData
End Property

Public Property Get iEspecieTitulo() As Integer
    iEspecieTitulo = mvariEspecieTitulo
End Property

Public Property Let lNumTitulo(ByVal vData As Long)
Attribute lNumTitulo.VB_Description = "Número do Título a receber"
    mvarlNumTitulo = vData
End Property

Public Property Get lNumTitulo() As Long
    lNumTitulo = mvarlNumTitulo
End Property

Public Property Let sCGCCedente(ByVal vData As String)
Attribute sCGCCedente.VB_Description = "No nosso caso o CGC da Filial Empresa da parcela a receber"
    mvarsCGCCedente = vData
End Property

Public Property Get sCGCCedente() As String
    sCGCCedente = mvarsCGCCedente
End Property

Public Property Let sCedente(ByVal vData As String)
    mvarsCedente = vData
End Property

Public Property Get sCedente() As String
    sCedente = mvarsCedente
End Property

Public Property Let sCGCSacadorAvalista(ByVal vData As String)
    mvarsCGCSacadorAvalista = vData
End Property

Public Property Get sCGCSacadorAvalista() As String
    sCGCSacadorAvalista = mvarsCGCSacadorAvalista
End Property

Public Property Let sSacadorAvalista(ByVal vData As String)
    mvarsSacadorAvalista = vData
End Property

Public Property Get sSacadorAvalista() As String
    sSacadorAvalista = mvarsSacadorAvalista
End Property

Public Property Set colBorderos(ByVal vData As Collection)
    Set mvarcolBorderos = vData
End Property

Public Property Get colBorderos() As Collection
    Set colBorderos = mvarcolBorderos
End Property

Public Property Set objTelaAtualizacao(ByVal vData As Object)
    Set mvarobjTelaAtualizacao = vData
End Property

Public Property Get objTelaAtualizacao() As Object
    Set objTelaAtualizacao = mvarobjTelaAtualizacao
End Property

Public Property Set objCobrador(ByVal vData As ClassCobrador)
    Set mvarobjCobrador = vData
End Property
Public Property Get objCobrador() As ClassCobrador
    Set objCobrador = mvarobjCobrador
End Property

Public Property Let lSequencialRegistro(ByVal vData As Long)
    mvarlSequencialRegistro = vData
End Property
Public Property Get lSequencialRegistro() As Long
    lSequencialRegistro = mvarlSequencialRegistro
End Property
Public Property Let sCodEmpresaBanco(ByVal vData As String)
    mvarsCodEmpresaBanco = vData
End Property

Public Property Get sCodEmpresaBanco() As String
    sCodEmpresaBanco = mvarsCodEmpresaBanco
End Property

Public Property Let lCNABProxSeqArqCobr(ByVal vData As Long)
    mvarlCNABProxSeqArqCobr = vData
End Property

Public Property Get lCNABProxSeqArqCobr() As Long
    lCNABProxSeqArqCobr = mvarlCNABProxSeqArqCobr
End Property

Public Property Set colOcorrRemParcRec(ByVal vData As Collection)
    Set mvarcolOcorrRemParcRec = vData
End Property

Public Property Get colOcorrRemParcRec() As Collection
    Set colOcorrRemParcRec = mvarcolOcorrRemParcRec
End Property

Public Property Let iCobrador(ByVal vData As Integer)
    mvariCobrador = vData
End Property

Public Property Get iCobrador() As Integer
    iCobrador = mvariCobrador
End Property

Public Property Let dtDataEmissaoArq(ByVal vData As Date)
    mvardtDataEmissaoArq = vData
End Property

Public Property Get dtDataEmissaoArq() As Date
    dtDataEmissaoArq = mvardtDataEmissaoArq
End Property

