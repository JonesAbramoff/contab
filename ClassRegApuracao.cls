VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRegApuracao"
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
Private mvarlNumIntDocLivFechado As Long
Private mvardSaldoCredorInicial As Double
Private mvardtDataEntregaGIA As Date
Private mvarsLocalEntregaGIA As String
Private mvarsObservacoes As String
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
Private mvarlNumeroLivro As Long
Private mvardtDataImpressao As Date
Private mvariFolhaInicial As Integer
Private mvardSaldoCredorFinal As Double 'local copy

Public Property Let dSaldoCredorFinal(ByVal vData As Double)
Attribute dSaldoCredorFinal.VB_Description = "Saldo Credor a transportar para Periodo Seguinte"
    mvardSaldoCredorFinal = vData
End Property

Public Property Get dSaldoCredorFinal() As Double
    dSaldoCredorFinal = mvardSaldoCredorFinal
End Property

Public Property Let iFolhaInicial(ByVal vData As Integer)
Attribute iFolhaInicial.VB_Description = "Apartir de que nº  folha a Apuração vai começar"
    mvariFolhaInicial = vData
End Property

Public Property Get iFolhaInicial() As Integer
    iFolhaInicial = mvariFolhaInicial
End Property

Public Property Let dtDataImpressao(ByVal vData As Date)
Attribute dtDataImpressao.VB_Description = "data em que o livro foi impresso definitivamente (ou enviado, se for arquivo)"
    mvardtDataImpressao = vData
End Property

Public Property Get dtDataImpressao() As Date
    dtDataImpressao = mvardtDataImpressao
End Property

Public Property Let lNumeroLivro(ByVal vData As Long)
Attribute lNumeroLivro.VB_Description = "Número do Livro"
    mvarlNumeroLivro = vData
End Property

Public Property Get lNumeroLivro() As Long
    lNumeroLivro = mvarlNumeroLivro
End Property

Public Property Let sTelContato(ByVal vData As String)
Attribute sTelContato.VB_Description = "Telefone de Contato"
    mvarsTelContato = vData
End Property

Public Property Get sTelContato() As String
    sTelContato = mvarsTelContato
End Property

Public Property Let sContato(ByVal vData As String)
Attribute sContato.VB_Description = "Contato"
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
Attribute sBairro.VB_Description = "Bairro da Filial Empresa"
    mvarsBairro = vData
End Property

Public Property Get sBairro() As String
    sBairro = mvarsBairro
End Property

Public Property Let sComplemento(ByVal vData As String)
Attribute sComplemento.VB_Description = "Complemento do Endereco da Filial Empresa"
    mvarsComplemento = vData
End Property

Public Property Get sComplemento() As String
    sComplemento = mvarsComplemento
End Property

Public Property Let lNumero(ByVal vData As Long)
Attribute lNumero.VB_Description = "Número da Filial Empresa"
    mvarlNumero = vData
End Property

Public Property Get lNumero() As Long
    lNumero = mvarlNumero
End Property

Public Property Let sLogradouro(ByVal vData As String)
Attribute sLogradouro.VB_Description = "Endereco da Filial Empresa"
    mvarsLogradouro = vData
End Property

Public Property Get sLogradouro() As String
    sLogradouro = mvarsLogradouro
End Property

Public Property Let sUF(ByVal vData As String)
Attribute sUF.VB_Description = "Estado da Filial Empresa"
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
Attribute sNome.VB_Description = "Nome da Filial Empresa"
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let sInscricaoEstadual(ByVal vData As String)
Attribute sInscricaoEstadual.VB_Description = "Inscricao Estadual da Filial Empresa"
    mvarsInscricaoEstadual = vData
End Property

Public Property Get sInscricaoEstadual() As String
    sInscricaoEstadual = mvarsInscricaoEstadual
End Property

Public Property Let sCGC(ByVal vData As String)
Attribute sCGC.VB_Description = "CGC da Filial Empresa"
    mvarsCGC = vData
End Property

Public Property Get sCGC() As String
    sCGC = mvarsCGC
End Property

Public Property Let sObservacoes(ByVal vData As String)
Attribute sObservacoes.VB_Description = "Observações da  Apuração"
    mvarsObservacoes = vData
End Property

Public Property Get sObservacoes() As String
    sObservacoes = mvarsObservacoes
End Property

Public Property Let sLocalEntregaGIA(ByVal vData As String)
Attribute sLocalEntregaGIA.VB_Description = "banco ou reparticao"
    mvarsLocalEntregaGIA = vData
End Property

Public Property Get sLocalEntregaGIA() As String
    sLocalEntregaGIA = mvarsLocalEntregaGIA
End Property

Public Property Let dtDataEntregaGIA(ByVal vData As Date)
Attribute dtDataEntregaGIA.VB_Description = "Data de Entrega da Guia"
    mvardtDataEntregaGIA = vData
End Property

Public Property Get dtDataEntregaGIA() As Date
    dtDataEntregaGIA = mvardtDataEntregaGIA
End Property

Public Property Let dSaldoCredorInicial(ByVal vData As Double)
Attribute dSaldoCredorInicial.VB_Description = "Saldo credor transportado do Periodo Anterior"
    mvardSaldoCredorInicial = vData
End Property

Public Property Get dSaldoCredorInicial() As Double
    dSaldoCredorInicial = mvardSaldoCredorInicial
End Property

Public Property Let lNumIntDocLivFechado(ByVal vData As Long)
Attribute lNumIntDocLivFechado.VB_Description = "Número Interno do Livro Fechado"
    mvarlNumIntDocLivFechado = vData
End Property

Public Property Get lNumIntDocLivFechado() As Long
    lNumIntDocLivFechado = mvarlNumIntDocLivFechado
End Property

Public Property Let dtDataFinal(ByVal vData As Date)
Attribute dtDataFinal.VB_Description = "Junto com a Filial e DataInicial é chave"
    mvardtDataFinal = vData
End Property

Public Property Get dtDataFinal() As Date
    dtDataFinal = mvardtDataFinal
End Property

Public Property Let dtDataInicial(ByVal vData As Date)
Attribute dtDataInicial.VB_Description = "Junto com a Filial e DataFinal é chave"
    mvardtDataInicial = vData
End Property

Public Property Get dtDataInicial() As Date
    dtDataInicial = mvardtDataInicial
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
Attribute iFilialEmpresa.VB_Description = "Filial Empresa da apuração"
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
Attribute lNumIntDoc.VB_Description = "Número Interno"
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property
