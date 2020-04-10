VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLivRegES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDocOrigem As Long
Private mvarlNumIntLivroFechado As Long
Private mvarlNumIntNF As Long
Private mvariStatusLivro As Integer
Private mvariTipo As Integer
Private mvarsCGC As String
Private mvarsInscricaoEstadual As String
Private mvarsNome As String
Private mvardtData As Date
Private mvarsUF As String
Private mvarsSerie As String
Private mvarlNumNotaFiscal As Long
Private mvariSituacao As Integer
Private mvariEmitente As Integer
Private mvariDestinatario As Integer
Private mvariOrigem As Integer
Private mvarlNumIntEmitente As Long
Private mvarlNumIntArqICMS As Long
Private mvarlNumIntRegApuracaoICMS As Long
Private mvarlNumIntRegApuracaoIPI As Long
Private mvariModelo As Integer
Private mvariTipoNumIntDocOrigem As Integer
Private mvarcolLinhas As Collection
Private mvarcolItensNF As Collection
Private mvarlNumIntDoc As Long
Private mvariFilialEmpresa As Integer
'local variable(s) to hold property value(s)
Private mvardtDataEmissao As Date 'local copy
'local variable(s) to hold property value(s)
Private mvariDestinoOrigem As Integer 'local copy
Private mvariFreteRespons As Integer
Private mvardPISValor As Double
Private mvardCOFINSValor As Double
Private mvariIEIsento As Integer

Public Property Let iIEIsento(ByVal vData As Integer)
    mvariIEIsento = vData
End Property

Public Property Get iIEIsento() As Integer
    iIEIsento = mvariIEIsento
End Property

Public Property Let dCOFINSValor(ByVal vData As Double)
    mvardCOFINSValor = vData
End Property

Public Property Get dCOFINSValor() As Double
    dCOFINSValor = mvardCOFINSValor
End Property

Public Property Let dPISValor(ByVal vData As Double)
    mvardPISValor = vData
End Property

Public Property Get dPISValor() As Double
    dPISValor = mvardPISValor
End Property

Public Property Let iFreteRespons(ByVal vData As Integer)
    mvariFreteRespons = vData
End Property

Public Property Get iFreteRespons() As Integer
    iFreteRespons = mvariFreteRespons
End Property

Public Property Let iDestinoOrigem(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iDestinoOrigem = 5
    mvariDestinoOrigem = vData
End Property


Public Property Get iDestinoOrigem() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iDestinoOrigem
    iDestinoOrigem = mvariDestinoOrigem
End Property



Public Property Let dtDataEmissao(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataEmissao = 5
    mvardtDataEmissao = vData
End Property


Public Property Get dtDataEmissao() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataEmissao
    dtDataEmissao = mvardtDataEmissao
End Property




Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Set colItensNF(ByVal vData As Collection)
Attribute colItensNF.VB_Description = "Coleção que flete a Tabela de Item NF que Vai para o Arquivo de ICMS"
    Set mvarcolItensNF = vData
End Property

Public Property Get colItensNF() As Collection
    Set colItensNF = mvarcolItensNF
End Property

Public Property Set colLinhas(ByVal vData As Collection)
    Set mvarcolLinhas = vData
End Property

Public Property Get colLinhas() As Collection
    Set colLinhas = mvarcolLinhas
End Property

Public Property Let iTipoNumIntDocOrigem(ByVal vData As Integer)
Attribute iTipoNumIntDocOrigem.VB_Description = "0:tabela nfiscal, 1:titulospag(ou titulospagbaixados), 2:nfspag(ou nfspagbaixadas)"
    mvariTipoNumIntDocOrigem = vData
End Property

Public Property Get iTipoNumIntDocOrigem() As Integer
    iTipoNumIntDocOrigem = mvariTipoNumIntDocOrigem
End Property

Public Property Let iModelo(ByVal vData As Integer)
Attribute iModelo.VB_Description = "Conforme a Tabela de Documentos Fiscais Pag 18 item 3.3"
    mvariModelo = vData
End Property

Public Property Get iModelo() As Integer
    iModelo = mvariModelo
End Property

Public Property Let lNumIntRegApuracaoIPI(ByVal vData As Long)
    mvarlNumIntRegApuracaoIPI = vData
End Property

Public Property Get lNumIntRegApuracaoIPI() As Long
    lNumIntRegApuracaoIPI = mvarlNumIntRegApuracaoIPI
End Property

Public Property Let lNumIntRegApuracaoICMS(ByVal vData As Long)
    mvarlNumIntRegApuracaoICMS = vData
End Property

Public Property Get lNumIntRegApuracaoICMS() As Long
    lNumIntRegApuracaoICMS = mvarlNumIntRegApuracaoICMS
End Property

Public Property Let lNumIntArqICMS(ByVal vData As Long)
Attribute lNumIntArqICMS.VB_Description = "identifica o ""arquivo"" gerado com info s/ICMS cnvenio 57/95"
    mvarlNumIntArqICMS = vData
End Property

Public Property Get lNumIntArqICMS() As Long
    lNumIntArqICMS = mvarlNumIntArqICMS
End Property

Public Property Let lNumIntEmitente(ByVal vData As Long)
Attribute lNumIntEmitente.VB_Description = "Código do Emitente"
    mvarlNumIntEmitente = vData
End Property

Public Property Get lNumIntEmitente() As Long
    lNumIntEmitente = mvarlNumIntEmitente
End Property

Public Property Let iOrigem(ByVal vData As Integer)
Attribute iOrigem.VB_Description = "0:Empresa:, 1:Cliente, 2:Fornecedor (Quem motivou a nota - útil para Nota Fiscal Interna de Entrada)"
    mvariOrigem = vData
End Property

Public Property Get iOrigem() As Integer
    iOrigem = mvariOrigem
End Property

Public Property Let iDestinatario(ByVal vData As Integer)
Attribute iDestinatario.VB_Description = "0:Empresa, 1:Cliente, 2:Fornecedor"
    mvariDestinatario = vData
End Property

Public Property Get iDestinatario() As Integer
    iDestinatario = mvariDestinatario
End Property

Public Property Let iEmitente(ByVal vData As Integer)
Attribute iEmitente.VB_Description = "0:Empresa, 1:Cliente, 2:Fornecedor"
    mvariEmitente = vData
End Property

Public Property Get iEmitente() As Integer
    iEmitente = mvariEmitente
End Property

Public Property Let iSituacao(ByVal vData As Integer)
Attribute iSituacao.VB_Description = "0:normal, 1:cancelada"
    mvariSituacao = vData
End Property

Public Property Get iSituacao() As Integer
    iSituacao = mvariSituacao
End Property

Public Property Let lNumNotaFiscal(ByVal vData As Long)
Attribute lNumNotaFiscal.VB_Description = "Número da Nota Fiscal"
    mvarlNumNotaFiscal = vData
End Property

Public Property Get lNumNotaFiscal() As Long
    lNumNotaFiscal = mvarlNumNotaFiscal
End Property

Public Property Let sSerie(ByVal vData As String)
Attribute sSerie.VB_Description = "Serie da Nota Fiscal"
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let sUF(ByVal vData As String)
Attribute sUF.VB_Description = "do remetente nas entradas e do destinatario nas saidas"
    mvarsUF = vData
End Property

Public Property Get sUF() As String
    sUF = mvarsUF
End Property

Public Property Let dtData(ByVal vData As Date)
Attribute dtData.VB_Description = "Data de entrada nas entradas ou Data de Emissao nas saidas"
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let sNome(ByVal vData As String)
Attribute sNome.VB_Description = "razao social do emitente / destinatario"
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let sInscricaoEstadual(ByVal vData As String)
Attribute sInscricaoEstadual.VB_Description = "Inscrição Estadual do remetente nas entradas e do destinatario nas saidas"
    mvarsInscricaoEstadual = vData
End Property

Public Property Get sInscricaoEstadual() As String
    sInscricaoEstadual = mvarsInscricaoEstadual
End Property

Public Property Let sCgc(ByVal vData As String)
Attribute sCgc.VB_Description = "CGC do remetente nas entradas e do destinatario nas saidas"
    mvarsCGC = vData
End Property

Public Property Get sCgc() As String
    sCgc = mvarsCGC
End Property

Public Property Let iTipo(ByVal vData As Integer)
Attribute iTipo.VB_Description = "0:entrada, 1:saida"
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let iStatusLivro(ByVal vData As Integer)
Attribute iStatusLivro.VB_Description = "0:original, 1:alterada, 2:original excluida, 3:alterada excluida"
    mvariStatusLivro = vData
End Property

Public Property Get iStatusLivro() As Integer
    iStatusLivro = mvariStatusLivro
End Property

Public Property Let lNumIntNF(ByVal vData As Long)
    mvarlNumIntNF = vData
End Property

Public Property Get lNumIntNF() As Long
    lNumIntNF = mvarlNumIntNF
End Property

Public Property Let lNumIntDocLivFechado(ByVal vData As Long)
    mvarlNumIntLivroFechado = vData
End Property

Public Property Get lNumIntDocLivFechado() As Long
    lNumIntDocLivFechado = mvarlNumIntLivroFechado
End Property

Public Property Let lNumIntDocOrigem(ByVal vData As Long)
    mvarlNumIntDocOrigem = vData
End Property

Public Property Get lNumIntDocOrigem() As Long
    lNumIntDocOrigem = mvarlNumIntDocOrigem
End Property
