VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGeracaoArqCNABPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarobjArq As ClassArqCNABPag
Private mvarobjCCI As New ClassContasCorrentesInternas
Private mvarsNomeArquivo As String
Private mvarlSequencialArq As Long
Private mvarlSequencialRegLote As Long
Private mvarlSequencialRegDetalhe As Long
Private mvarcolBaixasParcPag As New Collection
Private mvariBancoFavorecido As Integer
Private mvarsNomeFavorecido As String
Private mvarsAgenciaFavorecido As String
Private mvarlEnderecoFavorecido As Long
Private mvarsCGCFavorecido As String
Private mvardtDataVenctoParcela As Date
Private mvardValorPagto As Double
Private mvarobjBorderoPagto As New ClassBorderoPagto
Private mvarsContaFavorecido As String
Private mvarlQuantRegistrosLote As Long
Private mvardValorTotalLote As Double
Private mvarlQuantRegistrosArq As Long
Private mvarsCodEmpresaBanco As String
Private mvarsSiglaTitulo As String
Private mvarlNumTitulo As Long
Private mvardtDataEmissaoTitulo As Date
Private mvarsNossoNumero As String
Private mvarlQuantTitulos As Long
Private mvarobjTelaAtualizacao As Object 'local copy
Private mvarsDiretorio As String 'local copy
Private mvarobjBancoInfo As New ClassBancoInfo 'local copy

Private mvariTipoCobranca As Integer
Private mvariBancoCobrador As Integer
Private mvarsCodigoDeBarras As String 'Guarda o Numero Referênte ao Código da Barras

Public Property Let sCodigoDeBarras(ByVal vData As String)
    mvarsCodigoDeBarras = vData
End Property

Public Property Get sCodigoDeBarras() As String
    sCodigoDeBarras = mvarsCodigoDeBarras
End Property

Public Property Let iBancoCobrador(ByVal vData As Integer)
    mvariBancoCobrador = vData
End Property

Public Property Get iBancoCobrador() As Integer
    iBancoCobrador = mvariBancoCobrador
End Property

Public Property Let iTipoCobranca(ByVal vData As Integer)
    mvariTipoCobranca = vData
End Property

Public Property Get iTipoCobranca() As Integer
    iTipoCobranca = mvariTipoCobranca
End Property

Public Property Set objBancoInfo(ByVal vData As ClassBancoInfo)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.objBancoInfo = Form1
    Set mvarobjBancoInfo = vData
End Property

Public Property Get objBancoInfo() As ClassBancoInfo
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.objBancoInfo
    Set objBancoInfo = mvarobjBancoInfo
End Property

Public Property Let sDiretorio(ByVal vData As String)
    mvarsDiretorio = vData
End Property

Public Property Get sDiretorio() As String
    sDiretorio = mvarsDiretorio
End Property

Public Property Set objTelaAtualizacao(ByVal vData As Object)
    Set mvarobjTelaAtualizacao = vData
End Property

Public Property Get objTelaAtualizacao() As Object
    Set objTelaAtualizacao = mvarobjTelaAtualizacao
End Property

Public Property Let lQuantTitulos(ByVal vData As Long)
    mvarlQuantTitulos = vData
End Property

Public Property Get lQuantTitulos() As Long
    lQuantTitulos = mvarlQuantTitulos
End Property

Public Property Let sNossoNumero(ByVal vData As String)
    mvarsNossoNumero = vData
End Property

Public Property Get sNossoNumero() As String
    sNossoNumero = mvarsNossoNumero
End Property

Public Property Let dtDataEmissaoTitulo(ByVal vData As Date)
    mvardtDataEmissaoTitulo = vData
End Property

Public Property Get dtDataEmissaoTitulo() As Date
    dtDataEmissaoTitulo = mvardtDataEmissaoTitulo
End Property

Public Property Let lNumTitulo(ByVal vData As Long)
    mvarlNumTitulo = vData
End Property

Public Property Get lNumTitulo() As Long
    lNumTitulo = mvarlNumTitulo
End Property



Public Property Let sSiglaTitulo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSiglaTitulo = 5
    mvarsSiglaTitulo = vData
End Property


Public Property Get sSiglaTitulo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSiglaTitulo
    sSiglaTitulo = mvarsSiglaTitulo
End Property



Public Property Let sCodEmpresaBanco(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCodEmpresaBanco = 5
    mvarsCodEmpresaBanco = vData
End Property


Public Property Get sCodEmpresaBanco() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCodEmpresaBanco
    sCodEmpresaBanco = mvarsCodEmpresaBanco
End Property



Public Property Let lQuantRegistrosArq(ByVal vData As Long)
    mvarlQuantRegistrosArq = vData
End Property

Public Property Get lQuantRegistrosArq() As Long
    lQuantRegistrosArq = mvarlQuantRegistrosArq
End Property

Public Property Let dValorTotalLote(ByVal vData As Double)
    mvardValorTotalLote = vData
End Property

Public Property Get dValorTotalLote() As Double
    dValorTotalLote = mvardValorTotalLote
End Property

Public Property Let lQuantRegistrosLote(ByVal vData As Long)
    mvarlQuantRegistrosLote = vData
End Property

Public Property Get lQuantRegistrosLote() As Long
    lQuantRegistrosLote = mvarlQuantRegistrosLote
End Property

Public Property Let sContaFavorecido(ByVal vData As String)
    mvarsContaFavorecido = vData
End Property

Public Property Get sContaFavorecido() As String
    sContaFavorecido = mvarsContaFavorecido
End Property

Public Property Set objBorderoPagto(ByVal vData As ClassBorderoPagto)
    Set mvarobjBorderoPagto = vData
End Property

Public Property Get objBorderoPagto() As ClassBorderoPagto
    Set objBorderoPagto = mvarobjBorderoPagto
End Property

Public Property Let dValorPagto(ByVal vData As Double)
    mvardValorPagto = vData
End Property

Public Property Get dValorPagto() As Double
    dValorPagto = mvardValorPagto
End Property

Public Property Let dtDataVenctoParcela(ByVal vData As Date)
Attribute dtDataVenctoParcela.VB_Description = "Data de Vencimento da Parcela Baixada que está sendo tratada"
    mvardtDataVenctoParcela = vData
End Property

Public Property Get dtDataVenctoParcela() As Date
    dtDataVenctoParcela = mvardtDataVenctoParcela
End Property

Public Property Let sCGCFavorecido(ByVal vData As String)
    mvarsCGCFavorecido = vData
End Property

Public Property Get sCGCFavorecido() As String
    sCGCFavorecido = mvarsCGCFavorecido
End Property

Public Property Let lEnderecoFavorecido(ByVal vData As Long)
    mvarlEnderecoFavorecido = vData
End Property

Public Property Get lEnderecoFavorecido() As Long
    lEnderecoFavorecido = mvarlEnderecoFavorecido
End Property

Public Property Let sAgenciaFavorecido(ByVal vData As String)
    mvarsAgenciaFavorecido = vData
End Property

Public Property Get sAgenciaFavorecido() As String
    sAgenciaFavorecido = mvarsAgenciaFavorecido
End Property

Public Property Let sNomeFavorecido(ByVal vData As String)
Attribute sNomeFavorecido.VB_Description = "Nome do Fornecedor"
    mvarsNomeFavorecido = vData
End Property

Public Property Get sNomeFavorecido() As String
    sNomeFavorecido = mvarsNomeFavorecido
End Property

Public Property Let iBancoFavorecido(ByVal vData As Integer)
    mvariBancoFavorecido = vData
End Property

Public Property Get iBancoFavorecido() As Integer
    iBancoFavorecido = mvariBancoFavorecido
End Property

Public Property Set colBaixasParcPag(ByVal vData As Collection)
Attribute colBaixasParcPag.VB_Description = "Coleção que guarda as Parcelas Baixadas pelo Bordero que esta no objBordero."
    Set mvarcolBaixasParcPag = vData
End Property

Public Property Get colBaixasParcPag() As Collection
    Set colBaixasParcPag = mvarcolBaixasParcPag
End Property

Public Property Let lSequencialRegDetalhe(ByVal vData As Long)
    mvarlSequencialRegDetalhe = vData
End Property

Public Property Get lSequencialRegDetalhe() As Long
    lSequencialRegDetalhe = mvarlSequencialRegDetalhe
End Property

Public Property Let lSequencialRegLote(ByVal vData As Long)
    mvarlSequencialRegLote = vData
End Property

Public Property Get lSequencialRegLote() As Long
    lSequencialRegLote = mvarlSequencialRegLote
End Property

Public Property Let lSequencialArq(ByVal vData As Long)
    mvarlSequencialArq = vData
End Property

Public Property Get lSequencialArq() As Long
    lSequencialArq = mvarlSequencialArq
End Property

Public Property Let sNomeArquivo(ByVal vData As String)
Attribute sNomeArquivo.VB_Description = "Nome do arquivo a ser gerado"
    mvarsNomeArquivo = vData
End Property

Public Property Get sNomeArquivo() As String
    sNomeArquivo = mvarsNomeArquivo
End Property

Public Property Set objCCI(ByVal vData As ClassContasCorrentesInternas)
    Set mvarobjCCI = vData
End Property

Public Property Get objCCI() As ClassContasCorrentesInternas
    Set objCCI = mvarobjCCI
End Property

Public Property Set objArq(ByVal vData As ClassArqCNABPag)
    Set mvarobjArq = vData
End Property

Public Property Get objArq() As ClassArqCNABPag
    Set objArq = mvarobjArq
End Property

