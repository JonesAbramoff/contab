VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmCampos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariAtivo As Integer
Private mvarsDescricao As String
Private mvarsFormatacao As String
Private mvariImexivel As Integer
Private mvarsNome As String
Private mvarsNomeArq As String
Private mvarsNomeBd As String
Private mvariObrigatorio As Integer
Private mvariTamanho As Integer
Private mvariTipo As Integer
Private mvarsValDefault As String
Private mvarsValidacao As String
Private mvariPrecisao As Integer
Private mvariDecimais As Integer
Private mvariTamExibicao As Integer
Private mvarsTituloEntradaDados As String
Private mvarsTituloGrid As String
Private mvariSubTipo As Integer
Private mvariAlinhamento As Integer
Private mvariTestaIntegridade As Integer

Public Property Let iTestaIntegridade(ByVal vData As Integer)
    mvariTestaIntegridade = vData
End Property

Public Property Get iTestaIntegridade() As Integer
    iTestaIntegridade = mvariTestaIntegridade
End Property

Public Property Let iAlinhamento(ByVal vData As Integer)
    mvariAlinhamento = vData
End Property

Public Property Get iAlinhamento() As Integer
    iAlinhamento = mvariAlinhamento
End Property

Public Property Let iSubTipo(ByVal vData As Integer)
Attribute iSubTipo.VB_Description = "Servem para indicar como devem ser formatados os campos. Utilizado  nos browsers."
    mvariSubTipo = vData
End Property

Public Property Get iSubTipo() As Integer
    iSubTipo = mvariSubTipo
End Property

Public Property Let sTituloGrid(ByVal vData As String)
Attribute sTituloGrid.VB_Description = "Titulo do Campo utilizado nos grids"
    mvarsTituloGrid = vData
End Property

Public Property Get sTituloGrid() As String
    sTituloGrid = mvarsTituloGrid
End Property

Public Property Let sTituloEntradaDados(ByVal vData As String)
Attribute sTituloEntradaDados.VB_Description = "Titulo do Campo utilizado nas telas de entrada de dados"
    mvarsTituloEntradaDados = vData
End Property

Public Property Get sTituloEntradaDados() As String
    sTituloEntradaDados = mvarsTituloEntradaDados
End Property

Public Property Let iTamExibicao(ByVal vData As Integer)
Attribute iTamExibicao.VB_Description = "tamanho com que o campo será exibido"
    mvariTamExibicao = vData
End Property

Public Property Get iTamExibicao() As Integer
    iTamExibicao = mvariTamExibicao
End Property

Public Property Let iDecimais(ByVal vData As Integer)
Attribute iDecimais.VB_Description = "Número de Decimais"
    mvariDecimais = vData
End Property

Public Property Get iDecimais() As Integer
    iDecimais = mvariDecimais
End Property

Public Property Let iPrecisao(ByVal vData As Integer)
Attribute iPrecisao.VB_Description = "precisão do campo"
    mvariPrecisao = vData
End Property

Public Property Get iPrecisao() As Integer
    iPrecisao = mvariPrecisao
End Property

Public Property Let sValidacao(ByVal vData As String)
Attribute sValidacao.VB_Description = "Expressão de Validação do Campo"
    mvarsValidacao = vData
End Property

Public Property Get sValidacao() As String
    sValidacao = mvarsValidacao
End Property

Public Property Let sValDefault(ByVal vData As String)
Attribute sValDefault.VB_Description = "Valor Default do campo"
    mvarsValDefault = vData
End Property

Public Property Get sValDefault() As String
    sValDefault = mvarsValDefault
End Property

Public Property Let iTipo(ByVal vData As Integer)
Attribute iTipo.VB_Description = "Tipo do Campo"
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let iTamanho(ByVal vData As Integer)
Attribute iTamanho.VB_Description = "Tamanho do Campo"
    mvariTamanho = vData
End Property

Public Property Get iTamanho() As Integer
    iTamanho = mvariTamanho
End Property

Public Property Let iObrigatorio(ByVal vData As Integer)
Attribute iObrigatorio.VB_Description = "Indica se  é obrigatório o preenchimento do campo.  1=obrigatório  0=opcional"
    mvariObrigatorio = vData
End Property

Public Property Get iObrigatorio() As Integer
    iObrigatorio = mvariObrigatorio
End Property

Public Property Let sNomeBd(ByVal vData As String)
Attribute sNomeBd.VB_Description = "Nome do BD"
    mvarsNomeBd = vData
End Property

Public Property Get sNomeBd() As String
    sNomeBd = mvarsNomeBd
End Property

Public Property Let sNomeArq(ByVal vData As String)
Attribute sNomeArq.VB_Description = "Nome da Tabela"
    mvarsNomeArq = vData
End Property

Public Property Get sNomeArq() As String
    sNomeArq = mvarsNomeArq
End Property

Public Property Let sNome(ByVal vData As String)
Attribute sNome.VB_Description = "Nome do Campo"
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let iImexivel(ByVal vData As Integer)
Attribute iImexivel.VB_Description = "Indica se o campo pode ser alterado. 1=não pode ser alterado   0=pode ser alterado"
    mvariImexivel = vData
End Property

Public Property Get iImexivel() As Integer
    iImexivel = mvariImexivel
End Property

Public Property Let sFormatacao(ByVal vData As String)
Attribute sFormatacao.VB_Description = "Formatacao do Campo"
    mvarsFormatacao = vData
End Property

Public Property Get sFormatacao() As String
    sFormatacao = mvarsFormatacao
End Property

Public Property Let sDescricao(ByVal vData As String)
Attribute sDescricao.VB_Description = "Descrição do Campo"
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iAtivo(ByVal vData As Integer)
Attribute iAtivo.VB_Description = "Indica se o campo está ativo. 1=ativo  0=inativo"
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property

Private Sub Class_Initialize()

''    mvarsDescricao = String(STRING_DESCRICAO_CAMPO, 0)
''    mvarsFormatacao = String(STRING_FORMATACAO_CAMPO, 0)
''    mvarsNome = String(STRING_NOME_CAMPO, 0)
''    mvarsNomeArq = String(STRING_NOME_TABELA, 0)
''    mvarsNomeBd = String(STRING_NOME_BD, 0)
''    mvarsValDefault = String(STRING_VALOR_DEFAULT_CAMPO, 0)
''    mvarsValidacao = String(STRING_VALIDACAO_CAMPO, 0)
''    mvarsTituloEntradaDados = String(STRING_TITULO_ENTRADA_DADOS_CAMPO, 0)
''    mvarsTituloGrid = String(STRING_TITULO_GRID_CAMPO, 0)

End Sub
