VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLancamento_Detalhe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public sOrigem As String
Public iExercicio As Integer
Public iPeriodoLan As Integer
Public lDoc As Long
Public iSeq As Integer
Public iLote As Integer
Public iPeriodoLote As Integer
Public dtData As Date
Public sConta As String
Public sCcl As String
Public sHistorico As String
Public dValor As Double
Public iFilialEmpresa As Integer
Public lNumIntDoc As Long
Public iFilialCliForn As Integer
Public lCliForn As Long
Public iTransacao As Integer
Public iAglutina As Integer
Public sProduto As String 'utilizado por lançamentos de custo
Public iApropriaCRProd As Integer 'utilizado para indicar que o produto usa CRProd (movimento de produção)
Public lLoteOrigem As Long 'utilizado por lançamento pré-pendente
Public lContaSimples As Long
Public iSeqContraPartida As Integer
Public iSeqGrid As Integer 'guarda o sequencial no grid. Servirá para posteriormente acertar os sequenciais de contra partida já que os lançamentos zerados são jogados fora.
Public iEscaninho_Custo As Integer
Public dValorLivroAuxiliar As Double
Public iClienteFornecedor As Integer '1=indica que é um cliente associado ao lançamento 2=indica um fornecedor.
Public sDocOrigem As String 'guarda referencia para o documento que originou a contabilizacao
'reprocessamento
Public dQuantidade As Double 'guarda a quantidade do produto associada a esta contabilização. Só tem sentido se o campo produto estiver preenchido.
Public dtDataEstoque As Date 'guarda a data de estocagem associada a contabilidade. Servirá para o reprocessamento e o calculo do custo correto. Só tem sentido se o campo produto estiver preenchido.
Public iStatus As Integer
Public sModelo As String 'modelo de contabilização associado ao lançamento
Public iGerencial As Integer '0=Oficial 1=Gerencial 2=Gererencial e Oficial
Private mvarsUsuario As String
Private mvardtDataRegistro As Date
Private mvardHoraRegistro As Double

Public Property Let dHoraRegistro(ByVal vData As Double)
    mvardHoraRegistro = vData
End Property

Public Property Get dHoraRegistro() As Double
    dHoraRegistro = mvardHoraRegistro
End Property

Public Property Let dtDataRegistro(ByVal vData As Date)
    mvardtDataRegistro = vData
End Property

Public Property Get dtDataRegistro() As Date
    dtDataRegistro = mvardtDataRegistro
End Property

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

Private Sub Class_Initialize()
    sConta = String(STRING_CONTA, 0)
    sHistorico = String(STRING_HISTORICO, 0)
    sCcl = String(STRING_CCL, 0)
    sOrigem = String(STRING_ORIGEM, 0)
    sProduto = String(STRING_PRODUTO, 0)
End Sub
