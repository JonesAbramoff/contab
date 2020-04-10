VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassAdmMeioPagtoCondPagto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariAdmMeioPagto As Integer
Private mvariFilialEmpresa As Integer
Private mvariParcelamento As Integer
Private mvarsNomeParcelamento As String
Private mvariNumParcelas As Integer
Private mvariParcelasRecebto As Integer
Private mvardTaxa As Double
Private mvardValorMinimo As Double
Private mvardDesconto As Double
Private mvariJurosParcelamento As Integer
Private mvardJuros As Double
Private mvarcolParcelas As New Collection
Private mvardSaldo As Double
Private mvarsNomeAdmMeioPagto As String
Private mvariTipoCartao As Integer
Private mvariOrigem As Integer
Private mvariAtivo As Integer
Private mvardtDataLog As Date
Private mvariPreDatado As Integer
Private mvardtDataPreDatado As Date

Public Property Let dtDataPreDatado(ByVal vData As Date)
    mvardtDataPreDatado = vData
End Property

Public Property Get dtDataPreDatado() As Date
    dtDataPreDatado = mvardtDataPreDatado
End Property

Public Property Let iPreDatado(ByVal vData As Integer)
    mvariPreDatado = vData
End Property

Public Property Get iPreDatado() As Integer
    iPreDatado = mvariPreDatado
End Property

Public Property Let dtDataLog(ByVal vData As Date)
    mvardtDataLog = vData
End Property

Public Property Get dtDataLog() As Date
    dtDataLog = mvardtDataLog
End Property

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property

Public Property Let iOrigem(ByVal vData As Integer)
    mvariOrigem = vData
End Property

Public Property Get iOrigem() As Integer
    iOrigem = mvariOrigem
End Property

Public Property Let iTipoCartao(ByVal vData As Integer)
    mvariTipoCartao = vData
End Property

Public Property Get iTipoCartao() As Integer
    iTipoCartao = mvariTipoCartao
End Property

Public Property Let sNomeAdmMeioPagto(ByVal vData As String)
    mvarsNomeAdmMeioPagto = vData
End Property

Public Property Get sNomeAdmMeioPagto() As String
    sNomeAdmMeioPagto = mvarsNomeAdmMeioPagto
End Property

Public Property Let dSaldo(ByVal vData As Double)
    mvardSaldo = vData
End Property

Public Property Get dSaldo() As Double
    dSaldo = mvardSaldo
End Property

Public Property Set colParcelas(ByVal vData As Collection)
Attribute colParcelas.VB_Description = "Coleção que guarda as parcelas."
    Set mvarcolParcelas = vData
End Property

Public Property Get colParcelas() As Collection
    Set colParcelas = mvarcolParcelas
End Property

Public Property Let dJuros(ByVal vData As Double)
Attribute dJuros.VB_Description = "% de juros em cada parcela nesse parcelamento"
    mvardJuros = vData
End Property

Public Property Get dJuros() As Double
    dJuros = mvardJuros
End Property

Public Property Let iJurosParcelamento(ByVal vData As Integer)
Attribute iJurosParcelamento.VB_Description = "De quem é a responsabilidade do parcelamento"
    mvariJurosParcelamento = vData
End Property

Public Property Get iJurosParcelamento() As Integer
    iJurosParcelamento = mvariJurosParcelamento
End Property

Public Property Let dDesconto(ByVal vData As Double)
Attribute dDesconto.VB_Description = "% de desconto ofrecido nesse parcelamento"
    mvardDesconto = vData
End Property

Public Property Get dDesconto() As Double
    dDesconto = mvardDesconto
End Property

Public Property Let dValorMinimo(ByVal vData As Double)
Attribute dValorMinimo.VB_Description = "Valor Mínimo necessário para a aceitação do parcelamento."
    mvardValorMinimo = vData
End Property

Public Property Get dValorMinimo() As Double
    dValorMinimo = mvardValorMinimo
End Property

Public Property Let dTaxa(ByVal vData As Double)
Attribute dTaxa.VB_Description = "Taxa cobrada pela adm nesse parcelamento"
    mvardTaxa = vData
End Property

Public Property Get dTaxa() As Double
    dTaxa = mvardTaxa
End Property

Public Property Let iParcelasRecebto(ByVal vData As Integer)
Attribute iParcelasRecebto.VB_Description = "Em quantas parcelas a loja recebe da adm. para esse Meio de pagamento."
    mvariParcelasRecebto = vData
End Property

Public Property Get iParcelasRecebto() As Integer
    iParcelasRecebto = mvariParcelasRecebto
End Property

Public Property Let iNumParcelas(ByVal vData As Integer)
Attribute iNumParcelas.VB_Description = "Número de parcelas nesse parcelamento"
    mvariNumParcelas = vData
End Property

Public Property Get iNumParcelas() As Integer
    iNumParcelas = mvariNumParcelas
End Property

Public Property Let sNomeParcelamento(ByVal vData As String)
Attribute sNomeParcelamento.VB_Description = "Descreve o Parcelamento"
    mvarsNomeParcelamento = vData
End Property

Public Property Get sNomeParcelamento() As String
    sNomeParcelamento = mvarsNomeParcelamento
End Property

Public Property Let iParcelamento(ByVal vData As Integer)
Attribute iParcelamento.VB_Description = "Código do parcelamento"
    mvariParcelamento = vData
End Property

Public Property Get iParcelamento() As Integer
    iParcelamento = mvariParcelamento
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
Attribute iFilialEmpresa.VB_Description = "Filial Empresa em que o Meio de Pagamento é aceito"
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iAdmMeioPagto(ByVal vData As Integer)
Attribute iAdmMeioPagto.VB_Description = "Código do Meio de Pagamento"
    mvariAdmMeioPagto = vData
End Property

Public Property Get iAdmMeioPagto() As Integer
    iAdmMeioPagto = mvariAdmMeioPagto
End Property

Private Sub Class_Initialize()
    mvardtDataPreDatado = DATA_NULA
End Sub
