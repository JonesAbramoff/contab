VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassConfiguracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsGrupoReceita As String
Private mvarsGrupoDespesa As String
Private mvarsGrupoAtivo As String
Private mvarsGrupoPassivo As String
Private mvariIdAtualizacao As Integer
Private mvariUsoCcl As Integer
Private mvariDocPorPeriodo As Integer
Private mvariLotePorPeriodo As Integer
Private mvariTipoContaDefault As Integer
Private mvariNaturezaDefault As Integer
Private mvariHistPadrao As Integer
Private mvardtDataImplantacao As Date
Private mvarlDocAuto As Long
Private mvariContabCentralizada As Integer
Private mvariContabGerencial As Integer
Private mvariFilialAuxiliar As Integer

Public Property Let iFilialAuxiliar(ByVal vData As Integer)
    mvariFilialAuxiliar = vData
End Property

Public Property Get iFilialAuxiliar() As Integer
        iFilialAuxiliar = mvariFilialAuxiliar
End Property

Public Property Let iContabGerencial(ByVal vData As Integer)
    mvariContabGerencial = vData
End Property

Public Property Get iContabGerencial() As Integer
        iContabGerencial = mvariContabGerencial
End Property


Public Property Let iContabCentralizada(ByVal vData As Integer)
    mvariContabCentralizada = vData
End Property

Public Property Get iContabCentralizada() As Integer
        iContabCentralizada = mvariContabCentralizada
End Property

Public Property Let lDocAuto(ByVal vData As Long)
Attribute lDocAuto.VB_Description = "Número do ultimo documento automatico gerado pelo sistema"
    mvarlDocAuto = vData
End Property

Public Property Get lDocAuto() As Long
    lDocAuto = mvarlDocAuto
End Property

Public Property Let dtDataImplantacao(ByVal vData As Date)
Attribute dtDataImplantacao.VB_Description = "Data da Implantação do sistema."
    mvardtDataImplantacao = vData
End Property

Public Property Get dtDataImplantacao() As Date
    dtDataImplantacao = mvardtDataImplantacao
End Property

Public Property Let iHistPadrao(ByVal vData As Integer)
Attribute iHistPadrao.VB_Description = "Número do ultimo historico gerado automaticamente pelo sistema."
    mvariHistPadrao = vData
End Property

Public Property Get iHistPadrao() As Integer
        iHistPadrao = mvariHistPadrao
End Property

Public Property Let iNaturezaDefault(ByVal vData As Integer)
Attribute iNaturezaDefault.VB_Description = "Natureza da Conta default. 1=crédito 2=débito"
    mvariNaturezaDefault = vData
End Property

Public Property Get iNaturezaDefault() As Integer
    iNaturezaDefault = mvariNaturezaDefault
End Property

Public Property Let iTipoContaDefault(ByVal vData As Integer)
Attribute iTipoContaDefault.VB_Description = "Tipo de Conta default. 1=sintetica 3=analitica"
    mvariTipoContaDefault = vData
End Property

Public Property Get iTipoContaDefault() As Integer
    iTipoContaDefault = mvariTipoContaDefault
End Property

Public Property Let iLotePorPeriodo(ByVal vData As Integer)
Attribute iLotePorPeriodo.VB_Description = "Indica a inicialização do numero do LOTE por periodo ou exercicio. 0=exercicio 1=periodo"
    mvariLotePorPeriodo = vData
End Property

Public Property Get iLotePorPeriodo() As Integer
    iLotePorPeriodo = mvariLotePorPeriodo
End Property

Public Property Let iDocPorPeriodo(ByVal vData As Integer)
Attribute iDocPorPeriodo.VB_Description = "Indica a inicialização do numero do DOC (voucher) por periodo ou exercicio. 0=exercicio 1=periodo"
    mvariDocPorPeriodo = vData
End Property

Public Property Get iDocPorPeriodo() As Integer
    iDocPorPeriodo = mvariDocPorPeriodo
End Property

Public Property Let iUsoCcl(ByVal vData As Integer)
Attribute iUsoCcl.VB_Description = "Indica o uso ou não de centro de custo. 0=sem centro de custo 1=ccusto contábil  2=ccusto extra contábil"
    mvariUsoCcl = vData
End Property

Public Property Get iUsoCcl() As Integer
    iUsoCcl = mvariUsoCcl
End Property

Public Property Let iIDAtualizacao(ByVal vData As Integer)
Attribute iIDAtualizacao.VB_Description = "Ultimo número utilizado para  identificar um processo que solicitou a atualização."
    mvariIdAtualizacao = vData
End Property

Public Property Get iIDAtualizacao() As Integer
    iIDAtualizacao = mvariIdAtualizacao
End Property

Public Property Let sGrupoPassivo(ByVal vData As String)
Attribute sGrupoPassivo.VB_Description = "Grupo que compõe as contas de passivo"
    mvarsGrupoPassivo = vData
End Property

Public Property Get sGrupoPassivo() As String
    sGrupoPassivo = mvarsGrupoPassivo
End Property

Public Property Let sGrupoAtivo(ByVal vData As String)
Attribute sGrupoAtivo.VB_Description = "Grupo que compõe as contas de ativo"
    mvarsGrupoAtivo = vData
End Property

Public Property Get sGrupoAtivo() As String
    sGrupoAtivo = mvarsGrupoAtivo
End Property

Public Property Let sGrupoDespesa(ByVal vData As String)
Attribute sGrupoDespesa.VB_Description = "Grupo que compõe as contas de despesa"
    mvarsGrupoDespesa = vData
End Property

Public Property Get sGrupoDespesa() As String
    sGrupoDespesa = mvarsGrupoDespesa
End Property

Public Property Let sGrupoReceita(ByVal vData As String)
Attribute sGrupoReceita.VB_Description = "Grupo que compõe as contas de receita"
    mvarsGrupoReceita = vData
End Property

Public Property Get sGrupoReceita() As String
    sGrupoReceita = mvarsGrupoReceita
End Property



