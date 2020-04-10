VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPeriodosFilial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvariExercicio As Integer
Private mvariPeriodo As Integer
Private mvariLancamentos As Integer
Private mvariApurado As Integer
Private mvariLote As Integer
Private mvariFechado As Integer
Private mvardtDataApuracao As Date
Private mvarlDocApuracao As Long
Private mvarsNomeExterno As String
Private mvardtDataInicio As Date
Private mvardtDataFim As Date
Private mvarsOrigem As String

Public Property Let sOrigem(ByVal vData As String)
    mvarsOrigem = vData
End Property

Public Property Get sOrigem() As String
    sOrigem = mvarsOrigem
End Property

Public Property Let dtDataFim(ByVal vData As Date)
    mvardtDataFim = vData
End Property

Public Property Get dtDataFim() As Date
    dtDataFim = mvardtDataFim
End Property

Public Property Let dtDataInicio(ByVal vData As Date)
    mvardtDataInicio = vData
End Property

Public Property Get dtDataInicio() As Date
    dtDataInicio = mvardtDataInicio
End Property

Public Property Let sNomeExterno(ByVal vData As String)
    mvarsNomeExterno = vData
End Property

Public Property Get sNomeExterno() As String
    sNomeExterno = mvarsNomeExterno
End Property

Public Property Let lDocApuracao(ByVal vData As Long)
Attribute lDocApuracao.VB_Description = "Numero do ultimo documento de apuração do periodo  gerado."
    mvarlDocApuracao = vData
End Property

Public Property Get lDocApuracao() As Long
    lDocApuracao = mvarlDocApuracao
End Property

Public Property Let dtDataApuracao(ByVal vData As Date)
Attribute dtDataApuracao.VB_Description = "Data da ultima apuração."
    mvardtDataApuracao = vData
End Property

Public Property Get dtDataApuracao() As Date
    dtDataApuracao = mvardtDataApuracao
End Property

Public Property Let iFechado(ByVal vData As Integer)
Attribute iFechado.VB_Description = "0=aberto, 1=fechado (não admite lançamentos)"
    mvariFechado = vData
End Property

Public Property Get iFechado() As Integer
    iFechado = mvariFechado
End Property

Public Property Let iLote(ByVal vData As Integer)
Attribute iLote.VB_Description = "Código do ultimo Lote  de apuraçao do periodo"
    mvariLote = vData
End Property

Public Property Get iLote() As Integer
    iLote = mvariLote
End Property

Public Property Let iApurado(ByVal vData As Integer)
Attribute iApurado.VB_Description = "0=nao apurado, 1=apurado"
    mvariApurado = vData
End Property

Public Property Get iApurado() As Integer
    iApurado = mvariApurado
End Property

Public Property Let iLancamentos(ByVal vData As Integer)
Attribute iLancamentos.VB_Description = "Indica se os lançamentos do periodo estão no BD (Sim) ou não. 1= Sim 0=nao"
    mvariLancamentos = vData
End Property

Public Property Get iLancamentos() As Integer
    iLancamentos = mvariLancamentos
End Property

Public Property Let iPeriodo(ByVal vData As Integer)
Attribute iPeriodo.VB_Description = "Número Sequencial (inicialmente de 1 a 12)"
    mvariPeriodo = vData
End Property

Public Property Get iPeriodo() As Integer
    iPeriodo = mvariPeriodo
End Property

Public Property Let iExercicio(ByVal vData As Integer)
Attribute iExercicio.VB_Description = "Ano de instalação do Exercicio"
    mvariExercicio = vData
End Property

Public Property Get iExercicio() As Integer
    iExercicio = mvariExercicio
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
Attribute iFilialEmpresa.VB_Description = "identifica a filial ou estabelecimento da empresa do usuario"
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property



