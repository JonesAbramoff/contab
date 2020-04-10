VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPeriodo"
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
Private mvarsNomeExterno As String
Private mvardtDataInicio As Date
Private mvardtDataFim As Date
Private mvariLancamentos As Integer
Private mvariApurado As Integer
Private mvariLote As Integer
Private mvariFechado As Integer
Private mvardtDataApuracao As Date
Private mvarlDocApuracao As Long
Private mvariFechadoCTB As Integer

Public Property Let iFechadoCTB(ByVal vData As Integer)
    mvariFechadoCTB = vData
End Property

Public Property Get iFechadoCTB() As Integer
    iFechadoCTB = mvariFechadoCTB
End Property

Public Property Let lDocApuracao(ByVal vData As Long)
    mvarlDocApuracao = vData
End Property

Public Property Get lDocApuracao() As Long
    lDocApuracao = mvarlDocApuracao
End Property

Public Property Let dtDataApuracao(ByVal vData As Date)
    mvardtDataApuracao = vData
End Property

Public Property Get dtDataApuracao() As Date
    dtDataApuracao = mvardtDataApuracao
End Property

Public Property Let iFechado(ByVal vData As Integer)
    mvariFechado = vData
End Property

Public Property Get iFechado() As Integer
    iFechado = mvariFechado
End Property

Public Property Let iLote(ByVal vData As Integer)
    mvariLote = vData
End Property

Public Property Get iLote() As Integer
    iLote = mvariLote
End Property

Public Property Let iApurado(ByVal vData As Integer)
    mvariApurado = vData
End Property

Public Property Get iApurado() As Integer
    iApurado = mvariApurado
End Property

Public Property Let iLancamentos(ByVal vData As Integer)
    mvariLancamentos = vData
End Property

Public Property Get iLancamentos() As Integer
    iLancamentos = mvariLancamentos
End Property

Public Property Let dtDataFim(ByVal vData As Date)
Attribute dtDataFim.VB_Description = "Data em que o periodo terminou"
    mvardtDataFim = vData
End Property

Public Property Get dtDataFim() As Date
    dtDataFim = mvardtDataFim
End Property

Public Property Let dtDataInicio(ByVal vData As Date)
Attribute dtDataInicio.VB_Description = "Data em que o periodo iniciou."
    mvardtDataInicio = vData
End Property

Public Property Get dtDataInicio() As Date
    dtDataInicio = mvardtDataInicio
End Property

Public Property Let sNomeExterno(ByVal vData As String)
Attribute sNomeExterno.VB_Description = "Nome por extenso do periodo. Utilizado nas telas e relatórios"
    mvarsNomeExterno = vData
End Property

Public Property Get sNomeExterno() As String
    sNomeExterno = mvarsNomeExterno
End Property

Public Property Let iPeriodo(ByVal vData As Integer)
    mvariPeriodo = vData
End Property

Public Property Get iPeriodo() As Integer
    iPeriodo = mvariPeriodo
End Property

Public Property Let iExercicio(ByVal vData As Integer)
    mvariExercicio = vData
End Property

Public Property Get iExercicio() As Integer
    iExercicio = mvariExercicio
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Private Sub Class_Initialize()

    mvarsNomeExterno = String(STRING_PERIODO_NOME_EXTERNO, 0)

End Sub
