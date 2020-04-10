VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassExercicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariExercicio As Integer
Private mvariNumPeriodos As Integer
Private mvariStatus As Integer
Private mvariLoteApuracao As Integer
Private mvarsNomeExterno As String
Private mvardtDataInicio As Date
Private mvardtDataFim As Date
Private mvariFilialEmpresa As Integer

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let dtDataFim(ByVal vData As Date)
Attribute dtDataFim.VB_Description = "Data em que o periodo terminou"
    mvardtDataFim = vData
End Property

Public Property Get dtDataFim() As Date
    dtDataFim = mvardtDataFim
End Property

Public Property Let dtDataInicio(ByVal vData As Date)
Attribute dtDataInicio.VB_Description = "Data em que o periodo iniciou"
    mvardtDataInicio = vData
End Property

Public Property Get dtDataInicio() As Date
    dtDataInicio = mvardtDataInicio
End Property

Public Property Let sNomeExterno(ByVal vData As String)
Attribute sNomeExterno.VB_Description = "Nome descritivo a ser usado nas telas e relatórios."
    mvarsNomeExterno = vData
End Property

Public Property Get sNomeExterno() As String
    sNomeExterno = mvarsNomeExterno
End Property

Public Property Let iLoteApuracao(ByVal vData As Integer)
Attribute iLoteApuracao.VB_Description = "Número do último lote de apuração do exercicio"
    mvariLoteApuracao = vData
End Property

Public Property Get iLoteApuracao() As Integer
    iLoteApuracao = mvariLoteApuracao
End Property

Public Property Let iStatus(ByVal vData As Integer)
Attribute iStatus.VB_Description = "Status do Exercicio 1=aberto 2=apurado 3=fechado"
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let iNumPeriodos(ByVal vData As Integer)
Attribute iNumPeriodos.VB_Description = "Número de Periodos do Exercicio"
    mvariNumPeriodos = vData
End Property

Public Property Get iNumPeriodos() As Integer
    iNumPeriodos = mvariNumPeriodos
End Property

Public Property Let iExercicio(ByVal vData As Integer)
Attribute iExercicio.VB_Description = "Sequencial do exercicio"
    mvariExercicio = vData
End Property

Public Property Get iExercicio() As Integer
    iExercicio = mvariExercicio
End Property

Private Sub Class_Initialize()

    mvarsNomeExterno = String(STRING_PERIODO_NOME_EXTERNO, 0)

End Sub



