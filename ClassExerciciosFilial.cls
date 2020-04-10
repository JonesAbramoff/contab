VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassExerciciosFilial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvariExercicio As Integer
Private mvariStatus As Integer
Private mvariLoteApuracao As Integer
Private mvardtDataApuracao As Date
Private mvarlDocApuracao As Long

Public Property Let lDocApuracao(ByVal vData As Long)
Attribute lDocApuracao.VB_Description = "Número do ultimo documento de apuração utilizado"
    mvarlDocApuracao = vData
End Property

Public Property Get lDocApuracao() As Long
    lDocApuracao = mvarlDocApuracao
End Property

Public Property Let dtDataApuracao(ByVal vData As Date)
Attribute dtDataApuracao.VB_Description = "Data da ultima apuração"
    mvardtDataApuracao = vData
End Property

Public Property Get dtDataApuracao() As Date
    dtDataApuracao = mvardtDataApuracao
End Property

Public Property Let iLoteApuracao(ByVal vData As Integer)
Attribute iLoteApuracao.VB_Description = "Número do último lote de apuração do exercicio"
    mvariLoteApuracao = vData
End Property

Public Property Get iLoteApuracao() As Integer
    iLoteApuracao = mvariLoteApuracao
End Property

Public Property Let iStatus(ByVal vData As Integer)
Attribute iStatus.VB_Description = "Status do Exercicio 1=aberto 2=apurado"
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let iExercicio(ByVal vData As Integer)
Attribute iExercicio.VB_Description = "Sequencial do exercicio"
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

