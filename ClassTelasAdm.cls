VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTelasAdm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private gobjFormMsgErro As Object
Private gobjFormMsgErroBatch As Object
Private gobjFormMsgAviso As Object

Public Property Get objFormMsgErroBatch() As Object
    If gobjFormMsgErroBatch Is Nothing Then
        Set gobjFormMsgErroBatch = New FormMsgErroBatch
    End If
    Set objFormMsgErroBatch = gobjFormMsgErroBatch
End Property

Public Property Get objFormMsgErro() As Object
    If gobjFormMsgErro Is Nothing Then
        Set gobjFormMsgErro = New FormMsgErro
    End If
    Set objFormMsgErro = gobjFormMsgErro
End Property

Public Property Get objFormMsgAviso() As Object
    If gobjFormMsgAviso Is Nothing Then
        Set gobjFormMsgAviso = New FormMsgAviso
    End If
    Set objFormMsgAviso = gobjFormMsgAviso
End Property

Public Property Get objBrowseConfigura() As Object
    Set objBrowseConfigura = BrowseConfigura
End Property

Public Property Get objRelSelecionar() As Object
    Set objRelSelecionar = RelSelecionar
End Property

Public Property Get objReimpressao() As Object
    Set objReimpressao = Reimpressao
End Property

Public Property Get objFeriados() As Object
    Set objFeriados = Feriados
End Property

'Public Property Get objModulo() As Object
'    Set objModulo = Modulo
'End Property
'
'Public Property Get objEmpresaFilial() As Object
'    Set objEmpresaFilial = EmpresaFilial
'End Property
'
Public Property Get objCalendario() As Object
    Set objCalendario = Calendario
End Property
Public Property Get objCotacaoMoeda() As Object
    Set objCotacaoMoeda = CotacaoMoeda
End Property


Private Sub Class_Terminate()
    Set gobjFormMsgErro = Nothing
End Sub

Public Property Get objPropriedades() As Object
    Set objPropriedades = Propriedades
End Property

Public Property Get objCamposInvisiveis() As Object
    Set objCamposInvisiveis = CamposInvisiveis
End Property

Public Property Get objTrataAlteracao() As Object
    Set objTrataAlteracao = TrataAlteracao
End Property

Public Property Get objGrafico() As Object
    Set objGrafico = Grafico
End Property

Public Property Get objEnviarEmail() As Object
'Incluido por Tulio em 09/01/2002
    
    Set objEnviarEmail = EnviarEmail
End Property

Public Property Get objMoedas() As Object
    Set objMoedas = Moedas
End Property

Public Property Get objRelDlgExec() As Object
    
    Set objRelDlgExec = RelDlgExec
End Property

'#######################################
'Inserido por Wagner
Public Property Get objPlanilhasSelecionar() As Object
    Set objPlanilhasSelecionar = PlanilhasSelecionar
End Property

Public Property Get objOLAP() As Object
    Set objOLAP = OLAP
End Property

Public Property Get objWait() As Object
    Set objWait = Wait
End Property

Public Property Get objLocalizarGrid() As Object
    Set objLocalizarGrid = LocalizarGrid
End Property

Public Property Get objBackupConfig() As Object
    Set objBackupConfig = BackupConfig
End Property

Public Property Get objLogo() As Object
    Set objLogo = Logo
End Property
'#######################################
