VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form TelaAcompanhaBatch 
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4830
   Icon            =   "TelaAcompanhaBatch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3990
      Top             =   1560
   End
   Begin VB.CommandButton Cancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1695
      TabIndex        =   1
      Top             =   1905
      Width           =   1395
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   405
      Left            =   255
      TabIndex        =   0
      Top             =   1185
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label SPEDStatus2 
      Height          =   225
      Left            =   225
      TabIndex        =   6
      Top             =   720
      Width           =   4395
   End
   Begin VB.Label SpedStatus1 
      Height          =   225
      Left            =   225
      TabIndex        =   5
      Top             =   480
      Width           =   4395
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Processamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   255
      TabIndex        =   2
      Top             =   915
      Width           =   1305
   End
   Begin VB.Label TotReg 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   3270
      TabIndex        =   3
      Top             =   150
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número de Registros Processados:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   210
      Width           =   2985
   End
End
Attribute VB_Name = "TelaAcompanhaBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iCancelaBatch As Integer
Public dValorTotal As Double
Public dValorAtual As Double
Public sNomeArqParam As String

'Rotina Batch que a tela está acompanhando
Public iRotinaBatch As Integer

'Rotina de Atualizacao de Lote
Public iIdAtualizacao_Param As Integer

'Rotina de Apuracao de Exercicio
Public iFilialEmpresa As Integer
Public iExercicio As Integer
Public iLote As Integer
Public sHistorico As String
Public sContaResultado As String
Public colContasApuracao As Collection

'Rotina de Apuracao de Periodo
Public iPeriodo_Inicial As Integer
Public iPeriodo_Final As Integer
Public sContaPonte As String
Public iZeraRD As Integer

'Rotina Fechamento e Reabertura de Exercicio
Public sConta_Ativo_Inicial As String
Public sConta_Ativo_Final As String
Public sConta_Passivo_Inicial As String
Public sConta_Passivo_Final As String
Private GL_objKeepAlive As AdmKeepAlive

'Rotina Reprocessamento
Public iPeriodo As Integer

'Rotina de RateioOff
Public objRateioOffBatch As ClassRateioOffBatch

'Rotina Sped Diario
Public sDiretorio As String
Public dtDataIni As Date
Public dtDataFim As Date
Public lNumOrd As Long
Public sContaOutros As String

Public iIndSituacaoPer As Integer
Public iIndSitEspecial As Integer
Public iCodVersao As Integer
Public iIndNIRE As Integer
Public iFinalidade As Integer
Public sHashEscrSubst As String
Public sNIRESubst As String
Public iEmpGrandePorte As Integer

Public iTipoECD As Integer
Public sCodSCP As String
Public colSCPs As Collection

Private Sub Cancelar_Click()

Dim lErro As Long

On Error GoTo Erro_Cancelar_Click

    iCancelaBatch = CANCELA_BATCH
    
    Exit Sub

Erro_Cancelar_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174588)

    End Select

    Exit Sub

End Sub

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    TotReg.Caption = "0"
    ProgressBar1.Min = 0
    ProgressBar1.Max = 100
    Set GL_objKeepAlive = New AdmKeepAlive

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174589)

    End Select

    Exit Sub

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If iCancelaBatch <> CANCELA_BATCH Then
        iCancelaBatch = CANCELA_BATCH
        Cancel = 1
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Timer1_Timer()
    
Dim lErro As Long
Dim lteste As Long, sErro As String
Dim objTela As Object
Dim sTexto As String

On Error GoTo Erro_Timer1_Timer

    Timer1.Interval = 0
    
    gl_UltimoErro = SUCESSO
    
    Select Case iRotinaBatch
    
        Case ROTINA_ATUALIZACAO_BATCH
            sTexto = "ROTINA_ATUALIZACAO_BATCH"
            lErro = Rotina_Atualizacao_Int(iIdAtualizacao_Param)
            If lErro <> SUCESSO Or GL_lUltimoErro <> SUCESSO Then gError 20338
            
        Case ROTINA_APURA_EXERCICIO_BATCH
            sTexto = "ROTINA_APURA_EXERCICIO_BATCH"
            lErro = Rotina_Apura_Exercicio_Int(iFilialEmpresa, iExercicio, iLote, sHistorico, sContaResultado, colContasApuracao)
            If lErro <> SUCESSO Or GL_lUltimoErro <> SUCESSO Then gError 20341

        Case ROTINA_APURA_PERIODOS_BATCH
            sTexto = "ROTINA_APURA_PERIODOS_BATCH"
            lErro = Rotina_Apura_Periodos_Int(iFilialEmpresa, iExercicio, iPeriodo_Inicial, iPeriodo_Final, sContaResultado, sContaPonte, colContasApuracao, sHistorico, iZeraRD)
            If lErro <> SUCESSO Or GL_lUltimoErro <> SUCESSO Then gError 20344

        Case ROTINA_FECHAMENTO_EXERCICIO_BATCH
            sTexto = "ROTINA_FECHAMENTO_EXERCICIO_BATCH"
            lErro = Rotina_Fechamento_Exercicio_Int(iExercicio, sConta_Ativo_Inicial, sConta_Ativo_Final, sConta_Passivo_Inicial, sConta_Passivo_Final)
            If lErro <> SUCESSO Or GL_lUltimoErro <> SUCESSO Then gError 20346
        
        Case ROTINA_REABERTURA_EXERCICIO_BATCH
            sTexto = "ROTINA_REABERTURA_EXERCICIO_BATCH"
            lErro = Rotina_Reabertura_Exercicio_Int(iExercicio, sConta_Ativo_Inicial, sConta_Ativo_Final, sConta_Passivo_Inicial, sConta_Passivo_Final)
            If lErro <> SUCESSO Or GL_lUltimoErro <> SUCESSO Then gError 20356
            
        Case ROTINA_REPROCESSAMENTO_BATCH
            sTexto = "ROTINA_REPROCESSAMENTO_BATCH"
            lErro = Rotina_Reprocessamento_Int(iFilialEmpresa, iExercicio, iPeriodo)
            If lErro <> SUCESSO Or GL_lUltimoErro <> SUCESSO Then gError 20361
    
        Case ROTINA_RATEIOOFF_BATCH
            sTexto = "ROTINA_RATEIOOFF_BATCH"
            lErro = Rotina_RateioOff_Int(objRateioOffBatch)
            If lErro <> SUCESSO Or GL_lUltimoErro <> SUCESSO Then gError 36803
    
        Case ROTINA_DESAPURA_EXERCICIO_BATCH
            sTexto = "ROTINA_DESAPURA_EXERCICIO_BATCH"
            lErro = Rotina_Desapura_Exercicio_Int(iFilialEmpresa, iExercicio, iLote, sHistorico, sContaResultado, colContasApuracao)
            If lErro <> SUCESSO Or GL_lUltimoErro <> SUCESSO Then gError 188423
    
        Case ROTINA_SPED_DIARIO
            sTexto = "ROTINA_SPED_DIARIO"
        
            Set objTela = Me
        
            lErro = CF("Gera_Sped_Contabil_Diario", sDiretorio, iFilialEmpresa, dtDataIni, dtDataFim, lNumOrd, objTela, sContaOutros, SPED_CONTAB_TIPO_NORMAL, iIndSituacaoPer, iIndSitEspecial, iCodVersao, iIndNIRE, iFinalidade, sHashEscrSubst, sNIRESubst, iEmpGrandePorte, iTipoECD, sCodSCP, colSCPs)
            If lErro <> SUCESSO Or GL_lUltimoErro <> SUCESSO Then gError 203250
            
        Case ROTINA_FCONT
            sTexto = "ROTINA_FCONT"
            
            Set objTela = Me
        
            lErro = CF("Gera_Sped_Contabil_Diario", sDiretorio, iFilialEmpresa, dtDataIni, dtDataFim, lNumOrd, objTela, sContaOutros, SPED_CONTAB_TIPO_FCONT, iIndSituacaoPer, iIndSitEspecial)
            If lErro <> SUCESSO Or GL_lUltimoErro <> SUCESSO Then gError 203250
    
    End Select
    
    iCancelaBatch = CANCELA_BATCH
    
    Unload Me
    
    Exit Sub

Erro_Timer1_Timer:

'    If iCancelaBatch <> CANCELA_BATCH Then
'
'        sErro = "Houve algum tipo de erro. Verifique o conteúdo do arquivo de log. Seu nome encontra-se no arquivo \windows\adm100.ini ."
'        Call MsgBox(sErro, vbOKOnly, "SGE-Forprint")
'
'    End If
    
    Select Case gErr

        Case 20338, 20341, 20344, 20346, 20356, 20361, 36803, 188423, 203250

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174590)

    End Select
    
    If iCancelaBatch <> CANCELA_BATCH Then
        Call Rotina_ErrosBatch2(sTexto)
    End If
    
    iCancelaBatch = CANCELA_BATCH
    Unload Me

    Exit Sub
    
End Sub


Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub TotReg_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotReg, Source, X, Y)
End Sub

Private Sub TotReg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotReg, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Public Sub SPED_Status(ByVal sStatus1 As String, ByVal sStatus2 As String)
    SpedStatus1.Caption = sStatus1
    SPEDStatus2.Caption = sStatus2
    DoEvents
End Sub
