VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form TelaAcompanhaBatch 
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   Icon            =   "TelaAcompanha.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4635
      Top             =   2700
   End
   Begin VB.CommandButton Cancelar 
      Caption         =   "Cancelar"
      Height          =   735
      Left            =   2205
      TabIndex        =   1
      Top             =   2490
      Width           =   1395
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   405
      Left            =   585
      TabIndex        =   0
      Top             =   1395
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   714
      _Version        =   327682
      Appearance      =   1
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
      Left            =   600
      TabIndex        =   4
      Top             =   1125
      Width           =   1305
   End
   Begin VB.Label TotReg 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   3615
      TabIndex        =   3
      Top             =   360
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
      Left            =   585
      TabIndex        =   2
      Top             =   420
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

'Rotina Fechamento e Reabertura de Exercicio
Public sConta_Ativo_Inicial As String
Public sConta_Ativo_Final As String
Public sConta_Passivo_Inicial As String
Public sConta_Passivo_Final As String

'Rotina Reprocessamento
Public iPeriodo As Integer

'Private Sub Cancelar_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_Cancelar_Click
'
'    iCancelaBatch = CANCELA_BATCH
'
'    Exit Sub
'
'Erro_Cancelar_Click:
'
'    Select Case Err
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174585)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub Form_Load()
'
'Dim lErro As Long
'
'On Error GoTo Erro_Form_Load
'
'    TotReg.Caption = "0"
'    ProgressBar1.Min = 0
'    ProgressBar1.Max = 100
'
'    lErro_Chama_Tela = SUCESSO
'
'    Exit Sub
'
'Erro_Form_Load:
'
'    lErro_Chama_Tela = Err
'
'    Select Case Err
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174586)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
'    If iCancelaBatch <> CANCELA_BATCH Then
'        iCancelaBatch = CANCELA_BATCH
'        Cancel = 1
'    End If
'
'End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    End
'End Sub

'Private Sub Timer1_Timer()
'
'Dim lErro As Long
'Dim lteste As Long
'
'On Error GoTo Erro_Timer1_Timer
'
'    Timer1.Interval = 0
'
'    lErro = Sistema_Abrir_Batch(sNomeArqParam)
'    If lErro <> SUCESSO Then Error 27440
'
'    Select Case iRotinaBatch
'
'        Case ROTINA_ATUALIZACAO_BATCH
'            lErro = Rotina_Atualizacao_Int(iIdAtualizacao_Param)
'            If lErro <> SUCESSO Then Error 20338
'
'        Case ROTINA_APURA_EXERCICIO_BATCH
'            lErro = Rotina_Apura_Exercicio_Int(iFilialEmpresa, iExercicio, iLote, sHistorico, sContaResultado, colContasApuracao)
'            If lErro <> SUCESSO Then Error 20341
'
'        Case ROTINA_APURA_PERIODOS_BATCH
'            lErro = Rotina_Apura_Periodos_Int(iFilialEmpresa, iExercicio, iPeriodo_Inicial, iPeriodo_Final, sContaResultado, sContaPonte, colContasApuracao, sHistorico)
'            If lErro <> SUCESSO Then Error 20344
'
'        Case ROTINA_FECHAMENTO_EXERCICIO_BATCH
'            lErro = Rotina_Fechamento_Exercicio_Int(iExercicio, sConta_Ativo_Inicial, sConta_Ativo_Final, sConta_Passivo_Inicial, sConta_Passivo_Final)
'            If lErro <> SUCESSO Then Error 20346
'
'        Case ROTINA_REABERTURA_EXERCICIO_BATCH
'            lErro = Rotina_Reabertura_Exercicio_Int(iExercicio, sConta_Ativo_Inicial, sConta_Ativo_Final, sConta_Passivo_Inicial, sConta_Passivo_Final)
'            If lErro <> SUCESSO Then Error 20356
'
'        Case ROTINA_REPROCESSAMENTO_BATCH
'            lErro = Rotina_Reprocessamento_Int(iFilialEmpresa, iExercicio, iPeriodo)
'            If lErro <> SUCESSO Then Error 20361
'
'    End Select
'
'    iCancelaBatch = CANCELA_BATCH
'
'    Unload Me
'
'    Exit Sub
'
'Erro_Timer1_Timer:
'
'    Select Case Err
'
'        Case 20338, 20341, 20344, 20346, 20356, 20361, 27440
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174587)
'
'    End Select
'
'    iCancelaBatch = CANCELA_BATCH
'    Unload Me
'
'    Exit Sub
'
'End Sub
'

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

