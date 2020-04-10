VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form TelaAcompanhaBatchCPR2 
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   5700
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Log 
      BackColor       =   &H8000000F&
      Height          =   1965
      Left            =   135
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   2775
      Width           =   5445
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   600
      Top             =   4200
   End
   Begin VB.CommandButton BotaoCancelar 
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
      Height          =   300
      Left            =   1440
      TabIndex        =   10
      Top             =   4800
      Width           =   2880
   End
   Begin VB.Frame Frame1 
      Caption         =   "Processo"
      Height          =   975
      Left            =   150
      TabIndex        =   5
      Top             =   135
      Width           =   2670
      Begin VB.Label Label1 
         Caption         =   "Emails enviados:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   9
         Top             =   270
         Width           =   1860
      End
      Begin VB.Label Label3 
         Caption         =   "Total de emails:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   8
         Top             =   615
         Width           =   1860
      End
      Begin VB.Label ItensProc 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   270
         Left            =   1710
         TabIndex        =   7
         Top             =   285
         Width           =   525
      End
      Begin VB.Label TotalItens 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   270
         Left            =   1710
         TabIndex        =   6
         Top             =   615
         Width           =   525
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tempo"
      Height          =   975
      Left            =   2925
      TabIndex        =   0
      Top             =   135
      Width           =   2670
      Begin VB.Label Label4 
         Caption         =   "Restante:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   4
         Top             =   615
         Width           =   1860
      End
      Begin VB.Label Label6 
         Caption         =   "Decorrido:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   3
         Top             =   270
         Width           =   1860
      End
      Begin VB.Label TempoDecorrido 
         Alignment       =   1  'Right Justify
         Caption         =   "00:00:00"
         Height          =   270
         Left            =   1080
         TabIndex        =   2
         Top             =   285
         Width           =   1155
      End
      Begin VB.Label TempoRestante 
         Alignment       =   1  'Right Justify
         Caption         =   "00:00:00"
         Height          =   270
         Left            =   1080
         TabIndex        =   1
         Top             =   615
         Width           =   1155
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   150
      Top             =   4200
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   375
      Left            =   135
      TabIndex        =   11
      Top             =   1440
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   135
      TabIndex        =   14
      Top             =   2115
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label9 
      Caption         =   "Log:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   135
      TabIndex        =   20
      Top             =   2550
      Width           =   1185
   End
   Begin VB.Label Label8 
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   135
      TabIndex        =   18
      Top             =   1845
      Width           =   1185
   End
   Begin VB.Label Label7 
      Caption         =   "Email atual:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   135
      TabIndex        =   17
      Top             =   1170
      Width           =   1185
   End
   Begin VB.Label Label5 
      Caption         =   "Concluido:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3930
      TabIndex        =   15
      Top             =   1860
      Width           =   900
   End
   Begin VB.Label Percentual 
      Alignment       =   1  'Right Justify
      Caption         =   "0%"
      Height          =   270
      Left            =   4515
      TabIndex        =   16
      Top             =   1860
      Width           =   1065
   End
   Begin VB.Label Label2 
      Caption         =   "Concluido:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3930
      TabIndex        =   12
      Top             =   1185
      Width           =   900
   End
   Begin VB.Label PercentualEmail 
      Alignment       =   1  'Right Justify
      Caption         =   "0%"
      Height          =   270
      Left            =   4515
      TabIndex        =   13
      Top             =   1185
      Width           =   1065
   End
End
Attribute VB_Name = "TelaAcompanhaBatchCPR2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iCancelaBatch As Integer
Public sNomeArqParam As String

Public gcolEnvioDeEmail As Collection

Dim dTempoInicial As Double
Dim iTotalItens As Integer
Dim iItensProcessados As Integer
Dim dMediaTempoItem As Double
Dim dTempoEstimado As Double
Dim dPercConcluido As Double

Dim giFalhaNoEnvio As Integer

Private Sub BotaoCancelar_Click()

Dim lErro As Long

On Error GoTo Erro_Cancelar_Click

    iCancelaBatch = CANCELA_BATCH

    Exit Sub

Erro_Cancelar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196977)

    End Select

    Exit Sub
    
End Sub

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    dTempoInicial = CDbl(Time)
    iTotalItens = gcolEnvioDeEmail.Count
    TotalItens.Caption = gcolEnvioDeEmail.Count
    giFalhaNoEnvio = DESMARCADO
    
    Timer1.Interval = 1005
    Timer1.Enabled = True
   
    Timer2.Interval = 1000
    Timer2.Enabled = True
          
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196978)

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
Dim sErro As String
Dim objRotEnviodeEmail As New ClassRotEnviodeEmail

On Error GoTo Erro_Timer1_Timer

    Timer1.Interval = 0

    lErro = Sistema_Abrir_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 196979

    Set gcolModulo = New AdmColModulo
    
    lErro = CF("Modulos_Le_Empresa_Filial", glEmpresa, giFilialEmpresa, gcolModulo)
    If lErro <> SUCESSO Then gError 196980
            
    lErro = objRotEnviodeEmail.Rotina_Envia_Emails(gcolEnvioDeEmail)
    If lErro <> SUCESSO Then gError 196981

    iCancelaBatch = CANCELA_BATCH

    Unload Me

    Exit Sub

Erro_Timer1_Timer:

    If iCancelaBatch <> CANCELA_BATCH Then

'        sErro = "Houve algum tipo de erro. Verifique o arquivo de log de erros configurado em \windows\adm100.ini ."
'        Call MsgBox(sErro, vbOKOnly, "SGE-Forprint")

        Call Rotina_ErrosBatch
    
    End If

    Select Case gErr

        Case 196979 To 196981

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196982)

    End Select

    iCancelaBatch = CANCELA_BATCH
    Unload Me

    Exit Sub

End Sub

Public Function ProcessouItem() As Long

Dim dTempo As Double
Dim vbMsgBox As VbMsgBoxResult

On Error GoTo Erro_ProcessouItem

    dTempo = CDbl(Time)
    iItensProcessados = iItensProcessados + 1
    dMediaTempoItem = (dTempo - dTempoInicial) / iItensProcessados
    dTempoEstimado = (iTotalItens - iItensProcessados) * dMediaTempoItem
    dPercConcluido = iItensProcessados / iTotalItens
        
    Percentual.Caption = Format(dPercConcluido, "PERCENT")
    ItensProc.Caption = CStr(iItensProcessados)
    TempoRestante.Caption = Format(dTempoEstimado, "HH:MM:SS")
    ProgressBar1.Value = (iItensProcessados / iTotalItens) * 100
    
    DoEvents
    
    If (iCancelaBatch = CANCELA_BATCH) Or (BotaoCancelar.Enabled = False) Then

        'Imcomatível com o código de chamnada da tela        'SetWindowPos TelaAcompanhaBatchCPR2.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
        'vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_ENVIO_EMAIL")
        vbMsgBox = vbYes
        If vbMsgBox = vbYes Then gError 196983

        iCancelaBatch = 0

    End If
    
    If giFalhaNoEnvio = MARCADO Then gError 196983
    
    Log.Text = ""
    
    ProcessouItem = SUCESSO
    
    Exit Function
    
Erro_ProcessouItem:

    ProcessouItem = gErr
    
    Select Case gErr
    
        Case 196983
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196984)

    End Select
    
    Exit Function
    
End Function

Private Sub Timer2_Timer()

Dim dTempo As Double

    dTempo = CDbl(Time)
    
    TempoDecorrido.Caption = Format(dTempo - dTempoInicial, "HH:MM:SS")
    
End Sub

Public Function Trata_Progresso(ByVal lPercentCompete As Long) As Long

Dim dPercConcluido As Double

On Error GoTo Erro_Trata_Progresso

    dPercConcluido = lPercentCompete / 100

    PercentualEmail.Caption = Format(dPercConcluido, "PERCENT")
    ProgressBar2.Value = lPercentCompete

    Trata_Progresso = SUCESSO
    
    Exit Function
    
Erro_Trata_Progresso:

    Trata_Progresso = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196984)

    End Select
    
    Exit Function
    
End Function

Public Function Trata_Status(ByVal sStatus As String) As Long

On Error GoTo Erro_Trata_Status

    If Len(Trim(Log.Text)) > 0 Then Log.Text = Log.Text & vbNewLine
    Log.Text = Log.Text & sStatus
    
    Trata_Status = SUCESSO
    
    Exit Function
    
Erro_Trata_Status:

    Trata_Status = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196984)

    End Select
    
    Exit Function
    
End Function

Public Function Trata_Falha(ByVal sFalha As String) As Long

On Error GoTo Erro_Trata_Falha

    If Len(Trim(Log.Text)) > 0 Then Log.Text = Log.Text & vbNewLine
    Log.Text = Log.Text & sFalha
    
    giFalhaNoEnvio = MARCADO
    
    Call Rotina_Erro(vbOKOnly, sFalha, gErr, Error, 196984)
    
    Trata_Falha = SUCESSO
    
    Exit Function
    
Erro_Trata_Falha:

    Trata_Falha = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196984)

    End Select
    
    Exit Function
    
End Function

Public Function Trata_Sucesso() As Long

On Error GoTo Erro_Trata_Sucesso

    If Len(Trim(Log.Text)) > 0 Then Log.Text = Log.Text & vbNewLine
    Log.Text = Log.Text & "Email enviado com sucesso."
    
    Trata_Sucesso = SUCESSO
    
    Exit Function
    
Erro_Trata_Sucesso:

    Trata_Sucesso = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196984)

    End Select
    
    Exit Function
    
End Function
