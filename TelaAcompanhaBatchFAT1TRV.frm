VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form TelaAcompanhaBatchFAT1 
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
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
      Height          =   510
      Left            =   1665
      TabIndex        =   1
      Top             =   1800
      Width           =   1395
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3510
      Top             =   495
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   210
      TabIndex        =   0
      Top             =   1080
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid GridItens 
      Height          =   3030
      Left            =   4110
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1665
      Visible         =   0   'False
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   5345
      _Version        =   393216
      Rows            =   21
      Cols            =   4
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin MSMask.MaskEdBox CTBConta 
      Height          =   225
      Left            =   3720
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      AllowPrompt     =   -1  'True
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label Observacao 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   204
      TabIndex        =   7
      Top             =   468
      Width           =   2208
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
      Left            =   165
      TabIndex        =   6
      Top             =   135
      Width           =   2985
   End
   Begin VB.Label TotReg 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   3240
      TabIndex        =   5
      Top             =   75
      Width           =   1245
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
      Left            =   165
      TabIndex        =   4
      Top             =   810
      Width           =   1305
   End
End
Attribute VB_Name = "TelaAcompanhaBatchFAT1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iCancelaBatch As Integer
Public dValorTotal As Double
Public dValorAtual As Double
Public sNomeArqParam As String
Public giRelatorio As Integer

'Rotina Batch que a tela está acompanhando
Public iRotinaBatch As Integer

'Rotina de Geracao de Comissao Interna
Public gobjTRVGerComiInt As Object


Private Sub Cancelar_Click()

Dim lErro As Long

On Error GoTo Erro_Cancelar_Click

    iCancelaBatch = CANCELA_BATCH

    Exit Sub

Erro_Cancelar_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174597)

    End Select

    Exit Sub

End Sub

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    TotReg.Caption = "0"
    ProgressBar1.Min = 0
    ProgressBar1.Max = 100
    Timer1.Interval = 1
    iCancelaBatch = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174598)

    End Select

    Exit Sub

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If iCancelaBatch <> CANCELA_BATCH Then
        iCancelaBatch = CANCELA_BATCH
        Cancel = 1
        iCancelaBatch = 0
        Cancel = 0
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set gobjTRVGerComiInt = Nothing
End Sub

Private Sub Timer1_Timer()

Dim lErro As Long, sErro As String
Dim lteste As Long
Dim objRotGeraComiInt As New ClassRotGeraComiInt

On Error GoTo Erro_Timer1_Timer

    Timer1.Interval = 0

    
    Select Case iRotinaBatch
    
        Case ROTINA_TRVGERACOMIINT_BATCH
            lErro = objRotGeraComiInt.TRVGerComiInt_Gera(gobjTRVGerComiInt, giRelatorio)
            If lErro <> SUCESSO Or GL_lUltimoErro <> SUCESSO Then gError 129930
        'FIM
        '#################################################

    End Select

    iCancelaBatch = CANCELA_BATCH

    Unload Me

    Exit Sub

Erro_Timer1_Timer:

'    If iCancelaBatch <> CANCELA_BATCH Then
'
'        sErro = "Houve algum tipo de erro. Verifique o arquivo de log de erros configurado em \windows\adm100.ini ."
'        Call MsgBox(sErro, vbOKOnly, "SGE-Forprint")
'
'    End If
    
    Select Case gErr

        Case 15894, 25222, 25223, 55472, 83543, 92670, 129930, 129931, 129787

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174599)

    End Select
    
    If iCancelaBatch <> CANCELA_BATCH Then
        Call Rotina_ErrosBatch2("Importação de Dados do Sisweb")
    End If

    iCancelaBatch = CANCELA_BATCH
    Unload Me

    Exit Sub

End Sub


'Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label2, Source, X, Y)
'End Sub
'
'Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
'End Sub
'
'Private Sub TotReg_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(TotReg, Source, X, Y)
'End Sub
'
'Private Sub TotReg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(TotReg, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label1, Source, X, Y)
'End Sub
'
'Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
'End Sub
'
