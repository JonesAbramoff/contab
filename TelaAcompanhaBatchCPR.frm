VERSION 5.00
Begin VB.Form TelaAcompanhaBatchCPR 
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   4680
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
      Height          =   525
      Left            =   1500
      TabIndex        =   0
      Top             =   1080
      Width           =   1545
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4110
      Top             =   690
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
      Left            =   210
      TabIndex        =   2
      Top             =   360
      Width           =   2985
   End
   Begin VB.Label TotReg 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   3285
      TabIndex        =   1
      Top             =   300
      Width           =   1245
   End
End
Attribute VB_Name = "TelaAcompanhaBatchCPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iCancelaBatch As Integer
Public dValorAtual As Double
Public sNomeArqParam As String

'Rotina Batch que a tela está acompanhando
Public iRotinaBatch As Integer
Public objGeracaoArqICMS As ClassGeracaoArqICMS

Private Sub Cancelar_Click()

Dim lErro As Long

On Error GoTo Erro_Cancelar_Click

    iCancelaBatch = CANCELA_BATCH

    Exit Sub

Erro_Cancelar_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174594)

    End Select

    Exit Sub

End Sub

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    TotReg.Caption = "0"

    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174595)

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
Dim lteste As Long

On Error GoTo Erro_Timer1_Timer

    Timer1.Interval = 0

    lErro = Sistema_Abrir_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then Error 61497

    Set gcolModulo = New AdmColModulo
    
    lErro = CF("Modulos_Le_Empresa_Filial", glEmpresa, giFilialEmpresa, gcolModulo)
    If lErro <> SUCESSO Then Error 61498

    Select Case iRotinaBatch
    
        Case ROTINA_BACH_GERACAO_ARQ_ICMS
''            lErro = objGeracaoArqICMS.Rotina_Gerar_ICMS
''            If lErro <> SUCESSO Then Error 61499
            
    End Select

    iCancelaBatch = CANCELA_BATCH

    Unload Me

    Exit Sub

Erro_Timer1_Timer:

    Select Case Err

        Case 61497, 61498, 61499

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174596)

    End Select

    iCancelaBatch = CANCELA_BATCH
    Unload Me

    Exit Sub

End Sub

Sub Acompanha_Batch_CPR(dValorAtual As Double, iCancelou As Integer)

Dim lErro As Long

    lErro = DoEvents()
    
    TelaAcompanhaBatchCPR.TotReg.Caption = dValorAtual
    iCancelou = iCancelaBatch
    
End Sub

'Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label1, Source, X, Y)
'End Sub
'
'Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
'End Sub
'
'Private Sub TotReg_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(TotReg, Source, X, Y)
'End Sub
'
'Private Sub TotReg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(TotReg, Button, Shift, X, Y)
'End Sub

