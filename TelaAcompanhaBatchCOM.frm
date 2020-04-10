VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form TelaAcompanhaBatchCOM 
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4995
   Icon            =   "TelaAcompanhaBatchCOM.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3660
      Top             =   615
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
      Height          =   510
      Left            =   1815
      TabIndex        =   1
      Top             =   1920
      Width           =   1395
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   405
      Left            =   315
      TabIndex        =   0
      Top             =   1185
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   714
      _Version        =   393216
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
      Left            =   315
      TabIndex        =   2
      Top             =   930
      Width           =   1305
   End
   Begin VB.Label TotReg 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   3390
      TabIndex        =   3
      Top             =   195
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
      Left            =   315
      TabIndex        =   4
      Top             =   255
      Width           =   2985
   End
End
Attribute VB_Name = "TelaAcompanhaBatchCOM"
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
Const ROTINA_REQCOMPRAS_BAIXAR = 2

Public colReqComprasInfo As Collection
Const Rotina_ReqComprasBaixar_Batch = 0

Private Sub Cancelar_Click()

Dim lErro As Long

On Error GoTo Erro_Cancelar_Click

    iCancelaBatch = CANCELA_BATCH

    Exit Sub

Erro_Cancelar_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174591)

    End Select

    Exit Sub

End Sub

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    TotReg.Caption = "0"
    ProgressBar1.Min = 0
    ProgressBar1.Max = 100

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174592)

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

Dim lErro As Long, sErro As String
Dim lteste As Long

On Error GoTo Erro_Timer1_Timer

    Timer1.Interval = 0

    lErro = Sistema_Abrir_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 25222

    Set gcolModulo = New AdmColModulo

    lErro = CF("Modulos_Le_Empresa_Filial", glEmpresa, giFilialEmpresa, gcolModulo)
    If lErro <> SUCESSO Then gError 55472
    
    GL_lUltimoErro = SUCESSO

    Select Case iRotinaBatch
    
        Case ROTINA_REQCOMPRAS_BAIXAR
            lErro = Rotina_ReqComprasBaixar_Batch_Int(colReqComprasInfo)
            If lErro <> SUCESSO Or GL_lUltimoErro <> SUCESSO Then gError 63401
        
        Case ROTINA_CALCULO_PTOPEDIDO
            lErro = ParametrosPtoPed_Calcula()
            If lErro <> SUCESSO Or GL_lUltimoErro <> SUCESSO Then gError 69051
        
    End Select

    iCancelaBatch = CANCELA_BATCH

    Unload Me

    Exit Sub

Erro_Timer1_Timer:

    If iCancelaBatch <> CANCELA_BATCH Then

        sErro = "Houve algum tipo de erro. Verifique o arquivo de log de erros configurado em \windows\adm100.ini ."
        Call MsgBox(sErro, vbOKOnly, "SGE-Forprint")

    End If

    Select Case gErr

        Case 63401
        
        Case 69051
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174593)

    End Select

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
