VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form TelaAcompanhaBatch 
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   Icon            =   "TelaAcompanhaBatchInt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Log 
      Height          =   1665
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   3450
      Width           =   7815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Atualiza��o"
      Height          =   1140
      Left            =   225
      TabIndex        =   6
      Top             =   1920
      Width           =   7845
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   495
         Left            =   315
         TabIndex        =   7
         Top             =   495
         Width           =   7290
         _ExtentX        =   12859
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label NomeArqAtu2 
         Alignment       =   1  'Right Justify
         Height          =   195
         Left            =   6360
         TabIndex        =   15
         Top             =   225
         Width           =   1200
      End
      Begin VB.Label NomeArqAtu 
         Height          =   195
         Left            =   1770
         TabIndex        =   14
         Top             =   225
         Width           =   5835
      End
      Begin VB.Label LabelEtapaAtu 
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
         TabIndex        =   8
         Top             =   225
         Width           =   1305
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Importa��o"
      Height          =   1140
      Left            =   225
      TabIndex        =   3
      Top             =   660
      Width           =   7845
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   300
         TabIndex        =   4
         Top             =   495
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label NomeArqImp2 
         Alignment       =   1  'Right Justify
         Height          =   195
         Left            =   6585
         TabIndex        =   16
         Top             =   225
         Width           =   1005
      End
      Begin VB.Label NomeArqImp 
         Height          =   195
         Left            =   1785
         TabIndex        =   13
         Top             =   225
         Width           =   5745
      End
      Begin VB.Label LabelEtapaImp 
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
         TabIndex        =   5
         Top             =   225
         Width           =   1305
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4815
      Top             =   150
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
      Left            =   3420
      TabIndex        =   0
      Top             =   5235
      Width           =   1395
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "N�mero de Registros Atualizados:"
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
      Left            =   4215
      TabIndex        =   12
      Top             =   255
      Width           =   2880
   End
   Begin VB.Label TotReg 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   7140
      TabIndex        =   11
      Top             =   165
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   3195
      Width           =   390
   End
   Begin VB.Label TotArq 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   3120
      TabIndex        =   1
      Top             =   165
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "N�mero de Arquivos Importados:"
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
      Left            =   270
      TabIndex        =   2
      Top             =   240
      Width           =   2805
   End
End
Attribute VB_Name = "TelaAcompanhaBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Rotina Batch que a tela est� acompanhando
Public iRotinaBatch As Integer

Public iCancelaBatch As Integer
Public dValorTotalImp As Double
Public dValorAtualImp As Double
Public dValorTotalAtu As Double
Public dValorAtualAtu As Double
Public sNomeArqParam As String

Public objExportacaoAux As ClassArqExportacaoAux
Public objImportacaoAux As ClassArqImportacaoAux

Private Sub Cancelar_Click()

Dim lErro As Long

On Error GoTo Erro_Cancelar_Click

    iCancelaBatch = CANCELA_BATCH

    Exit Sub

Erro_Cancelar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189873)

    End Select

    Exit Sub

End Sub

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    TotArq.Caption = "0"
    TotReg.Caption = "0"
    ProgressBar1.Min = 0
    ProgressBar1.Max = 100
    dValorAtualImp = 0
    dValorAtualAtu = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189874)

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
    Set objExportacaoAux = Nothing
    Set objImportacaoAux = Nothing
End Sub

Private Sub Timer1_Timer()

Dim lErro As Long
Dim objIntegracao As New ClassIntegracao
Dim sErro As String

On Error GoTo Erro_Timer1_Timer

    Timer1.Interval = 0

'*** Para depurar, usando o BatchEst como .dll, o trecho abaixo deve estar comentado
    lErro = Sistema_Abrir_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 189875
'***

    Set gcolModulo = New AdmColModulo
    
    iCancelaBatch = 0
    
    lErro = CF("Modulos_Le_Empresa_Filial", glEmpresa, giFilialEmpresa, gcolModulo)
    If lErro <> SUCESSO Then gError 189876
    
    lErro = CF("Retorna_ColFiliais")
    If lErro <> SUCESSO Then gError 189877

    GL_lUltimoErro = SUCESSO
    
    Select Case iRotinaBatch
    
        Case ROTINA_IMPORTACAO_DADOS
    
            lErro = objIntegracao.Importa_Dados(objImportacaoAux)
            If lErro <> SUCESSO Or GL_lUltimoErro <> SUCESSO Then gError 189878

        Case ROTINA_EXPORTACAO_DADOS

            lErro = objIntegracao.Exporta_Dados(objExportacaoAux)
            If lErro <> SUCESSO Or GL_lUltimoErro <> SUCESSO Then gError 189879

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

        Case 189875 To 189879

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189880)

    End Select

    If iCancelaBatch <> CANCELA_BATCH Then
        Select Case iRotinaBatch
            Case ROTINA_IMPORTACAO_DADOS
                Call Rotina_ErrosBatch2("Importa��o de dados")
            Case ROTINA_EXPORTACAO_DADOS
                Call Rotina_ErrosBatch2("Exporta��o de dados")
        End Select
    End If

    iCancelaBatch = CANCELA_BATCH
    Unload Me

    Exit Sub

End Sub

'Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label1, Source, X, Y)
'End Sub
'
'Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
'End Sub
