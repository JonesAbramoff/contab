VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpPCConcorrenciaOcx 
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8985
   ScaleHeight     =   4845
   ScaleWidth      =   8985
   Begin VB.Frame Frame2 
      Caption         =   "Concorrências"
      Height          =   2475
      Left            =   240
      TabIndex        =   26
      Top             =   2160
      Width           =   8535
      Begin VB.CheckBox CheckPedidos 
         Caption         =   "Exibe Pedidos não emitidos"
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
         Left            =   1500
         TabIndex        =   14
         Top             =   2085
         Visible         =   0   'False
         Width           =   2910
      End
      Begin MSMask.MaskEdBox CodConcorrenciaDe 
         Height          =   300
         Left            =   1515
         TabIndex        =   6
         Top             =   255
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CodConcorrenciaAte 
         Height          =   300
         Left            =   5940
         TabIndex        =   7
         Top             =   315
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DescricaoDe 
         Height          =   300
         Left            =   1515
         TabIndex        =   12
         Top             =   690
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DescricaoAte 
         Height          =   300
         Left            =   5940
         TabIndex        =   13
         Top             =   690
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CodCompradorDe 
         Height          =   300
         Left            =   1515
         TabIndex        =   8
         Top             =   1680
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CodCompradorAte 
         Height          =   300
         Left            =   5940
         TabIndex        =   9
         Top             =   1725
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataDe 
         Height          =   315
         Left            =   2685
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1170
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDe 
         Height          =   315
         Left            =   1515
         TabIndex        =   10
         Top             =   1185
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataAte 
         Height          =   315
         Left            =   7125
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1170
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
         Height          =   315
         Left            =   5940
         TabIndex        =   11
         Top             =   1185
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Data Até:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5040
         TabIndex        =   38
         Top             =   1245
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Data De:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   705
         TabIndex        =   37
         Top             =   1245
         Width           =   780
      End
      Begin VB.Label LabelCodCompradorAte 
         AutoSize        =   -1  'True
         Caption         =   "Comprador Até:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4530
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   34
         Top             =   1725
         Width           =   1320
      End
      Begin VB.Label LabelCodCompradorDe 
         AutoSize        =   -1  'True
         Caption         =   "Comprador De:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   225
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   33
         Top             =   1710
         Width           =   1275
      End
      Begin VB.Label LabelDescDe 
         AutoSize        =   -1  'True
         Caption         =   "Descrição De:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   270
         TabIndex        =   30
         Top             =   780
         Width           =   1230
      End
      Begin VB.Label LabelDescAte 
         AutoSize        =   -1  'True
         Caption         =   "Descrição Até:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4575
         TabIndex        =   29
         Top             =   795
         Width           =   1275
      End
      Begin VB.Label LabelCodConcDe 
         AutoSize        =   -1  'True
         Caption         =   "Código De:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   540
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   28
         Top             =   315
         Width           =   960
      End
      Begin VB.Label LabelCodConcAte 
         AutoSize        =   -1  'True
         Caption         =   "Código Até:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4875
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   27
         Top             =   375
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filial Empresa"
      Height          =   1200
      Left            =   240
      TabIndex        =   21
      Top             =   915
      Width           =   6465
      Begin MSMask.MaskEdBox CodFilialDe 
         Height          =   300
         Left            =   1170
         TabIndex        =   2
         Top             =   345
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CodFilialAte 
         Height          =   300
         Left            =   4320
         TabIndex        =   3
         Top             =   330
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NomeDe 
         Height          =   300
         Left            =   1170
         TabIndex        =   4
         Top             =   780
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NomeAte 
         Height          =   300
         Left            =   4320
         TabIndex        =   5
         Top             =   780
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelNomeDe 
         AutoSize        =   -1  'True
         Caption         =   "Nome De:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   225
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   25
         Top             =   840
         Width           =   855
      End
      Begin VB.Label LabelNomeAte 
         AutoSize        =   -1  'True
         Caption         =   "Nome Até:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3360
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   24
         Top             =   825
         Width           =   900
      End
      Begin VB.Label LabelCodFilialDe 
         AutoSize        =   -1  'True
         Caption         =   "Código De:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   390
         Width           =   960
      End
      Begin VB.Label LabelCodFilialAte 
         AutoSize        =   -1  'True
         Caption         =   "Código Até:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3225
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   22
         Top             =   405
         Width           =   1005
      End
   End
   Begin VB.ComboBox ComboOrdenacao 
      Height          =   315
      ItemData        =   "RelOpPCConcorrenciaOcx.ctx":0000
      Left            =   1530
      List            =   "RelOpPCConcorrenciaOcx.ctx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6660
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpPCConcorrenciaOcx.ctx":0039
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpPCConcorrenciaOcx.ctx":0193
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpPCConcorrenciaOcx.ctx":031D
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpPCConcorrenciaOcx.ctx":084F
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   6975
      Picture         =   "RelOpPCConcorrenciaOcx.ctx":09CD
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1095
      Width           =   1635
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpPCConcorrenciaOcx.ctx":0ACF
      Left            =   1530
      List            =   "RelOpPCConcorrenciaOcx.ctx":0AD1
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   75
      Width           =   2595
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ordenados Por:"
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
      Left            =   195
      TabIndex        =   32
      Top             =   570
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Opção:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   210
      TabIndex        =   31
      Top             =   150
      Width           =   615
   End
End
Attribute VB_Name = "RelOpPCConcorrenciaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'RelOpPCConcorrencia
Const ORD_POR_CODIGO = 0
Const ORD_POR_DESCRICAO = 1
Const ORD_POR_DATA = 2
Const ORD_POR_COMPRADOR = 3


Private WithEvents objEventoCodConcDe As AdmEvento
Attribute objEventoCodConcDe.VB_VarHelpID = -1
Private WithEvents objEventoCodConcAte As AdmEvento
Attribute objEventoCodConcAte.VB_VarHelpID = -1
Private WithEvents objEventoCompradorDe As AdmEvento
Attribute objEventoCompradorDe.VB_VarHelpID = -1
Private WithEvents objEventoCompradorAte As AdmEvento
Attribute objEventoCompradorAte.VB_VarHelpID = -1
Private WithEvents objEventoCodFilialDe As AdmEvento
Attribute objEventoCodFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoCodFilialAte As AdmEvento
Attribute objEventoCodFilialAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeFilialDe As AdmEvento
Attribute objEventoNomeFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeFilialAte As AdmEvento
Attribute objEventoNomeFilialAte.VB_VarHelpID = -1

Dim iAlterado As Integer
Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 74579

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 74580

    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 74579

        Case 74580
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170690)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub Limpa_Tela_Rel()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Rel

    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 74581

    ComboOrdenacao.ListIndex = 0
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    CheckPedidos.Value = vbUnchecked
    
    Exit Sub

Erro_Limpa_Tela_Rel:

    Select Case gErr

        Case 74581

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170691)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_Rel

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCodConcDe = New AdmEvento
    Set objEventoCodConcAte = New AdmEvento
    Set objEventoCompradorDe = New AdmEvento
    Set objEventoCompradorAte = New AdmEvento
    Set objEventoCodFilialDe = New AdmEvento
    Set objEventoCodFilialAte = New AdmEvento
    Set objEventoNomeFilialDe = New AdmEvento
    Set objEventoNomeFilialAte = New AdmEvento

    ComboOrdenacao.ListIndex = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170692)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

    Set objEventoCodConcDe = Nothing
    Set objEventoCodConcAte = Nothing
    Set objEventoCompradorDe = Nothing
    Set objEventoCompradorAte = Nothing
    Set objEventoCodFilialDe = Nothing
    Set objEventoCodFilialAte = Nothing
    Set objEventoNomeFilialDe = Nothing
    Set objEventoNomeFilialAte = Nothing

End Sub

Private Sub CodCompradorAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodCompradorAte, iAlterado)
    
End Sub

Private Sub CodCompradorAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objComprador As New ClassComprador

On Error GoTo Erro_CodCompradorAte_Validate

    If Len(Trim(CodCompradorAte.Text)) > 0 Then

        'Lê o código informado
        lErro = CF("TP_Comprador_Le", CodCompradorAte, objComprador, 0)
        If lErro <> SUCESSO Then gError 74682
        
        CodCompradorAte.Text = objComprador.iCodigo
    End If

    Exit Sub

Erro_CodCompradorAte_Validate:

    Cancel = True

    Select Case gErr

        Case 74682

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170693)

    End Select

    Exit Sub

End Sub

Private Sub CodCompradorDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodCompradorDe, iAlterado)
    
End Sub

Private Sub CodCompradorDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objComprador As New ClassComprador

On Error GoTo Erro_CodCompradorDe_Validate

    If Len(Trim(CodCompradorDe.Text)) > 0 Then

        lErro = CF("TP_Comprador_Le", CodCompradorDe, objComprador, 0)
        If lErro <> SUCESSO Then gError 74681
        
        CodCompradorDe.Text = objComprador.iCodigo
        
    End If

    Exit Sub

Erro_CodCompradorDe_Validate:

    Cancel = True

    Select Case gErr

        Case 74681

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170694)

    End Select

    Exit Sub

End Sub

Private Sub CodConcorrenciaAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodConcorrenciaAte, iAlterado)
    
End Sub

Private Sub CodConcorrenciaDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodConcorrenciaDe, iAlterado)
    
End Sub

Private Sub CodFilialAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodFilialAte, iAlterado)
    
End Sub

Private Sub CodFilialDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodFilialDe, iAlterado)
    
End Sub

Private Sub DataAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)
    
End Sub

Private Sub DataDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)
    
End Sub


Private Sub LabelCodConcAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objConcorrencia As New ClassConcorrencia

On Error GoTo Erro_LabelCodConcAte_Click

    If Len(Trim(CodConcorrenciaAte.Text)) > 0 Then
        'Preenche com a Concorrencia da tela
        objConcorrencia.lCodigo = StrParaLong(CodConcorrenciaAte.Text)
    End If

    'Chama Tela ConcorrenciaLista
    Call Chama_Tela("ConcorrenciaLista", colSelecao, objConcorrencia, objEventoCodConcAte)

   Exit Sub

Erro_LabelCodConcAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170695)

    End Select

    Exit Sub

End Sub
Private Sub LabelCodConcDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objConcorrencia As New ClassConcorrencia

On Error GoTo Erro_LabelCodConcDe_Click

    If Len(Trim(CodConcorrenciaDe.Text)) > 0 Then
        'Preenche com o Pedido de Compra da tela
        objConcorrencia.lCodigo = StrParaLong(CodConcorrenciaDe.Text)
    End If

    'Chama Tela ConcorrenciaLista
    Call Chama_Tela("ConcorrenciaLista", colSelecao, objConcorrencia, objEventoCodConcDe)

   Exit Sub

Erro_LabelCodConcDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170696)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataDe.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then gError 74582

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 74582
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170697)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataAte.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError 74583

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 74583
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170698)

    End Select

    Exit Sub

End Sub




Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 74584

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 74584
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170699)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 74585

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 74585
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170700)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 74586

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 74586
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170701)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 74587

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 74587
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170702)

    End Select

    Exit Sub

End Sub


Private Sub LabelCodFilialDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelCodFilialDe_Click

    If Len(Trim(CodFilialDe.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.iCodFilial = StrParaInt(CodFilialDe.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoCodFilialDe)

   Exit Sub

Erro_LabelCodFilialDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170703)

    End Select

    Exit Sub

End Sub
Private Sub LabelCodFilialAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelCodFilialAte_Click

    If Len(Trim(CodFilialAte.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.iCodFilial = StrParaInt(CodFilialAte.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoCodFilialAte)

   Exit Sub

Erro_LabelCodFilialAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170704)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodCompradorAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objComprador As New ClassComprador

On Error GoTo Erro_LabelCodCompradorAte_Click

    If Len(Trim(CodCompradorAte.Text)) > 0 Then
        'Preenche com o comprador da tela
        objComprador.iCodigo = StrParaInt(CodCompradorAte.Text)
    End If

    'Chama Tela CompradoresLista
    Call Chama_Tela("CompradoresLista", colSelecao, objComprador, objEventoCompradorAte)

   Exit Sub

Erro_LabelCodCompradorAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170705)

    End Select

    Exit Sub

End Sub
Private Sub LabelCodCompradorDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objComprador As New ClassComprador

On Error GoTo Erro_LabelCodCompradorDe_Click

    If Len(Trim(CodCompradorDe.Text)) > 0 Then
        'Preenche com o comprador da tela
        objComprador.iCodigo = StrParaInt(CodCompradorDe.Text)
    End If

    'Chama Tela CompradoresLista
    Call Chama_Tela("CompradoresLista", colSelecao, objComprador, objEventoCompradorDe)

   Exit Sub

Erro_LabelCodCompradorDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170706)

    End Select

    Exit Sub

End Sub


Private Sub LabelNomeDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelNomeDe_Click

    If Len(Trim(NomeDe.Text)) > 0 Then
        'Preenche com o requisitante da tela
        objFilialEmpresa.sNome = NomeDe.Text
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoNomeFilialDe)

   Exit Sub

Erro_LabelNomeDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170707)

    End Select

    Exit Sub

End Sub

Private Sub LabelNomeAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelNomeAte_Click

    If Len(Trim(NomeAte.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.sNome = NomeAte.Text
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoNomeFilialAte)

   Exit Sub

Erro_LabelNomeAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170708)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodFilialAte_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    CodFilialAte.Text = CStr(objFilialEmpresa.iCodFilial)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeFilialDe_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    NomeDe.Text = objFilialEmpresa.sNome

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeFilialAte_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    NomeAte.Text = objFilialEmpresa.sNome

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCodFilialDe_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    CodFilialDe.Text = CStr(objFilialEmpresa.iCodFilial)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCodConcAte_evSelecao(obj1 As Object)

Dim objConcorrencia As New ClassConcorrencia

    Set objConcorrencia = obj1

    CodConcorrenciaAte.Text = CStr(objConcorrencia.lCodigo)

    Me.Show

End Sub

Private Sub objEventoCodConcDe_evSelecao(obj1 As Object)

Dim objConcorrencia As New ClassConcorrencia

    Set objConcorrencia = obj1

    CodConcorrenciaDe.Text = CStr(objConcorrencia.lCodigo)


    Me.Show

End Sub

Private Sub objEventoCompradorDe_evSelecao(obj1 As Object)

Dim objComprador As New ClassComprador

    Set objComprador = obj1

    CodCompradorDe.Text = CStr(objComprador.iCodigo)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCompradorAte_evSelecao(obj1 As Object)

Dim objComprador As New ClassComprador

    Set objComprador = obj1

    CodCompradorAte.Text = CStr(objComprador.iCodigo)

    Me.Show

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 74588

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 74589

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 74590

    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 74591

    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 74588
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 74589 To 74591

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170709)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 74592

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 74593

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 74592
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 74593

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170710)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 74594

    Select Case ComboOrdenacao.ListIndex

            Case ORD_POR_CODIGO
                
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "CodConcorrencia", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "PCCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemPedCompra", 1)
                
            
            Case ORD_POR_DESCRICAO

                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "DescConcorrencia", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "CodConcorrencia", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "PCCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemPedCompra", 1)
                
            Case ORD_POR_DATA
                
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "DataConcorrencia", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "CodConcorrencia", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "PCCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemPedCompra", 1)

            Case ORD_POR_COMPRADOR
                
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "CompradorCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "CodConcorrencia", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "PCCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemPedCompra", 1)
                
            Case Else
                gError 74954

    End Select

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 74594, 74954

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170711)

    End Select

    Exit Sub

End Sub


Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sCodFilial_I As String
Dim sCodFilial_F As String
Dim sNomeFilial_I As String
Dim sNomeFilial_F As String
Dim sDesc_I As String
Dim sDesc_F As String
Dim sCodConc_I As String
Dim sCodConc_F As String
Dim sCodComprador_I As String
Dim sCodComprador_F As String
Dim sCheck As String
Dim sOrdenacaoPor As String
Dim iOrdenacao As Long
Dim sOrd As String

On Error GoTo Erro_PreencherRelOp

    lErro = Formata_E_Critica_Parametros(sCodFilial_I, sCodFilial_F, sNomeFilial_I, sNomeFilial_F, sCodConc_I, sCodConc_F, sDesc_I, sDesc_F, sCodComprador_I, sCodComprador_F)
    If lErro <> SUCESSO Then gError 74595

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 74596

    lErro = objRelOpcoes.IncluirParametro("NCODFILIALINIC", sCodFilial_I)
    If lErro <> AD_BOOL_TRUE Then gError 74597

    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALINIC", NomeDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 74598

    lErro = objRelOpcoes.IncluirParametro("NCODCONCINIC", sCodConc_I)
    If lErro <> AD_BOOL_TRUE Then gError 74599

    lErro = objRelOpcoes.IncluirParametro("TDESCCONCINIC", DescricaoDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 74600

    lErro = objRelOpcoes.IncluirParametro("NCODCOMPINIC", sCodComprador_I)
    If lErro <> AD_BOOL_TRUE Then gError 74601

    'Preenche data inicial
    If Trim(DataDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATACONCINIC", DataDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATACONCINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 74602

    lErro = objRelOpcoes.IncluirParametro("NCODFILIALFIM", sCodFilial_F)
    If lErro <> AD_BOOL_TRUE Then gError 74603

    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALFIM", NomeAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 74604

    lErro = objRelOpcoes.IncluirParametro("NCODCONCFIM", sCodConc_F)
    If lErro <> AD_BOOL_TRUE Then gError 74605

    lErro = objRelOpcoes.IncluirParametro("TDESCCONCFIM", DescricaoAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 74606

    lErro = objRelOpcoes.IncluirParametro("NCODCOMPFIM", sCodComprador_F)
    If lErro <> AD_BOOL_TRUE Then gError 74607

    'Preenche data final
    If Trim(DataAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATACONCFIM", DataAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATACONCFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 74608

    'Exibe Pedidos Emitidos
    If CheckPedidos.Value Then
        sCheck = vbChecked
    Else
        sCheck = vbUnchecked
    End If

    lErro = objRelOpcoes.IncluirParametro("NPEDEMITIDO", sCheck)
    If lErro <> AD_BOOL_TRUE Then gError 74814
    
    Select Case ComboOrdenacao.ListIndex

            Case ORD_POR_CODIGO

                sOrdenacaoPor = "Codigo"

            Case ORD_POR_DESCRICAO

                sOrdenacaoPor = "Descricao"

            Case ORD_POR_DATA
                sOrdenacaoPor = "Data"

            Case ORD_POR_COMPRADOR
                sOrdenacaoPor = "Comprador"
                
            Case Else
                gError 74609

    End Select

    lErro = objRelOpcoes.IncluirParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> AD_BOOL_TRUE Then gError 74610

    sOrd = ComboOrdenacao.ListIndex
    lErro = objRelOpcoes.IncluirParametro("NORDENACAO", sOrd)
    If lErro <> AD_BOOL_TRUE Then gError 74611

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCodFilial_I, sCodFilial_F, sNomeFilial_I, sNomeFilial_F, sCodConc_I, sCodConc_F, sDesc_I, sDesc_F, sCodComprador_I, sCodComprador_F, sOrdenacaoPor, sOrd)
    If lErro <> SUCESSO Then gError 74612

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 74595 To 74612, 74814

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170712)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCodFilial_I As String, sCodFilial_F As String, sNomeFilial_I As String, sNomeFilial_F As String, sCodConc_I As String, sCodConc_F As String, sDesc_I As String, sDesc_F As String, sCodComprador_I As String, sCodComprador_F As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'critica Codigo da Filial Inicial e Final
    If CodFilialDe.Text <> "" Then
        sCodFilial_I = CStr(CodFilialDe.Text)
    Else
        sCodFilial_I = ""
    End If

    If CodFilialAte.Text <> "" Then
        sCodFilial_F = CStr(CodFilialAte.Text)
    Else
        sCodFilial_F = ""
    End If

    If sCodFilial_I <> "" And sCodFilial_F <> "" Then

        If StrParaLong(sCodFilial_I) > StrParaLong(sCodFilial_F) Then gError 74613

    End If

    If NomeDe.Text <> "" Then
        sNomeFilial_I = NomeDe.Text
    Else
        sNomeFilial_I = ""
    End If

    If NomeAte.Text <> "" Then
        sNomeFilial_F = NomeAte.Text
    Else
        sNomeFilial_F = ""
    End If

    If sNomeFilial_I <> "" And sNomeFilial_F <> "" Then
        If sNomeFilial_I > sNomeFilial_F Then gError 74614
    End If

    'critica CodigoConc Inicial e Final
    If CodConcorrenciaDe.Text <> "" Then
        sCodConc_I = CStr(CodConcorrenciaDe.Text)
    Else
        sCodConc_I = ""
    End If

    If CodConcorrenciaAte.Text <> "" Then
        sCodConc_F = CStr(CodConcorrenciaAte.Text)
    Else
        sCodConc_F = ""
    End If

    If sCodConc_I <> "" And sCodConc_F <> "" Then

        If StrParaLong(sCodConc_I) > StrParaLong(sCodConc_F) Then gError 74615

    End If

    'data inicial não pode ser maior que a final
    If Trim(DataDe.ClipText) <> "" And Trim(DataAte.ClipText) <> "" Then
    
         If CDate(DataDe.Text) > CDate(DataAte.Text) Then gError 74616
    
    End If
    
    'critica Comprador Inicial e Final
    If CodCompradorDe.Text <> "" Then
        sCodComprador_I = CStr(CodCompradorDe.Text)
    Else
        sCodComprador_I = ""
    End If

    If CodCompradorAte.Text <> "" Then
        sCodComprador_F = CStr(CodCompradorAte.Text)
    Else
        sCodComprador_F = ""
    End If

    If sCodComprador_I <> "" And sCodComprador_F <> "" Then

        If StrParaInt(sCodComprador_I) > StrParaInt(sCodComprador_F) Then gError 74617

    End If

    'critica Descricao Inicial e Final
    If DescricaoDe.Text <> "" Then
        sDesc_I = CStr(DescricaoDe.Text)
    Else
        sDesc_I = ""
    End If

    If DescricaoAte.Text <> "" Then
        sDesc_F = CStr(DescricaoAte.Text)
    Else
        sDesc_F = ""
    End If

    If sDesc_I <> "" And sDesc_F <> "" Then

        If sDesc_I > sDesc_F Then gError 74619

    End If


    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 74613
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            CodFilialDe.SetFocus

        Case 74614
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            NomeDe.SetFocus

        Case 74615
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PC_INICIAL_MAIOR", gErr)
            CodConcorrenciaDe.SetFocus

        Case 74616
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus
            
        Case 74617
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_INICIAL_MAIOR", gErr)
            CodCompradorDe.SetFocus

        Case 74619
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_INICIAL_MAIOR", gErr)
            DescricaoDe.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170713)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCodFilial_I As String, sCodFilial_F As String, sNomeFilial_I As String, sNomeFilial_F As String, sCodConc_I As String, sCodConc_F As String, sDesc_I As String, sDesc_F As String, sCodComprador_I As String, sCodComprador_F As String, sOrdenacaoPor As String, sOrd As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao


   If sCodFilial_I <> "" Then sExpressao = "FilEmpCodInic"

   If sCodFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpCodFim"

    End If

   If sNomeFilial_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpNomeInic"

    End If

    If sNomeFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpNomeFim"

    End If

    If sCodConc_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CodConc >= " & Forprint_ConvLong(StrParaLong(sCodConc_I))

    End If

    If sCodConc_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CodConc <= " & Forprint_ConvLong(StrParaLong(sCodConc_F))

    End If
    
    If Trim(DataDe.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DataConc >= " & Forprint_ConvData(CDate(DataDe.Text))

    End If

    If Trim(DataAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DataConc <= " & Forprint_ConvData(CDate(DataAte.Text))

    End If

    If sDesc_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DescConc >= " & Forprint_ConvTexto((sDesc_I))

    End If

    If sDesc_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DescConc <= " & Forprint_ConvTexto((sDesc_F))

    End If

    If sCodComprador_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CompCod >= " & Forprint_ConvInt(StrParaInt(sCodComprador_I))

    End If

    If sCodComprador_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CompCod <= " & Forprint_ConvInt(StrParaInt(sCodComprador_F))

    End If

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170714)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim sOrdenacaoPor As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 74620

    'pega Codigo Filial inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 74621

    CodFilialDe.Text = sParam
    Call CodFilialDe_Validate(bSGECancelDummy)

    'pega  Codigo Filial final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 74622

    CodFilialAte.Text = sParam
    Call CodFilialAte_Validate(bSGECancelDummy)

    'pega  Nome Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 74623

    NomeDe.Text = sParam
    Call NomeDe_Validate(bSGECancelDummy)

    'pega  Nome Final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 74624

    NomeAte.Text = sParam
    Call NomeAte_Validate(bSGECancelDummy)

    'pega  Codigo Conc inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODCONCINIC", sParam)
    If lErro <> SUCESSO Then gError 74625

    CodConcorrenciaDe.Text = sParam

    'pega  Codigo Conc final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODCONCFIM", sParam)
    If lErro <> SUCESSO Then gError 74626

    CodConcorrenciaAte.Text = sParam

    'pega Comprador Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODCOMPINIC", sParam)
    If lErro <> SUCESSO Then gError 74627

    CodCompradorDe.Text = sParam

    'pega Comprador Final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODCOMPFIM", sParam)
    If lErro <> SUCESSO Then gError 74628

    CodCompradorAte.Text = sParam

    'pega Descricao Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TDESCCONCINIC", sParam)
    If lErro <> SUCESSO Then gError 74629

    DescricaoDe.Text = sParam

    'pega Descricao Final e exibe
    lErro = objRelOpcoes.ObterParametro("TDESCCONCFIM", sParam)
    If lErro <> SUCESSO Then gError 74630

    DescricaoAte.Text = sParam
    
    'pega data  inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDATACONCINIC", sParam)
    If lErro <> SUCESSO Then gError 74631

    Call DateParaMasked(DataDe, CDate(sParam))

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DDATACONCFIM", sParam)
    If lErro <> SUCESSO Then gError 74632

    Call DateParaMasked(DataAte, CDate(sParam))

    lErro = objRelOpcoes.ObterParametro("NPEDEMITIDO", sParam)
    If lErro <> SUCESSO Then gError 74813

    If sParam = "1" Then
        CheckPedidos.Value = vbChecked
    Else
        CheckPedidos.Value = vbUnchecked
    End If
   
    lErro = objRelOpcoes.ObterParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> SUCESSO Then gError 74633

    Select Case sOrdenacaoPor

            Case "Codigo"

                ComboOrdenacao.ListIndex = ORD_POR_CODIGO

            Case "Descricao"
                
                ComboOrdenacao.ListIndex = ORD_POR_DESCRICAO

            Case "Data"
                
                ComboOrdenacao.ListIndex = ORD_POR_DATA

            Case "Comprador"
            
                ComboOrdenacao.ListIndex = ORD_POR_COMPRADOR
                
            Case Else
                gError 74634

    End Select

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 74620 To 74634, 74813

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170715)

    End Select

    Exit Function

End Function


Private Sub CodFilialDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_CodFilialDe_Validate

    If Len(Trim(CodFilialDe.Text)) > 0 Then

        objFilialEmpresa.iCodFilial = StrParaInt(CodFilialDe.Text)
        'Lê o código informado
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 74635

        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 74636

    End If

    Exit Sub

Erro_CodFilialDe_Validate:

    Cancel = True


    Select Case gErr

        Case 74635

        Case 74636
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170716)

    End Select

    Exit Sub

End Sub
Private Sub CodFilialAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_CodFilialAte_Validate

    If Len(Trim(CodFilialAte.Text)) > 0 Then

        objFilialEmpresa.iCodFilial = StrParaInt(CodFilialAte.Text)
        'Lê o código informado
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 74637

        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 74638

    End If

    Exit Sub

Erro_CodFilialAte_Validate:

    Cancel = True


    Select Case gErr

        Case 74637

        Case 74638
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170717)

    End Select

    Exit Sub

End Sub

Private Sub NomeDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim bAchou As Boolean
Dim colFiliais As New Collection

On Error GoTo Erro_NomeDe_Validate

    bAchou = False

    If Len(Trim(NomeDe.Text)) > 0 Then

        lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
        If lErro <> SUCESSO Then gError 74639

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = UCase(NomeDe.Text) Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 74640

        NomeDe.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeDe_Validate:

    Cancel = True

    Select Case gErr

        Case 74639

        Case 74640
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeDe.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170718)

    End Select

Exit Sub

End Sub

Private Sub NomeAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim bAchou As Boolean
Dim colFiliais As New Collection

On Error GoTo Erro_NomeAte_Validate

    bAchou = False
    If Len(Trim(NomeAte.Text)) > 0 Then

        lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
        If lErro <> SUCESSO Then gError 74641

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = UCase(NomeAte.Text) Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 74642

        NomeAte.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeAte_Validate:

    Cancel = True


    Select Case gErr

        Case 74641

        Case 74642
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeAte.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170719)

    End Select

Exit Sub

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)

End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

''    Parent.HelpContextID = IDH_RELOP_REQ
    Set Form_Load_Ocx = Me
    Caption = "Pedidos de Compra Emitidos por Concorrência"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpPCConcorrencia"

End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
    Height = UserControl.Height
End Property

Public Property Get Width() As Long
    Width = UserControl.Width
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ActiveControl
Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Sub Unload(objme As Object)

   RaiseEvent Unload

End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is CodConcorrenciaDe Then
            Call LabelCodConcDe_Click

        ElseIf Me.ActiveControl Is CodConcorrenciaAte Then
            Call LabelCodConcAte_Click

        ElseIf Me.ActiveControl Is CodFilialDe Then
            Call LabelCodFilialDe_Click

        ElseIf Me.ActiveControl Is CodFilialAte Then
            Call LabelCodFilialAte_Click

        ElseIf Me.ActiveControl Is NomeDe Then
            Call LabelNomeDe_Click

        ElseIf Me.ActiveControl Is NomeAte Then
            Call LabelNomeAte_Click

        ElseIf Me.ActiveControl Is CodCompradorDe Then
            Call LabelCodCompradorDe_Click

        ElseIf Me.ActiveControl Is CodCompradorAte Then
            Call LabelCodCompradorAte_Click

        End If

    End If

End Sub


Private Sub LabelCodFilialDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodFilialDe, Source, X, Y)
End Sub

Private Sub LabelCodFilialDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodFilialDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodFilialAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodFilialAte, Source, X, Y)
End Sub

Private Sub LabelCodFilialAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodFilialAte, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub LabelCodCompradorAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodCompradorAte, Source, X, Y)
End Sub

Private Sub LabelCodCompradorAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodCompradorAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodCompradorDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodCompradorDe, Source, X, Y)
End Sub

Private Sub LabelCodCompradorDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodCompradorDe, Button, Shift, X, Y)
End Sub

Private Sub LabelDescDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDescDe, Source, X, Y)
End Sub

Private Sub LabelDescDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDescDe, Button, Shift, X, Y)
End Sub

Private Sub LabelDescAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDescAte, Source, X, Y)
End Sub

Private Sub LabelDescAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDescAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodConcDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodConcDe, Source, X, Y)
End Sub

Private Sub LabelCodConcDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodConcDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodConcAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodConcAte, Source, X, Y)
End Sub

Private Sub LabelCodConcAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodConcAte, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeDe, Source, X, Y)
End Sub

Private Sub LabelNomeDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeAte, Source, X, Y)
End Sub

Private Sub LabelNomeAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeAte, Button, Shift, X, Y)
End Sub

