VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpPVReqComprasOcx 
   ClientHeight    =   4290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8100
   ScaleHeight     =   4290
   ScaleWidth      =   8100
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpPVReqComprasOcx.ctx":0000
      Left            =   930
      List            =   "RelOpPVReqComprasOcx.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   270
      Width           =   2805
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pedidos de Venda"
      Height          =   2160
      Left            =   135
      TabIndex        =   30
      Top             =   870
      Width           =   7065
      Begin VB.CheckBox CheckPVFaturados 
         Caption         =   "Inclui Pedidos de Venda Faturados"
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
         Left            =   165
         TabIndex        =   10
         Top             =   1800
         Visible         =   0   'False
         Width           =   4080
      End
      Begin VB.Frame Frame5 
         Caption         =   "Data de Emissão"
         Height          =   720
         Left            =   3105
         TabIndex        =   40
         Top             =   1065
         Width           =   3870
         Begin MSComCtl2.UpDown UpDownDataEmissaoDe 
            Height          =   315
            Left            =   1635
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   285
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmissaoDe 
            Height          =   315
            Left            =   450
            TabIndex        =   8
            Top             =   285
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataEmissaoAte 
            Height          =   315
            Left            =   3540
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   300
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmissaoAte 
            Height          =   315
            Left            =   2355
            TabIndex        =   9
            Top             =   300
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelDataAte 
            AutoSize        =   -1  'True
            Caption         =   "Até:"
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
            Left            =   1980
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   42
            Top             =   360
            Width           =   360
         End
         Begin VB.Label LabelDataDe 
            AutoSize        =   -1  'True
            Caption         =   "De:"
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
            Left            =   120
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   41
            Top             =   345
            Width           =   315
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Filial Empresa"
         Height          =   705
         Left            =   180
         TabIndex        =   37
         Top             =   285
         Width           =   2805
         Begin MSMask.MaskEdBox FilialEmpresaDe 
            Height          =   300
            Left            =   525
            TabIndex        =   2
            Top             =   270
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialEmpresaAte 
            Height          =   300
            Left            =   1935
            TabIndex        =   3
            Top             =   270
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label LabelEmpresaAte 
            AutoSize        =   -1  'True
            Caption         =   "Até:"
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
            Left            =   1485
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   39
            Top             =   330
            Width           =   360
         End
         Begin VB.Label LabelEmpresaDe 
            AutoSize        =   -1  'True
            Caption         =   "De:"
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
            Left            =   210
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   38
            Top             =   330
            Width           =   315
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Pedido de Venda"
         Height          =   720
         Left            =   180
         TabIndex        =   34
         Top             =   1065
         Width           =   2805
         Begin MSMask.MaskEdBox CodigoPVDe 
            Height          =   300
            Left            =   510
            TabIndex        =   6
            Top             =   285
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodigoPVAte 
            Height          =   300
            Left            =   1935
            TabIndex        =   7
            Top             =   285
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin VB.Label LabelCodigoDe 
            AutoSize        =   -1  'True
            Caption         =   "De:"
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
            TabIndex        =   36
            Top             =   345
            Width           =   315
         End
         Begin VB.Label LabelCodigoAte 
            AutoSize        =   -1  'True
            Caption         =   "Até:"
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
            Left            =   1515
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   35
            Top             =   345
            Width           =   360
         End
      End
      Begin VB.Frame FrameCcl 
         Caption         =   "Filial Faturamento"
         Height          =   690
         Left            =   3105
         TabIndex        =   31
         Top             =   285
         Width           =   3840
         Begin MSMask.MaskEdBox FilialFatAte 
            Height          =   300
            Left            =   2340
            TabIndex        =   5
            Top             =   285
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialFatDe 
            Height          =   300
            Left            =   465
            TabIndex        =   4
            Top             =   285
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label LabelFilialFatAte 
            AutoSize        =   -1  'True
            Caption         =   "Até:"
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
            Left            =   1965
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   33
            Top             =   345
            Width           =   360
         End
         Begin VB.Label LabelFilialFatDe 
            AutoSize        =   -1  'True
            Caption         =   "De:"
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
            Left            =   120
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   32
            Top             =   345
            Width           =   315
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Clientes"
      Height          =   1020
      Left            =   165
      TabIndex        =   23
      Top             =   3165
      Width           =   7770
      Begin VB.Frame FrameCodigo 
         Caption         =   "Código"
         Height          =   660
         Left            =   105
         TabIndex        =   27
         Top             =   240
         Width           =   2475
         Begin MSMask.MaskEdBox CodClienteDe 
            Height          =   300
            Left            =   480
            TabIndex        =   11
            Top             =   240
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodClienteAte 
            Height          =   300
            Left            =   1665
            TabIndex        =   12
            Top             =   240
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin VB.Label LabelClienteAte 
            AutoSize        =   -1  'True
            Caption         =   "Até:"
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
            Left            =   1290
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   29
            Top             =   300
            Width           =   360
         End
         Begin VB.Label LabelClienteDe 
            AutoSize        =   -1  'True
            Caption         =   "De:"
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
            Left            =   120
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   28
            Top             =   300
            Width           =   315
         End
      End
      Begin VB.Frame FrameNome 
         Caption         =   "Nome"
         Height          =   675
         Left            =   2670
         TabIndex        =   24
         Top             =   240
         Width           =   4980
         Begin MSMask.MaskEdBox NomeDe 
            Height          =   300
            Left            =   540
            TabIndex        =   13
            Top             =   210
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NomeAte 
            Height          =   300
            Left            =   3000
            TabIndex        =   14
            Top             =   210
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin VB.Label LabelNomeDe 
            AutoSize        =   -1  'True
            Caption         =   "De:"
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
            Left            =   165
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   26
            Top             =   270
            Width           =   315
         End
         Begin VB.Label LabelNomeAte 
            AutoSize        =   -1  'True
            Caption         =   "Até:"
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
            Left            =   2565
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   25
            Top             =   270
            Width           =   360
         End
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
      Left            =   3945
      Picture         =   "RelOpPVReqComprasOcx.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   120
      Width           =   1635
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5835
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   150
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpPVReqComprasOcx.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpPVReqComprasOcx.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpPVReqComprasOcx.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpPVReqComprasOcx.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOrdenacao 
      Height          =   315
      ItemData        =   "RelOpPVReqComprasOcx.ctx":0A9A
      Left            =   7320
      List            =   "RelOpPVReqComprasOcx.ctx":0AA7
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1935
      Visible         =   0   'False
      Width           =   1470
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
      Left            =   7290
      TabIndex        =   22
      Top             =   1620
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
      Left            =   225
      TabIndex        =   21
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpPVReqComprasOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'RelOpPVReqCompras
Const ORD_POR_CODIGO = 0
Const ORD_POR_EMISSAO = 1
Const ORD_POR_CLIENTE = 2

Private WithEvents objEventoCodigoPVDe As AdmEvento
Attribute objEventoCodigoPVDe.VB_VarHelpID = -1
Private WithEvents objEventoCodigoPVAte As AdmEvento
Attribute objEventoCodigoPVAte.VB_VarHelpID = -1
Private WithEvents objEventoCodClienteDe As AdmEvento
Attribute objEventoCodClienteDe.VB_VarHelpID = -1
Private WithEvents objEventoCodClienteAte As AdmEvento
Attribute objEventoCodClienteAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeDe As AdmEvento
Attribute objEventoNomeDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeAte As AdmEvento
Attribute objEventoNomeAte.VB_VarHelpID = -1
Private WithEvents objEventoFilialEmpresaDe As AdmEvento
Attribute objEventoFilialEmpresaDe.VB_VarHelpID = -1
Private WithEvents objEventoFilialEmpresaAte As AdmEvento
Attribute objEventoFilialEmpresaAte.VB_VarHelpID = -1
Private WithEvents objEventoFilialFatDe As AdmEvento
Attribute objEventoFilialFatDe.VB_VarHelpID = -1
Private WithEvents objEventoFilialFatAte As AdmEvento
Attribute objEventoFilialFatAte.VB_VarHelpID = -1

Dim iAlterado As Integer
Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 68806
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 68807

    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 68806
        
        Case 68807
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171944)

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
    If lErro <> SUCESSO Then gError 68808
    
    ComboOrdenacao.ListIndex = 0
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    CheckPVFaturados.Value = vbUnchecked
    
    Exit Sub
    
Erro_Limpa_Tela_Rel:
    
    Select Case gErr
    
        Case 68808
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171945)

    End Select

    Exit Sub
   
End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_Rel

End Sub


Public Sub Form_Load()

Dim lErro As Long
Dim sMascaraCcl As String

On Error GoTo Erro_Form_Load
    
        
    Set objEventoCodigoPVDe = New AdmEvento
    Set objEventoCodigoPVAte = New AdmEvento
    Set objEventoCodClienteDe = New AdmEvento
    Set objEventoCodClienteAte = New AdmEvento
    Set objEventoNomeDe = New AdmEvento
    Set objEventoNomeAte = New AdmEvento
    Set objEventoFilialEmpresaDe = New AdmEvento
    Set objEventoFilialEmpresaAte = New AdmEvento
    Set objEventoFilialFatDe = New AdmEvento
    Set objEventoFilialFatAte = New AdmEvento
        
    ComboOrdenacao.ListIndex = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171946)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoCodigoPVDe = Nothing
    Set objEventoCodigoPVAte = Nothing
    Set objEventoCodClienteDe = Nothing
    Set objEventoCodClienteAte = Nothing
    Set objEventoNomeDe = Nothing
    Set objEventoNomeAte = Nothing
    Set objEventoFilialEmpresaDe = Nothing
    Set objEventoFilialEmpresaAte = Nothing
    Set objEventoFilialFatDe = Nothing
    Set objEventoFilialFatAte = Nothing
    
End Sub

Private Sub CodClienteAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodClienteAte, iAlterado)
    
End Sub

Private Sub CodClienteDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodClienteDe, iAlterado)
    
End Sub

Private Sub CodigoPVAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoPVAte, iAlterado)
    
End Sub

Private Sub CodigoPVDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoPVDe, iAlterado)
    
End Sub




Private Sub DataEmissaoAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEmissaoAte, iAlterado)
    
End Sub

Private Sub DataEmissaoDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEmissaoDe, iAlterado)
    
End Sub

Private Sub FilialEmpresaAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(FilialEmpresaAte, iAlterado)
    
End Sub

Private Sub FilialEmpresaDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(FilialEmpresaDe, iAlterado)
    
End Sub

Private Sub FilialFatAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(FilialFatAte, iAlterado)
    
End Sub

Private Sub FilialFatDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(FilialFatDe, iAlterado)
    
End Sub

Private Sub LabelCodigoAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objPedVenda As New ClassPedidoDeVenda

On Error GoTo Erro_LabelCodigoAte_Click

    If Len(Trim(CodigoPVAte.Text)) > 0 Then
        'Preenche com o Pedido de Venda da tela
        objPedVenda.lCodigo = StrParaLong(CodigoPVAte.Text)
    End If

    'Chama Tela PedidoVendaTodosLista
    Call Chama_Tela("PedidoVendaTodosLista", colSelecao, objPedVenda, objEventoCodigoPVAte)

   Exit Sub

Erro_LabelCodigoAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171947)

    End Select

    Exit Sub

End Sub
Private Sub LabelCodigoDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objPedVenda As New ClassPedidoDeVenda

On Error GoTo Erro_LabelCodigoDe_Click

    If Len(Trim(CodigoPVDe.Text)) > 0 Then
        'Preenche com o Pedido de Venda da tela
        objPedVenda.lCodigo = StrParaLong(CodigoPVDe.Text)
    End If

    'Chama Tela PedidoVendaTodosLista
    Call Chama_Tela("PedidoVendaTodosLista", colSelecao, objPedVenda, objEventoCodigoPVDe)

   Exit Sub

Erro_LabelCodigoDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171948)

    End Select

    Exit Sub

End Sub


Private Sub DataEmissaoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissaoDe_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataEmissaoDe.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataEmissaoDe.Text)
    If lErro <> SUCESSO Then gError 68809

    Exit Sub
                   
Erro_DataEmissaoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 68809
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171949)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissaoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissaoAte_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataEmissaoAte.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataEmissaoAte.Text)
    If lErro <> SUCESSO Then gError 68810

    Exit Sub
                   
Erro_DataEmissaoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 68810
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171950)

    End Select

    Exit Sub

End Sub


Private Sub UpDownDataEmissaoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissaoDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEmissaoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 68811

    Exit Sub

Erro_UpDownDataEmissaoDe_DownClick:

    Select Case gErr

        Case 68811
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 171951)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEmissaoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissaoDe_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEmissaoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 68812

    Exit Sub

Erro_UpDownDataEmissaoDe_UpClick:

    Select Case gErr

        Case 68812
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 171952)

    End Select

    Exit Sub

End Sub
Private Sub UpDownDataEmissaoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissaoAte_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEmissaoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 68813

    Exit Sub

Erro_UpDownDataEmissaoAte_DownClick:

    Select Case gErr

        Case 68813
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 171953)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEmissaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissaoAte_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEmissaoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 68814

    Exit Sub

Erro_UpDownDataEmissaoAte_UpClick:

    Select Case gErr

        Case 68814
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 171954)

    End Select

    Exit Sub

End Sub

Private Sub LabelEmpresaDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelEmpresaDe_Click

    If Len(Trim(FilialEmpresaDe.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.iCodFilial = StrParaInt(FilialEmpresaDe.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoFilialEmpresaDe)

   Exit Sub

Erro_LabelEmpresaDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171955)

    End Select

    Exit Sub

End Sub
Private Sub LabelEmpresaAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelEmpresaAte_Click

    If Len(Trim(FilialEmpresaAte.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.iCodFilial = StrParaInt(FilialEmpresaAte.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoFilialEmpresaAte)

   Exit Sub

Erro_LabelEmpresaAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171956)

    End Select

    Exit Sub

End Sub
Private Sub LabelFilialFatDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelFilialFatDe_Click

    If Len(Trim(FilialFatDe.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.iCodFilial = StrParaInt(FilialFatDe.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoFilialFatDe)

   Exit Sub

Erro_LabelFilialFatDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171957)

    End Select

    Exit Sub

End Sub
Private Sub LabelFilialFatAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelFilialFatAte_Click

    If Len(Trim(FilialFatAte.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.iCodFilial = StrParaInt(FilialFatAte.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoFilialFatAte)

   Exit Sub

Erro_LabelFilialFatAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171958)

    End Select

    Exit Sub

End Sub


Private Sub LabelClienteAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCliente As New ClassCliente

On Error GoTo Erro_LabelClienteAte_Click

    If Len(Trim(CodClienteAte.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = StrParaLong(CodClienteAte.Text)
    End If

    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCodClienteAte)

   Exit Sub

Erro_LabelClienteAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171959)

    End Select

    Exit Sub

End Sub

Private Sub LabelClienteDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCliente As New ClassCliente

On Error GoTo Erro_LabelClienteDe_Click

    If Len(Trim(CodClienteDe.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = StrParaLong(CodClienteDe.Text)
    End If

    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCodClienteDe)

   Exit Sub

Erro_LabelClienteDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171960)

    End Select

    Exit Sub

End Sub


Private Sub LabelNomeDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCliente  As New ClassCliente

On Error GoTo Erro_LabelNomeDe_Click

    If Len(Trim(NomeDe.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.sNomeReduzido = NomeDe.Text
    End If

    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoNomeDe)

   Exit Sub

Erro_LabelNomeDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171961)

    End Select

    Exit Sub

End Sub
Private Sub LabelNomeAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCliente As New ClassCliente

On Error GoTo Erro_LabelNomeAte_Click

    If Len(Trim(NomeAte.Text)) > 0 Then
        'Preenche com o Cliente da tela
        objCliente.sNomeReduzido = NomeAte.Text
    End If

    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoNomeAte)

   Exit Sub

Erro_LabelNomeAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171962)

    End Select

    Exit Sub

End Sub


Private Sub objEventoFilialEmpresaAte_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    FilialEmpresaAte.Text = CStr(objFilialEmpresa.iCodFilial)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoFilialEmpresaDe_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    FilialEmpresaDe.Text = CStr(objFilialEmpresa.iCodFilial)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoFilialFatAte_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    FilialFatAte.Text = CStr(objFilialEmpresa.iCodFilial)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoFilialFatDe_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    FilialFatDe.Text = CStr(objFilialEmpresa.iCodFilial)

    Me.Show

    Exit Sub

End Sub
Private Sub objEventoCodigoPVAte_evSelecao(obj1 As Object)

Dim objPedVenda As New ClassPedidoDeVenda

    Set objPedVenda = obj1

    CodigoPVAte.Text = CStr(objPedVenda.lCodigo)

    Me.Show

End Sub

Private Sub objEventoCodigoPVDe_evSelecao(obj1 As Object)

Dim objPedVenda As New ClassPedidoDeVenda

    Set objPedVenda = obj1

    CodigoPVDe.Text = CStr(objPedVenda.lCodigo)

    Me.Show

End Sub


Private Sub objEventoCodClienteDe_evSelecao(obj1 As Object)

Dim objCliente As New ClassCliente

    Set objCliente = obj1

    CodClienteDe.Text = CStr(objCliente.lCodigo)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCodClienteAte_evSelecao(obj1 As Object)

Dim objCliente As New ClassCliente

    Set objCliente = obj1

    CodClienteAte.Text = CStr(objCliente.lCodigo)

    Me.Show

    Exit Sub

End Sub


Private Sub objEventoNomeAte_evSelecao(obj1 As Object)

Dim objCliente As New ClassCliente

    Set objCliente = obj1

    NomeAte.Text = objCliente.sNomeReduzido

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeDe_evSelecao(obj1 As Object)

Dim objCliente As New ClassCliente

    Set objCliente = obj1

    NomeDe.Text = objCliente.sNomeReduzido

    Me.Show

    Exit Sub

End Sub


Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 68815

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 68816

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 68817
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 68818
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 68815
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 68816 To 687818
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 171963)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 68819

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 68820

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr
    
        Case 68819
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 68820

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171964)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 68821

    Select Case ComboOrdenacao.ListIndex

            Case ORD_POR_CODIGO
                
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "CodigoPV", 1)
                
            Case ORD_POR_EMISSAO

                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "DataEmissao", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "CodigoPV", 1)
                
            Case ORD_POR_CLIENTE

                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ClienteCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "PVFilial", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "CodigoPV", 1)
                
            Case Else
                gError 74964

    End Select

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 68821, 74964

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171965)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sFilialEmpresa_I As String
Dim sFilialEmpresa_F As String
Dim sFilialFat_I As String
Dim sFilialFat_F As String
Dim sNome_I As String
Dim sNome_F As String
Dim sCodCliente_I As String
Dim sCodCliente_F As String
Dim sCodPV_I As String
Dim sCodPV_F As String
Dim sOrdenacaoPor As String
Dim iOrdenacao As Long
Dim sOrd As String
Dim sCheck As String

On Error GoTo Erro_PreencherRelOp
    
    lErro = Formata_E_Critica_Parametros(sFilialEmpresa_I, sFilialEmpresa_F, sFilialFat_I, sFilialFat_F, sCodPV_I, sCodPV_F, sNome_I, sNome_F, sCodCliente_I, sCodCliente_F)
    If lErro <> SUCESSO Then gError 68822

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 68823
         
    
    lErro = objRelOpcoes.IncluirParametro("NCODFILEMPINIC", sFilialEmpresa_I)
    If lErro <> AD_BOOL_TRUE Then gError 68824
         
    lErro = objRelOpcoes.IncluirParametro("NCODFILFATINIC", sFilialFat_I)
    If lErro <> AD_BOOL_TRUE Then gError 68825
    
    lErro = objRelOpcoes.IncluirParametro("NCODPVINIC", sCodPV_I)
    If lErro <> AD_BOOL_TRUE Then gError 68826
    
    lErro = objRelOpcoes.IncluirParametro("TNOMECLIINIC", NomeDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 68827
         
    lErro = objRelOpcoes.IncluirParametro("NCODCLIINIC", sCodCliente_I)
    If lErro <> AD_BOOL_TRUE Then gError 68828
    
    'Preenche dataemissao inicial
    If Trim(DataEmissaoDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DEMIINIC", DataEmissaoDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DEMIINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 68829
    
    lErro = objRelOpcoes.IncluirParametro("NCODFILEMPFIM", sFilialEmpresa_F)
    If lErro <> AD_BOOL_TRUE Then gError 68831
         
    lErro = objRelOpcoes.IncluirParametro("NCODFILFATFIM", sFilialFat_F)
    If lErro <> AD_BOOL_TRUE Then gError 68832
    
    lErro = objRelOpcoes.IncluirParametro("NCODPVFIM", sCodPV_F)
    If lErro <> AD_BOOL_TRUE Then gError 68833
    
    lErro = objRelOpcoes.IncluirParametro("TNOMECLIFIM", NomeAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 68834
         
    lErro = objRelOpcoes.IncluirParametro("NCODCLIFIM", sCodCliente_F)
    If lErro <> AD_BOOL_TRUE Then gError 68835
    
    'Preenche data de emissao Final
    If Trim(DataEmissaoAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DEMIFIM", DataEmissaoAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DEMIFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 68830
    
    
    'Inclui PedVenda Faturados
    If CheckPVFaturados.Value Then
        sCheck = 1
        
    Else
        sCheck = 0
    End If

    lErro = objRelOpcoes.IncluirParametro("NPEDVENFAT", sCheck)
    If lErro <> AD_BOOL_TRUE Then gError 68624
    
    Select Case ComboOrdenacao.ListIndex
        
            Case ORD_POR_CODIGO
            
                sOrdenacaoPor = "CodPV"
                    
            Case ORD_POR_EMISSAO
                sOrdenacaoPor = "DataEmissao"
                
            Case ORD_POR_CLIENTE
                sOrdenacaoPor = "Cliente"
                
            Case Else
                gError 68836
                  
    End Select

    lErro = objRelOpcoes.IncluirParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> AD_BOOL_TRUE Then gError 68837
   
    sOrd = ComboOrdenacao.ListIndex
    
    lErro = objRelOpcoes.IncluirParametro("NORDENACAO", sOrd)
    If lErro <> AD_BOOL_TRUE Then gError 68838
   
    If CheckPVFaturados.Value = 0 Then
        gobjRelatorio.sNomeTsk = "pvxrcom"
    Else
        gobjRelatorio.sNomeTsk = "pvtxrcom"
    End If
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sFilialEmpresa_I, sFilialEmpresa_F, sFilialFat_I, sFilialFat_F, sCodPV_I, sCodPV_F, sNome_I, sNome_F, sCodCliente_I, sCodCliente_F, sCheck, sOrdenacaoPor, sOrd)
    If lErro <> SUCESSO Then gError 68839

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 68822 To 68839
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171966)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sFilialEmpresa_I As String, sFilialEmpresa_F As String, sFilialFat_I As String, sFilialFat_F As String, sCodPV_I As String, sCodPV_F As String, sNome_I As String, sNome_F As String, sCodCliente_I As String, sCodCliente_F As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros
       
    'critica CodigoPV Inicial e Final
    If CodigoPVDe.Text <> "" Then
        sCodPV_I = CStr(CodigoPVDe.Text)
    Else
        sCodPV_I = ""
    End If
    
    If CodigoPVAte.Text <> "" Then
        sCodPV_F = CStr(CodigoPVAte.Text)
    Else
        sCodPV_F = ""
    End If
            
    If sCodPV_I <> "" And sCodPV_F <> "" Then
        
        If StrParaLong(sCodPV_I) > StrParaLong(sCodPV_F) Then gError 68840
        
    End If
    
    'critica CodigoFilialEmpresa Inicial e Final
    If FilialEmpresaDe.Text <> "" Then
        sFilialEmpresa_I = CStr(FilialEmpresaDe.Text)
    Else
        sFilialEmpresa_I = ""
    End If

    If FilialEmpresaAte.Text <> "" Then
        sFilialEmpresa_F = CStr(FilialEmpresaAte.Text)
    Else
        sFilialEmpresa_F = ""
    End If

    If sFilialEmpresa_I <> "" And sFilialEmpresa_F <> "" Then

        If StrParaInt(sFilialEmpresa_I) > StrParaInt(sFilialEmpresa_F) Then gError 68842

    End If
    
    'critica CodigoFilialFaturamento Inicial e Final
    If FilialFatDe.Text <> "" Then
        sFilialFat_I = CStr(FilialFatDe.Text)
    Else
        sFilialFat_I = ""
    End If

    If FilialFatAte.Text <> "" Then
        sFilialFat_F = CStr(FilialFatAte.Text)
    Else
        sFilialFat_F = ""
    End If

    If sFilialFat_I <> "" And sFilialFat_F <> "" Then

        If StrParaInt(sFilialFat_I) > StrParaInt(sFilialFat_F) Then gError 68843

    End If

    If NomeDe.Text <> "" Then
        sNome_I = NomeDe.Text
    Else
        sNome_I = ""
    End If
    
    If NomeAte.Text <> "" Then
        sNome_F = NomeAte.Text
    Else
        sNome_F = ""
    End If
    
    If sNome_I <> "" And sNome_F <> "" Then
        If sNome_I > sNome_F Then gError 68841
    End If
    
    
    'critica CodigoCliente Inicial e Final
    If CodClienteDe.Text <> "" Then
        sCodCliente_I = CStr(CodClienteDe.Text)
    Else
        sCodCliente_I = ""
    End If
    
    If CodClienteAte.Text <> "" Then
        sCodCliente_F = CStr(CodClienteAte.Text)
    Else
        sCodCliente_F = ""
    End If
            
    If sCodCliente_I <> "" And sCodCliente_F <> "" Then
        
        If StrParaLong(sCodCliente_I) > StrParaLong(sCodCliente_F) Then gError 68844
        
    End If
    
    'data de Envio inicial não pode ser maior que a final
    If Trim(DataEmissaoDe.ClipText) <> "" And Trim(DataEmissaoAte.ClipText) <> "" Then
    
         If CDate(DataEmissaoDe.Text) > CDate(DataEmissaoAte.Text) Then gError 68845
    
    End If
    
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 68840
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PV_INICIAL_MAIOR", gErr)
            CodigoPVDe.SetFocus
                
        Case 68841
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOMECLIENTE_INICIAL_MAIOR", gErr)
            NomeDe.SetFocus
            
        Case 68842
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            FilialEmpresaDe.SetFocus
            
        Case 68843
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            FilialFatDe.SetFocus
            
        Case 68844
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)
            CodClienteDe.SetFocus
            
        Case 68845
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_INICIAL_MAIOR", gErr)
            DataEmissaoDe.SetFocus
           
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171967)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sFilialEmpresa_I As String, sFilialEmpresa_F As String, sFilialFat_I As String, sFilialFat_F As String, sCodPV_I As String, sCodPV_F As String, sNome_I As String, sNome_F As String, sCodCliente_I As String, sCodCliente_F As String, sCheck As String, sOrdenacaoPor As String, sOrd As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   If sCodPV_I <> "" Then sExpressao = "CodPV >= " & Forprint_ConvLong(StrParaLong(sCodPV_I))

   If sCodPV_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CodPV <= " & Forprint_ConvLong(StrParaLong(sCodPV_F))

    End If

   If sNome_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CliNome >= " & Forprint_ConvTexto(sNome_I)

    End If
    
    If sNome_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CliNome <= " & Forprint_ConvTexto(sNome_F)

    End If
   
    If sCodCliente_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CliCod >= " & Forprint_ConvTexto((sCodCliente_I))

    End If
   
    If sCodCliente_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CliCod <= " & Forprint_ConvTexto((sCodCliente_F))

    End If
   
    If sFilialEmpresa_I <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpCod >= " & Forprint_ConvInt(StrParaInt(sFilialEmpresa_I))
    End If
    
    If sFilialEmpresa_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpCod <= " & Forprint_ConvInt(StrParaInt(sFilialEmpresa_F))

    End If

    If sFilialFat_I <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilFatCod >= " & Forprint_ConvInt(StrParaInt(sFilialFat_I))
    End If
    
    If sFilialFat_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilFatCod <= " & Forprint_ConvInt(StrParaInt(sFilialFat_F))

    End If

   If Trim(DataEmissaoDe.ClipText) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Emissao >= " & Forprint_ConvData(CDate(DataEmissaoDe.Text))
        
    End If
    
    If Trim(DataEmissaoAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Emissao <= " & Forprint_ConvData(CDate(DataEmissaoAte.Text))

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171968)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long, iTipoOrd As Integer, iAscendente As Integer
Dim sParam As String
Dim sOrdenacaoPor As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 68846
   
    'pega Codigo inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODPVINIC", sParam)
    If lErro <> SUCESSO Then gError 68847
    
    CodigoPVDe.Text = sParam
    
    'pega  Codigo final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODPVFIM", sParam)
    If lErro <> SUCESSO Then gError 68848
    
    CodigoPVAte.Text = sParam
                
    'pega  Nome Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMECLIINIC", sParam)
    If lErro <> SUCESSO Then gError 68849
                   
    NomeDe.Text = sParam
    Call NomeDe_Validate(bSGECancelDummy)
    
    'pega  Nome Final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMECLIFIM", sParam)
    If lErro <> SUCESSO Then gError 68850
                   
    NomeAte.Text = sParam
    Call NomeAte_Validate(bSGECancelDummy)
                        
    'pega CodigoFilialEmpresa inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILEMPINIC", sParam)
    If lErro <> SUCESSO Then gError 68851
    
    FilialEmpresaDe.Text = sParam
    Call FilialEmpresaDe_Validate(bSGECancelDummy)
    
    'pega  CodigoFilialEmpresa final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILEMPFIM", sParam)
    If lErro <> SUCESSO Then gError 68852
    
    FilialEmpresaAte.Text = sParam
    Call FilialEmpresaAte_Validate(bSGECancelDummy)
                
    'pega CodigoFilialFat Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILFATINIC", sParam)
    If lErro <> SUCESSO Then gError 68853
                   
    FilialFatDe.Text = sParam
    Call FilialFatDe_Validate(bSGECancelDummy)
    
    'pega CodigoFilialFat Final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILFATFIM", sParam)
    If lErro <> SUCESSO Then gError 68854
                   
    FilialFatAte.Text = sParam
    Call FilialFatAte_Validate(bSGECancelDummy)
                                            
    'pega CodigoCliente inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODCLIINIC", sParam)
    If lErro <> SUCESSO Then gError 68855
    
    CodClienteDe.Text = sParam
    Call CodClienteDe_Validate(bSGECancelDummy)
    
    'pega  CodigoCliente final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODCLIFIM", sParam)
    If lErro <> SUCESSO Then gError 68856
    
    CodClienteAte.Text = sParam
    Call CodClienteAte_Validate(bSGECancelDummy)
                                                           
    'pega DataEmissao inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DEMIINIC", sParam)
    If lErro <> SUCESSO Then gError 68857
    
    Call DateParaMasked(DataEmissaoDe, CDate(sParam))
    
    'pega data de emissao final e exibe
    lErro = objRelOpcoes.ObterParametro("DEMIFIM", sParam)
    If lErro <> SUCESSO Then gError 68858

    Call DateParaMasked(DataEmissaoAte, CDate(sParam))
   
    'pega 'Inclui PV Faturados' e exibe
    lErro = objRelOpcoes.ObterParametro("NPEDVENFAT", sParam)
    If lErro <> SUCESSO Then gError 72524

    If sParam = "1" Then
        CheckPVFaturados.Value = vbChecked
    Else
        CheckPVFaturados.Value = vbUnchecked
    End If
    
    lErro = objRelOpcoes.ObterParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> SUCESSO Then gError 68859
    
    Select Case sOrdenacaoPor
        
            Case "CodPV"
            
                ComboOrdenacao.ListIndex = ORD_POR_CODIGO
            
            Case "DataEmissao"
                
                ComboOrdenacao.ListIndex = ORD_POR_EMISSAO
                
            Case "Cliente"
            
                ComboOrdenacao.ListIndex = ORD_POR_CLIENTE
                
            Case Else
                gError 68860
                  
    End Select
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 68846 To 68860, 72524
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 171969)

    End Select

    Exit Function

End Function
Private Sub FilialEmpresaDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_FilialEmpresaDe_Validate

    If Len(Trim(FilialEmpresaDe.Text)) > 0 Then

        objFilialEmpresa.iCodFilial = StrParaInt(FilialEmpresaDe.Text)
        'Lê o código informado
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 68861
        
        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 68862

    End If

    Exit Sub

Erro_FilialEmpresaDe_Validate:

    Cancel = True


    Select Case gErr

        Case 68861

        Case 68862
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171970)

    End Select

    Exit Sub

End Sub
Private Sub FilialEmpresaAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_FilialEmpresaAte_Validate

    If Len(Trim(FilialEmpresaAte.Text)) > 0 Then

        objFilialEmpresa.iCodFilial = StrParaInt(FilialEmpresaAte.Text)
        'Lê o código informado
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 68863
        
        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 68864

    End If

    Exit Sub

Erro_FilialEmpresaAte_Validate:

    Cancel = True


    Select Case gErr

        Case 68863

        Case 68864
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171971)

    End Select

    Exit Sub

End Sub
Private Sub FilialFatAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_FilialFatAte_Validate

    If Len(Trim(FilialFatAte.Text)) > 0 Then

        objFilialEmpresa.iCodFilial = StrParaInt(FilialFatAte.Text)
        'Lê o código informado
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 68873
        
        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 68874

    End If

    Exit Sub

Erro_FilialFatAte_Validate:

    Cancel = True


    Select Case gErr

        Case 68873

        Case 68874
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171972)

    End Select

    Exit Sub

End Sub

Private Sub FilialFatDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_FilialFatDe_Validate

    If Len(Trim(FilialFatDe.Text)) > 0 Then

        objFilialEmpresa.iCodFilial = StrParaInt(FilialFatDe.Text)
        'Lê o código informado
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 68875
        
        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 68876

    End If

    Exit Sub

Erro_FilialFatDe_Validate:

    Cancel = True

    Select Case gErr

        Case 68875

        Case 68876
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171973)

    End Select

    Exit Sub

End Sub

Private Sub CodClienteDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_CodClienteDe_Validate

    If Len(Trim(CodClienteDe.Text)) > 0 Then

        objCliente.lCodigo = StrParaLong(CodClienteDe.Text)
        'Lê o código informado
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 68865

        'Se não encontrou o Requisitante ==> erro
        If lErro = 12293 Then gError 68866
        
    End If

    Exit Sub

Erro_CodClienteDe_Validate:

    Cancel = True


    Select Case gErr

        Case 68865

        Case 68866
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171974)

    End Select

    Exit Sub
    
End Sub

Private Sub CodClienteAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_CodClienteAte_Validate

    If Len(Trim(CodClienteAte.Text)) > 0 Then

        objCliente.lCodigo = StrParaLong(CodClienteAte.Text)
        'Lê o código informado
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 68867

        'Se não encontrou o Requisitante ==> erro
        If lErro = 12293 Then gError 68868
        
    End If

    Exit Sub

Erro_CodClienteAte_Validate:

    Cancel = True


    Select Case gErr

        Case 68867

        Case 68868
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171975)

    End Select

    Exit Sub
    
End Sub

Private Sub NomeDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_NomeDe_Validate

    
    If Len(Trim(NomeDe.Text)) > 0 Then

        objCliente.sNomeReduzido = NomeDe.Text
        'Lê o Cliente
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 68869
        
        If lErro = 12348 Then gError 68870
        
        NomeDe.Text = objCliente.sNomeReduzido

    End If
    
    Exit Sub

Erro_NomeDe_Validate:

    Cancel = True

    Select Case gErr

        Case 68869

        Case 68870
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, NomeDe.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171976)

    End Select

Exit Sub

End Sub
Private Sub NomeAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_NomeAte_Validate

    If Len(Trim(NomeAte.Text)) > 0 Then

        objCliente.sNomeReduzido = NomeAte.Text
        
        'Lê o Cliente
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 68871
        
        If lErro = 12348 Then gError 68872
        
        NomeAte.Text = objCliente.sNomeReduzido

    End If
    
    Exit Sub

Erro_NomeAte_Validate:

    Cancel = True

    Select Case gErr

        Case 68871

        Case 68872
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, NomeDe.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 171977)

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
    Caption = "Pedidos de Venda x Requisições de Compra"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpPVReqCompras"
    
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
        
        If Me.ActiveControl Is CodigoPVDe Then
            Call LabelCodigoDe_Click
            
        ElseIf Me.ActiveControl Is CodigoPVAte Then
            Call LabelCodigoAte_Click
            
        ElseIf Me.ActiveControl Is NomeDe Then
            Call LabelNomeDe_Click
            
        ElseIf Me.ActiveControl Is NomeAte Then
            Call LabelNomeAte_Click
            
        ElseIf Me.ActiveControl Is CodClienteDe Then
            Call LabelClienteDe_Click
            
        ElseIf Me.ActiveControl Is CodClienteAte Then
            Call LabelClienteAte_Click
            
        ElseIf Me.ActiveControl Is FilialEmpresaDe Then
            Call LabelEmpresaDe_Click
        
        ElseIf Me.ActiveControl Is FilialEmpresaAte Then
            Call LabelEmpresaAte_Click
        
        ElseIf Me.ActiveControl Is FilialFatDe Then
            Call LabelFilialFatDe_Click
        
        ElseIf Me.ActiveControl Is FilialFatAte Then
            Call LabelFilialFatAte_Click
        
        End If
    
    End If

End Sub


Private Sub LabelCodigoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoDe, Source, X, Y)
End Sub

Private Sub LabelCodigoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoAte, Source, X, Y)
End Sub

Private Sub LabelCodigoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoAte, Button, Shift, X, Y)
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





Private Sub LabelDataAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataAte, Source, X, Y)
End Sub

Private Sub LabelDataAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataAte, Button, Shift, X, Y)
End Sub

Private Sub LabelDataDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataDe, Source, X, Y)
End Sub

Private Sub LabelDataDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataDe, Button, Shift, X, Y)
End Sub

Private Sub LabelEmpresaAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelEmpresaAte, Source, X, Y)
End Sub

Private Sub LabelEmpresaAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelEmpresaAte, Button, Shift, X, Y)
End Sub

Private Sub LabelEmpresaDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelEmpresaDe, Source, X, Y)
End Sub

Private Sub LabelEmpresaDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelEmpresaDe, Button, Shift, X, Y)
End Sub

Private Sub LabelFilialFatAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilialFatAte, Source, X, Y)
End Sub

Private Sub LabelFilialFatAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilialFatAte, Button, Shift, X, Y)
End Sub

Private Sub LabelFilialFatDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilialFatDe, Source, X, Y)
End Sub

Private Sub LabelFilialFatDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilialFatDe, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteAte, Source, X, Y)
End Sub

Private Sub LabelClienteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteAte, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteDe, Source, X, Y)
End Sub

Private Sub LabelClienteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteDe, Button, Shift, X, Y)
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

