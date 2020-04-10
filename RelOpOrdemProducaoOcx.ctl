VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl RelOpOrdemProducaoOcx 
   ClientHeight    =   4950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9030
   KeyPreview      =   -1  'True
   ScaleHeight     =   4950
   ScaleWidth      =   9030
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6720
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpOrdemProducaoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpOrdemProducaoOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpOrdemProducaoOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpOrdemProducaoOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpOrdemProducaoOcx.ctx":0994
      Left            =   1410
      List            =   "RelOpOrdemProducaoOcx.ctx":0996
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   210
      Width           =   2916
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
      Left            =   4845
      Picture         =   "RelOpOrdemProducaoOcx.ctx":0998
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   3525
      Index           =   2
      Left            =   150
      TabIndex        =   9
      Top             =   1230
      Visible         =   0   'False
      Width           =   8670
      Begin VB.Frame FrameEmissao 
         Caption         =   "Emissão"
         Height          =   765
         Left            =   150
         TabIndex        =   40
         Top             =   1545
         Width           =   5640
         Begin MSMask.MaskEdBox DataInicial 
            Height          =   300
            Left            =   1095
            TabIndex        =   12
            Top             =   300
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataFinal 
            Height          =   300
            Left            =   3405
            TabIndex        =   13
            Top             =   285
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   300
            Left            =   2115
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   300
            Width           =   240
            _ExtentX        =   397
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   285
            Left            =   4410
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   285
            Width           =   240
            _ExtentX        =   397
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label dIni 
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
            Left            =   735
            TabIndex        =   44
            Top             =   330
            Width           =   315
         End
         Begin VB.Label dFim 
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
            Left            =   3000
            TabIndex        =   43
            Top             =   345
            Width           =   360
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Centro de Custo"
         Height          =   1245
         Left            =   135
         TabIndex        =   35
         Top             =   135
         Width           =   5640
         Begin MSMask.MaskEdBox CclInicial 
            Height          =   285
            Left            =   840
            TabIndex        =   10
            Top             =   330
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CclFinal 
            Height          =   270
            Left            =   840
            TabIndex        =   11
            Top             =   787
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   476
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin VB.Label DescCclInic 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2190
            TabIndex        =   39
            Top             =   330
            Width           =   3255
         End
         Begin VB.Label LabelCclDe 
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
            Height          =   195
            Left            =   420
            TabIndex        =   38
            Top             =   375
            Width           =   315
         End
         Begin VB.Label LabelCclAte 
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
            Height          =   195
            Left            =   405
            TabIndex        =   37
            Top             =   825
            Width           =   360
         End
         Begin VB.Label DescCclFim 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2190
            TabIndex        =   36
            Top             =   780
            Width           =   3255
         End
      End
      Begin VB.Frame FrameSituacao 
         Caption         =   "Situação"
         Height          =   705
         Left            =   135
         TabIndex        =   34
         Top             =   2565
         Width           =   5640
         Begin VB.OptionButton OpSituacao 
            Caption         =   "Sacramentadas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   3525
            TabIndex        =   16
            Top             =   285
            Width           =   1740
         End
         Begin VB.OptionButton OpSituacao 
            Caption         =   "Suspensas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   1950
            TabIndex        =   15
            Top             =   270
            Width           =   1245
         End
         Begin VB.OptionButton OpSituacao 
            Caption         =   "Todas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   780
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   270
            Width           =   960
         End
      End
      Begin MSComctlLib.TreeView TvwCcls 
         Height          =   3075
         Left            =   5895
         TabIndex        =   17
         Top             =   225
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   5424
         _Version        =   393217
         Indentation     =   453
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin VB.Label LabelCcl 
         Caption         =   "Centro de Custo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   45
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   3525
      Index           =   1
      Left            =   135
      TabIndex        =   1
      Top             =   1215
      Width           =   8670
      Begin VB.Frame FrameStatus 
         Caption         =   "Status"
         Height          =   630
         Left            =   150
         TabIndex        =   33
         Top             =   150
         Width           =   6660
         Begin VB.OptionButton OpStatus 
            Caption         =   "Encerradas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   4005
            TabIndex        =   4
            Top             =   195
            Width           =   1470
         End
         Begin VB.OptionButton OpStatus 
            Caption         =   "Abertas"
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
            Index           =   1
            Left            =   2715
            TabIndex        =   3
            Top             =   210
            Width           =   1215
         End
         Begin VB.OptionButton OpStatus 
            Caption         =   "Todas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   1380
            TabIndex        =   2
            Top             =   195
            Width           =   1110
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ordem de Produção"
         Height          =   810
         Left            =   150
         TabIndex        =   30
         Top             =   1065
         Width           =   6645
         Begin VB.TextBox OpFinal 
            Height          =   300
            Left            =   4140
            TabIndex        =   6
            Top             =   330
            Width           =   1680
         End
         Begin VB.TextBox OpInicial 
            Height          =   300
            Left            =   1065
            TabIndex        =   5
            Top             =   330
            Width           =   1680
         End
         Begin VB.Label LabelOpFinal 
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
            Left            =   3690
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   32
            Top             =   375
            Width           =   360
         End
         Begin VB.Label LabelOpInicial 
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
            Left            =   660
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   31
            Top             =   360
            Width           =   315
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Produtos"
         Height          =   1245
         Left            =   150
         TabIndex        =   24
         Top             =   2100
         Width           =   6675
         Begin MSMask.MaskEdBox ProdutoFinal 
            Height          =   300
            Left            =   1095
            TabIndex        =   8
            Top             =   795
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   529
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoInicial 
            Height          =   315
            Left            =   1095
            TabIndex        =   7
            Top             =   300
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label Label2 
            Caption         =   "Até o Nível:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3900
            TabIndex        =   29
            Top             =   1455
            Width           =   1110
         End
         Begin VB.Label DescProdInic 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2700
            TabIndex        =   28
            Top             =   300
            Width           =   3135
         End
         Begin VB.Label DescProdFim 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2700
            TabIndex        =   27
            Top             =   795
            Width           =   3135
         End
         Begin VB.Label LabelProdutoDe 
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
            Left            =   705
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   26
            Top             =   345
            Width           =   315
         End
         Begin VB.Label LabelProdutoAte 
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
            Left            =   675
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   25
            Top             =   825
            Width           =   360
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4020
      Left            =   120
      TabIndex        =   46
      Top             =   780
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   7091
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Parte 1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Parte 2"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Left            =   690
      TabIndex        =   47
      Top             =   255
      Width           =   615
   End
End
Attribute VB_Name = "RelOpOrdemProducaoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoOp As AdmEvento
Attribute objEventoOp.VB_VarHelpID = -1
Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio
Dim giProdInicial As Integer
Dim giCclInicial As Integer
Dim giOp_Inicial As Integer
Dim iFrameAtual As Integer

Private Sub Form_Load()

Dim lErro As Long
Dim sMascaraCclPadrao  As String

On Error GoTo Erro_Form_Load

    iFrameAtual = 1

    Set objEventoOp = New AdmEvento
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento

    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then Error 34141

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then Error 34142

    'Inicializa Máscara de Ccl
    sMascaraCclPadrao = String(STRING_CCL, 0)

    lErro = MascaraCcl(sMascaraCclPadrao)
    If lErro <> SUCESSO Then Error 54925

    CclInicial.Mask = sMascaraCclPadrao
    
    CclFinal.Mask = sMascaraCclPadrao

    'Inicializa a arvore de Centros de Custo
    lErro = Carga_Arvore_Ccl(TvwCcls.Nodes)
    If lErro <> SUCESSO Then Error 34144
    
    Call Define_Padrao

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = Err

    Select Case Err

        Case 34141, 34142, 34144, 54925

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170486)

    End Select

    Exit Sub

End Sub

Sub Define_Padrao()
'Preenche a tela com as opções padrão de FilialEmpresa

Dim lErro As Long

On Error GoTo Erro_Define_Padrao

    giProdInicial = 1
    giCclInicial = 1
    giOp_Inicial = 1

    OpSituacao(0).Value = True
    OpStatus(1).Value = True

    DataInicial.PromptInclude = False
    DataInicial.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataInicial.PromptInclude = True
    
    DataFinal.PromptInclude = False
    DataFinal.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataFinal.PromptInclude = True

    Exit Sub

Erro_Define_Padrao:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170487)

    End Select

    Exit Sub

End Sub

Private Sub CclInicial_GotFocus()

    giCclInicial = 1

End Sub

Private Sub CclFinal_GotFocus()
'mostra a arvore de Ccl

    giCclInicial = 0

End Sub

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal)

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial)

End Sub

Private Sub LabelOpInicial_Click()

Dim lErro As Long
Dim objOp As ClassOrdemDeProducao

On Error GoTo Erro_LabelOpInicial_Click

    giOp_Inicial = 1

    If Len(Trim(OpInicial.Text)) <> 0 Then

        Set objOp = New ClassOrdemDeProducao
        objOp.sCodigo = OpInicial.Text

    End If

    Call Chama_Browse_OP(objOp)

    Exit Sub

Erro_LabelOpInicial_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170488)

    End Select

    Exit Sub

End Sub

Private Sub LabelOpFinal_Click()

Dim lErro As Long
Dim colSelecao As Collection
Dim objOp As ClassOrdemDeProducao

On Error GoTo Erro_LabelOpFinal_Click

    giOp_Inicial = 0

    If Len(Trim(OpFinal.Text)) <> 0 Then

        Set objOp = New ClassOrdemDeProducao
        objOp.sCodigo = OpFinal.Text

    End If

    Call Chama_Browse_OP(objOp)

   Exit Sub

Erro_LabelOpFinal_Click:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170489)

    End Select

    Exit Sub

End Sub

Private Sub Chama_Browse_OP(objOp As ClassOrdemDeProducao)

Dim lErro As Long
Dim colSelecao As Collection
Dim iOpStatus As Integer
Dim iIndice As Integer

On Error GoTo Erro_Chama_Browse_OP

   'verifica status selecionado
    For iIndice = 0 To 2
        If OpStatus(iIndice).Value = True Then iOpStatus = iIndice
    Next

    Select Case iOpStatus

        Case 0
            Call Chama_Tela("OrdProdTodasListaModal", colSelecao, objOp, objEventoOp)

        Case 1
            Call Chama_Tela("OrdemProdListaModal", colSelecao, objOp, objEventoOp)

        Case 2
            Call Chama_Tela("OrdProdBaixadasListaModal", colSelecao, objOp, objEventoOp)

   End Select

   Exit Sub

Erro_Chama_Browse_OP:

    Select Case Err

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170490)

    End Select

    Exit Sub

End Sub

Private Sub objEventoOp_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOp As New ClassOrdemDeProducao

On Error GoTo Erro_objEventoOp_evSelecao

    Set objOp = obj1

    If giOp_Inicial = 1 Then

        OpInicial.Text = objOp.sCodigo
        
    Else

        OpFinal.Text = objOp.sCodigo

    End If

    Me.Show
    
    Exit Sub

Erro_objEventoOp_evSelecao:

    Select Case Err

       Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170491)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_GotFocus()

    giProdInicial = 1

End Sub

Private Sub ProdutoFinal_GotFocus()
    
    giProdInicial = 0

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim iIndice As Integer
Dim sDescCcl As String

On Error GoTo Erro_PreencherParametrosNaTela

 Call Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro Then Error 34149

    'pega Ordem de Producao Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TOPINIC", sParam)
    If lErro Then Error 34150

    OpInicial.Text = sParam

    'pega Ordem de Producao Final e exibe
    lErro = objRelOpcoes.ObterParametro("TOPFIM", sParam)
    If lErro Then Error 34151

    OpFinal.Text = sParam

    'pega Produto Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODINIC", sParam)
    If lErro Then Error 34152

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then Error 34153

    'pega parâmetro Produto Final e exibe
    lErro = objRelOpcoes.ObterParametro("TPRODFIM", sParam)
    If lErro Then Error 34154

    lErro = CF("Traz_Produto_MaskEd", sParam, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then Error 34155

   'pega Ccl Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCCLINIC", sParam)
    If lErro Then Error 34156

    If sParam <> "" Then
        
        lErro = Obtem_Descricao_Ccl(sParam, sDescCcl)
        If lErro <> SUCESSO Then Error 34157
    
        CclInicial.PromptInclude = False
        CclInicial.Text = sParam
        CclInicial.PromptInclude = True
        
        DescCclInic.Caption = sDescCcl
    
    End If
    
    'pega Ccl Final e exibe
    lErro = objRelOpcoes.ObterParametro("TCCLFIM", sParam)
    If lErro Then Error 34158

    If sParam <> "" Then
        
        lErro = Obtem_Descricao_Ccl(sParam, sDescCcl)
        If lErro <> SUCESSO Then Error 34159
    
        CclFinal.PromptInclude = False
        CclFinal.Text = sParam
        CclFinal.PromptInclude = True
        
        DescCclFim.Caption = sDescCcl

    End If
    
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DINIC", sParam)
    If lErro <> SUCESSO Then Error 34160
    
    If sParam <> "07/09/1822" Then Call DateParaMasked(DataInicial, StrParaDate(sParam))
    'DataInicial.PromptInclude = False
    'DataInicial.Text = sParam
    'DataInicial.PromptInclude = True

    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DFIM", sParam)
    If lErro <> SUCESSO Then Error 34161

    If sParam <> "07/09/1822" Then Call DateParaMasked(DataFinal, StrParaDate(sParam))
    'DataFinal.PromptInclude = False
    'DataFinal.Text = sParam
    'DataFinal.PromptInclude = True

    'Pega status e exibe
    lErro = objRelOpcoes.ObterParametro("NSTATUS", sParam)
    If lErro <> SUCESSO Then Error 34162

    OpStatus(CInt(sParam)) = True

    'Pega status e exibe
    lErro = objRelOpcoes.ObterParametro("NSITUACAO", sParam)
    If lErro <> SUCESSO Then Error 34163

    OpSituacao(CInt(sParam)) = True

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = Err

    Select Case Err

        Case 34149

        'erro ObterParametro
        Case 34150, 34151, 34152, 34154, 34156, 34158, 34160, 34161, 34162, 34163

        Case 34153, 34155

        Case 34157, 34159

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170492)

    End Select

    Exit Function

End Function

Private Sub Form_Unload(Cancel As Integer)

    Set objEventoOp = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    
End Sub
Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82448

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82449

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 82450

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 82448, 82450

        Case 82449
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170493)

    End Select

    Exit Sub

End Sub

Private Sub LabelProdutoAte_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoAte_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoFinal.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoFinal.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 82495

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoProduzivelLista", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 82495

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170494)

    End Select

    Exit Sub

End Sub

Private Sub LabelProdutoDe_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoDe_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoInicial.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoInicial.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 82494

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoProduzivelLista", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 82494

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170495)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 82400

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 82401

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 82402

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 82400, 82402

        Case 82401
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170496)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 29892
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then Error 34139

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
        
        Case 34139
        
        Case 29892
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170497)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub


Sub Limpar_Tela()

    Call Limpa_Tela(Me)

    DescProdInic.Caption = ""
    DescProdFim.Caption = ""
    DescCclInic.Caption = ""
    DescCclFim.Caption = ""

    ComboOpcoes.SetFocus

End Sub

Private Function Formata_E_Critica_Parametros(sProd_I As String, sProd_F As String, sCcl_I As String, sCcl_F As String) As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os parâmetros iniciais são maiores que os finais

Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim iCclPreenchida_I As Integer
Dim iCclPreenchida_F As Integer
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then Error 34165

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then Error 34166

    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambos os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then Error 34167

    End If

    'verifica se o Ccl Inicial é maior que o Ccl Final
    lErro = CF("Ccl_Formata", CclInicial.Text, sCcl_I, iCclPreenchida_I)
    If lErro Then Error 34168

    lErro = CF("Ccl_Formata", CclFinal.Text, sCcl_F, iCclPreenchida_F)
    If lErro Then Error 34169

    If (iCclPreenchida_I = CCL_PREENCHIDA) And (iCclPreenchida_F = CCL_PREENCHIDA) Then

        If sCcl_I > sCcl_F Then Error 34170

    End If

    'data inicial não pode ser maior que a data final
    
    If Trim(DataInicial.ClipText) <> "" And Trim(DataFinal.ClipText) <> "" Then
    
         If StrParaDate(DataInicial.Text) > StrParaDate(DataFinal.Text) Then Error 34171
    
    End If

    'ordem de produção inicial não pode ser maior que a final
    If Trim(OpInicial.Text) <> "" And Trim(OpFinal.Text) <> "" Then

        If OpInicial.Text > OpFinal.Text Then Error 34172

    End If

    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = Err

    Select Case Err

        Case 34165
            ProdutoInicial.SetFocus

        Case 34166
            ProdutoFinal.SetFocus

        Case 34167
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", Err)
            ProdutoInicial.SetFocus

        Case 34168
            CclInicial.SetFocus

        Case 34169
            CclFinal.SetFocus

        Case 34170
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_INICIAL_MAIOR", Err)
            CclInicial.SetFocus

        Case 34171
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
            CclInicial.SetFocus

        Case 34172
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OP_INICIAL_MAIOR", Err)
            OpInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170498)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    Limpar_Tela
    Call Define_Padrao

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sOp_I As String
Dim sOp_F As String
Dim sProd_I As String
Dim sProd_F As String
Dim sCcl_I As String
Dim sCcl_F As String
Dim sStatus As String
Dim sSituacao As String
Dim iIndice As Integer

On Error GoTo Erro_PreencherRelOp

    sProd_I = String(STRING_PRODUTO, 0)
    sProd_F = String(STRING_PRODUTO, 0)
    sCcl_I = String(STRING_CCL, 0)
    sCcl_F = String(STRING_CCL, 0)

    lErro = Formata_E_Critica_Parametros(sProd_I, sProd_F, sCcl_I, sCcl_F)
    If lErro <> SUCESSO Then Error 34178

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 34179

    sOp_I = OpInicial.Text
    lErro = objRelOpcoes.IncluirParametro("TOPINIC", sOp_I)
    If lErro <> AD_BOOL_TRUE Then Error 34180

    sOp_F = OpFinal.Text
    lErro = objRelOpcoes.IncluirParametro("TOPFIM", sOp_F)
    If lErro <> AD_BOOL_TRUE Then Error 34181

    lErro = objRelOpcoes.IncluirParametro("TPRODINIC", sProd_I)
    If lErro <> AD_BOOL_TRUE Then Error 34182

    lErro = objRelOpcoes.IncluirParametro("TPRODFIM", sProd_F)
    If lErro <> AD_BOOL_TRUE Then Error 34183

    lErro = objRelOpcoes.IncluirParametro("TCCLINIC", sCcl_I)
    If lErro <> AD_BOOL_TRUE Then Error 34184

    lErro = objRelOpcoes.IncluirParametro("TCCLFIM", sCcl_F)
    If lErro <> AD_BOOL_TRUE Then Error 34185

    If DataInicial.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DINIC", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 34186

    If DataFinal.ClipText <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then Error 34187

    'verifica opção selecionada
    For iIndice = 0 To 2
        If OpStatus(iIndice).Value = True Then sStatus = CStr(iIndice)
    Next

    lErro = objRelOpcoes.IncluirParametro("NSTATUS", sStatus)
    If lErro <> AD_BOOL_TRUE Then Error 34188

    'verifica opção selecionada
    For iIndice = 0 To 2
        If OpSituacao(iIndice).Value = True Then sSituacao = CStr(iIndice)
    Next
    
    If sStatus = "0" Then gobjRelatorio.sNomeTsk = "OrdProd"
    If sStatus = "1" Then gobjRelatorio.sNomeTsk = "OrdProdA"
    If sStatus = "2" Then gobjRelatorio.sNomeTsk = "OrdProdB"
    
    lErro = objRelOpcoes.IncluirParametro("NSITUACAO", sSituacao)
    If lErro <> AD_BOOL_TRUE Then Error 34189

    lErro = Monta_Expressao_Selecao(objRelOpcoes, sOp_I, sOp_F, sProd_I, sProd_F, sCcl_I, sCcl_F, sStatus, sSituacao)
    If lErro <> SUCESSO Then Error 34190

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 34178 To 34190

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170499)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then Error 34191

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then Error 34192

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Limpar_Tela

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 34191
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", Err)
            ComboOpcoes.SetFocus

        Case 34192

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170500)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 34193

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 34193

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170501)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then Error 34194

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then Error 34195

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then Error 34196

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 34194
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", Err)
            ComboOpcoes.SetFocus

        Case 34195, 34196

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170502)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    giProdInicial = 0

    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 34197
    
    If lErro <> SUCESSO Then Error 43265

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 34197

         Case 43265
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170503)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    giProdInicial = 1

    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then Error 34198
    
    If lErro <> SUCESSO Then Error 43266

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 34198

         Case 43266
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170504)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sOp_I As String, sOp_F As String, sProd_I As String, sProd_F As String, sCcl_I As String, sCcl_F As String, sStatus As String, sSituacao As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If sOp_I <> "" Then sExpressao = "OrdemProducao >= " & Forprint_ConvTexto(sOp_I)

    If sOp_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "OrdemProducao <= " & Forprint_ConvTexto(sOp_F)

    End If

    If sProd_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Produto >= " & Forprint_ConvTexto(sProd_I)

    End If

    If sProd_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Produto <= " & Forprint_ConvTexto(sProd_F)

    End If


    If sCcl_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Ccl >= " & Forprint_ConvTexto(sCcl_I)

    End If

    If sCcl_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Ccl <= " & Forprint_ConvTexto(sCcl_F)

    End If
    
''    If sExpressao <> "" Then sExpressao = sExpressao & " E "
''    sExpressao = "NSTATUS = " & Forprint_ConvInt(CInt(sStatus))
''
''    If sExpressao <> "" Then sExpressao = sExpressao & " E "
''    sExpressao = "NSITUACAO = " & Forprint_ConvInt(CInt(sSituacao))

     If Trim(DataInicial.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(StrParaDate(DataInicial.Text))

    End If
    
    If Trim(DataFinal.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(StrParaDate(DataFinal.Text))

    End If

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170505)

    End Select

    Exit Function

End Function

Function Obtem_Descricao_Ccl(sCcl As String, sDescCcl As String) As Long
'recebe em sCcl o Ccl no formato do Bd
'retorna em sDescCcl a descrição do Ccl ( que será formatado para tela )

Dim lErro As Long, iCclPreenchida As Integer
Dim objCcl As New ClassCcl
Dim sCopia As String

On Error GoTo Erro_Obtem_Descricao_Ccl

    sCopia = sCcl
    sDescCcl = String(STRING_CCL_DESCRICAO, 0)
    sCcl = String(STRING_CCL_MASK, 0)

    'determina qual Ccl deve ser lido
    objCcl.sCcl = sCopia

    lErro = Mascara_MascararCcl(sCopia, sCcl)
    If lErro <> SUCESSO Then Error 34199

    'verifica se a conta está preenchida
    lErro = CF("Ccl_Formata", sCcl, sCopia, iCclPreenchida)
    If lErro <> SUCESSO Then Error 34200

    If iCclPreenchida = CCL_PREENCHIDA Then

        'verifica se a Ccl existe
        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO Then Error 34201

        sDescCcl = objCcl.sDescCcl

    Else

        sCcl = ""
        sDescCcl = ""

    End If

    Obtem_Descricao_Ccl = SUCESSO

    Exit Function

Erro_Obtem_Descricao_Ccl:

    Obtem_Descricao_Ccl = Err

    Select Case Err

        Case 34199
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, sCopia)

        Case 34200

        Case 34201

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170506)

    End Select

    Exit Function

End Function

Function Ccl_Perde_Foco(Ccl As Object, Desc As Object) As Long
'recebe MaskEdBox do Ccl e o label da descrição

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim iCclPreenchida As Integer

On Error GoTo Erro_Ccl_Perde_Foco

    sCclFormatada = String(STRING_CCL, 0)

    Desc.Caption = ""

    lErro = CF("Ccl_Formata", Ccl.Text, sCclFormatada, iCclPreenchida)
    If lErro Then Error 34202

    If iCclPreenchida = CCL_PREENCHIDA Then

        'verifica se a Ccl existe
        objCcl.sCcl = sCclFormatada
        lErro = CF("Ccl_Le", objCcl)
        If lErro <> SUCESSO And lErro <> 5599 Then Error 34203

        If lErro = 5599 Then

            Ccl.Text = ""
            Ccl.SetFocus

            Error 34204

        End If

        Desc.Caption = objCcl.sDescCcl

    End If

    Ccl_Perde_Foco = SUCESSO

    Exit Function

Erro_Ccl_Perde_Foco:

    Ccl_Perde_Foco = Err

    Select Case Err

        Case 34202

        Case 34203

        Case 34204

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170507)

    End Select

    Exit Function

End Function

Private Function Carga_Arvore_Ccl(colNodes As Nodes) As Long
'move os dados de centro de custo/lucro do banco de dados para a arvore colNodes. /m

Dim objNode As Node
Dim colCcl As New Collection
Dim objCcl As ClassCcl
Dim lErro As Long
Dim sCclMascarado As String
Dim sCcl As String
Dim sCclPai As String
    
On Error GoTo Erro_Carga_Arvore_Ccl
    
    'leitura dos centro de custo/lucro no BD
    lErro = CF("Ccl_Le_Todos", colCcl)
    If lErro <> SUCESSO Then Error 34205
    
    'para cada centro de custo encontrado no bd
    For Each objCcl In colCcl
        
        sCclMascarado = String(STRING_CCL, 0)

        'coloca a mascara no centro de custo
        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then Error 34206

        sCcl = "C" & objCcl.sCcl

        sCclPai = String(STRING_CCL, 0)
        
        'retorna o centro de custo/lucro "pai" da centro de custo/lucro em questão, se houver
        lErro = Mascara_RetornaCclPai(objCcl.sCcl, sCclPai)
        If lErro <> SUCESSO Then Error 54514
        
        'se o centro de custo/lucro possui um centro de custo/lucro "pai"
        If Len(Trim(sCclPai)) > 0 Then

            sCclPai = "C" & sCclPai
            
            'adiciona o centro de custo como filho do centro de custo pai
            Set objNode = colNodes.Add(colNodes.Item(sCclPai), tvwChild, sCcl)

        Else
        
            'se o centro de custo/lucro não possui centro de custo/lucro "pai", adiciona na árvore sem pai
            Set objNode = colNodes.Add(, tvwLast, sCcl)
            
        End If
        
        'coloca o texto do nó que acabou de ser inserido
        objNode.Text = sCclMascarado & SEPARADOR & objCcl.sDescCcl
        
    Next
    
    Carga_Arvore_Ccl = SUCESSO

    Exit Function

Erro_Carga_Arvore_Ccl:

    Carga_Arvore_Ccl = Err

    Select Case Err

        Case 54514
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCclPai", Err, objCcl.sCcl)

        Case 34205

        Case 34206
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objCcl.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 170508)

    End Select
    
    Exit Function

End Function

''Function Carga_Arvore_Ccl(colNodes As Nodes) As Long
'''move os dados de centro de custo/lucro do banco de dados para a arvore colNodes.
''
''Dim objNode As Node
''Dim colCcl As New Collection
''Dim objCcl As ClassCcl
''Dim lErro As Long
''Dim sCclMascarado As String
''
''On Error GoTo Erro_Carga_Arvore_Ccl
''
''    lErro = CF("Ccl_Le_Todos",colCcl)
''    If lErro <> SUCESSO Then Error 34205
''
''    For Each objCcl In colCcl
''
''        sCclMascarado = String(STRING_CCL, 0)
''
''        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
''        If lErro <> SUCESSO Then Error 34206
''
''        Set objNode = colNodes.Add(, , "C" & objCcl.sCcl, sCclMascarado & SEPARADOR & objCcl.sDescCcl)
''
''    Next
''
''    Carga_Arvore_Ccl = SUCESSO
''
''    Exit Function
''
''Erro_Carga_Arvore_Ccl:
''
''    Carga_Arvore_Ccl = Err
''
''    Select Case Err
''
''        Case 34205
''
''        Case 34206
''            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objCcl.sCcl)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 170509)
''
''    End Select
''
''    Exit Function
''
''End Function

Private Sub CclFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CclFinal_Validate

    giCclInicial = 0

    lErro = Ccl_Perde_Foco(CclFinal, DescCclFim)
    If lErro <> SUCESSO Then Error 34207

    Exit Sub

Erro_CclFinal_Validate:

    Cancel = True


    Select Case Err

        Case 34207

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170510)

    End Select

    Exit Sub

End Sub

Private Sub CclInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CclInicial_Validate

    giCclInicial = 1

    lErro = Ccl_Perde_Foco(CclInicial, DescCclInic)
    If lErro <> SUCESSO Then Error 34208

    Exit Sub

Erro_CclInicial_Validate:

    Cancel = True


    Select Case Err

        Case 34208

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170511)

    End Select

    Exit Sub

End Sub

Private Sub TvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)

Dim sCcl As String
Dim sCclMascarado As String
Dim lErro As Long
Dim lPosicaoSeparador As Long

On Error GoTo Erro_TvwCcls_NodeClick

    sCcl = Right(Node.Key, Len(Node.Key) - 1)

    sCclMascarado = String(STRING_CCL, 0)

    lErro = Mascara_MascararCcl(sCcl, sCclMascarado)
    If lErro <> SUCESSO Then Error 34209

    If giCclInicial = 1 Then
        CclInicial.PromptInclude = False
        CclInicial.Text = sCclMascarado
        CclInicial.PromptInclude = True
    Else
        CclFinal.PromptInclude = False
        CclFinal.Text = sCclMascarado
        CclFinal.PromptInclude = True
    End If

    'Preenche a Descricao do centro de custo/lucro
    lPosicaoSeparador = InStr(Node.Text, SEPARADOR)

    If giCclInicial = 1 Then
        DescCclInic.Caption = Mid(Node.Text, lPosicaoSeparador + 1)
    Else
        DescCclFim.Caption = Mid(Node.Text, lPosicaoSeparador + 1)
    End If

    Exit Sub

Erro_TvwCcls_NodeClick:

    Select Case Err

        Case 34209
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 170512)

    End Select

    Exit Sub

End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim sDataFim As String
Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(DataFinal.ClipText) > 0 Then

        sDataFim = DataFinal.Text
        lErro = Data_Critica(sDataFim)
        If lErro <> SUCESSO Then Error 34210

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True


    Select Case Err

        Case 34210

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170513)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim sDataInic As String
Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(DataInicial.ClipText) > 0 Then

        sDataInic = DataInicial.Text
        lErro = Data_Critica(sDataInic)
        If lErro <> SUCESSO Then Error 34211

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True


    Select Case Err

        Case 34211

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170514)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(DataInicial, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 34212

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 34212
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170515)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataInicial, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 34213

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 34213
            DataInicial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170516)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    lErro = Data_Up_Down_Click(DataFinal, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 34214

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 34214
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170517)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    lErro = Data_Up_Down_Click(DataFinal, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 34215

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 34215
            DataFinal.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170518)

    End Select

    Exit Sub

End Sub

Private Function Valida_OrdProd(sCodigoOP As String) As Long

Dim objOp As New ClassOrdemDeProducao
Dim lErro As Long
Dim iIndice As Integer
Dim iOpStatus As Integer

On Error GoTo Erro_Valida_OrdProd

    objOp.iFilialEmpresa = giFilialEmpresa
    objOp.sCodigo = sCodigoOP
    
    'verifica status selecionado
    For iIndice = 0 To 2
       If OpStatus(iIndice).Value = True Then iOpStatus = iIndice
    Next

    'se a opção de status é "Abertas"
    If iOpStatus = 1 Then

        lErro = CF("OrdemDeProducao_Le_SemItens", objOp)
        If lErro <> SUCESSO And lErro <> 34455 Then Error 34278

        If lErro = 34455 Then Error 34279
    
    End If

    'Se a opção de status é "Encerradas"
    If iOpStatus = 2 Then

        lErro = CF("OPBaixada_Le_SemItens", objOp)
        If lErro <> SUCESSO And lErro <> 34459 Then Error 34280

        If lErro = 34459 Then Error 34281

    End If
    
    'Se a opção de status é "Todas"
    If iOpStatus = 0 Then

        lErro = CF("OrdemDeProducao_Le_SemItens", objOp)
        If lErro <> SUCESSO And lErro <> 34455 Then

            Error 34282

        Else
        
            If lErro <> SUCESSO Then

                lErro = CF("OPBaixada_Le_SemItens", objOp)
                If lErro <> SUCESSO And lErro <> 34459 Then Error 34283

                If lErro = 34459 Then Error 34284
            
            End If


        End If

    End If

    Valida_OrdProd = SUCESSO

    Exit Function

Erro_Valida_OrdProd:

    Valida_OrdProd = Err

    Select Case Err

        Case 34278
        
        Case 34279
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OP_ABERTA_INEXISTENTE", Err)
            
        Case 34280

        Case 34281
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OP_ENCERRADA_INEXISTENTE", Err)
            
        Case 34282, 34283
            
        Case 34284
           lErro = Rotina_Erro(vbOKOnly, "ERRO_OP_INEXISTENTE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170519)

    End Select

    Exit Function

End Function

Private Sub OpInicial_Validate(Cancel As Boolean)
Dim lErro As Long

On Error GoTo Erro_OpInicial_Validate

    giOp_Inicial = 1

    If Len(Trim(OpInicial.Text)) <> 0 Then

        lErro = Valida_OrdProd(OpInicial.Text)
        If lErro <> SUCESSO Then Error 34285
        
    End If

    Exit Sub

Erro_OpInicial_Validate:

    Cancel = True


    Select Case Err

        Case 34285
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170520)

    End Select

    Exit Sub

End Sub

Private Sub OpFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_OpFinal_Validate

    giOp_Inicial = 0

    If Len(Trim(OpFinal.Text)) <> 0 Then

        lErro = Valida_OrdProd(OpFinal.Text)
        If lErro <> SUCESSO Then Error 34286
    
    End If

    Exit Sub

Erro_OpFinal_Validate:

    Cancel = True


    Select Case Err

        Case 34286
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170521)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame4(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame4(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index

    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170522)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_ORDEM_PRODUCAO
    Set Form_Load_Ocx = Me
    Caption = "Ordens de Produção"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpOrdemProducao"
    
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
   ' Parent.UnloadDoFilho
    
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
        
        If Me.ActiveControl Is OpInicial Then
            Call LabelOpInicial_Click
        ElseIf Me.ActiveControl Is OpFinal Then
            Call LabelOpFinal_Click
        ElseIf Me.ActiveControl Is ProdutoInicial Then
            Call LabelProdutoDe_Click
        ElseIf Me.ActiveControl Is ProdutoFinal Then
            Call LabelProdutoAte_Click
        End If
    
    End If

End Sub





Private Sub LabelOpFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOpFinal, Source, X, Y)
End Sub

Private Sub LabelOpFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOpFinal, Button, Shift, X, Y)
End Sub

Private Sub LabelOpInicial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOpInicial, Source, X, Y)
End Sub

Private Sub LabelOpInicial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOpInicial, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub DescProdInic_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescProdInic, Source, X, Y)
End Sub

Private Sub DescProdInic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescProdInic, Button, Shift, X, Y)
End Sub

Private Sub DescProdFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescProdFim, Source, X, Y)
End Sub

Private Sub DescProdFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescProdFim, Button, Shift, X, Y)
End Sub

Private Sub LabelProdutoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProdutoDe, Source, X, Y)
End Sub

Private Sub LabelProdutoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProdutoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelProdutoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelProdutoAte, Source, X, Y)
End Sub

Private Sub LabelProdutoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelProdutoAte, Button, Shift, X, Y)
End Sub

Private Sub dIni_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dIni, Source, X, Y)
End Sub

Private Sub dIni_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dIni, Button, Shift, X, Y)
End Sub

Private Sub dFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(dFim, Source, X, Y)
End Sub

Private Sub dFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(dFim, Button, Shift, X, Y)
End Sub

Private Sub DescCclInic_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescCclInic, Source, X, Y)
End Sub

Private Sub DescCclInic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescCclInic, Button, Shift, X, Y)
End Sub

Private Sub LabelCclDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCclDe, Source, X, Y)
End Sub

Private Sub LabelCclDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCclDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCclAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCclAte, Source, X, Y)
End Sub

Private Sub LabelCclAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCclAte, Button, Shift, X, Y)
End Sub

Private Sub DescCclFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescCclFim, Source, X, Y)
End Sub

Private Sub DescCclFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescCclFim, Button, Shift, X, Y)
End Sub

Private Sub LabelCcl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCcl, Source, X, Y)
End Sub

Private Sub LabelCcl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCcl, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub


Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

