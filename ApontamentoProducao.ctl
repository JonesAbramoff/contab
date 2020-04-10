VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ApontamentoProducao 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleMode       =   0  'User
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5145
      Index           =   2
      Left            =   135
      TabIndex        =   23
      Top             =   750
      Visible         =   0   'False
      Width           =   9255
      Begin VB.Frame Frame4 
         Caption         =   "Ordem de Produção"
         Height          =   3240
         Left            =   60
         TabIndex        =   26
         Top             =   15
         Width           =   9150
         Begin VB.CheckBox TemApontamento 
            Height          =   315
            Left            =   2475
            TabIndex        =   57
            Top             =   1005
            Width           =   540
         End
         Begin VB.CommandButton BotaoAlterar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   6150
            Picture         =   "ApontamentoProducao.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Altera os dados do Apontamento no Grid"
            Top             =   2685
            Width           =   1380
         End
         Begin VB.CommandButton BotaoRemover 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   7650
            Picture         =   "ApontamentoProducao.ctx":1926
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Exclui os dados do Apontamento do Grid"
            Top             =   2685
            Width           =   1380
         End
         Begin MSMask.MaskEdBox Compet 
            Height          =   315
            Left            =   4590
            TabIndex        =   40
            Top             =   1005
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Etapa 
            Height          =   315
            Left            =   1185
            TabIndex        =   35
            Top             =   1020
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   3
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox UMedidaOP 
            Height          =   315
            Left            =   4965
            TabIndex        =   34
            Top             =   1605
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   5
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantidadeOP 
            Height          =   315
            Left            =   4005
            TabIndex        =   33
            Top             =   1605
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox VersaoProdOP 
            Height          =   315
            Left            =   3375
            TabIndex        =   32
            Top             =   1605
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoOP 
            Height          =   315
            Left            =   6195
            TabIndex        =   31
            Top             =   1005
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.TextBox DescProdutoOP 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   1695
            TabIndex        =   30
            Top             =   1605
            Width           =   1700
         End
         Begin VB.CheckBox Concluido 
            Height          =   315
            Left            =   405
            TabIndex        =   29
            Top             =   1020
            Width           =   795
         End
         Begin MSMask.MaskEdBox NumeroOP 
            Height          =   315
            Left            =   1695
            TabIndex        =   28
            Top             =   1020
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   6
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CT 
            Height          =   315
            Left            =   3000
            TabIndex        =   27
            Top             =   1005
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridApontamentos 
            Height          =   2040
            Left            =   90
            TabIndex        =   11
            Top             =   255
            Width           =   8955
            _ExtentX        =   15796
            _ExtentY        =   3598
            _Version        =   393216
            Rows            =   6
            Cols            =   3
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
         End
      End
      Begin VB.Frame FrameApontamentos 
         Caption         =   "Apontamento"
         Height          =   1815
         Left            =   60
         TabIndex        =   25
         Top             =   3285
         Width           =   9150
         Begin VB.Frame Frame2 
            Caption         =   "Apontamento Anterior"
            Height          =   675
            Left            =   90
            TabIndex        =   60
            Top             =   1050
            Width           =   8955
            Begin VB.Label LabelPercentualAnterior 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   7920
               TabIndex        =   66
               Top             =   255
               Width           =   885
            End
            Begin VB.Label LabelQuantidadeAnterior 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   4545
               TabIndex        =   65
               Top             =   255
               Width           =   1290
            End
            Begin VB.Label LabelDataAnterior 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1230
               TabIndex        =   64
               Top             =   255
               Width           =   1305
            End
            Begin VB.Label Label3 
               Caption         =   "% Concluído:"
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
               Left            =   6735
               TabIndex        =   63
               Top             =   285
               Width           =   1170
            End
            Begin VB.Label Label2 
               Caption         =   "Quantidade:"
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
               Left            =   3450
               TabIndex        =   62
               Top             =   285
               Width           =   1080
            End
            Begin VB.Label Label1 
               Caption         =   "Data:"
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
               Left            =   690
               TabIndex        =   61
               Top             =   285
               Width           =   525
            End
         End
         Begin VB.TextBox Observacao 
            Height          =   315
            Left            =   1320
            MaxLength       =   255
            TabIndex        =   15
            Top             =   675
            Width           =   7605
         End
         Begin MSMask.MaskEdBox PercConcluido 
            Height          =   315
            Left            =   8040
            TabIndex        =   14
            Top             =   255
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   315
            Left            =   4650
            TabIndex        =   13
            Top             =   255
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   315
            Left            =   1320
            TabIndex        =   12
            Top             =   255
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownData 
            Height          =   300
            Left            =   2625
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   255
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label LabelData 
            Caption         =   "Data:"
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
            Left            =   780
            TabIndex        =   59
            Top             =   285
            Width           =   525
         End
         Begin VB.Label LabelQuantidade 
            Caption         =   "Quantidade:"
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
            Left            =   3540
            TabIndex        =   38
            Top             =   285
            Width           =   1080
         End
         Begin VB.Label LabelObservacao 
            Caption         =   "Observação:"
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
            Left            =   150
            TabIndex        =   37
            Top             =   675
            Width           =   1125
         End
         Begin VB.Label LabelPercConcluido 
            Caption         =   "% Concluído:"
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
            Left            =   6825
            TabIndex        =   36
            Top             =   285
            Width           =   1170
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5115
      Index           =   1
      Left            =   150
      TabIndex        =   24
      Top             =   750
      Width           =   9210
      Begin VB.Frame FrameDataOP 
         Caption         =   "Data da O.P."
         Height          =   810
         Left            =   240
         TabIndex        =   54
         Top             =   4200
         Width           =   8790
         Begin MSMask.MaskEdBox DataOPInicial 
            Height          =   300
            Left            =   915
            TabIndex        =   9
            Top             =   315
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataOPInicial 
            Height          =   300
            Left            =   2070
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   315
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataOPFinal 
            Height          =   300
            Left            =   4095
            TabIndex        =   10
            Top             =   315
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataOPFinal 
            Height          =   300
            Left            =   5265
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   315
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
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
            Height          =   195
            Index           =   1
            Left            =   3660
            TabIndex        =   56
            Top             =   330
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
            Height          =   195
            Index           =   0
            Left            =   570
            TabIndex        =   55
            Top             =   330
            Width           =   315
         End
      End
      Begin VB.Frame FrameOP 
         Caption         =   "Ordem de Produção"
         Height          =   810
         Left            =   240
         TabIndex        =   51
         Top             =   1920
         Width           =   8790
         Begin VB.TextBox OpFinal 
            Height          =   300
            Left            =   4125
            MaxLength       =   6
            TabIndex        =   6
            Top             =   330
            Width           =   1680
         End
         Begin VB.TextBox OpInicial 
            Height          =   300
            Left            =   930
            MaxLength       =   6
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
            TabIndex        =   53
            Top             =   345
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
            Left            =   525
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   52
            Top             =   360
            Width           =   315
         End
      End
      Begin VB.Frame FrameCT 
         Caption         =   "Centros de Trabalho"
         Height          =   1395
         Left            =   240
         TabIndex        =   46
         Top             =   495
         Width           =   8790
         Begin MSMask.MaskEdBox CTInicial 
            Height          =   315
            Left            =   945
            TabIndex        =   3
            Top             =   360
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CTFinal 
            Height          =   315
            Left            =   945
            TabIndex        =   4
            Top             =   840
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label LabelCTAte 
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
            Height          =   255
            Left            =   555
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   50
            Top             =   885
            Width           =   435
         End
         Begin VB.Label LabelCTDe 
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
            Height          =   255
            Left            =   585
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   49
            Top             =   390
            Width           =   360
         End
         Begin VB.Label DescCTInicial 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2520
            TabIndex        =   48
            Top             =   360
            Width           =   5940
         End
         Begin VB.Label DescCTFinal 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2520
            TabIndex        =   47
            Top             =   840
            Width           =   5940
         End
      End
      Begin VB.Frame FrameProdutos 
         Caption         =   "Produtos"
         Height          =   1395
         Index           =   0
         Left            =   240
         TabIndex        =   41
         Top             =   2760
         Width           =   8790
         Begin MSMask.MaskEdBox ProdutoInicial 
            Height          =   315
            Left            =   915
            TabIndex        =   7
            Top             =   345
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoFinal 
            Height          =   315
            Left            =   915
            TabIndex        =   8
            Top             =   825
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label DescProdFim 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2490
            TabIndex        =   45
            Top             =   825
            Width           =   5940
         End
         Begin VB.Label DescProdInic 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2490
            TabIndex        =   44
            Top             =   345
            Width           =   5940
         End
         Begin VB.Label LabelProdutoDe 
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
            Height          =   255
            Left            =   555
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   43
            Top             =   375
            Width           =   360
         End
         Begin VB.Label LabelProdutoAte 
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
            Height          =   255
            Left            =   525
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   42
            Top             =   870
            Width           =   435
         End
      End
      Begin VB.OptionButton SoAbertos 
         Caption         =   "Só Apontamentos em Aberto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1455
         TabIndex        =   1
         Top             =   225
         Width           =   2745
      End
      Begin VB.OptionButton Todos 
         Caption         =   "Todos Apontamentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5295
         TabIndex        =   2
         Top             =   225
         Width           =   2190
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   7815
      ScaleHeight     =   480
      ScaleWidth      =   1530
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   75
      Width           =   1590
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ApontamentoProducao.ctx":324C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   570
         Picture         =   "ApontamentoProducao.ctx":33A6
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1050
         Picture         =   "ApontamentoProducao.ctx":38D8
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5520
      Left            =   105
      TabIndex        =   0
      Top             =   405
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   9737
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Apontamentos"
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
End
Attribute VB_Name = "ApontamentoProducao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim bRefreshApontamentos As Boolean

Dim iFrameAtual As Integer

Dim iTabPrincipalAlterado As Integer
Dim iQtdeAlterada As Integer

Dim colApontamentos As New Collection

Dim giOp_Inicial As Integer

'Grid Apontamentos
Dim objGridApontamentos As AdmGrid
Dim iGrid_Concluido_Col As Integer
Dim iGrid_Etapa_Col As Integer
Dim iGrid_NumeroOP_Col As Integer
Dim iGrid_TemApont_Col As Integer
Dim iGrid_CT_Col As Integer
Dim iGrid_Compet_Col As Integer
Dim iGrid_ProdutoOP_Col As Integer
Dim iGrid_DescProdutoOP_Col As Integer
Dim iGrid_VersaoProdOP_Col As Integer
Dim iGrid_QuantidadeOP_Col As Integer
Dim iGrid_UMedidaOP_Col As Integer

Private WithEvents objEventoCTInic As AdmEvento
Attribute objEventoCTInic.VB_VarHelpID = -1
Private WithEvents objEventoCTFim As AdmEvento
Attribute objEventoCTFim.VB_VarHelpID = -1

Private WithEvents objEventoOp As AdmEvento
Attribute objEventoOp.VB_VarHelpID = -1

Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1

Private Const TAB_Selecao = 1
Private Const TAB_Apontamentos = 2

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Apontamentos da Produção"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ApontamentoProducao"

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

Private Sub BotaoAlterar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoAlterar_Click

    'Se não tiver linha selecionada => Erro
    If GridApontamentos.Row = 0 Then gError 137728
    
    'Verifica se a linha selecionada está preenchida
    If Len(GridApontamentos.TextMatrix(GridApontamentos.Row, iGrid_Etapa_Col)) = 0 Then gError 137729

    'Verifica se Percentual está vazio
    If Len(Trim(PercConcluido.Text)) = 0 Then gError 137730
    
    'Verifica se a Quantidade está vazia
    If Len(Trim(Quantidade.Text)) = 0 Then gError 137731

    'Verifica se a Data está vazia
    If Len(Trim(Data.ClipText)) = 0 Then gError 137732
    
    'altera a coleção
    With colApontamentos.Item(GridApontamentos.Row)
    
        .dPercConcluido = StrParaDbl(PercConcluido.Text / 100)
        .dQuantidade = StrParaDbl(Quantidade.Text)
        .dtData = StrParaDate(Data.Text)
        .sObservacao = Observacao.Text
    
    End With
    
    'marca Apontamentos no Grid
    GridApontamentos.TextMatrix(GridApontamentos.Row, iGrid_TemApont_Col) = MARCADO
    
    'Atualiza na tela a checkbox marcada
    Call Grid_Refresh_Checkbox(objGridApontamentos)

    Exit Sub

Erro_BotaoAlterar_Click:

    Select Case gErr
    
        Case 137728
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 137729
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_PREENCHIDA", gErr)
    
        Case 137730
            Call Rotina_Erro(vbOKOnly, "ERRO_PERCENT_APONTAMENTO_NAO_PREENCHIDO", gErr)

        Case 137731
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_APONTAMENTO_NAO_PREENCHIDA", gErr)

        Case 137732
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_APONTAMENTO_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142952)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 137734

    'Limpa Tela
    Call Limpa_Tela_ApontamentoProducao

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 137734

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142953)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    Call Limpa_Tela_ApontamentoProducao

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142954)

    End Select

    Exit Sub

End Sub

Private Sub BotaoRemover_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoRemover_Click

    'Se não tiver linha selecionada => Erro
    If GridApontamentos.Row = 0 Then gError 137735
    
    'Verifica se a linha selecionada está preenchida
    If Len(GridApontamentos.TextMatrix(GridApontamentos.Row, iGrid_Etapa_Col)) = 0 Then gError 137736

    'limpa os campos do frame
    PercConcluido.Text = ""
    Quantidade.Text = ""
    
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
        
    Observacao.Text = ""

    LabelDataAnterior = ""
    LabelQuantidadeAnterior = ""
    LabelPercentualAnterior = ""

    'limpa obj da coleção
    With colApontamentos.Item(GridApontamentos.Row)
        .dPercConcluido = 0
        .dQuantidade = 0
        .sObservacao = ""
        .dtData = 0
    End With
    
    'desmarca Apontamentos no Grid
    GridApontamentos.TextMatrix(GridApontamentos.Row, iGrid_TemApont_Col) = DESMARCADO
    
    'Atualiza na tela a checkbox desmarcada
    Call Grid_Refresh_Checkbox(objGridApontamentos)

    Exit Sub

Erro_BotaoRemover_Click:

    Select Case gErr
    
        Case 137735
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 137736
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_PREENCHIDA", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142955)

    End Select

    Exit Sub

End Sub

Private Sub CTFinal_Change()
    
    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CTFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CTFinal, iAlterado)
    
End Sub

Private Sub CTFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_CTFinal_Validate

    DescCTFinal.Caption = ""

    'Verifica se CTFinal não está preenchido
    If Len(Trim(CTFinal.Text)) <> 0 Then
    
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
        
        'Procura pela empresa toda
        objCentrodeTrabalho.iFilialEmpresa = giFilialEmpresa
        
        'Verifica sua existencia
        lErro = CF("TP_CentrodeTrabalho_Le", CTFinal, objCentrodeTrabalho)
        If lErro <> SUCESSO Then gError 137737
                
        DescCTFinal.Caption = objCentrodeTrabalho.sDescricao
           
    End If
    
    Exit Sub

Erro_CTFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 137737
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142956)

    End Select

    Exit Sub

End Sub

Private Sub CTInicial_Change()
    
    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CTInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CTInicial, iAlterado)
    
End Sub

Private Sub CTInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_CTInicial_Validate

    DescCTInicial.Caption = ""

    'Verifica se CTInicial não está preenchido
    If Len(Trim(CTInicial.Text)) <> 0 Then

        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
        
        'Procura pela empresa toda
        objCentrodeTrabalho.iFilialEmpresa = giFilialEmpresa
        
        'Verifica sua existencia
        lErro = CF("TP_CentrodeTrabalho_Le", CTInicial, objCentrodeTrabalho)
        If lErro <> SUCESSO Then gError 137738
                
        DescCTInicial.Caption = objCentrodeTrabalho.sDescricao
       
    End If
    
    Exit Sub

Erro_CTInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 137738
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142957)

    End Select

    Exit Sub

End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long
Dim lIntervalo As Long

On Error GoTo Erro_Data_Validate

    'Verifica se Data está preenchida
    If Len(Trim(Data.ClipText)) <> 0 Then

        'Critica a Data
        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError 137739
        
        'Verifica qual é o intervalo entre as datas (data de início de operações MRP)
        lIntervalo = DateDiff("d", gobjEST.dtDataInicioMRP, StrParaDate(Data.Text))
        
        'Se o intervalo for negativo -> Erro
        If lIntervalo < 0 Then gError 137740
        
    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case 137739
        
        Case 137740
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_MENOR_DATAINICIO_MRP", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142958)

    End Select

    Exit Sub

End Sub

Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)
    
End Sub

Private Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataOPFinal_Change()
    
    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub


Private Sub GridApontamentos_LostFocus()

    Call Grid_Libera_Foco(objGridApontamentos)

End Sub

Private Sub LabelCTAte_Click()

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCTAte

    'Verifica se o CTFinal foi preenchido
    If Len(Trim(CTFinal.Text)) <> 0 Then
            
        objCentrodeTrabalho.sNomeReduzido = CTFinal.Text
        
    End If

    Call Chama_Tela("CentrodeTrabalhoLista", colSelecao, objCentrodeTrabalho, objEventoCTFim)

    Exit Sub

Erro_LabelCTAte:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142959)

    End Select

    Exit Sub

End Sub

Private Sub LabelCTDe_Click()

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCTDe

    'Verifica se o CTInicial foi preenchido
    If Len(Trim(CTInicial.Text)) <> 0 Then
    
        objCentrodeTrabalho.sNomeReduzido = CTInicial.Text
        
    End If

    Call Chama_Tela("CentrodeTrabalhoLista", colSelecao, objCentrodeTrabalho, objEventoCTInic)

    Exit Sub

Erro_LabelCTDe:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142960)

    End Select

    Exit Sub

End Sub


Private Sub LabelOpFinal_Click()

Dim lErro As Long
Dim colSelecao As Collection
Dim objOP As New ClassOrdemDeProducao
Dim sSelecao As String

On Error GoTo Erro_LabelOpFinal_Click

    giOp_Inicial = 0
    
    If Len(Trim(OpFinal.Text)) <> 0 Then

        objOP.sCodigo = OpFinal.Text

    End If
    sSelecao = "Tipo = 0"
    
    'lista as OP's
    Call Chama_Tela("OrdemProducaoLista", colSelecao, objOP, objEventoOp, sSelecao)
    
    Exit Sub

Erro_LabelOpFinal_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142961)

    End Select

    Exit Sub

End Sub

Private Sub LabelOpInicial_Click()

Dim lErro As Long
Dim colSelecao As Collection
Dim objOP As New ClassOrdemDeProducao
Dim sSelecao As String

On Error GoTo Erro_LabelOpInicial_Click

    giOp_Inicial = 1

    If Len(Trim(OpInicial.Text)) <> 0 Then
    
        objOP.sCodigo = OpInicial.Text

    End If

    sSelecao = "Tipo = 0"
    
    'lista as OP's
    Call Chama_Tela("OrdemProducaoLista", colSelecao, objOP, objEventoOp, sSelecao)
    
    Exit Sub

Erro_LabelOpInicial_Click:

    Select Case gErr

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142962)

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
        If lErro <> SUCESSO Then gError 137741

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoProduzivelLista", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 137741
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142963)

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
        If lErro <> SUCESSO Then gError 137742

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoProduzivelLista", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 137742
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142964)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCTFim_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_objEventoCTFim_evSelecao

    Set objCentrodeTrabalho = obj1

    CTFinal.Text = objCentrodeTrabalho.sNomeReduzido
        
    Call CTFinal_Validate(bSGECancelDummy)
        
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCTFim_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142965)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCTInic_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_objEventoCTInic_evSelecao

    Set objCentrodeTrabalho = obj1

    CTInicial.Text = objCentrodeTrabalho.sNomeReduzido
        
    Call CTInicial_Validate(bSGECancelDummy)
        
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCTInic_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142966)

    End Select

    Exit Sub

End Sub

Private Sub objEventoOp_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOP As New ClassOrdemDeProducao

On Error GoTo Erro_objEventoOp_evSelecao

    Set objOP = obj1

    If giOp_Inicial = 1 Then

        OpInicial.Text = objOP.sCodigo
        
    Else

        OpFinal.Text = objOP.sCodigo

    End If

    Me.Show
    
    Exit Sub

Erro_objEventoOp_evSelecao:

    Select Case Err

       Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142967)

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 137743

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 137744

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 137745

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 137743, 137745
            'erro tratado na rotina chamada

        Case 137744
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142968)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 137746

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 137747

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 137748

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 137746, 137748
            'erro tratado na rotina chamada

        Case 137747
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142969)

    End Select

    Exit Sub

End Sub

Private Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub OpFinal_Change()
    
    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub OpFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_OpFinal_Validate

    giOp_Inicial = 0

    If Len(Trim(OpFinal.Text)) <> 0 Then


    End If

    Exit Sub

Erro_OpFinal_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142970)

    End Select

    Exit Sub

End Sub

Private Sub OpInicial_Change()
    
    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub OpInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_OpInicial_Validate

    giOp_Inicial = 1

    If Len(Trim(OpFinal.Text)) <> 0 Then

    
    End If

    Exit Sub

Erro_OpInicial_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142971)

    End Select

    Exit Sub

End Sub

Private Sub PercConcluido_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dQtdeTotalOP As Double
Dim dPercConcluido As Double
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_PercConcluido_Validate

    'Veifica se PercConcluido está preenchida
    If Len(Trim(PercConcluido.Text)) <> 0 Then
        
        'Se o percentual atual é menor que o percentual que já estava apontado...
        If StrParaDbl(Val(PercConcluido.Text)) < StrParaDbl(Val(LabelPercentualAnterior.Caption)) And iQtdeAlterada = REGISTRO_ALTERADO Then
        
            'Pergunta ao usuário se confirma o percentual digitado
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_PERCCONCLUIDO_MENOR")
        
            If vbMsgRes = vbNo Then gError 137749
        
        End If
        
        dQtdeTotalOP = StrParaDbl(GridApontamentos.TextMatrix(GridApontamentos.Row, iGrid_QuantidadeOP_Col))
        
        dPercConcluido = StrParaDbl(Val(PercConcluido.Text) / 100)
        
        Quantidade.Text = Formata_Estoque(dQtdeTotalOP * dPercConcluido)
        
    End If
    iQtdeAlterada = 0
    
    Exit Sub

Erro_PercConcluido_Validate:

    Cancel = True

    Select Case gErr
    
        Case 137749

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142972)

    End Select

    Exit Sub

End Sub

Private Sub PercConcluido_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(PercConcluido, iAlterado)
    
End Sub

Private Sub PercConcluido_Change()

    iAlterado = REGISTRO_ALTERADO
    iQtdeAlterada = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoFinal_Change()
    
    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ProdutoFinal, iAlterado)
    
End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 137750
    
    If lErro <> SUCESSO Then gError 137751
  
    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 137750
            'erro tratado na rotina chamada

        Case 137751
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142973)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoInicial_Change()
    
    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ProdutoInicial, iAlterado)
    
End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 137752
    
    If lErro <> SUCESSO Then gError 137753

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 137752
            'erro tratado na rotina chamada
            
        Case 137753
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142974)

    End Select

    Exit Sub

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dQtdeTotalOP As Double
Dim dPercConcluido As Double
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Quantidade_Validate

    'Verifica se Quantidade está preenchida
    If Len(Trim(Quantidade.Text)) <> 0 Then

        'Critica a Quantidade do Apontamento
        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 137754
        
        'Se a quantidade atual é menor que a quantidade que já estava apontada...
        If StrParaDbl(Quantidade.Text) < StrParaDbl(LabelQuantidadeAnterior.Caption) And iQtdeAlterada = REGISTRO_ALTERADO Then
        
            'Pergunta ao usuário se confirma a quantidade digitada
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_QUANTIDADE_MENOR")
        
            If vbMsgRes = vbNo Then gError 137839
        
        End If
        
        dQtdeTotalOP = StrParaDbl(GridApontamentos.TextMatrix(GridApontamentos.Row, iGrid_QuantidadeOP_Col))
        
        dPercConcluido = StrParaDbl(Quantidade.Text) / dQtdeTotalOP
        
        Quantidade.Text = Formata_Estoque(StrParaDbl(Quantidade.Text))
        PercConcluido.Text = Format(dPercConcluido * 100, "###.##")

    End If
    iQtdeAlterada = 0

    Exit Sub

Erro_Quantidade_Validate:

    Cancel = True

    Select Case gErr

        Case 137754, 137839

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142975)

    End Select

    Exit Sub

End Sub

Private Sub Quantidade_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Quantidade, iAlterado)
    
End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO
    iQtdeAlterada = REGISTRO_ALTERADO

End Sub

Private Sub SoAbertos_Click()
    
    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Opcao_BeforeClick

    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
    
    'Se estava no tab de seleção e está passando para outro tab
    If iFrameAtual = TAB_Selecao Then
    
        'Valida a seleção
        lErro = ValidaSelecao()
        If lErro <> SUCESSO Then gError 137755
    
    End If

    Exit Sub

Erro_Opcao_BeforeClick:

    Cancel = True

    Select Case gErr

        Case 137755

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142976)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
        
        'Se Frame selecionado foi o de Apontamentos
        If TabStrip1.SelectedItem.Index = TAB_Apontamentos Then
            If iTabPrincipalAlterado = REGISTRO_ALTERADO Then
                Call Trata_TabApontamentos
            End If
        
        End If
        
    End If

    Exit Sub

End Sub

Private Sub Todos_Click()
    
    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty(True, UserControl.Enabled, True)
End Sub

Private Sub Unload(objme As Object)
   RaiseEvent Unload
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****
    
Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoCTInic = Nothing
    Set objEventoCTFim = Nothing

    Set objEventoOp = Nothing
    
    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing
    
    Set objGridApontamentos = Nothing

    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142977)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    iFrameAtual = 1
    
    Set objEventoCTInic = New AdmEvento
    Set objEventoCTFim = New AdmEvento
    
    Set objEventoOp = New AdmEvento
    
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento

    'inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then gError 134587

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then gError 134588
        
    'GridApontamentos
    Set objGridApontamentos = New AdmGrid
    
    'tela em questão
    Set objGridApontamentos.objForm = Me
    
    lErro = Inicializa_GridApontamentos(objGridApontamentos)
    If lErro <> SUCESSO Then gError 137758
    
    SoAbertos.Value = True
    Todos.Value = False
    
    iTabPrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = 0
    iQtdeAlterada = 0
    bRefreshApontamentos = True

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142978)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142979)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataOPFinal_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataOPFinal_DownClick

    DataOPFinal.SetFocus

    If Len(DataOPFinal.ClipText) > 0 Then

        sData = DataOPFinal.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 137759

        DataOPFinal.Text = sData

    End If

    Exit Sub

Erro_UpDownDataOPFinal_DownClick:

    Select Case gErr

        Case 137759

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142980)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataOPFinal_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataOPFinal_UpClick

    DataOPFinal.SetFocus

    If Len(Trim(DataOPFinal.ClipText)) > 0 Then

        sData = DataOPFinal.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 137760

        DataOPFinal.Text = sData

    End If

    Exit Sub

Erro_UpDownDataOPFinal_UpClick:

    Select Case gErr

        Case 137760

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142981)

    End Select

    Exit Sub

End Sub

Private Sub DataOPInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataOPInicial, iAlterado)
    
End Sub

Private Sub DataOPInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataOPInicial_Validate

    If Len(Trim(DataOPInicial.ClipText)) <> 0 Then

        lErro = Data_Critica(DataOPInicial.Text)
        If lErro <> SUCESSO Then gError 137761

    End If

    Exit Sub

Erro_DataOPInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 137761

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142982)

    End Select

    Exit Sub

End Sub

Private Sub DataOPInicial_Change()
    
    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownDataOPInicial_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataOPInicial_DownClick

    DataOPInicial.SetFocus

    If Len(DataOPInicial.ClipText) > 0 Then

        sData = DataOPInicial.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 137762

        DataOPInicial.Text = sData

    End If

    Exit Sub

Erro_UpDownDataOPInicial_DownClick:

    Select Case gErr

        Case 137762

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142983)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataOPInicial_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataOPInicial_UpClick

    DataOPInicial.SetFocus

    If Len(Trim(DataOPInicial.ClipText)) > 0 Then

        sData = DataOPInicial.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 137763

        DataOPInicial.Text = sData

    End If

    Exit Sub

Erro_UpDownDataOPInicial_UpClick:

    Select Case gErr

        Case 137763

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142984)

    End Select

    Exit Sub

End Sub

Private Sub DataOPFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataOPFinal, iAlterado)
    
End Sub

Private Sub DataOPFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataOPFinal_Validate

    If Len(Trim(DataOPFinal.ClipText)) <> 0 Then

        lErro = Data_Critica(DataOPFinal.Text)
        If lErro <> SUCESSO Then gError 137764

    End If

    Exit Sub

Erro_DataOPFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 137764

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142985)

    End Select

    Exit Sub

End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable

    If objControl.Name = "Concluido" Then
    
        If Len(GridApontamentos.TextMatrix(GridApontamentos.Row, iGrid_NumeroOP_Col)) <> 0 Then
    
            objControl.Enabled = True
        
        Else
        
            objControl.Enabled = False
        
        End If

    Else

        objControl.Enabled = False
    
    End If
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 142986)

    End Select

    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
            
        'Verifica qual a coluna do Grid em questão
        Select Case objGridInt.objGrid.Col
            
            Case iGrid_Concluido_Col
            
                lErro = Saida_Celula_Concluido(objGridInt)
                If lErro <> SUCESSO Then gError 137765
    
        End Select
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 137766

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 137765
            'erros tratatos nas rotinas chamadas
        
        Case 137766
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142987)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridApontamentos(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Concluido")
    objGrid.colColuna.Add ("Etapa")
    objGrid.colColuna.Add ("O.P.")
    objGrid.colColuna.Add ("Apont.")
    objGrid.colColuna.Add ("CT")
    objGrid.colColuna.Add ("Compet.")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("Versão")
    objGrid.colColuna.Add ("U.M.")
    objGrid.colColuna.Add ("Quantidade")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Concluido.Name)
    objGrid.colCampo.Add (Etapa.Name)
    objGrid.colCampo.Add (NumeroOP.Name)
    objGrid.colCampo.Add (TemApontamento.Name)
    objGrid.colCampo.Add (CT.Name)
    objGrid.colCampo.Add (Compet.Name)
    objGrid.colCampo.Add (ProdutoOP.Name)
    objGrid.colCampo.Add (DescProdutoOP.Name)
    objGrid.colCampo.Add (VersaoProdOP.Name)
    objGrid.colCampo.Add (UMedidaOP.Name)
    objGrid.colCampo.Add (QuantidadeOP.Name)

    'Colunas do Grid
    iGrid_Concluido_Col = 1
    iGrid_Etapa_Col = 2
    iGrid_NumeroOP_Col = 3
    iGrid_TemApont_Col = 4
    iGrid_CT_Col = 5
    iGrid_Compet_Col = 6
    iGrid_ProdutoOP_Col = 7
    iGrid_DescProdutoOP_Col = 8
    iGrid_VersaoProdOP_Col = 9
    iGrid_UMedidaOP_Col = 10
    iGrid_QuantidadeOP_Col = 11

    objGrid.objGrid = GridApontamentos
    
    objGrid.iLinhasVisiveis = 5
    
    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Largura da primeira coluna
    GridApontamentos.ColWidth(0) = 250

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    Call Grid_Inicializa(objGrid)

    Inicializa_GridApontamentos = SUCESSO

End Function

Private Sub GridApontamentos_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGridApontamentos, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGridApontamentos, iAlterado)
        End If

End Sub

Private Sub GridApontamentos_GotFocus()
    
    Call Grid_Recebe_Foco(objGridApontamentos)

End Sub

Private Sub GridApontamentos_EnterCell()

    Call Grid_Entrada_Celula(objGridApontamentos, iAlterado)

End Sub

Private Sub GridApontamentos_LeaveCell()
    
    Call Saida_Celula(objGridApontamentos)

End Sub

Private Sub GridApontamentos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer


    Call Grid_Trata_Tecla(KeyAscii, objGridApontamentos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridApontamentos, iAlterado)
    End If

End Sub

Private Sub GridApontamentos_RowColChange()

    Call Grid_RowColChange(objGridApontamentos)
    
    Call Mostra_Apontamentos

End Sub

Private Sub GridApontamentos_Scroll()

    Call Grid_Scroll(objGridApontamentos)

End Sub

Private Sub Concluido_Click()

Dim iClick As Integer
Dim lErro As Long

On Error GoTo Erro_Concluido_Click

    iAlterado = REGISTRO_ALTERADO
    
    bRefreshApontamentos = False
    Call Grid_Refresh_Checkbox(objGridApontamentos)
    bRefreshApontamentos = True
    
    'Muda o valor de concluído na coleção
    colApontamentos.Item(GridApontamentos.Row).iConcluido = StrParaInt(GridApontamentos.TextMatrix(GridApontamentos.Row, iGrid_Concluido_Col))
    
    'Se mudou para MARCADO...
    If colApontamentos.Item(GridApontamentos.Row).iConcluido = MARCADO Then
    
        'e se ainda não chegou nos 100% concluído...
        If colApontamentos.Item(GridApontamentos.Row).dPercConcluido < 1 Then
            
            'Altera os valores na coleção para 100%
            colApontamentos.Item(GridApontamentos.Row).dtData = gdtDataAtual
            colApontamentos.Item(GridApontamentos.Row).dQuantidade = StrParaDbl(GridApontamentos.TextMatrix(GridApontamentos.Row, iGrid_QuantidadeOP_Col))
            colApontamentos.Item(GridApontamentos.Row).dPercConcluido = 1
        
            'Altera na tela também
            Data.PromptInclude = False
            Data.Text = Format(colApontamentos.Item(GridApontamentos.Row).dtData, "dd/mm/yy")
            Data.PromptInclude = True
            
            Quantidade.Text = Formata_Estoque(colApontamentos.Item(GridApontamentos.Row).dQuantidade)
            PercConcluido.Text = CStr(colApontamentos.Item(GridApontamentos.Row).dPercConcluido * 100)
            
            'Posiciona no campo Quantidade para possíveis alterações
            Quantidade.SetFocus
        
        End If

    End If

    Exit Sub

Erro_Concluido_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142988)

    End Select

    Exit Sub

End Sub

Private Sub Concluido_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridApontamentos)

End Sub

Private Sub Concluido_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridApontamentos)

End Sub

Private Sub Concluido_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridApontamentos.objControle = Concluido
    lErro = Grid_Campo_Libera_Foco(objGridApontamentos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub NumeroOP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NumeroOP_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridApontamentos)

End Sub

Private Sub NumeroOP_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridApontamentos)

End Sub

Private Sub NumeroOP_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridApontamentos.objControle = NumeroOP
    lErro = Grid_Campo_Libera_Foco(objGridApontamentos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TemApontamento_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TemApontamento_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridApontamentos)

End Sub

Private Sub TemApontamento_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridApontamentos)

End Sub

Private Sub TemApontamento_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridApontamentos.objControle = TemApontamento
    lErro = Grid_Campo_Libera_Foco(objGridApontamentos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CT_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CT_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridApontamentos)

End Sub

Private Sub CT_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridApontamentos)

End Sub

Private Sub CT_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridApontamentos.objControle = CT
    lErro = Grid_Campo_Libera_Foco(objGridApontamentos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Compet_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Compet_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridApontamentos)

End Sub

Private Sub Compet_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridApontamentos)

End Sub

Private Sub Compet_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridApontamentos.objControle = Compet
    lErro = Grid_Campo_Libera_Foco(objGridApontamentos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ProdutoOP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoOP_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridApontamentos)

End Sub

Private Sub ProdutoOP_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridApontamentos)

End Sub

Private Sub ProdutoOP_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridApontamentos.objControle = ProdutoOP
    lErro = Grid_Campo_Libera_Foco(objGridApontamentos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescProdutoOP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescProdutoOP_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridApontamentos)

End Sub

Private Sub DescProdutoOP_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridApontamentos)

End Sub

Private Sub DescProdutoOP_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridApontamentos.objControle = DescProdutoOP
    lErro = Grid_Campo_Libera_Foco(objGridApontamentos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub VersaoProdOP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub VersaoProdOP_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridApontamentos)

End Sub

Private Sub VersaoProdOP_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridApontamentos)

End Sub

Private Sub VersaoProdOP_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridApontamentos.objControle = VersaoProdOP
    lErro = Grid_Campo_Libera_Foco(objGridApontamentos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantidadeOP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantidadeOP_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridApontamentos)

End Sub

Private Sub QuantidadeOP_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridApontamentos)

End Sub

Private Sub QuantidadeOP_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridApontamentos.objControle = QuantidadeOP
    lErro = Grid_Campo_Libera_Foco(objGridApontamentos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UMedidaOP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UMedidaOP_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridApontamentos)

End Sub

Private Sub UMedidaOP_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridApontamentos)

End Sub

Private Sub UMedidaOP_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridApontamentos.objControle = UMedidaOP
    lErro = Grid_Campo_Libera_Foco(objGridApontamentos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Etapa_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Etapa_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridApontamentos)

End Sub

Private Sub Etapa_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridApontamentos)

End Sub

Private Sub Etapa_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridApontamentos.objControle = Etapa
    lErro = Grid_Campo_Libera_Foco(objGridApontamentos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    Data.SetFocus

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 137767

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 137767

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142989)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_UpClick

    Data.SetFocus

    If Len(Trim(Data.ClipText)) > 0 Then

        sData = Data.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 137768

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 137768

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142990)

    End Select

    Exit Sub

End Sub

Function Limpa_Tela_ApontamentoProducao() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_ApontamentoProducao
        
    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    Call Grid_Limpa(objGridApontamentos)
    
    Set colApontamentos = New Collection
    
    DescCTInicial.Caption = ""
    DescCTFinal.Caption = ""
    
    DescProdInic.Caption = ""
    DescProdFim.Caption = ""
    
    SoAbertos.Value = True
    Todos.Value = False

    iAlterado = 0
    iQtdeAlterada = 0

    Limpa_Tela_ApontamentoProducao = SUCESSO

    Exit Function

Erro_Limpa_Tela_ApontamentoProducao:

    Limpa_Tela_ApontamentoProducao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142991)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Concluido(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Concluido do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Concluido

    Set objGridInt.objControle = Concluido

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 137769

    Saida_Celula_Concluido = SUCESSO

    Exit Function

Erro_Saida_Celula_Concluido:

    Saida_Celula_Concluido = gErr

    Select Case gErr
        
        Case 137769
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142992)

    End Select

    Exit Function

End Function

Private Function Trata_TabApontamentos() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_TabApontamentos
        
    'Verifica se tem seleção e Preenche o Grid
    lErro = Traz_Apontamentos_Selecionados()
    If lErro <> SUCESSO Then gError 137770
        
    iTabPrincipalAlterado = 0
    
    Trata_TabApontamentos = SUCESSO

    Exit Function

Erro_Trata_TabApontamentos:

    Trata_TabApontamentos = gErr

    Select Case gErr

        Case 137770

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142993)

    End Select

    Exit Function

End Function

Function Traz_Apontamentos_Selecionados() As Long

Dim lErro As Long
Dim objApontamentoSeleciona As New ClassApontamentoSeleciona

On Error GoTo Erro_Traz_Apontamentos_Selecionados

    GL_objMDIForm.MousePointer = vbHourglass

    'Limpa o GridApontamentos
    Call Grid_Limpa(objGridApontamentos)
    
    'Preenche o objApontamentoSeleciona conforme dados da tela
    lErro = Move_TabSelecao_Memoria(objApontamentoSeleciona)
    If lErro <> SUCESSO Then gError 137772
    
    'Lê o Plano Mestre de Produção, seus Itens e os Planos Operacionais
    lErro = CF("ApontamentoProducao_ObterDadosGrid", objApontamentoSeleciona)
    If lErro <> SUCESSO And lErro <> 137773 Then gError 137774
    
    'Não encontrou dados...
    If lErro <> SUCESSO Then gError 137775

    'Preenche o GridApontamentos com os dados obtidos
    lErro = Grid_Apontamentos_Preenche(objApontamentoSeleciona.colPlanoOperacional)
    If lErro <> SUCESSO Then gError 137776
    
    GL_objMDIForm.MousePointer = vbDefault

    Traz_Apontamentos_Selecionados = SUCESSO
    
    Exit Function
    
Erro_Traz_Apontamentos_Selecionados:
    
    Traz_Apontamentos_Selecionados = gErr
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
    Case 137772, 137774, 137776 'Tratado na rotina chamada
    
    Case 137775
            Call Rotina_Erro(vbOKOnly, "ERRO_SEM_APONTAMENTOS_ENCONTRADOS", gErr)
    
    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142994)

    End Select

    Exit Function

End Function

Function Move_TabSelecao_Memoria(objApontamentoSeleciona As ClassApontamentoSeleciona) As Long

Dim lErro As Long
Dim sProduto As String
Dim iProdPreenchido As Integer
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_Move_TabSelecao_Memoria

    'Move Filtro Principal para Memória
    If SoAbertos.Value = True Then
        objApontamentoSeleciona.iSoAbertos = SoAbertos.Value
    Else
        objApontamentoSeleciona.iTodos = Todos.Value
    End If
    
    'Move Centro de Trabalho para Memória
    If Len(Trim(CTInicial.Text)) <> 0 Then
    
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
    
        objCentrodeTrabalho.sNomeReduzido = CTInicial.Text
        
        'Lê o CentrodeTrabalho que está sendo Passado
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 137777
        
        objApontamentoSeleciona.lCTInicial = objCentrodeTrabalho.lNumIntDoc
        
    End If
    
    If Len(Trim(CTFinal.Text)) <> 0 Then
    
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
    
        objCentrodeTrabalho.sNomeReduzido = CTFinal.Text
        
        'Lê o CentrodeTrabalho que está sendo Passado
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 137778
        
        objApontamentoSeleciona.lCTFinal = objCentrodeTrabalho.lNumIntDoc
        
    End If
    
    'Move Produto para Memória
    If Len(Trim(ProdutoInicial.Text)) <> 0 Then
    
        lErro = CF("Produto_Formata", ProdutoInicial.Text, sProduto, iProdPreenchido)
        If lErro <> SUCESSO Then gError 137779

        objApontamentoSeleciona.sProdutoInicial = sProduto
        
    End If

    If Len(Trim(ProdutoFinal.Text)) <> 0 Then
    
        lErro = CF("Produto_Formata", ProdutoFinal.Text, sProduto, iProdPreenchido)
        If lErro <> SUCESSO Then gError 137779

        objApontamentoSeleciona.sProdutoFinal = sProduto
        
    End If

    'Move Código OP para Memória
    If Len(Trim(OpInicial.Text)) <> 0 Then
        objApontamentoSeleciona.sOPInicial = OpInicial.Text
    End If

    If Len(Trim(OpFinal.Text)) <> 0 Then
        objApontamentoSeleciona.sOPFinal = OpFinal.Text
    End If
        
    'Move Data OP para Memória
    objApontamentoSeleciona.dtDataOPInicial = StrParaDate(DataOPInicial.Text)
    objApontamentoSeleciona.dtDataOPFinal = StrParaDate(DataOPFinal.Text)

    Move_TabSelecao_Memoria = SUCESSO

    Exit Function

Erro_Move_TabSelecao_Memoria:

    Move_TabSelecao_Memoria = gErr

    Select Case gErr
    
        Case 137777 To 137779
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142995)

    End Select

    Exit Function

End Function

Private Function Grid_Apontamentos_Preenche(colPlanoOperacional As Collection) As Long
'Preenche o Grid de Apontamentos com os dados de colPlanoOperacional

Dim lErro As Long
Dim iLinha As Integer
Dim iIndice As Integer
Dim objPlanoOperacional As New ClassPlanoOperacional
Dim objApontamento As New ClassApontamentoProducao
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes
Dim objCompetencias As ClassCompetencias

On Error GoTo Erro_Grid_Apontamentos_Preenche

    Set colApontamentos = New Collection

    'altera o numero de linhas do Grid para o numero de Itens da coleção passada
    If GridApontamentos.Rows <= colPlanoOperacional.Count Then
        GridApontamentos.Rows = colPlanoOperacional.Count + 1
    End If

    iLinha = 0

    'Percorre todas os Planos Operacionais da Coleção
    For Each objPlanoOperacional In colPlanoOperacional

        iLinha = iLinha + 1
        
        Set objApontamento = objPlanoOperacional.objApontamento
        
        'Passa para a tela os dados do Plano Operacional em questão
        GridApontamentos.TextMatrix(iLinha, iGrid_Concluido_Col) = objApontamento.iConcluido
        GridApontamentos.TextMatrix(iLinha, iGrid_Etapa_Col) = objPlanoOperacional.iNivel & SEPARADOR & objPlanoOperacional.iSeq
        GridApontamentos.TextMatrix(iLinha, iGrid_NumeroOP_Col) = objPlanoOperacional.sCodOPOrigem
        GridApontamentos.TextMatrix(iLinha, iGrid_TemApont_Col) = objPlanoOperacional.iTemApontamento
        
        If objPlanoOperacional.lNumIntDocCT > 0 Then
            
            Set objCentrodeTrabalho = New ClassCentrodeTrabalho
            
            objCentrodeTrabalho.lNumIntDoc = objPlanoOperacional.lNumIntDocCT
            
            'Lê o CentrodeTrabalho que está sendo Passado
            lErro = CF("CentrodeTrabalho_Le_NumIntDoc", objCentrodeTrabalho)
            If lErro <> SUCESSO And lErro <> 134590 Then gError 137780
            
            GridApontamentos.TextMatrix(iLinha, iGrid_CT_Col) = objCentrodeTrabalho.sNomeReduzido
        
        End If
        
        If objPlanoOperacional.lNumIntDocOper <> 0 Then
        
            Set objOrdemProducaoOperacoes = New ClassOrdemProducaoOperacoes
            
            objOrdemProducaoOperacoes.lNumIntDoc = objPlanoOperacional.lNumIntDocOper
            
            lErro = CF("OrdemDeProducao_Le_Oper_NumIntDoc", objOrdemProducaoOperacoes)
            If lErro <> SUCESSO And lErro <> 137039 Then gError 137781
            
            If lErro = SUCESSO Then
                
                Set objCompetencias = New ClassCompetencias
                
                objCompetencias.lNumIntDoc = objOrdemProducaoOperacoes.lNumIntDocCompet
                
                lErro = CF("Competencias_Le_NumIntDoc", objCompetencias)
                If lErro <> SUCESSO And lErro <> 134336 Then gError 137782
                
                GridApontamentos.TextMatrix(iLinha, iGrid_Compet_Col) = objCompetencias.sNomeReduzido
                
            End If
            
        End If
        
        sProdutoFormatado = objPlanoOperacional.sProduto

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'Mascara produto
        lErro = Mascara_RetornaProdutoEnxuto(sProdutoFormatado, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 137787

        ProdutoOP.PromptInclude = False
        ProdutoOP.Text = sProdutoMascarado
        ProdutoOP.PromptInclude = True
        
        'le o produto para obter sua descricao
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 137784
        
        If lErro = 28030 Then gError 137785
            
        'se o produto não for ativo ==> Erro
        If objProduto.iAtivo <> PRODUTO_ATIVO Then gError 137786

        GridApontamentos.TextMatrix(iLinha, iGrid_ProdutoOP_Col) = ProdutoOP.Text
        GridApontamentos.TextMatrix(iLinha, iGrid_DescProdutoOP_Col) = objProduto.sDescricao
        GridApontamentos.TextMatrix(iLinha, iGrid_VersaoProdOP_Col) = objPlanoOperacional.sVersao
        GridApontamentos.TextMatrix(iLinha, iGrid_QuantidadeOP_Col) = Formata_Estoque(objPlanoOperacional.dQuantidade)
        GridApontamentos.TextMatrix(iLinha, iGrid_UMedidaOP_Col) = objPlanoOperacional.sUM
        
        colApontamentos.Add objApontamento
        
    Next

    Call Grid_Refresh_Checkbox(objGridApontamentos)

    'Passa para o Obj o número de Planos Operacionais passados pela Coleção
    objGridApontamentos.iLinhasExistentes = colPlanoOperacional.Count
    
    Grid_Apontamentos_Preenche = SUCESSO
    
    Exit Function

Erro_Grid_Apontamentos_Preenche:

    Grid_Apontamentos_Preenche = gErr
    
    Select Case gErr
    
        Case 137780 To 137784, 137787
                    
        Case 137785
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            
        Case 137786
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142996)

    End Select
    
    Exit Function

End Function

Function ValidaSelecao() As Long

Dim objCTInicial As ClassCentrodeTrabalho
Dim objCTFinal As ClassCentrodeTrabalho
Dim sProd_I As String
Dim sProd_F As String
Dim iProdPreenchido_I As Integer
Dim iProdPreenchido_F As Integer
Dim lErro As Long

On Error GoTo Erro_ValidaSelecao

    'Valida Centros de Trabalho
    If Len(Trim(CTInicial.Text)) <> 0 And Len(Trim(CTFinal.Text)) <> 0 Then
    
        Set objCTInicial = New ClassCentrodeTrabalho
    
        objCTInicial.sNomeReduzido = CTInicial.Text
        
        'Lê CT Inicial pelo NomeReduzido
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCTInicial)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 137788
                
        Set objCTFinal = New ClassCentrodeTrabalho
        
        objCTFinal.sNomeReduzido = CTFinal.Text
        
        'Lê CT Final pelo NomeReduzido
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCTFinal)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 137789
                
        'codigo do CT inicial não pode ser maior que o final
        If objCTInicial.lCodigo > objCTFinal.lCodigo Then gError 137790
        
    End If
    
    'Valida OPs
    'ordem de produção inicial não pode ser maior que a final
    If Len(Trim(OpInicial.Text)) <> 0 And Len(Trim(OpFinal.Text)) <> 0 Then

        If OpInicial.Text > OpFinal.Text Then gError 137791

    End If

    'Valida Produtos
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError 137792
    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError 137793
    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""

    'se ambas os produtos estão preenchidos, o produto inicial não pode ser maior que o final
    If iProdPreenchido_I = PRODUTO_PREENCHIDO And iProdPreenchido_F = PRODUTO_PREENCHIDO Then

        If sProd_I > sProd_F Then gError 137794
    
    End If
    
    'Valida Datas de OPs
    'data da OP inicial não pode ser maior que a final
    If Len(Trim(DataOPInicial.ClipText)) <> 0 And Len(Trim(DataOPFinal.ClipText)) <> 0 Then
        
        If StrParaDate(DataOPInicial.Text) > StrParaDate(DataOPFinal.Text) Then gError 137795
    
    End If
    
    ValidaSelecao = SUCESSO
    
    Exit Function
    
Erro_ValidaSelecao:

    ValidaSelecao = gErr

    Select Case gErr
    
        Case 137790
            Call Rotina_Erro(vbOKOnly, "ERRO_CT_INICIAL_MAIOR", gErr)
    
        Case 137791
            Call Rotina_Erro(vbOKOnly, "ERRO_OP_INICIAL_MAIOR", gErr)
            
        Case 137795
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAOP_INICIAL_MAIOR", gErr)
            
        Case 137788, 137789
            'erros tratados nas rotinas chamadas
            
        Case 137792
            ProdutoInicial.SetFocus

        Case 137793
            ProdutoFinal.SetFocus

        Case 137794
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INICIAL_MAIOR", gErr)
                        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142997)
    
    End Select

    Exit Function

End Function

Function Mostra_Apontamentos() As Long

Dim lErro As Long

On Error GoTo Erro_Mostra_Apontamentos
    
    'se tem uma linha selecionada do grid e é para renovar o apontamento ...
    If GridApontamentos.Row > 0 And bRefreshApontamentos = True Then
        
        'coloca caption do frame padrão
        FrameApontamentos.Caption = "Apontamento"
    
        'limpa os campos do frame
        PercConcluido.Text = ""
        Quantidade.Text = ""
        
        Data.PromptInclude = False
        Data.Text = ""
        Data.PromptInclude = True
            
        Observacao.Text = ""
        
        LabelPercentualAnterior.Caption = ""
        LabelQuantidadeAnterior.Caption = ""
        LabelDataAnterior.Caption = ""
        
        'e se esta linha está preenchida
        If Len(GridApontamentos.TextMatrix(GridApontamentos.Row, iGrid_Etapa_Col)) > 0 Then
        
            'coloca OP e Etapa no caption do frame
            FrameApontamentos.Caption = "Apontamento (Nº OP: " & GridApontamentos.TextMatrix(GridApontamentos.Row, iGrid_NumeroOP_Col) & " - Etapa: " & GridApontamentos.TextMatrix(GridApontamentos.Row, iGrid_Etapa_Col) & ")"
            
            'se tem apontamentos... exibe-os
            If GridApontamentos.TextMatrix(GridApontamentos.Row, iGrid_TemApont_Col) = MARCADO Then
            
                LabelPercentualAnterior.Caption = Format((colApontamentos.Item(GridApontamentos.Row).dPercConcluido * 100), "#0.#0\%")
                LabelQuantidadeAnterior.Caption = Formata_Estoque(colApontamentos.Item(GridApontamentos.Row).dQuantidade)
                LabelDataAnterior.Caption = Format(colApontamentos.Item(GridApontamentos.Row).dtData, "dd/mm/yyyy")
                            
                Observacao.Text = colApontamentos.Item(GridApontamentos.Row).sObservacao
                
            Else
            
                If GridApontamentos.TextMatrix(GridApontamentos.Row, iGrid_Concluido_Col) = MARCADO Then
                    
                    PercConcluido.Text = CStr(colApontamentos.Item(GridApontamentos.Row).dPercConcluido * 100)
                    Quantidade.Text = Formata_Estoque(colApontamentos.Item(GridApontamentos.Row).dQuantidade)
                    Observacao.Text = colApontamentos.Item(GridApontamentos.Row).sObservacao
                    
                End If
                
            End If
            
            Data.PromptInclude = False
            Data.Text = Format(gdtDataAtual, "dd/mm/yy")
            Data.PromptInclude = True
            
        End If
        
        iQtdeAlterada = 0
    
    End If
    
    Mostra_Apontamentos = SUCESSO
    
    Exit Function
    
Erro_Mostra_Apontamentos:

    Mostra_Apontamentos = gErr

    Select Case gErr
                            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142998)
    
    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objPO As ClassPlanoOperacional
Dim objApontamentoProducao As New ClassApontamentoProducao
Dim iIndice As Integer
Dim ColPO As New Collection

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Consiste os dados e verifica se já existia Apontamento
    For iIndice = 1 To colApontamentos.Count
    
        Set objApontamentoProducao = colApontamentos.Item(iIndice)
        
        If objApontamentoProducao.iConcluido = MARCADO Then
                                
            'Verifica se Tem Apontamento está desmarcado
            If CInt(GridApontamentos.TextMatrix(iIndice, iGrid_TemApont_Col)) = DESMARCADO Then gError 137796
        
        End If
        
        Set objPO = New ClassPlanoOperacional
        
        objPO.lNumIntDoc = objApontamentoProducao.lNumIntDocPO
        objPO.iTemApontamento = CInt(GridApontamentos.TextMatrix(iIndice, iGrid_TemApont_Col))
        
        Set objPO.objApontamento = objApontamentoProducao
        
        ColPO.Add objPO
        
    Next
                
    'Grava os ApontamentoProducao no Banco de Dados
    lErro = CF("ApontamentoProducao_Grava", ColPO)
    If lErro <> SUCESSO Then gError 137797
                
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:
    
    Gravar_Registro = gErr
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
        Case 137796
            Call Rotina_Erro(vbOKOnly, "ERRO_TEM_APONTAMENTO_GRID_NAO_PREENCHIDO", gErr, iIndice)
        
        Case 137797
            'erros tratados nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142999)
        
    End Select
    
    Exit Function
    
End Function


