VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl EtapasDaProducao 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9450
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9450
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   -15
      TabIndex        =   0
      Top             =   60
      Width           =   9435
      Begin VB.Frame Frame2 
         Caption         =   "Operação"
         Height          =   2640
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   2700
         Width           =   9195
         Begin VB.TextBox Observacao 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1380
            MaxLength       =   255
            TabIndex        =   3
            Top             =   1815
            Width           =   7635
         End
         Begin MSMask.MaskEdBox CodigoCTPadrao 
            Height          =   315
            Left            =   1380
            TabIndex        =   2
            Top             =   1410
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label LabelRepeticao 
            AutoSize        =   -1  'True
            Caption         =   "Número de Repetições:"
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
            Height          =   195
            Left            =   6585
            TabIndex        =   67
            Top             =   2235
            Width           =   2010
         End
         Begin VB.Label Repeticao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8595
            TabIndex        =   66
            Top             =   2190
            Width           =   420
         End
         Begin VB.Label NumMaxMaqPorOper 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6090
            TabIndex        =   64
            Top             =   2205
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Número Máximo de Máquinas:"
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
            Height          =   195
            Left            =   3510
            TabIndex        =   63
            Top             =   2235
            Width           =   2550
         End
         Begin VB.Label CodigoCompetencia 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1380
            TabIndex        =   62
            Top             =   1020
            Width           =   2445
         End
         Begin VB.Label VersaoLabel 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1380
            TabIndex        =   60
            Top             =   630
            Width           =   1665
         End
         Begin VB.Label LabelDetVersao 
            Caption         =   "Versão:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   615
            TabIndex        =   59
            Top             =   690
            Width           =   690
         End
         Begin VB.Label ProdutoLabel 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1380
            TabIndex        =   58
            Top             =   225
            Width           =   7635
         End
         Begin VB.Label LabelDetProduto 
            Caption         =   "Produto:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   555
            TabIndex        =   57
            Top             =   255
            Width           =   810
         End
         Begin VB.Label QtdeLabel 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6015
            TabIndex        =   56
            Top             =   630
            Width           =   1470
         End
         Begin VB.Label LabelDetQtde 
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
            Height          =   210
            Left            =   4890
            TabIndex        =   55
            Top             =   690
            Width           =   1050
         End
         Begin VB.Label UMLabel 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   8235
            TabIndex        =   54
            Top             =   630
            Width           =   780
         End
         Begin VB.Label LabelDetUM 
            Caption         =   "U.M.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   7725
            TabIndex        =   53
            Top             =   690
            Width           =   480
         End
         Begin VB.Label DescricaoCTPadrao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3870
            TabIndex        =   29
            Top             =   1410
            Width           =   5145
         End
         Begin VB.Label DescricaoCompetencia 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3870
            TabIndex        =   28
            Top             =   1020
            Width           =   5145
         End
         Begin VB.Label LabelNivel 
            AutoSize        =   -1  'True
            Caption         =   "Nível:"
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
            Height          =   195
            Left            =   765
            TabIndex        =   27
            Top             =   2235
            Width           =   540
         End
         Begin VB.Label Sequencial 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2970
            TabIndex        =   26
            Top             =   2205
            Width           =   420
         End
         Begin VB.Label Nivel 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1380
            TabIndex        =   25
            Top             =   2205
            Width           =   420
         End
         Begin VB.Label LabelSeq 
            AutoSize        =   -1  'True
            Caption         =   "Seqüencial:"
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
            Height          =   195
            Left            =   1905
            TabIndex        =   24
            Top             =   2235
            Width           =   1020
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
            Height          =   330
            Left            =   165
            TabIndex        =   23
            Top             =   1830
            Width           =   1155
         End
         Begin VB.Label CTLabel 
            Caption         =   "C. Trabalho:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   330
            Left            =   210
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   22
            Top             =   1440
            Width           =   1110
         End
         Begin VB.Label CompetenciaLabel 
            Caption         =   "Competência:"
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
            Left            =   105
            TabIndex        =   21
            Top             =   1050
            Width           =   1155
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   2685
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   2700
         Visible         =   0   'False
         Width           =   9210
         Begin VB.TextBox StatusMRP 
            BackColor       =   &H8000000F&
            Height          =   530
            Left            =   3930
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   110
            Width           =   5205
         End
         Begin MSMask.MaskEdBox Horas 
            Height          =   315
            Left            =   5595
            TabIndex        =   32
            Top             =   1560
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   315
            Left            =   6450
            TabIndex        =   31
            Top             =   1560
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.TextBox TaxaProducao 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   7665
            TabIndex        =   14
            Top             =   1560
            Width           =   1185
         End
         Begin VB.TextBox NomeMaquina 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   2700
            TabIndex        =   13
            Top             =   1560
            Width           =   2040
         End
         Begin VB.TextBox OPCodigoMRP 
            Height          =   315
            Left            =   690
            MaxLength       =   6
            TabIndex        =   6
            Top             =   900
            Width           =   1125
         End
         Begin VB.CheckBox MRP 
            Enabled         =   0   'False
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
            Left            =   75
            TabIndex        =   5
            Top             =   90
            Width           =   225
         End
         Begin VB.Frame FrameFinal 
            Caption         =   "Datas"
            Height          =   1155
            Left            =   15
            TabIndex        =   11
            Top             =   1485
            Width           =   2160
            Begin MSMask.MaskEdBox DataFinal 
               Height          =   315
               Left            =   645
               TabIndex        =   7
               Top             =   690
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataFinal 
               Height          =   300
               Left            =   1785
               TabIndex        =   9
               TabStop         =   0   'False
               Top             =   690
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataInicio 
               Height          =   315
               Left            =   660
               TabIndex        =   47
               Top             =   240
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataInicio 
               Height          =   300
               Left            =   1800
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   240
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin VB.Label LabelDataInicio 
               Caption         =   "Início:"
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
               Left            =   75
               TabIndex        =   49
               Top             =   270
               Width           =   525
            End
            Begin VB.Label LabelDataFinal 
               Caption         =   "Final:"
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
               TabIndex        =   12
               Top             =   720
               Width           =   495
            End
         End
         Begin MSMask.MaskEdBox QuantidadeMaquina 
            Height          =   315
            Left            =   4740
            TabIndex        =   15
            Top             =   1560
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "0"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridMaquinas 
            Height          =   1755
            Left            =   2235
            TabIndex        =   8
            Top             =   885
            Width           =   6930
            _ExtentX        =   12224
            _ExtentY        =   3096
            _Version        =   393216
         End
         Begin VB.Label LabelMRP 
            Caption         =   "Plano Mestre de Produção"
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
            Left            =   360
            TabIndex        =   61
            Top             =   140
            Width           =   2445
         End
         Begin VB.Label LabelMaquinas 
            Caption         =   "Maquinas:"
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
            Left            =   2250
            TabIndex        =   19
            Top             =   660
            Width           =   945
         End
         Begin VB.Label LabelStatusMRP 
            Caption         =   "Status:"
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
            Left            =   3255
            TabIndex        =   18
            Top             =   140
            Width           =   645
         End
         Begin VB.Label LabelOP 
            Caption         =   "O.P.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   180
            TabIndex        =   16
            Top             =   930
            Width           =   510
         End
      End
      Begin VB.CommandButton BotaoCancela 
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
         Height          =   360
         Left            =   4820
         TabIndex        =   52
         Top             =   5490
         Width           =   1350
      End
      Begin VB.CommandButton BotaoOK 
         Caption         =   "OK"
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
         Left            =   2985
         TabIndex        =   46
         Top             =   5490
         Width           =   1350
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ordem de Produção"
         Height          =   2370
         Left            =   75
         TabIndex        =   33
         Top             =   -30
         Width           =   4305
         Begin VB.CheckBox ProduzLogo 
            Caption         =   "Produz Logo"
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
            Left            =   2835
            TabIndex        =   65
            Top             =   705
            Width           =   1410
         End
         Begin VB.TextBox Codigo 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   1185
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   37
            Top             =   240
            Width           =   1305
         End
         Begin VB.TextBox Quantidade 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   1185
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   36
            Top             =   1515
            Width           =   1350
         End
         Begin VB.TextBox UM 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   3480
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   35
            Top             =   1515
            Width           =   720
         End
         Begin VB.TextBox DataNecessidade 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   3015
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   34
            Top             =   1920
            Width           =   1200
         End
         Begin MSMask.MaskEdBox Prioridade 
            Height          =   315
            Left            =   3525
            TabIndex        =   38
            Top             =   225
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   2
            Format          =   "##"
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin VB.Label Label6 
            Caption         =   "Descrição:"
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
            Left            =   225
            TabIndex        =   51
            Top             =   1110
            Width           =   930
         End
         Begin VB.Label LabelDescProd 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1185
            TabIndex        =   50
            Top             =   1080
            Width           =   3015
         End
         Begin VB.Label CodigoOPLabel 
            AutoSize        =   -1  'True
            Caption         =   "Código OP:"
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
            Left            =   180
            TabIndex        =   45
            Top             =   270
            Width           =   975
         End
         Begin VB.Label LabelProduto 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1185
            TabIndex        =   44
            Top             =   660
            Width           =   1605
         End
         Begin VB.Label Label3 
            Caption         =   "Produto:"
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
            Left            =   420
            TabIndex        =   43
            Top             =   690
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Prioridade:"
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
            Left            =   2565
            TabIndex        =   42
            Top             =   270
            Width           =   975
         End
         Begin VB.Label LabelUMedida 
            Caption         =   "UM:"
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
            Left            =   3090
            TabIndex        =   41
            Top             =   1545
            Width           =   345
         End
         Begin VB.Label Label4 
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
            Height          =   225
            Left            =   120
            TabIndex        =   40
            Top             =   1545
            Width           =   1020
         End
         Begin VB.Label Label1 
            Caption         =   "Data da Necessidade:"
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
            Left            =   1050
            TabIndex        =   39
            Top             =   1965
            Width           =   1920
         End
      End
      Begin VB.Frame FrameRoteiro 
         Caption         =   "Roteiro de Fabricação:"
         Height          =   2385
         Left            =   4425
         TabIndex        =   30
         Top             =   -30
         Width           =   4950
         Begin MSComctlLib.TreeView Roteiro 
            Height          =   2085
            Left            =   45
            TabIndex        =   1
            Top             =   210
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   3678
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   354
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            FullRowSelect   =   -1  'True
            Appearance      =   1
         End
      End
      Begin MSComctlLib.TabStrip TabStrip2 
         Height          =   3045
         Left            =   75
         TabIndex        =   4
         Top             =   2370
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   5371
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Detalhe"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Plano Mestre de Produção"
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
End
Attribute VB_Name = "EtapasDaProducao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjPMPItem As New ClassPMPItens
Dim gobjPMP As New ClassPMP
Dim gobjPMPItemCopia As New ClassPMPItens
Dim gobjPO As ClassPlanoOperacional

Dim iFrameAtualOper As Integer

'Grid de Maquinas
Dim objGridMaquinas As AdmGrid
Dim iGrid_NomeMaquina_Col As Integer
Dim iGrid_QuantidadeMaquina_Col As Integer
Dim iGrid_Data_Col As Integer
Dim iGrid_Horas_Col As Integer
Dim iGrid_TaxaProducao_Col As Integer

Dim colComponentes As New Collection
Dim iProxChave As Integer

'variaveis auxiliares para recalculo de nivel e sequencial
Dim aNivelSequencial(NIVEL_MAXIMO_OPERACOES) As Integer 'para cada nivel guarda o maior sequencial
Dim aSeqPai(NIVEL_MAXIMO_OPERACOES) As Integer 'para cada nivel guarda o SeqPai

Dim iUltimoNivel As Integer

Private WithEvents objEventoCentroDeTrabalho As AdmEvento
Attribute objEventoCentroDeTrabalho.VB_VarHelpID = -1

Public iAlterado As Integer

'**** inicio do trecho a ser copiado *****
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Etapas da Produção"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "EtapasDaProducao"

End Function

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

Private Sub BotaoCancela_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoCancela_Click
       
    'Nao mexer no obj da tela
    giRetornoTela = vbOK
    
    Unload Me
    
    Exit Sub
    
Erro_BotaoCancela_Click:

    Select Case gErr
    
        Case 137633
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159646)

    End Select
    
    Exit Sub

End Sub


Private Sub ProduzLogo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

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
        If Me.ActiveControl Is CodigoCTPadrao Then
            Call CTLabel_Click
        End If
    End If

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    CodigoCTPadrao.Enabled = True
    
    'Indica se a tela não foi carregada corretamente
    giRetornoTela = vbAbort
    
    Set objEventoCentroDeTrabalho = New AdmEvento
    
    iFrameAtualOper = 1
    
    'Grid de Maquinas
    Set objGridMaquinas = New AdmGrid
    
    'tela em questão
    Set objGridMaquinas.objForm = Me
    
    lErro = Inicializa_GridMaquinas(objGridMaquinas)
    If lErro <> SUCESSO Then gError 137634
    
    MRP.Value = vbUnchecked
    
    lErro = Habilita_MRP()
    If lErro <> SUCESSO Then gError 137635
    
    If gobjEST.iTemRepeticoesOper = MARCADO Then
        Repeticao.Visible = True
        LabelRepeticao.Visible = True
    Else
        Repeticao.Visible = False
        LabelRepeticao.Visible = False
    End If
    
    iAlterado = 0
    
    'Sinaliza que o Form_Load ocorreu com sucesso
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 137634, 137635
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159647)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objPMP As ClassPMP, Optional objPMPItens As ClassPMPItens, Optional objPO As ClassPlanoOperacional) As Long 'Alterado por Wagner #MRP#

Dim lErro As Long
Dim iIndice As Integer
Dim objPOAux As ClassPlanoOperacional

On Error GoTo Erro_Trata_Parametros

    If Not (objPMP Is Nothing) Then
        Set gobjPMP = objPMP
    End If
    
    If Not (objPMPItens Is Nothing) Then
    
        Set gobjPMPItemCopia = objPMPItens
        
        lErro = PMPItem_Cria_Copia(objPMPItens, gobjPMPItem)
        If lErro <> SUCESSO Then gError 137636
    
        'traz dados da OP para a tela
        lErro = Traz_EtapasDaProducao_Tela(objPMPItens)
        If lErro <> SUCESSO Then gError 137637
            
        Call ComandoSeta_Fechar(Me.Name)
                
    End If

    If Not (objPO Is Nothing) Then
    
        iIndice = 1
        
        For Each objPOAux In objPMPItens.ColPO
                        
            If objPOAux.lNumIntDoc = objPO.lNumIntDoc Then Exit For
            
            iIndice = iIndice + 1

        Next
    
        'selecionar a raiz
        Set Roteiro.SelectedItem = Roteiro.Nodes.Item(iIndice)
        Roteiro.SelectedItem.Selected = True
        
        'e carregar as operações pertinentes
        Call Roteiro_NodeClick(Roteiro.Nodes.Item(iIndice))
    
    End If

    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    giRetornoTela = vbCancel

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 137636, 137637
            'erros tratados nas rotinas chamadas
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159648)
    
    End Select
    
    Exit Function
        
End Function

Private Sub BotaoOK_Click()
    
Dim lErro As Long
    
On Error GoTo Erro_BotaoOK_Click
    
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 137638
    
    'Indica que saiu da tela de forma legal
    giRetornoTela = vbOK
    
    iAlterado = 0
    
    'Fecha a tela
    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case gErr

        Case 137638
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159649)

    End Select

    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long

On Error GoTo Erro_Gravar_Registro
    
    lErro = Move_EtapasDaProducao_Memoria()
    If lErro <> SUCESSO Then gError 136212
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr
    
        Case 136212
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159650)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoCentroDeTrabalho = Nothing
    
    Set gobjPMPItem = Nothing
    Set gobjPMP = Nothing
    Set gobjPO = Nothing
    Set gobjPMPItemCopia = Nothing
    
    Set colComponentes = Nothing
    
    Set objGridMaquinas = Nothing

End Sub

Function Traz_EtapasDaProducao_Tela(objPMPItens As ClassPMPItens) As Long

Dim lErro As Long
Dim objItemOP As New ClassItemOP
Dim objProdutos As ClassProduto
Dim sProdutoFormatado As String

On Error GoTo Erro_Traz_EtapasDaProducao_Tela

    Codigo.Text = objPMPItens.sCodOPOrigem

    lErro = Mascara_RetornaProdutoTela(objPMPItens.sProduto, sProdutoFormatado)
    If lErro <> SUCESSO Then gError 137639

    Set objProdutos = New ClassProduto

    objProdutos.sCodigo = objPMPItens.sProduto

    lErro = CF("Produto_Le", objProdutos)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 137640

    LabelProduto.Caption = sProdutoFormatado
    
    LabelDescProd = objProdutos.sDescricao
    
    Quantidade.Text = Formata_Estoque(objPMPItens.dQuantidade)
    
    ProduzLogo.Value = objPMPItens.iProduzLogo
    
    UM.Text = objPMPItens.sUM
    
    Prioridade.PromptInclude = False
    Prioridade.Text = objPMPItens.iPrioridade
    Prioridade.PromptInclude = True
    
    DataNecessidade.Text = Format(objPMPItens.dtDataNecessidade, "dd/mm/yyyy")
    
    lErro = Trata_Arvore(objPMPItens.sProduto, objPMPItens.sVersao, objPMPItens.sUM, objPMPItens.dQuantidade)
    If lErro <> SUCESSO Then gError 137641
        
    Traz_EtapasDaProducao_Tela = SUCESSO

    Exit Function

Erro_Traz_EtapasDaProducao_Tela:

    Traz_EtapasDaProducao_Tela = gErr

    Select Case gErr
    
        Case 137639 To 137641

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159651)

    End Select

    Exit Function

End Function

Function Move_EtapasDaProducao_Memoria() As Long

Dim lErro As Long
Dim objItemOPOperacoes As New ClassItemOP
Dim objOrdemProducaoOperacoes As New ClassOrdemProducaoOperacoes
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Move_EtapasDaProducao_Memoria

    If gobjPMPItem.iPrioridade <> StrParaInt(Prioridade.Text) Then
        gobjPO.iAlterado = REGISTRO_ALTERADO
    End If
    
    If gobjPMPItem.iProduzLogo <> ProduzLogo.Value Then
        gobjPMPItem.iAlterado = REGISTRO_ALTERADO
    End If
    
    gobjPMPItem.dQuantidade = StrParaDbl(Quantidade.Text)
    gobjPMPItem.iPrioridade = StrParaInt(Prioridade.Text)
    gobjPMPItem.iProduzLogo = ProduzLogo.Value
    
    If gobjPMPItem.ColPO.Count <> 0 Then
    
        If gobjPMPItem.dtDataNecessidade < gobjPMPItem.ColPO.Item(1).dtDataFim Then
        
            'Pergunta ao usuário se confirma a alteração dos dados
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DATAFIMPRODUCAO_MAIOR_DATANECESSIDADE")
            If vbMsgRes = vbNo Then gError 136789
            
        End If
        
        gobjPMPItem.objItemOP.dtDataFimProd = gobjPMPItem.ColPO.Item(1).dtDataFim
        gobjPMPItem.objItemOP.dtDataInicioProd = gobjPMPItem.ColPO.Item(1).dtDataInicio
        
    End If
    
    If Not (gobjPO Is Nothing) Then
    
        lErro = Move_MRP_Memoria(gobjPO)
        If lErro <> SUCESSO Then gError 136790
        
    End If

    Set gobjPMPItemCopia.ColPO = New Collection
    
    'Retorna a cópia original do Plano Mestre
    lErro = PMPItem_Cria_Copia(gobjPMPItem, gobjPMPItemCopia)
    If lErro <> SUCESSO Then gError 137633
    
    Move_EtapasDaProducao_Memoria = SUCESSO

    Exit Function

Erro_Move_EtapasDaProducao_Memoria:

    Move_EtapasDaProducao_Memoria = gErr

    Select Case gErr
    
        Case 136789, 136790, 137898

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159652)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable

    'Pesquisa o controle da coluna em questão
    Select Case objControl.Name

        Case Horas.Name ',QuantidadeMaquina.Name
            
'            If MRP.Value = vbChecked Then
'
'                If Len(Trim(GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_NomeMaquina_Col))) <> 0 Then
'                    objControl.Enabled = True
'                Else
'                    objControl.Enabled = False
'                End If
'
'            Else
                objControl.Enabled = False
'            End If
                
         Case Else
            
            objControl.Enabled = False
        
    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159653)

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
    
        'Verifica se é o GridMaquinas
        If objGridInt.objGrid.Name = GridMaquinas.Name Then
    
            Select Case GridMaquinas.Col
    
                Case iGrid_QuantidadeMaquina_Col
    
                    lErro = Saida_Celula_QuantidadeMaquina(objGridInt)
                    If lErro <> SUCESSO Then gError 137642
     
                Case iGrid_Horas_Col
    
                    lErro = Saida_Celula_Horas(objGridInt)
                    If lErro <> SUCESSO Then gError 138322
     
     
            End Select
                   
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 137643

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 137642, 138322
        
        Case 137643
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159654)

    End Select

    Exit Function

End Function

Function Trata_Arvore(ByVal sProdutoRaiz As String, ByVal sVersao As String, ByVal sUMedida As String, ByVal dQuantidade As Double) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Trata_Arvore

    'limpa Operacoes
    lErro = Limpa_Operacoes()
    If lErro <> SUCESSO Then gError 137644
    
    'limpa a Arvore
    lErro = Limpa_Arvore_Roteiro()
    If lErro <> SUCESSO Then gError 137645
        
    If Len(Trim(Codigo.Text)) <> 0 And Len(Trim(sProdutoRaiz)) <> 0 Then
    
        If Len(Trim(sUMedida)) = 0 Then gError 137646
        
        If dQuantidade = 0 Then gError 137647
        
        lErro = CF("Produto_Formata", sProdutoRaiz, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 137648
        
        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 137649
        
        'Monta a árvore de operações
        lErro = Carrega_Arvore()
        If lErro <> SUCESSO Then gError 137650
        
    End If
    
    Trata_Arvore = SUCESSO
    
    Exit Function
    
Erro_Trata_Arvore:

    Trata_Arvore = gErr
    
    Select Case gErr
    
        Case 137646
            Call Rotina_Erro(vbOKOnly, "ERRO_UM_ITEM_VAZIO", gErr)
        
        Case 137647
            Call Rotina_Erro(vbOKOnly, "ERRO_QTDE_ITEM_VAZIA", gErr)
            
        Case 137649
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOOP_NAO_PREENCHIDO", gErr, 1)
    
        Case 137644, 137645, 137648, 137650
            'erros tratados nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159655)
            
    End Select
    
    Exit Function

End Function

Private Sub CodigoCTPadrao_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub


Private Sub CodigoCTPadrao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodigoCTPadrao, iAlterado)
    
End Sub

Private Sub CodigoCTPadrao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objCTCompetencias As New ClassCTCompetencias
Dim objCompetencias As ClassCompetencias
Dim bCompetenciaCadastrada As Boolean
Dim lCTAnterior As Long

On Error GoTo Erro_CodigoCTPadrao_Validate

    DescricaoCTPadrao.Caption = ""
    
    'Verifica se CodigoCTPadrao não está preenchido
    If Len(Trim(CodigoCTPadrao.Text)) <> 0 Then

        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
        
        'Procura pela empresa toda
        objCentrodeTrabalho.iFilialEmpresa = EMPRESA_TODA
        
        'Verifica sua existencia
        lErro = CF("TP_CentrodeTrabalho_Le", CodigoCTPadrao, objCentrodeTrabalho)
        If lErro <> SUCESSO Then gError 137652
        
        Set objCompetencias = New ClassCompetencias
        
        objCompetencias.sNomeReduzido = CodigoCompetencia.Caption
        
        'Lê a Competencia pelo NomeReduzido para verificar seu NumIntDoc
        lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
        If lErro <> SUCESSO And lErro <> 134937 Then gError 137654
    
        If lErro <> SUCESSO Then gError 137655
        
        lErro = CF("CentrodeTrabalho_Le_CTCompetencias", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134453 Then gError 137656
    
        bCompetenciaCadastrada = False
        
        For Each objCTCompetencias In objCentrodeTrabalho.colCompetencias
        
            If objCTCompetencias.lNumIntDocCompet = objCompetencias.lNumIntDoc Then
            
                bCompetenciaCadastrada = True
                Exit For
                
            End If
        
        Next
            
        If bCompetenciaCadastrada = False Then gError 137657
            
        DescricaoCTPadrao.Caption = objCentrodeTrabalho.sDescricao
        
        '############################################
        'Inserido por Wagner
        If gobjPO.lNumIntDocCT <> objCentrodeTrabalho.lNumIntDoc Then
        
            gobjPO.iAlterado = REGISTRO_ALTERADO
            
            lCTAnterior = gobjPO.lNumIntDocCT
            
            gobjPO.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
        
            lErro = Acerta_Data_Inicio(gobjPO)
            If lErro <> SUCESSO Then gError 138320
                    
        End If
        '############################################
                       
    End If
       
    Exit Sub

Erro_CodigoCTPadrao_Validate:

    Cancel = True

    gobjPO.lNumIntDocCT = lCTAnterior

    Select Case gErr
    
        Case 137652, 137653, 137654, 137656, 138320
            'erros tratados nas rotinas chamadas
                                
        Case 137655, 137657
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPETENCIA_NAO_CADASTRADA_CT", gErr, objCentrodeTrabalho.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159656)

    End Select

    Exit Sub

End Sub

Private Sub CTLabel_Click()

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim colSelecao As New Collection

On Error GoTo Erro_CTLabel

    If CodigoCTPadrao.Enabled Then

        'Verifica se o CodigoCTPadrao foi preenchido
        If Len(Trim(CodigoCTPadrao.Text)) <> 0 Then
            
            objCentrodeTrabalho.sNomeReduzido = CodigoCTPadrao.Text
            
        End If
    
        Call Chama_Tela_Modal("CentrodeTrabalhoLista", colSelecao, objCentrodeTrabalho, objEventoCentroDeTrabalho)

    End If

    Exit Sub

Erro_CTLabel:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159657)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCentroDeTrabalho_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objCentrodeTrabalho = obj1

    CodigoCTPadrao.Text = objCentrodeTrabalho.sNomeReduzido
        
    Call CodigoCTPadrao_Validate(bSGECancelDummy)
        
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159658)

    End Select

    Exit Sub

End Sub

Private Sub Roteiro_NodeClick(ByVal Node As MSComctlLib.Node)

Dim lErro As Long
Dim objOrdemProducaoOperacoes As New ClassOrdemProducaoOperacoes
Dim objNode As Node
Dim objPO As ClassPlanoOperacional

On Error GoTo Erro_Roteiro_NodeClick

    Set objNode = Roteiro.SelectedItem

    Set objOrdemProducaoOperacoes = colComponentes.Item(objNode.Tag)
    
    '########################################
    'Inserido por Wagner
    Call Limpa_MRP
    
    For Each objPO In gobjPMPItem.ColPO
        If objPO.iNivel = objOrdemProducaoOperacoes.iNivel And objOrdemProducaoOperacoes.iSeqArvore = objPO.iSeq Then
            
            Set gobjPO = objPO
            
            lErro = Preenche_MRP(objPO)
            If lErro <> SUCESSO Then gError 136403
            Exit For
        End If
    Next
    '########################################

    lErro = Preenche_Operacoes(objOrdemProducaoOperacoes)
    If lErro <> SUCESSO Then gError 137658
    
    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_Roteiro_NodeClick:

    Select Case gErr

        Case 137658, 136403 'Inserido por Wagner

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159659)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip2_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip2.SelectedItem.Index <> iFrameAtualOper Then

        If TabStrip_PodeTrocarTab(iFrameAtualOper, TabStrip2, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame2(TabStrip2.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame2(iFrameAtualOper).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtualOper = TabStrip2.SelectedItem.Index
        
    End If

End Sub

Private Sub UpDownDataFinal_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFinal_DownClick

    DataFinal.SetFocus

    If Len(DataFinal.ClipText) > 0 Then

        sData = DataFinal.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 137659

        DataFinal.Text = sData
        
        Call DataFinal_Validate(bSGECancelDummy) 'Inserido por Wagner

    End If

    Exit Sub

Erro_UpDownDataFinal_DownClick:

    Select Case gErr

        Case 137659

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159660)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataFinal_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFinal_UpClick

    DataFinal.SetFocus

    If Len(Trim(DataFinal.ClipText)) > 0 Then

        sData = DataFinal.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 137660

        DataFinal.Text = sData

        Call DataFinal_Validate(bSGECancelDummy) 'Inserido por Wagner

    End If

    Exit Sub

Erro_UpDownDataFinal_UpClick:

    Select Case gErr

        Case 137660

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159661)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataInicio_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicio_DownClick

    DataInicio.SetFocus

    If Len(DataInicio.ClipText) > 0 Then

        sData = DataInicio.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 137661

        DataInicio.Text = sData

        Call DataInicio_Validate(bSGECancelDummy) 'Inserido por Wagner

    End If

    Exit Sub

Erro_UpDownDataInicio_DownClick:

    Select Case gErr

        Case 137661

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159662)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataInicio_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicio_UpClick

    DataInicio.SetFocus

    If Len(Trim(DataInicio.ClipText)) > 0 Then

        sData = DataInicio.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 137662

        DataInicio.Text = sData

        Call DataInicio_Validate(bSGECancelDummy) 'Inserido por Wagner

    End If

    Exit Sub

Erro_UpDownDataInicio_UpClick:

    Select Case gErr

        Case 137662

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159663)

    End Select

    Exit Sub

End Sub

Function Carrega_Arvore() As Long
'preenche a treeview Roteiro com a composicao de objRoteirosDeFabricacao
   
Dim objNode As Node
Dim lErro As Long, sChave As String, sChaveTvw As String
Dim iIndice As Integer
Dim sTexto As String
Dim objPO As New ClassPlanoOperacional
Dim objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes
Dim objCompetencias As ClassCompetencias
Dim objProduto As ClassProduto
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objOperacaoPai As ClassOrdemProducaoOperacoes

On Error GoTo Erro_Carrega_Arvore

    'Para cada Plano Operacional no Item do Plano Mestre
    For Each objPO In gobjPMPItem.ColPO
    
        Set objOrdemProducaoOperacoes = New ClassOrdemProducaoOperacoes
        
        For Each objOrdemProducaoOperacoes In gobjPMPItem.objItemOP.colOrdemProducaoOperacoes
        
            If objOrdemProducaoOperacoes.lNumIntDoc = objPO.lNumIntDocOper Then
                Exit For
            End If
                
        Next
                             
        Set objCompetencias = New ClassCompetencias
        
        objCompetencias.lNumIntDoc = objOrdemProducaoOperacoes.lNumIntDocCompet
        
        lErro = CF("Competencias_Le_NumIntDoc", objCompetencias)
        If lErro <> SUCESSO And lErro <> 134336 Then gError 137663
        
        'prepara texto que identificará a nova Operação que está sendo incluida
        sTexto = objCompetencias.sNomeReduzido
        
        Set objProduto = New ClassProduto
        
        objProduto.sCodigo = objOrdemProducaoOperacoes.sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 137664
        
        sTexto = sTexto & " (" & objProduto.sNomeReduzido
        
        If objPO.lNumIntDocCT > 0 Then
        
            Set objCentrodeTrabalho = New ClassCentrodeTrabalho
            
            objCentrodeTrabalho.lNumIntDoc = objPO.lNumIntDocCT
            
            lErro = CF("CentroDeTrabalho_Le_NumIntDoc", objCentrodeTrabalho)
            If lErro <> SUCESSO And lErro <> 134590 Then gError 137665
            
            If lErro = SUCESSO Then
        
                sTexto = sTexto & " - " & objCentrodeTrabalho.sNomeReduzido
                
            End If
           
        End If
        
        sTexto = sTexto & ")"
        
        'prepara uma chave para relacionar colComponentes ao node que está sendo incluido
        Call Calcula_Proxima_Chave(sChaveTvw)
        
        sChave = sChaveTvw
        sChaveTvw = sChaveTvw & objCompetencias.lCodigo
        
        If objOrdemProducaoOperacoes.iNivel = 0 Then
        
            Set objNode = Roteiro.Nodes.Add(, tvwFirst, sChaveTvw, sTexto)
                    
        Else
        
            For Each objOperacaoPai In gobjPMPItem.objItemOP.colOrdemProducaoOperacoes
                If objOrdemProducaoOperacoes.iSeqPai = objOperacaoPai.iSeq Then
                    Exit For
                End If
            Next
        
            Set objNode = Roteiro.Nodes.Add(objOperacaoPai.iPosicaoArvore, tvwChild, sChaveTvw, sTexto)
        
        End If
        
        objOrdemProducaoOperacoes.iPosicaoArvore = objNode.Index
                
        Roteiro.Nodes.Item(objNode.Index).Expanded = True
        
        colComponentes.Add objOrdemProducaoOperacoes, sChave
        
        objNode.Tag = sChave
                                        
    Next
    
    'se houver árvore ...
    If Roteiro.Nodes.Count > 0 Then
        
        'selecionar a raiz
        Set Roteiro.SelectedItem = Roteiro.Nodes.Item(1)
        Roteiro.SelectedItem.Selected = True
        
        'e carregar as operações pertinentes
        Call Roteiro_NodeClick(Roteiro.Nodes.Item(1))
        
        CodigoCTPadrao.Enabled = True
        Observacao.Enabled = True
        
    Else
    
        CodigoCTPadrao.Enabled = False
        Observacao.Enabled = False
        
    End If
    
    Carrega_Arvore = SUCESSO

    Exit Function

Erro_Carrega_Arvore:

    Carrega_Arvore = gErr

    Select Case gErr

        Case 137663 To 137665, 137903
            'erros tratados nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159664)

    End Select

    Exit Function

End Function

Function Preenche_Operacoes(objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes) As Long
'preenche as tabs de Detalhes e Produção à partir dos dados de objOrdemProducaoOperacoes

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProdutoMascarado As String
Dim objCompetencias As ClassCompetencias
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_Preenche_Operacoes

    'Limpa as Tabs de Detalhes e Insumos
    lErro = Limpa_Operacoes()
    If lErro <> SUCESSO Then gError 137666

    Nivel.Caption = objOrdemProducaoOperacoes.iNivel
    Sequencial.Caption = objOrdemProducaoOperacoes.iSeqArvore
    
    Set objProduto = New ClassProduto
    
    objProduto.sCodigo = objOrdemProducaoOperacoes.sProduto
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 137667
    
    lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 137668
    
    ProdutoLabel.Caption = sProdutoMascarado & SEPARADOR & objProduto.sDescricao
    VersaoLabel.Caption = objOrdemProducaoOperacoes.sVersao
    QtdeLabel.Caption = Formata_Estoque(gobjPO.dQuantidade)
    UMLabel.Caption = gobjPO.sUM
    
    Set objCompetencias = New ClassCompetencias
    
    objCompetencias.lNumIntDoc = objOrdemProducaoOperacoes.lNumIntDocCompet
    
    lErro = CF("Competencias_Le_NumIntDoc", objCompetencias)
    If lErro <> SUCESSO And lErro <> 134336 Then gError 137669
    
    CodigoCompetencia.Caption = objCompetencias.sNomeReduzido
    DescricaoCompetencia.Caption = objCompetencias.sDescricao
    
    If objOrdemProducaoOperacoes.iNumMaxMaqPorOper <> 0 Then
        NumMaxMaqPorOper.Caption = CStr(objOrdemProducaoOperacoes.iNumMaxMaqPorOper)
    End If
    
    If objOrdemProducaoOperacoes.iNumRepeticoes <> 0 Then
        Repeticao.Caption = CStr(objOrdemProducaoOperacoes.iNumRepeticoes)
    End If
    
    If gobjPO.lNumIntDocCT <> 0 Then
        
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
        
        objCentrodeTrabalho.lNumIntDoc = gobjPO.lNumIntDocCT
        
        lErro = CF("CentroDeTrabalho_Le_NumIntDoc", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134590 Then gError 137670
        
        CodigoCTPadrao.PromptInclude = False
        CodigoCTPadrao.Text = objCentrodeTrabalho.sNomeReduzido
        CodigoCTPadrao.PromptInclude = True
        
        DescricaoCTPadrao.Caption = objCentrodeTrabalho.sDescricao
    
    End If
    
    Observacao.Text = objOrdemProducaoOperacoes.sObservacao
    
    iAlterado = 0

    Preenche_Operacoes = SUCESSO

    Exit Function

Erro_Preenche_Operacoes:

    Preenche_Operacoes = gErr

    Select Case gErr
    
        Case 137666 To 137670
            'erros tratados nas rotinas chamadas
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159665)

    End Select

    Exit Function

End Function

Private Sub Roteiro_Collapse(ByVal Node As MSComctlLib.Node)
    Roteiro_NodeClick Node
End Sub

Function Limpa_Operacoes() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Operacoes
        
    Nivel.Caption = ""
    Sequencial.Caption = ""
    
    ProdutoLabel.Caption = ""
    VersaoLabel.Caption = ""
    QtdeLabel.Caption = ""
    UMLabel.Caption = ""
    
    CodigoCompetencia.Caption = ""
    DescricaoCompetencia.Caption = ""
    
    CodigoCTPadrao.PromptInclude = False
    CodigoCTPadrao.Text = ""
    CodigoCTPadrao.PromptInclude = True
    
    DescricaoCTPadrao.Caption = ""
    
    Observacao.Text = ""
    
    NumMaxMaqPorOper.Caption = ""
    
    iAlterado = 0

    Limpa_Operacoes = SUCESSO

    Exit Function

Erro_Limpa_Operacoes:

    Limpa_Operacoes = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159666)

    End Select

    Exit Function

End Function

Function Limpa_Arvore_Roteiro() As Long
'Limpa a Arvore do Roteiro

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Arvore_Roteiro

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 137671

    Roteiro.Nodes.Clear
    Set colComponentes = New Collection
    
    iProxChave = 1

    Limpa_Arvore_Roteiro = SUCESSO

    Exit Function

Erro_Limpa_Arvore_Roteiro:

    Limpa_Arvore_Roteiro = gErr
    
    Select Case gErr

        Case 137671

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159667)

    End Select

    Exit Function

End Function

Sub Calcula_Proxima_Chave(sChave As String)

Dim iNumero As Integer

    iNumero = iProxChave
    iProxChave = iProxChave + 1
    sChave = "X" & Right$(CStr(100000 + iNumero), 5)

End Sub

Private Function Move_Operacoes_Memoria(ByVal objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes, ByVal objCompetencias As ClassCompetencias, ByVal objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Operacoes_Memoria
        
    objCompetencias.sNomeReduzido = CodigoCompetencia.Caption
    
    'Verifica a Competencia no BD a partir do Código
    lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
    If lErro <> SUCESSO And lErro <> 134937 Then gError 137672

    objOrdemProducaoOperacoes.lNumIntDocCompet = objCompetencias.lNumIntDoc
    
    If Len(Trim(CodigoCTPadrao.Text)) <> 0 Then
            
        objCentrodeTrabalho.sNomeReduzido = CodigoCTPadrao.Text
        
        'Lê o CentrodeTrabalho que está sendo Passado
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 137673
        
        objOrdemProducaoOperacoes.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
    
    End If
    
    If Len(Trim(Observacao.Text)) <> 0 Then objOrdemProducaoOperacoes.sObservacao = Observacao.Text
    
    Move_Operacoes_Memoria = SUCESSO

    Exit Function

Erro_Move_Operacoes_Memoria:

    Move_Operacoes_Memoria = gErr

    Select Case gErr

        Case 137672, 137673
            'erros tratados nas rotinas chamadas
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159668)

    End Select

    Exit Function

End Function

Sub Recalcula_Nivel_Sequencial()
'(re)calcula niveis e sequencias de toda a estrutura
'deve ser chamada apos a remocao de algum node

Dim iIndice As Integer

    If Roteiro.Nodes.Count = 0 Then Exit Sub

    For iIndice = LBound(aNivelSequencial) To UBound(aNivelSequencial)
        aNivelSequencial(iIndice) = 0
    Next

    iUltimoNivel = 0

    'chamar rotina que recalcula recursivamente os campos nivel e sequencial (Nivel e SeqArvore)
    Call Calcula_Nivel_Sequencial(Roteiro.Nodes.Item(1), 0, 0)

End Sub

Sub Calcula_Nivel_Sequencial(objNode As Node, iNivel As Integer, iPosicaoAtual As Integer)
'parte recursiva do recalculo de nivel e sequencial, atuando a partir do node passado
'iNivel informa o nivel deste node

Dim objOrdemProducaoOperacoes As New ClassOrdemProducaoOperacoes
Dim sChave1 As String

    sChave1 = objNode.Tag

    Set objOrdemProducaoOperacoes = colComponentes.Item(sChave1)

    aNivelSequencial(iNivel) = aNivelSequencial(iNivel) + 1

    iPosicaoAtual = iPosicaoAtual + 1
    aSeqPai(iNivel) = iPosicaoAtual

    objOrdemProducaoOperacoes.iSeqArvore = aNivelSequencial(iNivel)

    If iNivel > 0 Then
        objOrdemProducaoOperacoes.iSeqRoteiroPai = aSeqPai(iNivel - 1)
    Else
        objOrdemProducaoOperacoes.iSeqRoteiroPai = 0
    End If
    
    objOrdemProducaoOperacoes.iSeqRoteiro = iPosicaoAtual

    objOrdemProducaoOperacoes.iNivelRoteiro = iNivel
    
    colComponentes.Remove sChave1
    colComponentes.Add objOrdemProducaoOperacoes, sChave1

    If objNode.Children > 0 Then
        Call Calcula_Nivel_Sequencial(objNode.Child, iNivel + 1, iPosicaoAtual)
    End If

    If objNode.Index <> objNode.LastSibling.Index Then Call Calcula_Nivel_Sequencial(objNode.Next, iNivel, iPosicaoAtual)

    If iNivel > iUltimoNivel Then iUltimoNivel = iNivel
   
End Sub

Private Function Habilita_MRP() As Long
        
Dim lErro As Long

On Error GoTo Erro_Habilita_MRP

    If MRP.Value = vbChecked Then
    
        LabelMRP.Enabled = True
        
        LabelDataInicio.Enabled = True
        DataInicio.Enabled = True
        UpDownDataInicio.Enabled = True
     
        LabelDataFinal.Enabled = True
        DataFinal.Enabled = True
        UpDownDataFinal.Enabled = True
                
        LabelOP.Enabled = True
        OPCodigoMRP.Enabled = True
        LabelStatusMRP.Enabled = True
        
        LabelMaquinas.Enabled = True
                    
    Else
    
        LabelMRP.Enabled = False
    
        LabelDataInicio.Enabled = False
        DataInicio.PromptInclude = False
        DataInicio.Text = ""
        DataInicio.PromptInclude = True
        DataInicio.Enabled = False
        UpDownDataInicio.Enabled = False
        
        LabelDataFinal.Enabled = False
        DataFinal.PromptInclude = False
        DataFinal.Text = ""
        DataFinal.PromptInclude = True
        DataFinal.Enabled = False
        UpDownDataFinal.Enabled = False
        
        LabelOP.Enabled = False
        OPCodigoMRP.Text = ""
        OPCodigoMRP.Enabled = False
        LabelStatusMRP.Enabled = False
        StatusMRP.Text = ""
        
        LabelMaquinas.Enabled = False
        
        Call Grid_Limpa(objGridMaquinas)
            
    End If
    
    Habilita_MRP = SUCESSO
    
    Exit Function
    
Erro_Habilita_MRP:

    Habilita_MRP = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159669)

    End Select

    Exit Function

End Function

Private Sub GridMaquinas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridMaquinas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMaquinas, iAlterado)
    End If

End Sub

Private Sub GridMaquinas_GotFocus()
    
    Call Grid_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub GridMaquinas_EnterCell()

    Call Grid_Entrada_Celula(objGridMaquinas, iAlterado)

End Sub

Private Sub GridMaquinas_LeaveCell()
    
    Call Saida_Celula(objGridMaquinas)

End Sub

Private Sub GridMaquinas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridMaquinas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMaquinas, iAlterado)
    End If

End Sub

Private Sub GridMaquinas_Scroll()

    Call Grid_Scroll(objGridMaquinas)

End Sub

Private Function Inicializa_GridMaquinas(objGrid As AdmGrid) As Long
'Inserido por Jorge Specian - 10/05/2005

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Maquina")
    objGrid.colColuna.Add ("Qtd Maq")
    objGrid.colColuna.Add ("Horas")
    objGrid.colColuna.Add ("Data")
    objGrid.colColuna.Add ("Tx.Prod.")

    'Controles que participam do Grid
    objGrid.colCampo.Add (NomeMaquina.Name)
    objGrid.colCampo.Add (QuantidadeMaquina.Name)
    objGrid.colCampo.Add (Horas.Name)
    objGrid.colCampo.Add (Data.Name)
    objGrid.colCampo.Add (TaxaProducao.Name)

    'Colunas do Grid
    iGrid_NomeMaquina_Col = 1
    iGrid_QuantidadeMaquina_Col = 2
    iGrid_Horas_Col = 3
    iGrid_Data_Col = 4
    iGrid_TaxaProducao_Col = 5

    objGrid.objGrid = GridMaquinas

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 4

    'Largura da primeira coluna
    GridMaquinas.ColWidth(0) = 250

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    Call Grid_Inicializa(objGrid)

    Inicializa_GridMaquinas = SUCESSO

End Function

Private Function Saida_Celula_QuantidadeMaquina(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade de máquinas do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iQuantAnterior As Integer
Dim objPOMaquinas As ClassPOMaquinas

On Error GoTo Erro_Saida_Celula_QuantidadeMaquina

    Set objGridInt.objControle = QuantidadeMaquina
    
    'se a quantidade foi preenchida
    If Len(QuantidadeMaquina.ClipText) > 0 Then

        lErro = Valor_Inteiro_Critica(QuantidadeMaquina.Text)
        If lErro <> SUCESSO Then gError 136405
        
        Set objPOMaquinas = gobjPO.colAlocacaoMaquinas.Item(GridMaquinas.Row)
        
        If objPOMaquinas.iQuantidade <> StrParaInt(QuantidadeMaquina.Text) Then
        
            iQuantAnterior = objPOMaquinas.iQuantidade

            If StrParaInt(QuantidadeMaquina.Text) > StrParaInt(NumMaxMaqPorOper.Caption) Then gError 141615
                
            objPOMaquinas.iQuantidade = StrParaInt(QuantidadeMaquina.Text)
            
            GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_QuantidadeMaquina_Col) = CStr(QuantidadeMaquina.Text)
            
            lErro = Acerta_Data_Inicio(gobjPO)
            If lErro <> SUCESSO Then gError 136408
            
            objPOMaquinas.iQuantidade = iQuantAnterior
        
        End If
                
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 136794

    Saida_Celula_QuantidadeMaquina = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantidadeMaquina:

    Saida_Celula_QuantidadeMaquina = gErr

    gobjPO.colAlocacaoMaquinas.Item(GridMaquinas.Row).iQuantidade = iQuantAnterior

    Select Case gErr

        Case 136405, 136408, 136794
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 141615
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_SUPERIOR_MAXIMA", gErr, StrParaInt(QuantidadeMaquina.Text), StrParaInt(NumMaxMaqPorOper.Caption))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159670)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Horas(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade de máquinas do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dHoraAnterior As Double
Dim objPOMaquinas As ClassPOMaquinas

On Error GoTo Erro_Saida_Celula_Horas

    Set objGridInt.objControle = Horas
    
    'se a quantidade de horas foi preenchida
    If Len(Horas.ClipText) > 0 Then

        lErro = Valor_Double_Critica(Horas.Text)
        If lErro <> SUCESSO Then gError 137674
        
        Set objPOMaquinas = gobjPO.colAlocacaoMaquinas.Item(GridMaquinas.Row)
        
        If Abs(objPOMaquinas.dHorasMaquina - StrParaDbl(Horas.Text)) > QTDE_ESTOQUE_DELTA Then
        
            dHoraAnterior = objPOMaquinas.dHorasMaquina
        
            objPOMaquinas.dHorasMaquina = StrParaDbl(Horas.Text)
            
            gobjPO.iAlterado = REGISTRO_ALTERADO
            
            GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_Horas_Col) = Formata_Estoque(Horas.Text)
                       
            lErro = Acerta_Data_Inicio(gobjPO)
            If lErro <> SUCESSO And lErro <> 134941 Then gError 136404
        
            objPOMaquinas.dHorasMaquina = dHoraAnterior
        
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 137675

    Saida_Celula_Horas = SUCESSO

    Exit Function

Erro_Saida_Celula_Horas:

    gobjPO.colAlocacaoMaquinas.Item(GridMaquinas.Row).dHorasMaquina = dHoraAnterior

    Saida_Celula_Horas = gErr

    Select Case gErr

        Case 137674, 137675, 136404
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159671)

    End Select

    Exit Function

End Function
Public Sub Show()
'    Parent.Show
'    Parent.SetFocus
End Sub

Private Sub NomeMaquina_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub NomeMaquina_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Private Sub NomeMaquina_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = NomeMaquina
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Private Sub QuantidadeMaquina_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub QuantidadeMaquina_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Private Sub QuantidadeMaquina_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = QuantidadeMaquina
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Horas_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub Horas_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Private Sub Horas_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = Horas
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Data_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub Data_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = Data
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub
Private Sub TaxaProducao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub TaxaProducao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Private Sub TaxaProducao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = TaxaProducao
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    'Verifica se Quantidade está preenchida
    If Len(Trim(Quantidade.Text)) <> 0 Then

        'Critica a Quantidade
        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 137676

    End If

    Exit Sub

Erro_Valor_Validate:

    Cancel = True

    Select Case gErr

        Case 137676

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159672)

    End Select

    Exit Sub

End Sub

Private Sub Quantidade_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Quantidade, iAlterado)
    
End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub


Private Sub Prioridade_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    'Verifica se Prioridade está preenchida
    If Len(Trim(Prioridade.Text)) <> 0 Then
    
        lErro = Valor_Inteiro_Critica(Prioridade.Text)
        If lErro <> SUCESSO Then gError 137677
    
    End If

    Exit Sub

Erro_Valor_Validate:

    Cancel = True

    Select Case gErr

        Case 137677

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159673)

    End Select

    Exit Sub

End Sub

Private Sub Prioridade_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Prioridade, iAlterado)
    
End Sub

Private Sub Prioridade_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub DataNecessidade_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    'Verifica se DataNecessidade está preenchida
    If Len(Trim(DataNecessidade.Text)) <> 0 Then

        'Critica a DataNecessidade
        lErro = Data_Critica(DataNecessidade.Text)
        If lErro <> SUCESSO Then gError 136416

    End If

    Exit Sub

Erro_Valor_Validate:

    Cancel = True

    Select Case gErr

        Case 136416

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159674)

    End Select

    Exit Sub

End Sub

Private Sub DataNecessidade_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

'###############################################################################
'Inserido por Wagner
Function Preenche_MRP(objPO As ClassPlanoOperacional) As Long
'preenche as tabs de Detalhes, Insumos e Produção à partir dos dados de objOrdemProducaoOperacoes

Dim lErro As Long
Dim objPOMaquinas As ClassPOMaquinas
Dim iIndice As Integer
Dim objTaxa As ClassTaxaDeProducao
Dim objMaquina As ClassMaquinas

On Error GoTo Erro_Preenche_MRP

    MRP.Value = Checked
    
    Call Grid_Limpa(objGridMaquinas)
    
    lErro = Habilita_MRP()
    If lErro <> SUCESSO Then gError 137678
    
    DataInicio.PromptInclude = False
    DataInicio.Text = Format(objPO.dtDataInicio, "dd/mm/yy")
    DataInicio.PromptInclude = True
    
    DataFinal.PromptInclude = False
    DataFinal.Text = Format(objPO.dtDataFim, "dd/mm/yy")
    DataFinal.PromptInclude = True
    
    OPCodigoMRP.Text = objPO.sCodOPOrigem
    
    If objPO.iStatus <> 0 Then StatusMRP.Text = objPO.iStatus & SEPARADOR & objPO.sDescErro
    
    iIndice = 0
    
    For Each objPOMaquinas In objPO.colAlocacaoMaquinas
    
        iIndice = iIndice + 1
        
        Set objTaxa = objPOMaquinas.objTaxaProducao
        Set objMaquina = New ClassMaquinas
        
        objMaquina.lNumIntDoc = objPOMaquinas.lNumIntDocMaq
        
        lErro = CF("Maquinas_Le_NumIntDoc", objMaquina)
        If lErro <> SUCESSO And lErro <> 106353 Then gError 136409
    
        GridMaquinas.TextMatrix(iIndice, iGrid_Data_Col) = objPOMaquinas.dtData
        GridMaquinas.TextMatrix(iIndice, iGrid_Horas_Col) = Formata_Estoque(objPOMaquinas.dHorasMaquina)
        GridMaquinas.TextMatrix(iIndice, iGrid_NomeMaquina_Col) = objMaquina.sNomeReduzido
        GridMaquinas.TextMatrix(iIndice, iGrid_QuantidadeMaquina_Col) = Format(objPOMaquinas.iQuantidade, "###")
        If objTaxa.iTipo = ITEM_TIPO_TAXAPRODUCAO_FIXO Then
            GridMaquinas.TextMatrix(iIndice, iGrid_TaxaProducao_Col) = Formata_Estoque(objTaxa.dTempoOperacao) & " /" & objTaxa.sUMTempo
        Else
            GridMaquinas.TextMatrix(iIndice, iGrid_TaxaProducao_Col) = Formata_Estoque(objTaxa.dQuantidade / objTaxa.dTempoOperacao) & " " & objTaxa.sUMProduto & "/" & objTaxa.sUMTempo
        End If
    Next
    
    objGridMaquinas.iLinhasExistentes = iIndice

    Preenche_MRP = SUCESSO

    Exit Function

Erro_Preenche_MRP:

    Preenche_MRP = gErr

    Select Case gErr
    
        Case 136409, 137678
            'erros tratados nas rotinas chamadas
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159675)

    End Select

    Exit Function

End Function

Function Limpa_MRP() As Long
'Limpa os dados relativos a simulação do MRP

Dim lErro As Long

On Error GoTo Erro_Limpa_MRP

    MRP.Value = vbUnchecked
    MRP.Enabled = False
    
    Call Habilita_MRP

    DataInicio.PromptInclude = False
    DataInicio.Text = ""
    DataInicio.PromptInclude = True
    
    DataFinal.PromptInclude = False
    DataFinal.Text = ""
    DataFinal.PromptInclude = True
    
    OPCodigoMRP.Text = ""
    
    StatusMRP.Text = ""
    
    Call Grid_Limpa(objGridMaquinas)
    
    Limpa_MRP = SUCESSO

    Exit Function

Erro_Limpa_MRP:

    Limpa_MRP = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159676)

    End Select

    Exit Function

End Function

Function Move_MRP_Memoria(objPO As ClassPlanoOperacional) As Long
'preenche as tabs de Detalhes, Insumos e Produção à partir dos dados de objOrdemProducaoOperacoes

Dim lErro As Long
Dim objPOMaquinas As ClassPOMaquinas
Dim iIndice As Integer

On Error GoTo Erro_Move_MRP_Memoria
    
    objPO.iAlterado = REGISTRO_ALTERADO
    
    objPO.dtDataFim = StrParaDate(DataFinal.Text)
    objPO.dtDataInicio = StrParaDate(DataInicio.Text)
    
    objPO.sCodOPOrigem = OPCodigoMRP.Text

    iIndice = 0
    
    For Each objPOMaquinas In objPO.colAlocacaoMaquinas
    
        iIndice = iIndice + 1
        
        objPOMaquinas.iQuantidade = StrParaInt(GridMaquinas.TextMatrix(iIndice, iGrid_QuantidadeMaquina_Col))
    
    Next
    
    objGridMaquinas.iLinhasExistentes = iIndice

    Move_MRP_Memoria = SUCESSO

    Exit Function

Erro_Move_MRP_Memoria:

    Move_MRP_Memoria = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159677)

    End Select

    Exit Function

End Function

Function Acerta_Data_Inicio(objPO As ClassPlanoOperacional) As Long

Dim lErro As Long
Dim objCT As New ClassCentrodeTrabalho
Dim iNodeAnterior As Integer
Dim dtDataFimAnterior As Date
Dim dtDataIniAnterior As Date

On Error GoTo Erro_Acerta_Data_Inicio

    If Len(Trim(CodigoCTPadrao.Text)) <> 0 Then
       
        dtDataFimAnterior = objPO.dtDataFim
        dtDataIniAnterior = objPO.dtDataInicio
   
        lErro = Move_MRP_Memoria(objPO)
        If lErro <> SUCESSO Then gError 136410
    
        objCT.sNomeReduzido = CodigoCTPadrao.Text
        objCT.iFilialEmpresa = giFilialEmpresa
        
        'Lê o CentrodeTrabalho que está sendo Passado
        lErro = CF("CentrodeTrabalho_Le_Completo", objCT)
        If lErro <> SUCESSO And lErro <> 137210 Then gError 136411
                               
        'Acerta datas das etapas filhas
        lErro = CF("PlanoOperacional_AcertaDatas", gobjPMP, gobjPMPItem, objPO, objCT, MRP_ACERTA_POR_DATA_FIM)
        If lErro <> SUCESSO Then gError 136412
        
        lErro = Preenche_MRP(objPO)
        If lErro <> SUCESSO Then gError 136843
        
    End If
    
    'Guarda o indice do nó que está
    iNodeAnterior = Roteiro.SelectedItem.Index
    
    'Remonta a árvore novamente
    lErro = Trata_Arvore(gobjPMPItem.sProduto, gobjPMPItem.sVersao, gobjPMPItem.sUM, gobjPMPItem.dQuantidade)
    If lErro <> SUCESSO Then gError 138221
    
    'selecionar o nó que estava antes
    Set Roteiro.SelectedItem = Roteiro.Nodes.Item(iNodeAnterior)
    Roteiro.SelectedItem.Selected = True
    
    'e carregar as operações pertinentes
    Call Roteiro_NodeClick(Roteiro.Nodes.Item(iNodeAnterior))
    
    Acerta_Data_Inicio = SUCESSO

    Exit Function

Erro_Acerta_Data_Inicio:

    Acerta_Data_Inicio = gErr

    objPO.dtDataFim = dtDataFimAnterior
    objPO.dtDataInicio = dtDataIniAnterior

    Select Case gErr
    
        Case 136410 To 136412, 136843, 138221
        
        Case 136791
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAFIMPRODUCAO_MAIOR_DATANECESSIDADE", gErr)
   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159678)

    End Select

    Exit Function

End Function

Function Acerta_Data_Fim(objPO As ClassPlanoOperacional) As Long

Dim lErro As Long
Dim objCT As New ClassCentrodeTrabalho
Dim iNodeAnterior As Integer
Dim dtDataFimAnterior As Date
Dim dtDataIniAnterior As Date

On Error GoTo Erro_Acerta_Data_Fim

    If Len(Trim(CodigoCTPadrao.Text)) <> 0 Then

        dtDataIniAnterior = objPO.dtDataInicio
        dtDataFimAnterior = objPO.dtDataFim
        
        lErro = Move_MRP_Memoria(objPO)
        If lErro <> SUCESSO Then gError 136413
        
        objCT.sNomeReduzido = CodigoCTPadrao.Text
        objCT.iFilialEmpresa = giFilialEmpresa
        
        'Lê o CentrodeTrabalho que está sendo Passado
        lErro = CF("CentrodeTrabalho_Le_Completo", objCT)
        If lErro <> SUCESSO And lErro <> 137210 Then gError 136414
                       
        'Acerta data das etapas filhas
        lErro = CF("PlanoOperacional_AcertaDatas", gobjPMP, gobjPMPItem, objPO, objCT, MRP_ACERTA_POR_DATA_INICIO)
        If lErro <> SUCESSO Then gError 136415

        lErro = Preenche_MRP(objPO)
        If lErro <> SUCESSO Then gError 136842

    End If
    
    'Guarda o indice do nó que está
    iNodeAnterior = Roteiro.SelectedItem.Index
    
    'Remonta a árvore novamente
    lErro = Trata_Arvore(gobjPMPItem.sProduto, gobjPMPItem.sVersao, gobjPMPItem.sUM, gobjPMPItem.dQuantidade)
    If lErro <> SUCESSO Then gError 138220
    
    'selecionar o nó que estava antes
    Set Roteiro.SelectedItem = Roteiro.Nodes.Item(iNodeAnterior)
    Roteiro.SelectedItem.Selected = True
    
    'e carregar as operações pertinentes
    Call Roteiro_NodeClick(Roteiro.Nodes.Item(iNodeAnterior))
        
    Acerta_Data_Fim = SUCESSO

    Exit Function

Erro_Acerta_Data_Fim:

    Acerta_Data_Fim = gErr

    objPO.dtDataFim = dtDataFimAnterior
    objPO.dtDataInicio = dtDataIniAnterior

    Select Case gErr
    
        Case 136413 To 136415, 136842, 138220
    
        Case 136792
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAFIMPRODUCAO_MAIOR_DATANECESSIDADE", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159679)

    End Select

    Exit Function

End Function

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(Trim(DataFinal.ClipText)) > 0 Then
    
        lErro = Data_Critica(DataFinal.Text)
        If lErro <> SUCESSO Then gError 136416

        'Se mudou a hora
        If StrParaDate(DataFinal.Text) <> gobjPO.dtDataFim Then
        
            gobjPO.iAlterado = REGISTRO_ALTERADO
            
            'Acerta a data Inicial e hora Inicial da etapa e as datas das etapas descendentes
            lErro = Acerta_Data_Inicio(gobjPO)
            If lErro <> SUCESSO Then gError 136417
            
            'Os status deixam de ser válidos
            Call Limpa_MRP_Status
            
        End If
        
    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr
    
        Case 136416, 136417
            'erros tratados nas rotinas chamadas
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159680)

    End Select

    Exit Sub

End Sub

Private Sub DataInicio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicio_Validate

    If Len(Trim(DataInicio.ClipText)) > 0 Then
    
        lErro = Data_Critica(DataInicio.Text)
        If lErro <> SUCESSO Then gError 136420
        
        'Se mudou a data
        If StrParaDate(DataInicio.Text) <> gobjPO.dtDataInicio Then
        
            gobjPO.iAlterado = REGISTRO_ALTERADO
            
            'Acerta a data final e hora final da etapa e as datas das etapas descendentes
            lErro = Acerta_Data_Fim(gobjPO)
            If lErro <> SUCESSO Then gError 136421
            
            'Os status deixam de ser válidos
            Call Limpa_MRP_Status
            
        End If

    End If

    Exit Sub

Erro_DataInicio_Validate:

    Cancel = True

    Select Case gErr
    
        Case 136420, 136421
            'erros tratados nas rotinas chamadas
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159681)

    End Select

    Exit Sub
    
End Sub

Private Sub DataInicio_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataInicio, iAlterado)
End Sub

Private Sub DataFinal_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataFinal, iAlterado)
End Sub

Private Sub Limpa_MRP_Status()
'Os status deixam de ser válidos

Dim lErro As Long
Dim objPO As ClassPlanoOperacional

On Error GoTo Erro_Limpa_MRP_Status

    'Tanto o item do Plano Mestre quanto as etapas perdem o seu status
    gobjPMPItem.iStatus = 0
    gobjPMPItem.sDescErro = ""
    
    StatusMRP.Text = ""
    
    For Each objPO In gobjPMPItem.ColPO
    
        objPO.iStatus = 0
        objPO.sDescErro = ""
    
    Next

    Exit Sub

Erro_Limpa_MRP_Status:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159682)

    End Select

    Exit Sub
    
End Sub

Private Sub OPCodigoMRP_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_OPCodigoMRP_Validate

    If gobjPO.sCodOPOrigem <> OPCodigoMRP.Text Then
    
        'Não é permitido mudar o código da OP ancestral, só
        'das SubOps
        If gobjPO.iNivel = 0 Then gError 136424
        
        'Coloca no objeto global o código da subOP
        gobjPO.sCodOPOrigem = OPCodigoMRP.Text

    End If

    Exit Sub

Erro_OPCodigoMRP_Validate:

    Cancel = True

    Select Case gErr
    
        Case 136424
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERAR_CODIGO_OP", gErr, gobjPO.sCodOPOrigem)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159683)

    End Select

    Exit Sub
    
End Sub

Function DataFinal_Valida(objPO As ClassPlanoOperacional) As Long

Dim lErro As Long
Dim objPOPai As ClassPlanoOperacional

On Error GoTo Erro_DataFinal_Valida

    'Para cada etapa
    For Each objPOPai In gobjPMPItem.ColPO
    
        'Verifica se a etapa corrente tem pai
        If objPOPai.lNumIntDoc = objPO.lNumIntDocPOPai Then
            
            'Se a etapa filha termina depois da etapa pai começar
            'Erro
            If objPOPai.dtDataInicio < objPO.dtDataFim Then gError 136425
        
        End If
    
    Next

    DataFinal_Valida = SUCESSO

    Exit Function

Erro_DataFinal_Valida:

    DataFinal_Valida = gErr

    Select Case gErr
    
        Case 136425, 136426
            Call Rotina_Erro(vbOKOnly, "ERRO_PO_DATA_FINAL_INVALIDA", gErr, Error, objPO.iNivel, objPO.iSeq, objPOPai.iNivel, objPOPai.iSeq)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159684)

    End Select

    Exit Function

End Function
'#############################################################################

Function PMPItem_Cria_Copia(ByVal objPMPItem As ClassPMPItens, ByVal objPMPItemCopia As ClassPMPItens) As Long
'Cria uma cópia do obj Global caso o botão cancelar seja clicado

Dim lErro As Long
Dim objPO As ClassPlanoOperacional
Dim objPOCopia As ClassPlanoOperacional
Dim obj

On Error GoTo Erro_PMPItem_Cria_Copia

    'Esse obj não serão modificados
    Set objPMPItemCopia.objItemOP = objPMPItem.objItemOP
    Set objPMPItemCopia.objOP = objPMPItem.objOP

    'Copia atributos
    objPMPItemCopia.dQuantidade = objPMPItem.dQuantidade
    objPMPItemCopia.dtDataNecessidade = objPMPItem.dtDataNecessidade
    objPMPItemCopia.iAlterado = objPMPItem.iAlterado
    objPMPItemCopia.iFilialCli = objPMPItem.iFilialCli
    objPMPItemCopia.iFilialEmpresa = objPMPItem.iFilialEmpresa
    objPMPItemCopia.iOrdem = objPMPItem.iOrdem
    objPMPItemCopia.iPrioridade = objPMPItem.iPrioridade
    objPMPItemCopia.iStatus = objPMPItem.iStatus
    objPMPItemCopia.lCliente = objPMPItem.lCliente
    objPMPItemCopia.lCodGeracao = objPMPItem.lCodGeracao
    objPMPItemCopia.lNumIntDoc = objPMPItem.lNumIntDoc
    objPMPItemCopia.sCodOPOrigem = objPMPItem.sCodOPOrigem
    objPMPItemCopia.sDescErro = objPMPItem.sDescErro
    objPMPItemCopia.sProduto = objPMPItem.sProduto
    objPMPItemCopia.sUM = objPMPItem.sUM
    objPMPItemCopia.sVersao = objPMPItem.sVersao
    objPMPItemCopia.lUltimoProxPO = objPMPItem.lUltimoProxPO
    objPMPItemCopia.iProduzLogo = objPMPItem.iProduzLogo

    For Each objPO In objPMPItem.ColPO
    
        Set objPOCopia = New ClassPlanoOperacional
    
        'Não precisa copiar a coleção, só apontar porque quando modificada ela
        'recebe um novo endereço de memória deixando os dados do endereço anterior
        'sem alteração
        Set objPOCopia.colAlocacaoMaquinas = objPO.colAlocacaoMaquinas
        
        'Não é modificado nessa função
        Set objPOCopia.colOPFilhas = objPO.colOPFilhas
        Set objPOCopia.colRCFilhas = objPO.colRCFilhas
        Set objPOCopia.objApontamento = objPO.objApontamento
        Set objPOCopia.objOP = objPO.objOP
    
        'Copia atributos
        objPOCopia.dQtdTotal = objPO.dQtdTotal
        objPOCopia.dQuantidade = objPO.dQuantidade
        objPOCopia.dtDataFim = objPO.dtDataFim
        objPOCopia.dtDataInicio = objPO.dtDataInicio
        objPOCopia.dTempoGasto = objPOCopia.dTempoGasto
        objPOCopia.iAlterado = objPO.iAlterado
        objPOCopia.iFilialEmpresa = objPO.iFilialEmpresa
        objPOCopia.iNivel = objPO.iNivel
        objPOCopia.iSeq = objPO.iSeq
        objPOCopia.iStatus = objPO.iStatus
        objPOCopia.lNumIntDoc = objPO.lNumIntDoc
        objPOCopia.lNumIntDocCT = objPO.lNumIntDocCT
        objPOCopia.lNumIntDocOper = objPO.lNumIntDocOper
        objPOCopia.lNumIntDocPMP = objPO.lNumIntDocPMP
        objPOCopia.lNumIntDocPOPai = objPO.lNumIntDocPOPai
        objPOCopia.sCodOPOrigem = objPO.sCodOPOrigem
        objPOCopia.sDescErro = objPO.sDescErro
        objPOCopia.sProduto = objPO.sProduto
        objPOCopia.sUM = objPO.sUM
        objPOCopia.sVersao = objPO.sVersao
        objPOCopia.iNumMaxMaqPorOper = objPO.iNumMaxMaqPorOper
        objPOCopia.iNumRepeticoes = objPO.iNumRepeticoes
        
        objPMPItemCopia.ColPO.Add objPOCopia
    
    Next

    PMPItem_Cria_Copia = SUCESSO

    Exit Function

Erro_PMPItem_Cria_Copia:

    PMPItem_Cria_Copia = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159685)

    End Select

    Exit Function

End Function

Private Function Altera_Arvore() As Long

Dim lErro As Long
Dim objNode As Node
Dim sChave As String
Dim objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes
Dim objCompetencias As ClassCompetencias
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objProduto As ClassProduto
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sTexto As String

On Error GoTo Erro_Altera_Arvore

    Set objNode = Roteiro.SelectedItem

    If objNode Is Nothing Then gError 137899
    If objNode.Selected = False Then gError 137900
    
    Set objOrdemProducaoOperacoes = colComponentes.Item(objNode.Tag)
    Set objCompetencias = New ClassCompetencias
    Set objCentrodeTrabalho = New ClassCentrodeTrabalho
    
    'preenche objOperacoes à partir dos dados da tela
    lErro = Move_Operacoes_Memoria(objOrdemProducaoOperacoes, objCompetencias, objCentrodeTrabalho)
    If lErro <> SUCESSO Then gError 137901

    sChave = objNode.Tag
        
    'prepara texto que identificará a nova Operação que está sendo incluida
    
    sTexto = objCompetencias.sNomeReduzido
            
    Set objProduto = New ClassProduto
    
    objProduto.sCodigo = objOrdemProducaoOperacoes.sProduto
        
    'Le Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 137902
            
    sTexto = sTexto & " (" & objProduto.sNomeReduzido

    If Len(Trim(CodigoCTPadrao.ClipText)) <> 0 Then
       sTexto = sTexto & " - " & objCentrodeTrabalho.sNomeReduzido
    End If
        
    sTexto = sTexto & ")"

    objNode.Text = sTexto

    colComponentes.Remove (sChave)
    colComponentes.Add objOrdemProducaoOperacoes, sChave

    Call Recalcula_Nivel_Sequencial

    Altera_Arvore = SUCESSO

    Exit Function

Erro_Altera_Arvore:

    Altera_Arvore = gErr

    Select Case gErr

        Case 137899, 137900
            Call Rotina_Erro(vbOKOnly, "AVISO_SELECIONAR_ESTRUTURA_ROTEIRO", gErr)

        Case 137901, 137902
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159686)

    End Select

    Exit Function

End Function
