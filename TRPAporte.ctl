VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TRPAporte 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4875
      Index           =   1
      Left            =   135
      TabIndex        =   23
      Top             =   960
      Width           =   9240
      Begin VB.Frame Frame10 
         Caption         =   "Dados do Cliente"
         Height          =   1005
         Left            =   0
         TabIndex        =   82
         Top             =   1440
         Width           =   9195
         Begin VB.CommandButton BotaoAbrirCli 
            Caption         =   "Abrir"
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
            Left            =   2445
            TabIndex        =   8
            Top             =   195
            Width           =   660
         End
         Begin VB.CommandButton BotaoOutrosAportes 
            Caption         =   "Outros Aportes desse Cliente"
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
            Left            =   6450
            TabIndex        =   9
            Top             =   195
            Width           =   2640
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fil. Emp:"
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
            Index           =   11
            Left            =   3225
            TabIndex        =   90
            Top             =   255
            Width           =   750
         End
         Begin VB.Label FilialEmpresa 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3990
            TabIndex        =   89
            Top             =   225
            Width           =   1830
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
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
            Index           =   10
            Left            =   630
            TabIndex        =   88
            Top             =   255
            Width           =   645
         End
         Begin VB.Label CodigoCliente 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1335
            TabIndex        =   87
            Top             =   225
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ:"
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
            Index           =   51
            Left            =   6480
            TabIndex        =   86
            Top             =   630
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Razão Social:"
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
            Index           =   52
            Left            =   75
            TabIndex        =   85
            Top             =   615
            Width           =   1200
         End
         Begin VB.Label RazaoSocial 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1335
            TabIndex        =   84
            Top             =   585
            Width           =   4485
         End
         Begin VB.Label CNPJ 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7065
            TabIndex        =   83
            Top             =   585
            Width           =   2025
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Previsão"
         Height          =   930
         Left            =   0
         TabIndex        =   34
         Top             =   2475
         Width           =   9195
         Begin VB.Frame Frame8 
            Caption         =   "Período da Previsão"
            Height          =   645
            Left            =   135
            TabIndex        =   38
            Top             =   180
            Width           =   5085
            Begin MSComCtl2.UpDown UpDownPrevDe 
               Height          =   300
               Left            =   2310
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   255
               Width           =   225
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox PrevDataDe 
               Height          =   300
               Left            =   1185
               TabIndex        =   10
               Top             =   255
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox PrevDataAte 
               Height          =   300
               Left            =   3225
               TabIndex        =   12
               Top             =   240
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownPrevAte 
               Height          =   300
               Left            =   4395
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   240
               Width           =   225
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin VB.Label Label1 
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
               Index           =   3
               Left            =   795
               TabIndex        =   40
               Top             =   285
               Width           =   315
            End
            Begin VB.Label Label1 
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
               Index           =   4
               Left            =   2820
               TabIndex        =   39
               Top             =   285
               Width           =   360
            End
         End
         Begin MSMask.MaskEdBox Previsao 
            Height          =   315
            Left            =   7065
            TabIndex        =   14
            Top             =   180
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Realizado US$:"
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
            Index           =   6
            Left            =   4695
            TabIndex        =   37
            Top             =   585
            Width           =   2325
         End
         Begin VB.Label Realizado 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   7065
            TabIndex        =   36
            Top             =   525
            Width           =   2025
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Prev. Venda US$:"
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
            Height          =   315
            Index           =   5
            Left            =   5250
            TabIndex        =   35
            Top             =   240
            Width           =   1770
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Outros"
         Height          =   1440
         Left            =   0
         TabIndex        =   31
         Top             =   3435
         Width           =   9210
         Begin VB.ComboBox Historico 
            Height          =   315
            Left            =   1320
            TabIndex        =   15
            Top             =   210
            Width           =   7830
         End
         Begin VB.TextBox Observacao 
            Height          =   735
            Left            =   1320
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   16
            Top             =   615
            Width           =   7815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
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
            Index           =   8
            Left            =   45
            TabIndex        =   33
            Top             =   645
            Width           =   1200
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Histórico:"
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
            Index           =   7
            Left            =   195
            TabIndex        =   32
            Top             =   240
            Width           =   1050
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Identificação"
         Height          =   1380
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   9195
         Begin VB.ComboBox Moeda 
            Height          =   315
            Left            =   7080
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   615
            Width           =   2040
         End
         Begin VB.ComboBox Tipo 
            Height          =   315
            Left            =   1350
            TabIndex        =   4
            Top             =   615
            Width           =   4500
         End
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2235
            Picture         =   "TRPAporte.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Numeração Automática"
            Top             =   255
            Width           =   300
         End
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   7080
            TabIndex        =   7
            Top             =   975
            Width           =   2040
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   315
            Left            =   1350
            TabIndex        =   0
            Top             =   240
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   300
            Left            =   1350
            TabIndex        =   6
            Top             =   960
            Width           =   4485
            _ExtentX        =   7911
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataEmissao 
            Height          =   315
            Left            =   4275
            TabIndex        =   2
            Top             =   240
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataEmissao 
            Height          =   300
            Left            =   5595
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label SaldoInvest 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   7080
            TabIndex        =   81
            Top             =   225
            Width           =   2040
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Moeda:"
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
            Height          =   195
            Index           =   2
            Left            =   6405
            TabIndex        =   30
            Top             =   660
            Width           =   690
         End
         Begin VB.Label LabelCodigo 
            Alignment       =   1  'Right Justify
            Caption         =   "Código:"
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
            Height          =   315
            Left            =   480
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   29
            Top             =   270
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo:"
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
            Height          =   315
            Index           =   1
            Left            =   465
            TabIndex        =   28
            Top             =   660
            Width           =   810
         End
         Begin VB.Label LabelCliente 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
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
            Height          =   195
            Left            =   630
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   27
            Top             =   1005
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Filial:"
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
            Height          =   195
            Index           =   13
            Left            =   6570
            TabIndex        =   26
            Top             =   1035
            Width           =   555
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Data de Emissão:"
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
            Height          =   315
            Index           =   0
            Left            =   2730
            TabIndex        =   25
            Top             =   270
            Width           =   1515
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Saldo Total:"
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
            Index           =   9
            Left            =   4710
            TabIndex        =   80
            Top             =   285
            Width           =   2325
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4860
      Index           =   3
      Left            =   120
      TabIndex        =   49
      Top             =   915
      Visible         =   0   'False
      Width           =   9255
      Begin VB.CommandButton BotaoVerCred 
         Caption         =   "Créditos utilizados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   3
         Left            =   4635
         TabIndex        =   68
         Top             =   4230
         Width           =   2235
      End
      Begin VB.CommandButton BotaoConsultarDocumentoD 
         Caption         =   "Consultar Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   6945
         TabIndex        =   69
         Top             =   4230
         Width           =   2235
      End
      Begin VB.Frame Frame6 
         Caption         =   "Diretos"
         Height          =   3810
         Left            =   60
         TabIndex        =   50
         Top             =   285
         Width           =   9135
         Begin VB.ComboBox Forma 
            Height          =   315
            ItemData        =   "TRPAporte.ctx":00EA
            Left            =   2760
            List            =   "TRPAporte.ctx":00F4
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   1110
            Width           =   2655
         End
         Begin MSMask.MaskEdBox DocD 
            Height          =   315
            Left            =   6345
            TabIndex        =   51
            Top             =   765
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorD 
            Height          =   315
            Left            =   465
            TabIndex        =   53
            Top             =   870
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataVencimento 
            Height          =   315
            Left            =   1860
            TabIndex        =   54
            Top             =   495
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridD 
            Height          =   1200
            Left            =   75
            TabIndex        =   55
            Top             =   195
            Width           =   8985
            _ExtentX        =   15849
            _ExtentY        =   2117
            _Version        =   393216
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4935
      Index           =   4
      Left            =   105
      TabIndex        =   71
      Top             =   945
      Visible         =   0   'False
      Width           =   9255
      Begin VB.CommandButton BotaoHistSFCD 
         Caption         =   "Histórico de Utilização Detalhado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   4575
         TabIndex        =   72
         Top             =   4230
         Width           =   2235
      End
      Begin VB.Frame Frame7 
         Caption         =   "Sobre fatura por cumprimento de meta"
         Height          =   3810
         Left            =   60
         TabIndex        =   74
         Top             =   285
         Width           =   9135
         Begin MSMask.MaskEdBox DataValidadeAteSFC 
            Height          =   315
            Left            =   4080
            TabIndex        =   75
            Top             =   600
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercSFC 
            Height          =   315
            Left            =   1050
            TabIndex        =   76
            Top             =   615
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorSFC 
            Height          =   315
            Left            =   5505
            TabIndex        =   77
            Top             =   570
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataValidadeDeSFC 
            Height          =   315
            Left            =   2745
            TabIndex        =   78
            Top             =   585
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridSFC 
            Height          =   330
            Left            =   75
            TabIndex        =   79
            Top             =   195
            Width           =   8985
            _ExtentX        =   15849
            _ExtentY        =   582
            _Version        =   393216
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
      End
      Begin VB.CommandButton BotaoHistSFC 
         Caption         =   "Histórico de Utilização Resumido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   6915
         TabIndex        =   73
         Top             =   4230
         Width           =   2235
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4935
      Index           =   2
      Left            =   105
      TabIndex        =   41
      Top             =   915
      Visible         =   0   'False
      Width           =   9255
      Begin VB.CommandButton BotaoHistSFD 
         Caption         =   "Histórico de Utilização Detalhado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   4620
         TabIndex        =   66
         Top             =   4230
         Width           =   2235
      End
      Begin VB.CommandButton BotaoHistSF 
         Caption         =   "Histórico de Utilização Resumido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   6915
         TabIndex        =   67
         Top             =   4230
         Width           =   2235
      End
      Begin VB.Frame Frame5 
         Caption         =   "Sobre fatura"
         Height          =   3810
         Left            =   60
         TabIndex        =   42
         Top             =   285
         Width           =   9135
         Begin MSMask.MaskEdBox DataValidadeAte 
            Height          =   315
            Left            =   4080
            TabIndex        =   43
            Top             =   600
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorSF 
            Height          =   315
            Left            =   1050
            TabIndex        =   44
            Top             =   615
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Saldo 
            Height          =   315
            Left            =   5505
            TabIndex        =   45
            Top             =   570
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Percentual 
            Height          =   315
            Left            =   6945
            TabIndex        =   46
            Top             =   630
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataValidadeDe 
            Height          =   315
            Left            =   2745
            TabIndex        =   47
            Top             =   585
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridSF 
            Height          =   1200
            Left            =   75
            TabIndex        =   48
            Top             =   195
            Width           =   8985
            _ExtentX        =   15849
            _ExtentY        =   2117
            _Version        =   393216
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4950
      Index           =   5
      Left            =   105
      TabIndex        =   56
      Top             =   885
      Visible         =   0   'False
      Width           =   9255
      Begin VB.CommandButton BotaoConsultarDocumentoC 
         Caption         =   "Consultar Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   6945
         TabIndex        =   70
         Top             =   4230
         Width           =   2235
      End
      Begin VB.Frame Frame9 
         Caption         =   "Condicionados"
         Height          =   3810
         Left            =   60
         TabIndex        =   57
         Top             =   285
         Width           =   9135
         Begin VB.ComboBox FormaC 
            Height          =   315
            ItemData        =   "TRPAporte.ctx":0116
            Left            =   3345
            List            =   "TRPAporte.ctx":0120
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   825
            Width           =   1305
         End
         Begin VB.ComboBox Base 
            Height          =   315
            ItemData        =   "TRPAporte.ctx":0142
            Left            =   1095
            List            =   "TRPAporte.ctx":014C
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   600
            Width           =   1620
         End
         Begin MSMask.MaskEdBox ValorC 
            Height          =   315
            Left            =   240
            TabIndex        =   60
            Top             =   285
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox StatusC 
            Height          =   315
            Left            =   5670
            TabIndex        =   61
            Top             =   915
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DocC 
            Height          =   315
            Left            =   5520
            TabIndex        =   62
            Top             =   540
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercentualC 
            Height          =   315
            Left            =   7230
            TabIndex        =   63
            Top             =   555
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataPagto 
            Height          =   315
            Left            =   4245
            TabIndex        =   64
            Top             =   600
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridC 
            Height          =   1095
            Left            =   75
            TabIndex        =   65
            Top             =   225
            Width           =   8985
            _ExtentX        =   15849
            _ExtentY        =   1931
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   7320
      ScaleHeight     =   450
      ScaleWidth      =   2025
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   30
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "TRPAporte.ctx":0179
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "TRPAporte.ctx":02D3
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "TRPAporte.ctx":045D
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "TRPAporte.ctx":098F
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5370
      Left            =   75
      TabIndex        =   22
      Top             =   570
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   9472
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Inicial"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pagtos sobre fatura"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pagtos diretos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pagto sobre fatura por meta"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pagtos diretos por meta"
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
Attribute VB_Name = "TRPAporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variáveis Globais
Dim iFrameAtual As Integer
Dim iAlterado As Integer

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1

Dim gobjAporte As ClassTRPAportes
Dim gobjAportePagtoCond As ClassTRPAportePagtoCond
Dim gobjAportePagtoDireto As ClassTRPAportePagtoDireto
Dim gobjAportePagtoSF As ClassTRPAportePagtoFat
Dim gobjAportePagtoSFC As ClassTRPAportePagtoFatC
Dim giTipoAporte As Integer

Dim sClienteAnt As String
Dim dtDataDeAnt As Date
Dim dtDataAteAnt As Date

Dim iFilialCliAnt As Integer

'GridSF
Dim objGridSF As AdmGrid
Dim iGrid_ValorSF_Col As Integer
Dim iGrid_ValidadeDe_Col As Integer
Dim iGrid_ValidadeAte_Col As Integer
Dim iGrid_Percentual_Col As Integer
Dim iGrid_Saldo_Col As Integer

'GridSFC
Dim objGridSFC As AdmGrid
Dim iGrid_ValorSFC_Col As Integer
Dim iGrid_ValidadeDeSFC_Col As Integer
Dim iGrid_ValidadeAteSFC_Col As Integer
Dim iGrid_PercSFC_Col As Integer

Dim objGridD As AdmGrid
Dim iGrid_ValorD_Col As Integer
Dim iGrid_Vencimento_Col As Integer
Dim iGrid_Forma_Col As Integer
Dim iGrid_DocD_Col As Integer

Dim objGridC As AdmGrid
Dim iGrid_Base_Col As Integer
Dim iGrid_PercentualC_Col As Integer
Dim iGrid_Pagamento_Col As Integer
Dim iGrid_FormaC_Col As Integer
Dim iGrid_DocC_Col As Integer
Dim iGrid_StatusC_Col As Integer
Dim iGrid_ValorC_Col As Integer

Private Const FRAME_IDENTIFICACAO = 1
Private Const FRAME_PAGTOSSF = 2
Private Const FRAME_PAGTOSD = 3
Private Const FRAME_PAGTOSSFC = 4
Private Const FRAME_PAGTOSDC = 5

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Aportes"
    Call Form_Load

End Function

Public Function Name() As String
    Name = "TRPAporte"
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

Private Sub Filial_Click()
    Call Filial_Validate(bSGECancelDummy)
End Sub

Private Sub PrevDataDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PrevDataAte_Change()
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

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoCodigo = Nothing
    Set objEventoCliente = Nothing
    
    Set objGridSF = Nothing
    Set objGridSFC = Nothing
    Set objGridC = Nothing
    Set objGridD = Nothing
    
    Set gobjAporte = Nothing
    Set gobjAportePagtoCond = Nothing
    Set gobjAportePagtoDireto = Nothing
    Set gobjAportePagtoSF = Nothing

    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190564)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCodigo = New AdmEvento
    Set objEventoCliente = New AdmEvento
    
    Set gobjAporte = New ClassTRPAportes
    
    iAlterado = 0
    
    Set objGridSF = New AdmGrid
    Set objGridSFC = New AdmGrid
    Set objGridC = New AdmGrid
    Set objGridD = New AdmGrid
    
    lErro = Inicializa_Grid_SF(objGridSF)
    If lErro <> SUCESSO Then gError 190565
    
    lErro = Inicializa_Grid_SFC(objGridSFC)
    If lErro <> SUCESSO Then gError 190565
    
    lErro = Inicializa_Grid_D(objGridD)
    If lErro <> SUCESSO Then gError 190566
    
    lErro = Inicializa_Grid_C(objGridC)
    If lErro <> SUCESSO Then gError 190567
    
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TRPTIPOAPORTE_OCR, Tipo)
    If lErro <> SUCESSO Then gError 190568
    
    Historico.Clear
    lErro = CF("Carrega_Combo_Historico", Historico, "TRPAportes", STRING_TRP_APORTE_HISTORICO)
    If lErro <> SUCESSO Then gError 190569
    
    'carrega a combo de Moedas
    lErro = Carrega_Moeda()
    If lErro <> SUCESSO Then gError 190570
    
    lErro = CF("Carrega_Combo_Base", Base)
    If lErro <> SUCESSO Then gError 190741
        
    lErro = CF("Carrega_Combo_FormaPagto", Forma)
    If lErro <> SUCESSO Then gError 190742
        
    lErro = CF("Carrega_Combo_FormaPagto", FormaC)
    If lErro <> SUCESSO Then gError 190743
    
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    
    iFrameAtual = 1

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 190565 To 190570, 190741 To 190743

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190571)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objTRPAporte As ClassTRPAportes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objTRPAporte Is Nothing) Then

        lErro = Traz_TRPAporte_Tela(objTRPAporte)
        If lErro <> SUCESSO Then gError 190572

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 190572

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190573)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objTRPAporte As ClassTRPAportes) As Long

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iLinha As Integer
Dim objAportePagtoD As ClassTRPAportePagtoDireto
Dim objAportePagtoC As ClassTRPAportePagtoCond
Dim objAportePagtoSF As ClassTRPAportePagtoFat
Dim objAportePagtoSFC As ClassTRPAportePagtoFatC
Dim iNumPagto As Integer

On Error GoTo Erro_Move_Tela_Memoria

    objCliente.sNomeReduzido = Cliente.Text

    'Lê o Cliente através do Nome Reduzido
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 190574

    objTRPAporte.lCodigo = StrParaLong(Codigo.Text)
    objTRPAporte.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objTRPAporte.lCliente = objCliente.lCodigo
    objTRPAporte.iFilialCliente = Codigo_Extrai(Filial.Text)
    objTRPAporte.iMoeda = Codigo_Extrai(Moeda)
    objTRPAporte.iTipo = Codigo_Extrai(Tipo)
    objTRPAporte.dPrevValor = StrParaDbl(Previsao.Text)
    objTRPAporte.dtPrevDataDe = StrParaDate(PrevDataDe.Text)
    objTRPAporte.dtPrevDataAte = StrParaDate(PrevDataAte.Text)
    objTRPAporte.sObservacao = Observacao.Text
    objTRPAporte.sHistorico = Historico.Text
    
    For iLinha = 1 To objGridD.iLinhasExistentes
    
        iNumPagto = iNumPagto + 1
        
        Set objAportePagtoD = New ClassTRPAportePagtoDireto
    
        objAportePagtoD.dValor = StrParaDbl(GridD.TextMatrix(iLinha, iGrid_ValorD_Col))
        objAportePagtoD.iFormaPagto = Codigo_Extrai(GridD.TextMatrix(iLinha, iGrid_Forma_Col))
        objAportePagtoD.dtVencimento = StrParaDate(GridD.TextMatrix(iLinha, iGrid_Vencimento_Col))
        objAportePagtoD.lNumIntDoc = gobjAporte.colPagtoDireto.Item(iLinha).lNumIntDoc
        objAportePagtoD.lNumIntDocAporte = gobjAporte.colPagtoDireto.Item(iLinha).lNumIntDocAporte
        objAportePagtoD.lNumIntDocDestino = gobjAporte.colPagtoDireto.Item(iLinha).lNumIntDocDestino
        objAportePagtoD.iTipoDocDestino = gobjAporte.colPagtoDireto.Item(iLinha).iTipoDocDestino
        objAportePagtoD.iSeq = iLinha
    
        If objAportePagtoD.dValor = 0 Then gError 190700
        If objAportePagtoD.iFormaPagto = 0 Then gError 190701
        If objAportePagtoD.dtVencimento = DATA_NULA Then gError 190702
    
        objTRPAporte.colPagtoDireto.Add objAportePagtoD
    
    Next
    
    For iLinha = 1 To objGridC.iLinhasExistentes
        
        iNumPagto = iNumPagto + 1
        
        Set objAportePagtoC = New ClassTRPAportePagtoCond
    
        objAportePagtoC.iBase = Codigo_Extrai(GridC.TextMatrix(iLinha, iGrid_Base_Col))
        objAportePagtoC.dPercentual = PercentParaDbl(GridC.TextMatrix(iLinha, iGrid_PercentualC_Col))
        objAportePagtoC.iFormaPagto = Codigo_Extrai(GridC.TextMatrix(iLinha, iGrid_FormaC_Col))
        objAportePagtoC.dtDataPagto = StrParaDate(GridC.TextMatrix(iLinha, iGrid_Pagamento_Col))
        objAportePagtoC.lNumIntDoc = gobjAporte.colPagtoCondicionados.Item(iLinha).lNumIntDoc
        objAportePagtoC.lNumIntDocAporte = gobjAporte.colPagtoCondicionados.Item(iLinha).lNumIntDocAporte
        objAportePagtoC.lNumIntDocDestino = gobjAporte.colPagtoCondicionados.Item(iLinha).lNumIntDocDestino
        objAportePagtoC.iTipoDocDestino = gobjAporte.colPagtoCondicionados.Item(iLinha).iTipoDocDestino
        objAportePagtoC.dValor = gobjAporte.colPagtoCondicionados.Item(iLinha).dValor
        objAportePagtoC.iSeq = iLinha

        If objAportePagtoC.dPercentual = 0 Then gError 190703
        If objAportePagtoC.iFormaPagto = 0 Then gError 190704
        If objAportePagtoC.iBase = 0 Then gError 190705
        If objAportePagtoC.dtDataPagto = DATA_NULA Then gError 190706
        
        objTRPAporte.colPagtoCondicionados.Add objAportePagtoC
    
    Next
    
    For iLinha = 1 To objGridSF.iLinhasExistentes
        
        iNumPagto = iNumPagto + 1
        
        Set objAportePagtoSF = New ClassTRPAportePagtoFat
    
        objAportePagtoSF.dPercentual = PercentParaDbl(GridSF.TextMatrix(iLinha, iGrid_Percentual_Col))
        objAportePagtoSF.dSaldo = StrParaDbl(GridSF.TextMatrix(iLinha, iGrid_Saldo_Col))
        objAportePagtoSF.dValor = StrParaDbl(GridSF.TextMatrix(iLinha, iGrid_ValorSF_Col))
        objAportePagtoSF.dtValidadeDe = StrParaDate(GridSF.TextMatrix(iLinha, iGrid_ValidadeDe_Col))
        objAportePagtoSF.dtValidadeAte = StrParaDate(GridSF.TextMatrix(iLinha, iGrid_ValidadeAte_Col))
        
        objAportePagtoSF.lNumIntDoc = gobjAporte.colPagtoSobreFatura.Item(iLinha).lNumIntDoc
        objAportePagtoSF.lNumIntDocAporte = gobjAporte.colPagtoSobreFatura.Item(iLinha).lNumIntDocAporte
        objAportePagtoSF.iSeq = iLinha

        If objAportePagtoSF.dPercentual = 0 Then gError 190707
        If objAportePagtoSF.dValor = 0 Then gError 190708
        If objAportePagtoSF.dtValidadeDe = DATA_NULA Then gError 196479
        If objAportePagtoSF.dtValidadeAte = DATA_NULA Then gError 196480
        If objAportePagtoSF.dtValidadeAte < objAportePagtoSF.dtValidadeDe Then gError 196481

        objTRPAporte.colPagtoSobreFatura.Add objAportePagtoSF
    
    Next
    
    For iLinha = 1 To objGridSFC.iLinhasExistentes
        
        iNumPagto = iNumPagto + 1
        
        Set objAportePagtoSFC = New ClassTRPAportePagtoFatC
    
        objAportePagtoSFC.dPercentual = PercentParaDbl(GridSFC.TextMatrix(iLinha, iGrid_PercSFC_Col))
        objAportePagtoSFC.dValor = StrParaDbl(GridSFC.TextMatrix(iLinha, iGrid_ValorSFC_Col))
        objAportePagtoSFC.dtValidadeDe = StrParaDate(GridSFC.TextMatrix(iLinha, iGrid_ValidadeDeSFC_Col))
        objAportePagtoSFC.dtValidadeAte = StrParaDate(GridSFC.TextMatrix(iLinha, iGrid_ValidadeAteSFC_Col))
        
        objAportePagtoSFC.lNumIntDoc = gobjAporte.colPagtoSobreFaturaCond.Item(iLinha).lNumIntDoc
        objAportePagtoSFC.lNumIntDocAporte = gobjAporte.colPagtoSobreFaturaCond.Item(iLinha).lNumIntDocAporte
        objAportePagtoSFC.iSeq = iLinha

        If objAportePagtoSFC.dPercentual = 0 Then gError 190707
        If objAportePagtoSFC.dtValidadeDe = DATA_NULA Then gError 196482
        If objAportePagtoSFC.dtValidadeAte = DATA_NULA Then gError 196483
        If objAportePagtoSFC.dtValidadeAte < objAportePagtoSFC.dtValidadeDe Then gError 196484

        objTRPAporte.colPagtoSobreFaturaCond.Add objAportePagtoSFC
    
    Next
    
    If iNumPagto = 0 Then gError 190699

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 190574
        
        Case 190699 'ERRO_APORTE_SEM_PAGTOS
            Call Rotina_Erro(vbOKOnly, "ERRO_APORTE_SEM_PAGTOS", gErr)
            
        Case 190700, 190708 'ERRO_VALOR_NAO_PREENCHIDO_GRID
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO_GRID", gErr, iLinha)

        Case 190701, 190704 'ERRO_FORMAPAGTO_NAO_PREENCHIDA_GRID
            Call Rotina_Erro(vbOKOnly, "ERRO_FORMAPAGTO_NAO_PREENCHIDA_GRID", gErr, iLinha)

        Case 190702 'ERRO_DATAVENCIMENTO_NAO_PREENCHIDO_GRID
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_NAO_PREENCHIDO_GRID", gErr, iLinha)

        Case 190703, 190707 'ERRO_PERCENTUAL_NAO_PREENCHIDO_GRID
            Call Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_NAO_PREENCHIDO_GRID", gErr, iLinha)

        Case 190705 'ERRO_BASE_NAO_PREENCHIDO_GRID
            Call Rotina_Erro(vbOKOnly, "ERRO_BASE_NAO_PREENCHIDO_GRID", gErr, iLinha)
        
        Case 190706 'ERRO_GRID_DATA_NAO_PREENCHIDA
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_DATA_NAO_PREENCHIDA", gErr, iLinha)
            
        Case 196479
            Call Rotina_Erro(vbOKOnly, "ERRO_VALIDADEDE_APORTEPAGTOFAT_NAO_PREENCHIDA", gErr, iLinha)
            
        Case 196480
            Call Rotina_Erro(vbOKOnly, "ERRO_VALIDADEATE_APORTEPAGTOFAT_NAO_PREENCHIDA", gErr, iLinha)
            
        Case 196481
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVALIDADE_INICIAL_MAIOR", gErr, iLinha)
            
        Case 196482
            Call Rotina_Erro(vbOKOnly, "ERRO_VALIDADEDE_APORTEPAGTOFATC_NAO_PREENCHIDA", gErr, iLinha)
            
        Case 196483
            Call Rotina_Erro(vbOKOnly, "ERRO_VALIDADEATE_APORTEPAGTOFATC_NAO_PREENCHIDA", gErr, iLinha)
            
        Case 196484
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVALIDADE_INICIAL_MAIOR", gErr, iLinha)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190575)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TRPAportes"

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", StrParaLong(Codigo.Text), 0, "Codigo"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 190576

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190577)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objTRPAporte As New ClassTRPAportes

On Error GoTo Erro_Tela_Preenche

    objTRPAporte.lCodigo = colCampoValor.Item("Codigo").vValor

    If objTRPAporte.lCodigo <> 0 Then
    
        lErro = Traz_TRPAporte_Tela(objTRPAporte)
        If lErro <> SUCESSO Then gError 190578
        
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 190578

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190579)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTRPAporte As New ClassTRPAportes
Dim objTRPAporteBD As New ClassTRPAportes
Dim vbResult As VbMsgBoxResult

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Codigo.Text)) = 0 Then gError 190580
    If Len(Trim(DataEmissao.Text)) = 0 Then gError 190581
    If Codigo_Extrai(Tipo.Text) = 0 Then gError 190582
    If Len(Trim(Moeda.Text)) = 0 Then gError 190583
    If Len(Trim(Cliente.Text)) = 0 Then gError 190584
    If Codigo_Extrai(Filial.Text) = 0 Then gError 190585
    
    If StrParaDate(PrevDataDe.Text) <> DATA_NULA Or StrParaDate(PrevDataAte.Text) <> DATA_NULA Or StrParaDbl(Previsao.Text) > 0 Or objGridC.iLinhasExistentes > 0 Then
    
        If StrParaDate(PrevDataDe.Text) = DATA_NULA Then gError 190586
        If StrParaDate(PrevDataAte.Text) = DATA_NULA Then gError 190587
        If StrParaDbl(Previsao.Text) = 0 Then gError 190588
        
        If StrParaDate(PrevDataDe.Text) > StrParaDate(PrevDataAte.Text) Then gError 190589
    
    End If
    '#####################

    'Preenche o objTRPAporte
    lErro = Move_Tela_Memoria(objTRPAporte)
    If lErro <> SUCESSO Then gError 190590

'    lErro = Trata_Alteracao(objTRPAporte, objTRPAporte.lCodigo)
'    If lErro <> SUCESSO Then gError 190591

    objTRPAporteBD.lCodigo = objTRPAporte.lCodigo

    lErro = CF("TRPAportes_Le", objTRPAporteBD, False)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192370

    If lErro = SUCESSO Then
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_APORTE_JA_CADASTRADO")
        If vbResult = vbNo Then gError 192371
    End If
    
    'Grava o/a TRPAporte no Banco de Dados
    lErro = CF("TRPAportes_Grava", objTRPAporte)
    If lErro <> SUCESSO Then gError 190592
    
    Historico.Clear
    lErro = CF("Carrega_Combo_Historico", Historico, "TRPAportes", STRING_TRP_APORTE_HISTORICO)
    If lErro <> SUCESSO Then gError 190593

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 190580
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            'Codigo.SetFocus

        Case 190581
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDO", gErr)
            'DataEmissao.SetFocus

        Case 190582
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_PREENCHIDO", gErr)
            'Tipo.SetFocus

        Case 190583
            Call Rotina_Erro(vbOKOnly, "ERRO_SIMBOLOMOEDA_NAO_PREENCHIDO", gErr)
            'Moeda.SetFocus

        Case 190584 'Cliente
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            'Cliente.SetFocus

        Case 190585 'Filial
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
            'Filial.SetFocus
            
        Case 190586
            Call Rotina_Erro(vbOKOnly, "ERRO_PREENCHIMENTO_PREVISAO_INCOMPLETO", gErr)
            'PrevDataDe.SetFocus
        
        Case 190587
            Call Rotina_Erro(vbOKOnly, "ERRO_PREENCHIMENTO_PREVISAO_INCOMPLETO", gErr)
            'PrevDataAte.SetFocus
            
        Case 190588
            Call Rotina_Erro(vbOKOnly, "ERRO_PREENCHIMENTO_PREVISAO_INCOMPLETO", gErr)
            'Previsao.SetFocus
            
        Case 190589
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)
            'PrevDataDe.SetFocus

        Case 190590 To 190593, 192370, 192371

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190594)

    End Select

    Exit Function

End Function

Function Limpa_Tela_TRPAporte() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_TRPAporte

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    Call Grid_Limpa(objGridC)
    Call Grid_Limpa(objGridD)
    Call Grid_Limpa(objGridSF)
    Call Grid_Limpa(objGridSFC)
    
    Filial.Clear
    Historico.Text = ""
    Tipo.ListIndex = -1
    Moeda.ListIndex = -1
    
    Realizado.Caption = ""
    
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    
    SaldoInvest.Caption = ""
    
    Set gobjAporte = New ClassTRPAportes
    
    RazaoSocial.Caption = ""
    CodigoCliente.Caption = ""
    CNPJ.Caption = ""
    FilialEmpresa.Caption = ""
    
    iFilialCliAnt = 0

    iAlterado = 0

    Limpa_Tela_TRPAporte = SUCESSO

    Exit Function

Erro_Limpa_Tela_TRPAporte:

    Limpa_Tela_TRPAporte = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190595)

    End Select

    Exit Function

End Function

Function Traz_TRPAporte_Tela(objTRPAporte As ClassTRPAportes) As Long

Dim lErro As Long
Dim objAportePagtoD As ClassTRPAportePagtoDireto
Dim objAportePagtoC As ClassTRPAportePagtoCond
Dim objAportePagtoSF As ClassTRPAportePagtoFat
Dim objAportePagtoSFC As ClassTRPAportePagtoFatC
Dim iLinha As Integer
Dim bExisteDestino As Boolean
Dim lNumTitulo As Long
Dim sDoc As String

On Error GoTo Erro_Traz_TRPAporte_Tela

    Call Limpa_Tela_TRPAporte
    
    If objTRPAporte.lCodigo <> 0 Then
        Codigo.PromptInclude = False
        Codigo.Text = CStr(objTRPAporte.lCodigo)
        Codigo.PromptInclude = True
    End If

    'Lê o TRPAporte que está sendo Passado
    lErro = CF("TRPAportes_Le", objTRPAporte)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190596
    
    If lErro = SUCESSO Then
        
        Cliente.Text = CStr(objTRPAporte.lCliente)
        Call Cliente_Validate(bSGECancelDummy)

        If objTRPAporte.dtDataEmissao <> DATA_NULA Then
            DataEmissao.PromptInclude = False
            DataEmissao.Text = Format(objTRPAporte.dtDataEmissao, "dd/mm/yy")
            DataEmissao.PromptInclude = True
        End If

        Call Combo_Seleciona_ItemData(Moeda, objTRPAporte.iMoeda)
        Call Combo_Seleciona_ItemData(Tipo, objTRPAporte.iTipo)

        Observacao.Text = objTRPAporte.sObservacao
        Historico.Text = objTRPAporte.sHistorico

        If objTRPAporte.dtPrevDataDe <> DATA_NULA Then
            PrevDataDe.PromptInclude = False
            PrevDataDe.Text = Format(objTRPAporte.dtPrevDataDe, "dd/mm/yy")
            PrevDataDe.PromptInclude = True
        End If
        
        If objTRPAporte.dtPrevDataAte <> DATA_NULA Then
            PrevDataAte.PromptInclude = False
            PrevDataAte.Text = Format(objTRPAporte.dtPrevDataAte, "dd/mm/yy")
            PrevDataAte.PromptInclude = True
        End If
        
        Previsao.Text = Format(objTRPAporte.dPrevValor, "STANDARD")
        
        iLinha = 0
        For Each objAportePagtoD In objTRPAporte.colPagtoDireto
        
            iLinha = iLinha + 1
            
            GridD.TextMatrix(iLinha, iGrid_ValorD_Col) = Format(objAportePagtoD.dValor, "STANDARD")
            GridD.TextMatrix(iLinha, iGrid_Vencimento_Col) = Format(objAportePagtoD.dtVencimento, "dd/mm/yyyy")
            
            Call Combo_Seleciona_ItemData(Forma, objAportePagtoD.iFormaPagto)
            
            GridD.TextMatrix(iLinha, iGrid_Forma_Col) = Forma.Text
        
            lErro = CF("Verifica_Existencia_Doc_TRP", objAportePagtoD.lNumIntDocDestino, objAportePagtoD.iTipoDocDestino, bExisteDestino, lNumTitulo, sDoc)
            If lErro <> SUCESSO Then gError 190597
        
            GridD.TextMatrix(iLinha, iGrid_DocD_Col) = sDoc
        
        Next
        
        objGridD.iLinhasExistentes = objTRPAporte.colPagtoDireto.Count
        
        iLinha = 0
        For Each objAportePagtoC In objTRPAporte.colPagtoCondicionados
            iLinha = iLinha + 1
            
            If objAportePagtoC.dValor > 0 Then
                GridC.TextMatrix(iLinha, iGrid_ValorC_Col) = Format(objAportePagtoC.dValor, "STANDARD")
            End If
            
            GridC.TextMatrix(iLinha, iGrid_PercentualC_Col) = Format(objAportePagtoC.dPercentual, "Percent")
            GridC.TextMatrix(iLinha, iGrid_Pagamento_Col) = Format(objAportePagtoC.dtDataPagto, "dd/mm/yyyy")
        
            Call Combo_Seleciona_ItemData(Base, objAportePagtoC.iBase)
            
            GridC.TextMatrix(iLinha, iGrid_Base_Col) = Base.Text
            
            Call Combo_Seleciona_ItemData(Forma, objAportePagtoC.iFormaPagto)
            
            GridC.TextMatrix(iLinha, iGrid_FormaC_Col) = Forma.Text
            
            Select Case objAportePagtoC.iStatus
            
                Case STATUS_TRP_OCR_LIBERADO
                    GridC.TextMatrix(iLinha, iGrid_StatusC_Col) = STATUS_TRP_OCR_LIBERADO_TEXTO
                
                Case STATUS_TRP_OCR_BLOQUEADO
                    GridC.TextMatrix(iLinha, iGrid_StatusC_Col) = STATUS_TRP_OCR_BLOQUEADO_TEXTO
            
            End Select
            
            lErro = CF("Verifica_Existencia_Doc_TRP", objAportePagtoC.lNumIntDocDestino, objAportePagtoC.iTipoDocDestino, bExisteDestino, lNumTitulo, sDoc)
            If lErro <> SUCESSO Then gError 190598
        
            GridC.TextMatrix(iLinha, iGrid_DocC_Col) = sDoc
        
        Next
        
        objGridC.iLinhasExistentes = objTRPAporte.colPagtoCondicionados.Count
        
        iLinha = 0
        For Each objAportePagtoSF In objTRPAporte.colPagtoSobreFatura
            iLinha = iLinha + 1
            
            GridSF.TextMatrix(iLinha, iGrid_ValorSF_Col) = Format(objAportePagtoSF.dValor, "STANDARD")
            GridSF.TextMatrix(iLinha, iGrid_ValidadeDe_Col) = Format(objAportePagtoSF.dtValidadeDe, "dd/mm/yyyy")
            GridSF.TextMatrix(iLinha, iGrid_ValidadeAte_Col) = Format(objAportePagtoSF.dtValidadeAte, "dd/mm/yyyy")
            GridSF.TextMatrix(iLinha, iGrid_Percentual_Col) = Format(objAportePagtoSF.dPercentual, "Percent")
            GridSF.TextMatrix(iLinha, iGrid_Saldo_Col) = Format(objAportePagtoSF.dSaldo, "STANDARD")
        
        Next
        
        objGridSF.iLinhasExistentes = objTRPAporte.colPagtoSobreFatura.Count
        
        
        iLinha = 0
        For Each objAportePagtoSFC In objTRPAporte.colPagtoSobreFaturaCond
            iLinha = iLinha + 1
            
            GridSFC.TextMatrix(iLinha, iGrid_ValorSFC_Col) = Format(objAportePagtoSFC.dValor, "STANDARD")
            GridSFC.TextMatrix(iLinha, iGrid_ValidadeDeSFC_Col) = Format(objAportePagtoSFC.dtValidadeDe, "dd/mm/yyyy")
            GridSFC.TextMatrix(iLinha, iGrid_ValidadeAteSFC_Col) = Format(objAportePagtoSFC.dtValidadeAte, "dd/mm/yyyy")
            GridSFC.TextMatrix(iLinha, iGrid_PercSFC_Col) = Format(objAportePagtoSFC.dPercentual, "Percent")
        
        Next
        
        objGridSFC.iLinhasExistentes = objTRPAporte.colPagtoSobreFaturaCond.Count
        
    End If
    
    Set gobjAporte = objTRPAporte
    
    Call Atualiza_Realizado
    Call Calcula_SaldoInvest

    iAlterado = 0

    Traz_TRPAporte_Tela = SUCESSO

    Exit Function

Erro_Traz_TRPAporte_Tela:

    Traz_TRPAporte_Tela = gErr

    Select Case gErr

        Case 190596 To 190598

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190599)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 190600

    'Limpa Tela
    Call Limpa_Tela_TRPAporte

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 190600

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190601)

    End Select

    Exit Sub

End Sub

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190602)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 190603

    Call Limpa_Tela_TRPAporte

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 190603

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190604)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTRPAporte As New ClassTRPAportes
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Codigo.Text)) = 0 Then gError 190605
    '#####################

    objTRPAporte.lCodigo = StrParaLong(Codigo.Text)

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_TRPAPORTE", objTRPAporte.lCodigo)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("TRPAportes_Exclui", objTRPAporte)
        If lErro <> SUCESSO Then gError 190606

        'Limpa Tela
        Call Limpa_Tela_TRPAporte

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 190605
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 190606

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190607)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica se Codigo está preenchida
    If Len(Trim(Codigo.Text)) <> 0 Then

       'Critica a Codigo
       lErro = Long_Critica(Codigo.Text)
       If lErro <> SUCESSO Then gError 190607

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 190607

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190608)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub

Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataEmissao_DownClick

    DataEmissao.SetFocus

    If Len(DataEmissao.ClipText) > 0 Then

        sData = DataEmissao.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 190609

        DataEmissao.Text = sData

    End If

    Exit Sub

Erro_UpDownDataEmissao_DownClick:

    Select Case gErr

        Case 190609

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190610)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataEmissao_UpClick

    DataEmissao.SetFocus

    If Len(Trim(DataEmissao.ClipText)) > 0 Then

        sData = DataEmissao.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 190611

        DataEmissao.Text = sData

    End If

    Exit Sub

Erro_UpDownDataEmissao_UpClick:

    Select Case gErr

        Case 190611

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190612)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)
    
End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    If Len(Trim(DataEmissao.ClipText)) <> 0 Then

        lErro = Data_Critica(DataEmissao.Text)
        If lErro <> SUCESSO Then gError 190613

    End If

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True

    Select Case gErr

        Case 190613

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190614)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Moeda_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Moeda_Validate



    Exit Sub

Erro_Moeda_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190615)

    End Select

    Exit Sub

End Sub

Private Sub Moeda_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Tipo_Validate


    Exit Sub

Erro_Tipo_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190616)

    End Select

    Exit Sub

End Sub

Private Sub Tipo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownPrevDe_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownPrevDe_DownClick

    PrevDataDe.SetFocus

    If Len(PrevDataDe.ClipText) > 0 Then

        sData = PrevDataDe.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 190617

        PrevDataDe.Text = sData

    End If

    Call Atualiza_Realizado

    Exit Sub

Erro_UpDownPrevDe_DownClick:

    Select Case gErr

        Case 190617

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190618)

    End Select

    Exit Sub

End Sub

Private Sub UpDownPrevDe_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownPrevDe_UpClick

    PrevDataDe.SetFocus

    If Len(Trim(PrevDataDe.ClipText)) > 0 Then

        sData = PrevDataDe.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 190619

        PrevDataDe.Text = sData

    End If

    Call Atualiza_Realizado

    Exit Sub

Erro_UpDownPrevDe_UpClick:

    Select Case gErr

        Case 190619

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190620)

    End Select

    Exit Sub

End Sub

Private Sub UpDownPrevAte_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownPrevAte_DownClick

    PrevDataAte.SetFocus

    If Len(PrevDataAte.ClipText) > 0 Then

        sData = PrevDataAte.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 190621

        PrevDataAte.Text = sData

    End If

    Call Atualiza_Realizado

    Exit Sub

Erro_UpDownPrevAte_DownClick:

    Select Case gErr

        Case 190621

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190622)

    End Select

    Exit Sub

End Sub

Private Sub UpDownPrevAte_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownPrevAte_UpClick

    PrevDataAte.SetFocus

    If Len(Trim(PrevDataAte.ClipText)) > 0 Then

        sData = PrevDataAte.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 190623

        PrevDataAte.Text = sData

    End If

    Call Atualiza_Realizado

    Exit Sub

Erro_UpDownPrevAte_UpClick:

    Select Case gErr

        Case 190623

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190624)

    End Select

    Exit Sub

End Sub

Private Sub Observacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Observacao_Validate

    'Verifica se Observacao está preenchida
    If Len(Trim(Observacao.Text)) <> 0 Then

    End If

    Exit Sub

Erro_Observacao_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190625)

    End Select

    Exit Sub

End Sub

Private Sub Observacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Historico_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Historico_Validate

    'Verifica se Historico está preenchida
    If Len(Trim(Historico.Text)) <> 0 Then

       If Len(Historico.Text) > STRING_TRP_APORTE_HISTORICO Then gError 190626

    End If

    Exit Sub

Erro_Historico_Validate:

    Cancel = True

    Select Case gErr
    
        Case 190626
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_HISTORICO", gErr, STRING_TRP_APORTE_HISTORICO, Len(Historico.Text))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190627)

    End Select

    Exit Sub

End Sub

Private Sub Historico_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Historico_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTRPAporte As ClassTRPAportes

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objTRPAporte = obj1

    'Mostra os dados do TRPAporte na tela
    lErro = Traz_TRPAporte_Tela(objTRPAporte)
    If lErro <> SUCESSO Then gError 190628

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 190628

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190629)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objTRPAporte As New ClassTRPAportes
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then

        objTRPAporte.lCodigo = Codigo.Text

    End If

    Call Chama_Tela("TRPAportesLista", colSelecao, objTRPAporte, objEventoCodigo)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190630)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_SF(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid ItensRequisicoes

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_SF

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Percentual")
    objGridInt.colColuna.Add ("Validade De")
    objGridInt.colColuna.Add ("Validade Até")
    objGridInt.colColuna.Add ("Saldo")

    'campos de edição do grid
    objGridInt.colCampo.Add (ValorSF.Name)
    objGridInt.colCampo.Add (Percentual.Name)
    objGridInt.colCampo.Add (DataValidadeDe.Name)
    objGridInt.colCampo.Add (DataValidadeAte.Name)
    objGridInt.colCampo.Add (Saldo.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_ValorSF_Col = 1
    iGrid_Percentual_Col = 2
    iGrid_ValidadeDe_Col = 3
    iGrid_ValidadeAte_Col = 4
    iGrid_Saldo_Col = 5

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridSF

    'Linhas do grid
    objGridInt.objGrid.Rows = 20 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 9
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_SF = SUCESSO

    Exit Function

Erro_Inicializa_Grid_SF:

    Inicializa_Grid_SF = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 190631)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_SFC(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid ItensRequisicoes

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_SFC

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Percentual")
    objGridInt.colColuna.Add ("Validade De")
    objGridInt.colColuna.Add ("Validade Até")
    objGridInt.colColuna.Add ("Valor Investido")

    'campos de edição do grid
    objGridInt.colCampo.Add (PercSFC.Name)
    objGridInt.colCampo.Add (DataValidadeDeSFC.Name)
    objGridInt.colCampo.Add (DataValidadeAteSFC.Name)
    objGridInt.colCampo.Add (ValorSFC.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_PercSFC_Col = 1
    iGrid_ValidadeDeSFC_Col = 2
    iGrid_ValidadeAteSFC_Col = 3
    iGrid_ValorSFC_Col = 4

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridSFC

    'Linhas do grid
    objGridInt.objGrid.Rows = 20 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 9
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_SFC = SUCESSO

    Exit Function

Erro_Inicializa_Grid_SFC:

    Inicializa_Grid_SFC = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 190631)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_D(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid ItensRequisicoes

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_D

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Valor R$")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Forma de Pagto")
    objGridInt.colColuna.Add ("Doc.")

    'campos de edição do grid
    objGridInt.colCampo.Add (ValorD.Name)
    objGridInt.colCampo.Add (DataVencimento.Name)
    objGridInt.colCampo.Add (Forma.Name)
    objGridInt.colCampo.Add (DocD.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_ValorD_Col = 1
    iGrid_Vencimento_Col = 2
    iGrid_Forma_Col = 3
    iGrid_DocD_Col = 4

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridD

    'Linhas do grid
    objGridInt.objGrid.Rows = 20 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 9
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_D = SUCESSO

    Exit Function

Erro_Inicializa_Grid_D:

    Inicializa_Grid_D = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 190632)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_C(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid ItensRequisicoes

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_C

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Base")
    objGridInt.colColuna.Add ("%")
    objGridInt.colColuna.Add ("Pagto")
    objGridInt.colColuna.Add ("Forma Pagto")
    objGridInt.colColuna.Add ("Status")
    objGridInt.colColuna.Add ("Valor R$")
    objGridInt.colColuna.Add ("Doc.")

    'campos de edição do grid
    objGridInt.colCampo.Add (Base.Name)
    objGridInt.colCampo.Add (PercentualC.Name)
    objGridInt.colCampo.Add (DataPagto.Name)
    objGridInt.colCampo.Add (FormaC.Name)
    objGridInt.colCampo.Add (StatusC.Name)
    objGridInt.colCampo.Add (ValorC.Name)
    objGridInt.colCampo.Add (DocC.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Base_Col = 1
    iGrid_PercentualC_Col = 2
    iGrid_Pagamento_Col = 3
    iGrid_FormaC_Col = 4
    iGrid_StatusC_Col = 5
    iGrid_ValorC_Col = 6
    iGrid_DocC_Col = 7

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridC

    'Linhas do grid
    objGridInt.objGrid.Rows = 20 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 9

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_C = SUCESSO

    Exit Function

Erro_Inicializa_Grid_C:

    Inicializa_Grid_C = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 190633)

    End Select

    Exit Function

End Function

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long
Dim iLinha As Integer
Dim iFrameAnterior

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index = iFrameAtual Then Exit Sub

    If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

    'Torna Frame correspondente ao Tab selecionado visivel
    Frame1(TabStrip1.SelectedItem.Index).Visible = True
    'Torna Frame atual invisivel
    Frame1(iFrameAtual).Visible = False
    'Armazena novo valor de iFrameAtual
    iFrameAtual = TabStrip1.SelectedItem.Index
    
    Call Calcula_SaldoInvest

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190634)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objClienteTRP As New ClassClienteTRP
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Cliente_Validate

    'Verifica se o Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        'Busca o Cliente no BD
        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then gError 190635
           
        RazaoSocial.Caption = objCliente.sRazaoSocial
        CodigoCliente.Caption = CStr(objCliente.lCodigo)
           
        lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 190636

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)
        
        If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)
               
        lErro = CF("Cliente_Le_Customizado", objCliente, True)
        If lErro <> SUCESSO Then gError 198634
        
        If Not (objCliente.objInfoUsu Is Nothing) Then

            objFilialEmpresa.iCodFilial = objClienteTRP.iFilialEmpresa
            
            lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
            If lErro <> SUCESSO Then gError 198635
                
            FilialEmpresa.Caption = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome

        End If

    'Se não estiver preenchido
    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        'Limpa a Combo de Filiais
        Filial.Clear
        
        RazaoSocial.Caption = ""
        CodigoCliente.Caption = ""
        CNPJ.Caption = ""
        FilialEmpresa.Caption = ""

    End If
    
    Call Filial_Validate(bSGECancelDummy)

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 190635, 190636, 198634, 198635

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190637)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_GotFocus()
    Call MaskEdBox_TrataGotFocus(Cliente, iAlterado)
End Sub

Private Sub Cliente_Change()
    iAlterado = REGISTRO_ALTERADO
    Call Atualiza_Realizado
End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult
Dim objCliente As New ClassCliente

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida ou alterada
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text <> Filial.List(Filial.ListIndex) Then

        'Tenta selecionar na combo
        lErro = Combo_Seleciona(Filial, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 190637
    
        'Se não encontrou o CÓDIGO
        If lErro = 6730 Then
    
            'Verifica se o cliente foi digitado
            If Len(Trim(Cliente.Text)) = 0 Then gError 190638
    
            sCliente = Cliente.Text
            objFilialCliente.iCodFilial = iCodigo
    
            'Pesquisa se existe Filial com o código extraído
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
            If lErro <> SUCESSO And lErro <> 17660 Then gError 190639
    
            If lErro = 17660 Then
    
                'Lê o Cliente
                objCliente.sNomeReduzido = sCliente
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO And lErro <> 12348 Then gError 190640
    
                'Se encontrou o Cliente
                If lErro = SUCESSO Then
                    
                    objFilialCliente.lCodCliente = objCliente.lCodigo
    
                    gError 190641
                
                End If
                
            End If
            
            If iCodigo <> 0 Then
            
                'Coloca na tela a Filial lida
                Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome
            
            Else
                
                objCliente.lCodigo = 0
                objFilialCliente.iCodFilial = 0
                
            End If
            
        'Não encontrou a STRING
        ElseIf lErro = 6731 Then
            
            'trecho incluido por Leo em 17/04/02
            objCliente.sNomeReduzido = Cliente.Text
            
            'Lê o Cliente
            lErro = CF("Cliente_Le_NomeReduzido", objCliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 190642
            
            If lErro = SUCESSO Then gError 190643
            
        End If
    
    End If
    
    sCliente = Cliente.Text
    objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)

    'Pesquisa se existe Filial com o código extraído
    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
    If lErro <> SUCESSO And lErro <> 17660 Then gError 190639
    
    If objFilialCliente.sCgc <> "" Then
        CNPJ.Caption = objFilialCliente.sCgc
        Call Formata_CNPJ
    Else
        CNPJ.Caption = ""
    End If

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 190637, 190639

        Case 190638
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 190640, 190642 'tratado na rotina chamada

        Case 190641
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 190643
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190644)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche campo Cliente
    Cliente.Text = objCliente.sNomeReduzido

    'Executa o Validate
    Call Cliente_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Public Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objCliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub GridD_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridD, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridD, iAlterado)
    End If

End Sub

Private Sub GridD_GotFocus()
    Call Grid_Recebe_Foco(objGridD)
End Sub

Private Sub GridD_EnterCell()
    Call Grid_Entrada_Celula(objGridD, iAlterado)
End Sub

Private Sub GridD_LeaveCell()
    Call Saida_Celula(objGridD)
End Sub

Private Sub GridD_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridD, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridD, iAlterado)
    End If

End Sub

Private Sub GridD_RowColChange()
    Call Grid_RowColChange(objGridD)
End Sub

Private Sub GridD_Scroll()
    Call Grid_Scroll(objGridD)
End Sub

Private Sub GridD_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iItemAtual As Integer
Dim iLinhasExistentesAnt As Integer
Dim vbMsgRes As VbMsgBoxResult
    
On Error GoTo Erro_GridD_KeyDown

    'Guarda o número de linhas existentes e a linha atual
    iLinhasExistentesAnt = objGridD.iLinhasExistentes
    iItemAtual = GridD.Row
    
    lErro = Remove_Linha(objGridD, iItemAtual, KeyCode)
    If lErro <> SUCESSO Then gError 190645
    
    Call Grid_Trata_Tecla1(KeyCode, objGridD)

    Exit Sub

Erro_GridD_KeyDown:

    Select Case gErr

        Case 190645

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190646)

    End Select

    Exit Sub

End Sub

Private Sub GridD_LostFocus()
    Call Grid_Libera_Foco(objGridD)
End Sub

Private Sub GridC_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridC, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridC, iAlterado)
    End If

End Sub

Private Sub GridC_GotFocus()
    Call Grid_Recebe_Foco(objGridC)
End Sub

Private Sub GridC_EnterCell()
    Call Grid_Entrada_Celula(objGridC, iAlterado)
End Sub

Private Sub GridC_LeaveCell()
    Call Saida_Celula(objGridC)
End Sub

Private Sub GridC_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridC, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridC, iAlterado)
    End If

End Sub

Private Sub GridC_RowColChange()
    Call Grid_RowColChange(objGridC)
End Sub

Private Sub GridC_Scroll()
    Call Grid_Scroll(objGridC)
End Sub

Private Sub GridC_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iItemAtual As Integer
Dim iLinhasExistentesAnt As Integer
Dim vbMsgRes As VbMsgBoxResult
    
On Error GoTo Erro_GridC_KeyDown

    'Guarda o número de linhas existentes e a linha atual
    iLinhasExistentesAnt = objGridC.iLinhasExistentes
    iItemAtual = GridC.Row
    
    lErro = Remove_Linha(objGridC, iItemAtual, KeyCode)
    If lErro <> SUCESSO Then gError 190647
    
    Call Grid_Trata_Tecla1(KeyCode, objGridC)
    
    Exit Sub

Erro_GridC_KeyDown:

    Select Case gErr

        Case 190647

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190648)

    End Select

    Exit Sub

End Sub

Private Sub GridC_LostFocus()
    Call Grid_Libera_Foco(objGridC)
End Sub

Private Sub GridSF_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridSF, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridSF, iAlterado)
    End If

End Sub

Private Sub GridSF_GotFocus()
    Call Grid_Recebe_Foco(objGridSF)
End Sub

Private Sub GridSF_EnterCell()
    Call Grid_Entrada_Celula(objGridSF, iAlterado)
End Sub

Private Sub GridSF_LeaveCell()
    Call Saida_Celula(objGridSF)
End Sub

Private Sub GridSF_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridSF, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridSF, iAlterado)
    End If

End Sub

Private Sub GridSF_RowColChange()
    Call Grid_RowColChange(objGridSF)
End Sub

Private Sub GridSF_Scroll()
    Call Grid_Scroll(objGridSF)
End Sub

Private Sub GridSF_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iItemAtual As Integer
Dim iLinhasExistentesAnt As Integer
Dim vbMsgRes As VbMsgBoxResult
    
On Error GoTo Erro_GridSF_KeyDown

    'Guarda o número de linhas existentes e a linha atual
    iLinhasExistentesAnt = objGridSF.iLinhasExistentes
    iItemAtual = GridSF.Row
    
    lErro = Remove_Linha(objGridSF, iItemAtual, KeyCode)
    If lErro <> SUCESSO Then gError 190649
    
    Call Grid_Trata_Tecla1(KeyCode, objGridSF)

    Exit Sub

Erro_GridSF_KeyDown:

    Select Case gErr

        Case 190649

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190650)

    End Select

    Exit Sub

End Sub

Private Sub GridSF_LostFocus()
    Call Grid_Libera_Foco(objGridSF)
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then
    
        'Verifica qual é o grid
        If objGridInt.objGrid.Name = GridSF.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
            
                Case iGrid_ValidadeDe_Col
                
                    lErro = Saida_Celula_Data(objGridInt, DataValidadeDe)
                    If lErro <> SUCESSO Then gError 190651

                Case iGrid_ValidadeAte_Col
                
                    lErro = Saida_Celula_Data(objGridInt, DataValidadeAte)
                    If lErro <> SUCESSO Then gError 190652

                Case iGrid_ValorSF_Col
                
                    lErro = Saida_Celula_ValorSF(objGridInt)
                    If lErro <> SUCESSO Then gError 190653

                Case iGrid_Percentual_Col
                
                    lErro = Saida_Celula_Percentual(objGridInt, Percentual)
                    If lErro <> SUCESSO Then gError 190654

                Case iGrid_Saldo_Col
                
                    lErro = Saida_Celula_Valor(objGridInt, Saldo)
                    If lErro <> SUCESSO Then gError 190655

            End Select
            
        ElseIf objGridInt.objGrid.Name = GridSFC.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
            
                Case iGrid_ValidadeDeSFC_Col
                
                    lErro = Saida_Celula_Data(objGridInt, DataValidadeDeSFC)
                    If lErro <> SUCESSO Then gError 190651

                Case iGrid_ValidadeAteSFC_Col
                
                    lErro = Saida_Celula_Data(objGridInt, DataValidadeAteSFC)
                    If lErro <> SUCESSO Then gError 190652

                Case iGrid_ValorSFC_Col
                
                    lErro = Saida_Celula_Valor(objGridInt, ValorSFC)
                    If lErro <> SUCESSO Then gError 190653

                Case iGrid_PercSFC_Col
                
                    lErro = Saida_Celula_Percentual(objGridInt, PercSFC)
                    If lErro <> SUCESSO Then gError 190654

            End Select
            
        ElseIf objGridInt.objGrid.Name = GridD.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_ValorD_Col
                
                    lErro = Saida_Celula_Valor(objGridInt, ValorD)
                    If lErro <> SUCESSO Then gError 190656
                    
                Case iGrid_Vencimento_Col
                
                    lErro = Saida_Celula_Data(objGridInt, DataVencimento)
                    If lErro <> SUCESSO Then gError 190657

                Case iGrid_Forma_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, Forma)
                    If lErro <> SUCESSO Then gError 190658
            
            End Select
        
        ElseIf objGridInt.objGrid.Name = GridC.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_PercentualC_Col
                
                    lErro = Saida_Celula_Valor(objGridInt, PercentualC)
                    If lErro <> SUCESSO Then gError 190659
                    
                Case iGrid_Pagamento_Col
                
                    lErro = Saida_Celula_Data(objGridInt, DataPagto)
                    If lErro <> SUCESSO Then gError 190660

                Case iGrid_FormaC_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, FormaC)
                    If lErro <> SUCESSO Then gError 190661

                Case iGrid_Base_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, Base)
                    If lErro <> SUCESSO Then gError 190662

            End Select
                         
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 190663

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 190651 To 190662
            'erros tratatos nas rotinas chamadas
        
        Case 190663
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190664)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable
              
    Select Case objControl.Name
    
        Case DataValidadeAte.Name, DataValidadeDe.Name, Percentual.Name
            objControl.Enabled = True
            
        Case DataValidadeAteSFC.Name, DataValidadeDeSFC.Name, PercSFC.Name
            objControl.Enabled = True
            
        Case ValorSF.Name
                If Abs(StrParaDbl(GridSF.TextMatrix(iLinha, iGrid_ValorSF_Col)) - StrParaDbl(GridSF.TextMatrix(iLinha, iGrid_Saldo_Col))) > QTDE_ESTOQUE_DELTA Then
                    objControl.Enabled = False
                Else
                    objControl.Enabled = True
                End If
            
        Case ValorD.Name, DataVencimento.Name, Forma.Name
        
            If Len(Trim(GridD.TextMatrix(iLinha, iGrid_DocD_Col))) > 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case Base.Name, PercentualC.Name, DataPagto.Name, FormaC.Name
    
            If Len(Trim(GridC.TextMatrix(iLinha, iGrid_DocC_Col))) > 0 Then
                objControl.Enabled = False
            Else
                If gobjAporte.colPagtoCondicionados.Count >= iLinha Then
                    If gobjAporte.colPagtoCondicionados.Item(iLinha).iStatus <> STATUS_TRP_OCR_BLOQUEADO Then
                        objControl.Enabled = False
                    Else
                        objControl.Enabled = True
                    End If
                Else
                    objControl.Enabled = True
                End If
            End If
            
        Case StatusC.Name, DocC.Name, DocD.Name, ValorC.Name
            objControl.Enabled = False
     
        Case Else
            objControl.Enabled = False
            
    End Select
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 190665)

    End Select

    Exit Sub

End Sub

Function Carrega_Moeda() As Long

Dim lErro As Long
Dim objMoeda As ClassMoedas
Dim colMoedas As New Collection

On Error GoTo Erro_Carrega_Moeda
    
    lErro = CF("Moedas_Le_Todas", colMoedas)
    If lErro <> SUCESSO Then gError 190666
    
    'se não existem moedas cadastradas
    If colMoedas.Count = 0 Then gError 190667
    
    For Each objMoeda In colMoedas
    
        Moeda.AddItem objMoeda.iCodigo & SEPARADOR & objMoeda.sNome
        Moeda.ItemData(Moeda.NewIndex) = objMoeda.iCodigo
    
    Next

    Carrega_Moeda = SUCESSO
    
    Exit Function
    
Erro_Carrega_Moeda:

    Carrega_Moeda = gErr
    
    Select Case gErr
    
        Case 190666
        
        Case 190667
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDAS_NAO_CADASTRADAS", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190668)
    
    End Select

End Function

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click
    
    lErro = CF("Config_ObterAutomatico", "FATConfig", "NUM_PROX_TRPAPORTES", "TRPOcorrencias", "Codigo", lCodigo)
    If lErro <> SUCESSO Then gError 190669
    
    Codigo.PromptInclude = False
    Codigo.Text = CStr(lCodigo)
    Codigo.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr
        
        Case 190669

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190670)
    
    End Select

    Exit Sub
    
End Sub

Public Sub ValorSF_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorSF_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridSF)
End Sub

Public Sub ValorSF_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridSF)
End Sub

Public Sub ValorSF_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridSF.objControle = ValorSF
    lErro = Grid_Campo_Libera_Foco(objGridSF)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub DataValidadeDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataValidadeDe_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridSF)
End Sub

Public Sub DataValidadeDe_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridSF)
End Sub

Public Sub DataValidadeDe_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridSF.objControle = DataValidadeDe
    lErro = Grid_Campo_Libera_Foco(objGridSF)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub DataValidadeAte_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataValidadeAte_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridSF)
End Sub

Public Sub DataValidadeAte_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridSF)
End Sub

Public Sub DataValidadeAte_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridSF.objControle = DataValidadeAte
    lErro = Grid_Campo_Libera_Foco(objGridSF)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub Saldo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Saldo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridSF)
End Sub

Public Sub Saldo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridSF)
End Sub

Public Sub Saldo_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridSF.objControle = Saldo
    lErro = Grid_Campo_Libera_Foco(objGridSF)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub Percentual_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Percentual_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridSF)
End Sub

Public Sub Percentual_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridSF)
End Sub

Public Sub Percentual_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridSF.objControle = Percentual
    lErro = Grid_Campo_Libera_Foco(objGridSF)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub ValorSFC_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorSFC_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridSFC)
End Sub

Public Sub ValorSFC_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridSFC)
End Sub

Public Sub ValorSFC_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridSFC.objControle = ValorSFC
    lErro = Grid_Campo_Libera_Foco(objGridSFC)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub DataValidadeDeSFC_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataValidadeDeSFC_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridSFC)
End Sub

Public Sub DataValidadeDeSFC_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridSFC)
End Sub

Public Sub DataValidadeDeSFC_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridSFC.objControle = DataValidadeDeSFC
    lErro = Grid_Campo_Libera_Foco(objGridSFC)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub DataValidadeAteSFC_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataValidadeAteSFC_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridSFC)
End Sub

Public Sub DataValidadeAteSFC_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridSFC)
End Sub

Public Sub DataValidadeAteSFC_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridSFC.objControle = DataValidadeAteSFC
    lErro = Grid_Campo_Libera_Foco(objGridSFC)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub PercSFC_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercSFC_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridSFC)
End Sub

Public Sub PercSFC_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridSFC)
End Sub

Public Sub PercSFC_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridSFC.objControle = PercSFC
    lErro = Grid_Campo_Libera_Foco(objGridSFC)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub ValorD_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorD_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridD)
End Sub

Public Sub ValorD_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridD)
End Sub

Public Sub ValorD_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridD.objControle = ValorD
    lErro = Grid_Campo_Libera_Foco(objGridD)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub DataVencimento_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataVencimento_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridD)
End Sub

Public Sub DataVencimento_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridD)
End Sub

Public Sub DataVencimento_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridD.objControle = DataVencimento
    lErro = Grid_Campo_Libera_Foco(objGridD)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub Forma_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Forma_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridD)
End Sub

Public Sub Forma_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridD)
End Sub

Public Sub Forma_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridD.objControle = Forma
    lErro = Grid_Campo_Libera_Foco(objGridD)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub DocD_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DocD_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridD)
End Sub

Public Sub DocD_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridD)
End Sub

Public Sub DocD_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridD.objControle = DocD
    lErro = Grid_Campo_Libera_Foco(objGridD)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub ValorC_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorC_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridC)
End Sub

Public Sub ValorC_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridC)
End Sub

Public Sub ValorC_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridC.objControle = ValorC
    lErro = Grid_Campo_Libera_Foco(objGridC)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub Base_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Base_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridC)
End Sub

Public Sub Base_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridC)
End Sub

Public Sub Base_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridC.objControle = Base
    lErro = Grid_Campo_Libera_Foco(objGridC)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub FormaC_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FormaC_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridC)
End Sub

Public Sub FormaC_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridC)
End Sub

Public Sub FormaC_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridC.objControle = FormaC
    lErro = Grid_Campo_Libera_Foco(objGridC)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub DataPagto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataPagto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridC)
End Sub

Public Sub DataPagto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridC)
End Sub

Public Sub DataPagto_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridC.objControle = DataPagto
    lErro = Grid_Campo_Libera_Foco(objGridC)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub StatusC_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub StatusC_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridC)
End Sub

Public Sub StatusC_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridC)
End Sub

Public Sub StatusC_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridC.objControle = StatusC
    lErro = Grid_Campo_Libera_Foco(objGridC)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub DocC_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DocC_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridC)
End Sub

Public Sub DocC_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridC)
End Sub

Public Sub DocC_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridC.objControle = DocC
    lErro = Grid_Campo_Libera_Foco(objGridC)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub PercentualC_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercentualC_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridC)
End Sub

Public Sub PercentualC_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridC)
End Sub

Public Sub PercentualC_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridC.objControle = PercentualC
    lErro = Grid_Campo_Libera_Foco(objGridC)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Function Saida_Celula_Data(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'Faz a crítica da célula Data que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Data

    Set objGridInt.objControle = objControle

    If Len(Trim(objControle.ClipText)) > 0 Then
    
        'Critica a Data informada
        lErro = Data_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError 190671

        'verifica se precisa preencher o grid com uma nova linha
        lErro = Adiciona_Linha(objGridInt)
        If lErro <> SUCESSO Then gError 190672

        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 190673

    Saida_Celula_Data = SUCESSO

    Exit Function

Erro_Saida_Celula_Data:

    Saida_Celula_Data = gErr

    Select Case gErr

        Case 190671 To 190673
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190674)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Function Saida_Celula_Valor(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'Faz a crítica da célula Data que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = objControle

    If Len(Trim(objControle.Text)) > 0 Then
    
        'Critica o valor informado
        lErro = Valor_Positivo_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError 190675

        objControle.Text = Format(objControle.Text, "STANDARD")

        'verifica se precisa preencher o grid com uma nova linha
        lErro = Adiciona_Linha(objGridInt)
        If lErro <> SUCESSO Then gError 190676
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 190677

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr

        Case 190675 To 190677
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190678)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Function Saida_Celula_ValorSF(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Data que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ValorSF

    Set objGridInt.objControle = ValorSF

    If Len(Trim(ValorSF.Text)) > 0 Then
    
        'Critica o valor informado
        lErro = Valor_Positivo_Critica(ValorSF.Text)
        If lErro <> SUCESSO Then gError 190679

        ValorSF.Text = Format(ValorSF.Text, "STANDARD")

        'verifica se precisa preencher o grid com uma nova linha
        lErro = Adiciona_Linha(objGridInt)
        If lErro <> SUCESSO Then gError 190680
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 190681

    GridSF.TextMatrix(GridSF.Row, iGrid_Saldo_Col) = ValorSF.Text

    Saida_Celula_ValorSF = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorSF:

    Saida_Celula_ValorSF = gErr

    Select Case gErr

        Case 190679 To 190681
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190682)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Function Saida_Celula_Percentual(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'Faz a crítica da célula Data que está deixando de ser a corrente

Dim lErro As Long
Dim dPercent As Double

On Error GoTo Erro_Saida_Celula_Percentual

    Set objGridInt.objControle = objControle

    If Len(Trim(objControle.Text)) > 0 Then
    
        'Critica a porcentagem
        lErro = Porcentagem_Critica_Negativa(objControle.Text)
        If lErro <> SUCESSO Then gError 190683

        dPercent = StrParaDbl(objControle.Text)

        'se for igual a 100% -> erro
        If dPercent = 100 Then gError 190684

        objControle.Text = Format(dPercent, "Fixed")
        
        'verifica se precisa preencher o grid com uma nova linha
        lErro = Adiciona_Linha(objGridInt)
        If lErro <> SUCESSO Then gError 190685
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 190686

    Saida_Celula_Percentual = SUCESSO

    Exit Function

Erro_Saida_Celula_Percentual:

    Saida_Celula_Percentual = gErr

    Select Case gErr

        Case 190683, 190685, 190686
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 190684
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190687)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Padrao(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Padrao

    Set objGridInt.objControle = objControle
    
    If Len(Trim(objControle.Text)) > 0 Then
    
        lErro = Adiciona_Linha(objGridInt)
        If lErro <> SUCESSO Then gError 190688
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 190689

    Saida_Celula_Padrao = SUCESSO

    Exit Function

Erro_Saida_Celula_Padrao:

    Saida_Celula_Padrao = gErr

    Select Case gErr

        Case 190688, 190689
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190690)

    End Select

    Exit Function

End Function

Public Function Adiciona_Linha(ByVal objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objAportePagtoD As New ClassTRPAportePagtoDireto
Dim objAportePagtoC As New ClassTRPAportePagtoCond
Dim objAportePagtoSF As New ClassTRPAportePagtoFat
Dim objAportePagtoSFC As New ClassTRPAportePagtoFatC

On Error GoTo Erro_Adiciona_Linha
              
    'verifica se precisa preencher o grid com uma nova linha
    If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    
        Select Case objGridInt.objGrid.Name
        
            Case GridSF.Name
                gobjAporte.colPagtoSobreFatura.Add objAportePagtoSF
                
            Case GridD.Name
                gobjAporte.colPagtoDireto.Add objAportePagtoD
        
            Case GridC.Name
            
                GridC.TextMatrix(GridC.Row, iGrid_StatusC_Col) = STATUS_TRP_OCR_BLOQUEADO_TEXTO
                objAportePagtoC.iStatus = STATUS_TRP_OCR_BLOQUEADO
            
                gobjAporte.colPagtoCondicionados.Add objAportePagtoC
        
            Case GridSFC.Name
                gobjAporte.colPagtoSobreFaturaCond.Add objAportePagtoSFC
        
        End Select
        
    End If
    
    Adiciona_Linha = SUCESSO
        
    Exit Function

Erro_Adiciona_Linha:

    Adiciona_Linha = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 190691)

    End Select

    Exit Function

End Function

Public Function Remove_Linha(ByVal objGridInt As AdmGrid, ByVal iLinha As Integer, ByVal iKeyCode As Integer) As Long

Dim lErro As Long
Dim sDoc As String

On Error GoTo Erro_Remove_Linha

    If iKeyCode = vbKeyDelete Then
              
        Select Case objGridInt.objGrid.Name
        
            Case GridSF.Name
            
                'Se já teve alteração no saldo não pode excluir
                If Abs(StrParaDbl(GridSF.TextMatrix(iLinha, iGrid_ValorSF_Col)) - StrParaDbl(GridSF.TextMatrix(iLinha, iGrid_Saldo_Col))) > DELTA_VALORMONETARIO Then gError 190692
            
                gobjAporte.colPagtoSobreFatura.Remove (iLinha)
                
            Case GridD.Name
            
                sDoc = GridD.TextMatrix(iLinha, iGrid_DocD_Col)
                
                'Se já deu origem a outro documento não pode excluir
                If Codigo_Extrai(GridD.TextMatrix(iLinha, iGrid_Forma_Col)) <> 2 Then
                    If Len(Trim(sDoc)) > 0 Then gError 190693
                End If
            
                gobjAporte.colPagtoDireto.Remove (iLinha)
        
            Case GridC.Name
            
                sDoc = GridC.TextMatrix(iLinha, iGrid_DocC_Col)
                
                'Se já deu origem a outro documento não pode excluir
                If Codigo_Extrai(GridC.TextMatrix(iLinha, iGrid_FormaC_Col)) <> 2 Then
                    If Len(Trim(sDoc)) > 0 Then gError 190694
                End If
                
                'If gobjAporte.colPagtoCondicionados.iItem(iLinha).iStatus = STATUS_TRP_OCR_LIBERADO Then gError 190695
                
                gobjAporte.colPagtoCondicionados.Remove (iLinha)
                
            Case GridSFC.Name
            
                'Se já teve alteração no saldo não pode excluir
                If StrParaDbl(GridSFC.TextMatrix(iLinha, iGrid_ValorSFC_Col)) > DELTA_VALORMONETARIO Then gError 190692
            
                gobjAporte.colPagtoSobreFaturaCond.Remove (iLinha)
        
        End Select
        
    End If
    
    Remove_Linha = SUCESSO
        
    Exit Function

Erro_Remove_Linha:

    Remove_Linha = gErr

    Select Case gErr
    
        Case 190692
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PAGTO_SALDO_DIF_VALOR", gErr)
    
        Case 190693, 190694
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_EXISTE_DOC_DESTINO", gErr, sDoc)
        
        'Case 190695 'Não permitir excluir um pagto cond liberado ???
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 190696)

    End Select

    Exit Function

End Function

Private Sub Previsao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Previsao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Previsao_Validate

    'Veifica se Previsao está preenchida
    If Len(Trim(Previsao.Text)) <> 0 Then

       'Critica a Previsao
       lErro = Valor_Positivo_Critica(Previsao.Text)
       If lErro <> SUCESSO Then gError 190697
        
    End If
    

    Exit Sub

Erro_Previsao_Validate:

    Cancel = True

    Select Case gErr

        Case 190697

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190698)

    End Select

    Exit Sub
    
End Sub

Public Function Atualiza_Realizado() As Long

Dim lErro As Long
Dim dValor As Double
Dim objCliente As New ClassCliente
Dim dValorUSS As Double

On Error GoTo Erro_Atualiza_Realizado

    If sClienteAnt <> Cliente.Text Or StrParaDate(PrevDataDe.Text) <> dtDataDeAnt Or StrParaDate(PrevDataAte.Text) <> dtDataAteAnt Then

        sClienteAnt = Cliente.Text
        dtDataDeAnt = StrParaDate(PrevDataDe.Text)
        dtDataAteAnt = StrParaDate(PrevDataAte.Text)

        If Len(Trim(Cliente.Text)) > 0 And StrParaDate(PrevDataDe.Text) <> DATA_NULA And StrParaDate(PrevDataAte.Text) <> DATA_NULA Then
                  
            objCliente.sNomeReduzido = Cliente.Text
        
            'Lê o Cliente através do Nome Reduzido
            lErro = CF("Cliente_Le_NomeReduzido", objCliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 190713
        
            lErro = CF("Vouchers_Le_Periodo_Cliente", objCliente.lCodigo, StrParaDate(PrevDataDe.Text), StrParaDate(PrevDataAte.Text), dValor, dValorUSS)
            If lErro <> SUCESSO Then gError 190714
            
            Realizado.Caption = Format(dValorUSS, "STANDARD")
        
        Else
            Realizado.Caption = ""
        End If
        
    End If
    
    Atualiza_Realizado = SUCESSO
        
    Exit Function

Erro_Atualiza_Realizado:

    Atualiza_Realizado = gErr

    Select Case gErr
    
        Case 190713, 190714
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 190715)

    End Select

    Exit Function

End Function

Private Sub PrevDataDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(PrevDataDe, iAlterado)
    
End Sub

Private Sub PrevDataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PrevDataDe_Validate

    If Len(Trim(PrevDataDe.ClipText)) <> 0 Then

        lErro = Data_Critica(PrevDataDe.Text)
        If lErro <> SUCESSO Then gError 190716

    End If
    
    Call Atualiza_Realizado

    Exit Sub

Erro_PrevDataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 190716

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190717)

    End Select

    Exit Sub

End Sub

Private Sub PrevDataAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(PrevDataAte, iAlterado)
    
End Sub

Private Sub PrevDataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PrevDataAte_Validate

    If Len(Trim(PrevDataAte.ClipText)) <> 0 Then

        lErro = Data_Critica(PrevDataAte.Text)
        If lErro <> SUCESSO Then gError 190718

    End If
    
    Call Atualiza_Realizado

    Exit Sub

Erro_PrevDataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 190718

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190719)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Cliente Then Call LabelCliente_Click
        If Me.ActiveControl Is Codigo Then Call LabelCodigo_Click
    
    End If
    
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub
Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub
Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub
Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub
Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub
Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Private Sub GridD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If GridD.Row > 0 And Not (gobjAporte Is Nothing) Then
    
        If GridD.Row <= gobjAporte.colPagtoDireto.Count Then

            'Verifica se foi o botao direito do mouse que foi pressionado
            If Button = vbRightButton Then
        
                'Seta objTela como a Tela de Baixas a Receber
                Set PopUpMenuPagtoAporte.objTela = Me
                
                Set gobjAportePagtoDireto = gobjAporte.colPagtoDireto.Item(GridD.Row)
                giTipoAporte = FORMAPAGTO_TRP_APORTE_TIPOPAGTO_DIRETO
        
                'Chama o Menu PopUp
                PopUpMenuPagtoAporte.PopupMenu PopUpMenuPagtoAporte.mnuGrid, vbPopupMenuRightButton
        
                'Limpa o objTela
                Set PopUpMenuPagtoAporte.objTela = Nothing
        
            End If
            
        End If
    
    End If
    
End Sub

Private Sub GridC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If GridC.Row > 0 And Not (gobjAporte Is Nothing) Then
    
        If GridC.Row <= gobjAporte.colPagtoDireto.Count Then
    
            'Verifica se foi o botao direito do mouse que foi pressionado
            If Button = vbRightButton Then
        
                'Seta objTela como a Tela de Baixas a Receber
                Set PopUpMenuPagtoAporte.objTela = Me
                
                Set gobjAportePagtoCond = gobjAporte.colPagtoCondicionados.Item(GridC.Row)
                giTipoAporte = FORMAPAGTO_TRP_APORTE_TIPOPAGTO_COND
                
                'Chama o Menu PopUp
                PopUpMenuPagtoAporte.PopupMenu PopUpMenuPagtoAporte.mnuGrid, vbPopupMenuRightButton
        
                'Limpa o objTela
                Set PopUpMenuPagtoAporte.objTela = Nothing
        
            End If
        
        End If
        
    End If
    
End Sub

Public Function mnuTvwAbrirDestino_Click() As Long

Dim lErro As Long
Dim objObjeto As Object
Dim sTela As String
Dim lNumIntDocDestino As Long
Dim iTipoDocDestino As Integer
Dim bExisteDestino As Boolean
Dim lNumTitulo As Long
Dim sDoc As String

On Error GoTo Erro_mnuTvwAbrirDestino_Click

    If giTipoAporte = FORMAPAGTO_TRP_APORTE_TIPOPAGTO_DIRETO Then
        lNumIntDocDestino = gobjAportePagtoDireto.lNumIntDocDestino
        iTipoDocDestino = gobjAportePagtoDireto.iTipoDocDestino
    Else
        lNumIntDocDestino = gobjAportePagtoCond.lNumIntDocDestino
        iTipoDocDestino = gobjAportePagtoCond.iTipoDocDestino
    End If

    lErro = CF("Verifica_Existencia_Doc_TRP", lNumIntDocDestino, iTipoDocDestino, bExisteDestino, lNumTitulo, sDoc)
    If lErro <> SUCESSO Then gError 192377

    If Not bExisteDestino Then gError 192378
    
    Select Case iTipoDocDestino
    
        Case TRP_TIPO_DOC_DESTINO_CREDFORN
            sTela = TRP_TIPO_DOC_DESTINO_CREDFORN_TELA
            Set objObjeto = New ClassCreditoPagar
            
        Case TRP_TIPO_DOC_DESTINO_DEBCLI
            sTela = TRP_TIPO_DOC_DESTINO_DEBCLI_TELA
            Set objObjeto = New ClassDebitoRecCli
    
        Case TRP_TIPO_DOC_DESTINO_TITPAG
            sTela = TRP_TIPO_DOC_DESTINO_TITPAG_TELA
            Set objObjeto = New ClassTituloPagar
    
        Case TRP_TIPO_DOC_DESTINO_TITREC
            sTela = TRP_TIPO_DOC_DESTINO_TITREC_TELA
            Set objObjeto = New ClassTituloReceber
            
        Case TRP_TIPO_DOC_DESTINO_NFSPAG
            sTela = TRP_TIPO_DOC_DESTINO_NFSPAG_TELA
            Set objObjeto = New ClassNFsPag
    
    End Select
    
    objObjeto.lNumIntDoc = lNumIntDocDestino
    
    Call Chama_Tela(sTela, objObjeto)
    
    mnuTvwAbrirDestino_Click = SUCESSO
    
    Exit Function

Erro_mnuTvwAbrirDestino_Click:

    mnuTvwAbrirDestino_Click = gErr

    Select Case gErr
    
        Case 192377
        
        Case 192378
            Call Rotina_Erro(vbOKOnly, "ERRO_PAGTO_SEM_DOC_ASSOCIADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192376)

    End Select

    Exit Function
    
End Function

Private Sub GridSF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If GridSF.Row > 0 And Not (gobjAporte Is Nothing) Then
    
        If GridSF.Row <= gobjAporte.colPagtoSobreFatura.Count Then
    
            'Verifica se foi o botao direito do mouse que foi pressionado
            If Button = vbRightButton Then
        
                'Seta objTela como a Tela de Baixas a Receber
                Set PopUpMenuPagtoAporteSF.objTela = Me
                
                Set gobjAportePagtoSF = gobjAporte.colPagtoSobreFatura.Item(GridSF.Row)
                
                'Chama o Menu PopUp
                PopUpMenuPagtoAporteSF.PopupMenu PopUpMenuPagtoAporteSF.mnuGrid, vbPopupMenuRightButton
        
                'Limpa o objTela
                Set PopUpMenuPagtoAporteSF.objTela = Nothing
        
            End If
        
        End If
        
    End If
    
End Sub

Public Function mnuGridHistorico_Click(Optional ByVal bDetalhado As Boolean = False) As Long

Dim lErro As Long
Dim objVoucher As New ClassTRPVouchers
Dim colSelecao As New Collection
Dim sNomeBrowse As String

On Error GoTo Erro_mnuGridHistorico_Click

    If bDetalhado Then
        sNomeBrowse = "TRPAPortesPagtoFatHistLista"
    Else
        sNomeBrowse = "TRPAPortesPagtoFatHistResLista"
    End If

    If iFrameAtual = FRAME_PAGTOSSF Then
     
        colSelecao.Add gobjAportePagtoSF.lNumIntDoc
        colSelecao.Add TRP_TIPO_APORTE_SOBREFATURA
         
        Call Chama_Tela(sNomeBrowse, colSelecao, objVoucher, Nothing, "NumIntDocPagtoAporteFat = ? AND TipoPagtoAporte = ?")
    
    ElseIf iFrameAtual = FRAME_PAGTOSSFC Then
    
        colSelecao.Add gobjAportePagtoSFC.lNumIntDoc
        colSelecao.Add TRP_TIPO_APORTE_SOBREFATURA_COND
         
        Call Chama_Tela(sNomeBrowse, colSelecao, objVoucher, Nothing, "NumIntDocPagtoAporteFat = ? AND TipoPagtoAporte = ?")
    
    End If
    
    mnuGridHistorico_Click = SUCESSO
    
    Exit Function

Erro_mnuGridHistorico_Click:

    mnuGridHistorico_Click = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192880)

    End Select

    Exit Function
    
End Function

Private Sub BotaoHistSF_Click()

On Error GoTo Erro_BotaoHistSF_Click

    If GridSF.Row = 0 Then gError 192885
    
    If GridSF.Row > 0 And Not (gobjAporte Is Nothing) Then
    
        If GridSF.Row <= gobjAporte.colPagtoSobreFatura.Count Then

            Set gobjAportePagtoSF = gobjAporte.colPagtoSobreFatura.Item(GridSF.Row)
                
            Call mnuGridHistorico_Click
        
        End If
        
    End If

    Exit Sub

Erro_BotaoHistSF_Click:

    Select Case gErr
    
        Case 192885
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192886)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoHistSFD_Click()

On Error GoTo Erro_BotaoHistSFD_Click

    If GridSF.Row = 0 Then gError 192885
    
    If GridSF.Row > 0 And Not (gobjAporte Is Nothing) Then
    
        If GridSF.Row <= gobjAporte.colPagtoSobreFatura.Count Then

            Set gobjAportePagtoSF = gobjAporte.colPagtoSobreFatura.Item(GridSF.Row)
                
            Call mnuGridHistorico_Click(True)
        
        End If
        
    End If

    Exit Sub

Erro_BotaoHistSFD_Click:

    Select Case gErr
    
        Case 192885
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192886)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoConsultarDocumentoC_Click()

On Error GoTo Erro_BotaoConsultarDocumentoC_Click

    If GridC.Row = 0 Then gError 192883
    
    If GridC.Row > 0 And Not (gobjAporte Is Nothing) Then
    
        If GridC.Row <= gobjAporte.colPagtoCondicionados.Count Then
                
            Set gobjAportePagtoCond = gobjAporte.colPagtoCondicionados.Item(GridC.Row)
            giTipoAporte = FORMAPAGTO_TRP_APORTE_TIPOPAGTO_COND

            Call mnuTvwAbrirDestino_Click
            
        End If
    
    End If
    
    Exit Sub

Erro_BotaoConsultarDocumentoC_Click:

    Select Case gErr
    
        Case 192883
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192884)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoConsultarDocumentoD_Click()

On Error GoTo Erro_BotaoConsultarDocumentoD_Click

    If GridD.Row = 0 Then gError 192881

    If GridD.Row > 0 And Not (gobjAporte Is Nothing) Then
    
        If GridD.Row <= gobjAporte.colPagtoDireto.Count Then
                
            Set gobjAportePagtoDireto = gobjAporte.colPagtoDireto.Item(GridD.Row)
            giTipoAporte = FORMAPAGTO_TRP_APORTE_TIPOPAGTO_DIRETO

            Call mnuTvwAbrirDestino_Click
            
        End If
    
    End If
    
    Exit Sub

Erro_BotaoConsultarDocumentoD_Click:

    Select Case gErr
    
        Case 192881
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192882)

    End Select

    Exit Sub
    
End Sub

Private Sub GridSFC_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridSFC, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridSFC, iAlterado)
    End If

End Sub

Private Sub GridSFC_GotFocus()
    Call Grid_Recebe_Foco(objGridSFC)
End Sub

Private Sub GridSFC_EnterCell()
    Call Grid_Entrada_Celula(objGridSFC, iAlterado)
End Sub

Private Sub GridSFC_LeaveCell()
    Call Saida_Celula(objGridSFC)
End Sub

Private Sub GridSFC_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridSFC, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridSFC, iAlterado)
    End If

End Sub

Private Sub GridSFC_RowColChange()
    Call Grid_RowColChange(objGridSFC)
End Sub

Private Sub GridSFC_Scroll()
    Call Grid_Scroll(objGridSFC)
End Sub

Private Sub GridSFC_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iItemAtual As Integer
Dim iLinhasExistentesAnt As Integer
Dim vbMsgRes As VbMsgBoxResult
    
On Error GoTo Erro_GridSFC_KeyDown

    'Guarda o número de linhas existentes e a linha atual
    iLinhasExistentesAnt = objGridSFC.iLinhasExistentes
    iItemAtual = GridSFC.Row
    
    lErro = Remove_Linha(objGridSFC, iItemAtual, KeyCode)
    If lErro <> SUCESSO Then gError 190649
    
    Call Grid_Trata_Tecla1(KeyCode, objGridSFC)

    Exit Sub

Erro_GridSFC_KeyDown:

    Select Case gErr

        Case 190649

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190650)

    End Select

    Exit Sub

End Sub

Private Sub GridSFC_LostFocus()
    Call Grid_Libera_Foco(objGridSFC)
End Sub

Private Sub BotaoHistSFC_Click()

On Error GoTo Erro_BotaoHistSFC_Click

    If GridSFC.Row = 0 Then gError 192885
    
    If GridSFC.Row > 0 And Not (gobjAporte Is Nothing) Then
    
        If GridSFC.Row <= gobjAporte.colPagtoSobreFaturaCond.Count Then

            Set gobjAportePagtoSFC = gobjAporte.colPagtoSobreFaturaCond.Item(GridSFC.Row)
                
            Call mnuGridHistorico_Click
        
        End If
        
    End If

    Exit Sub

Erro_BotaoHistSFC_Click:

    Select Case gErr
    
        Case 192885
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192886)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoHistSFCD_Click()

On Error GoTo Erro_BotaoHistSFCD_Click

    If GridSFC.Row = 0 Then gError 192885
    
    If GridSFC.Row > 0 And Not (gobjAporte Is Nothing) Then
    
        If GridSFC.Row <= gobjAporte.colPagtoSobreFaturaCond.Count Then

            Set gobjAportePagtoSFC = gobjAporte.colPagtoSobreFaturaCond.Item(GridSFC.Row)
                
            Call mnuGridHistorico_Click(True)
        
        End If
        
    End If

    Exit Sub

Erro_BotaoHistSFCD_Click:

    Select Case gErr
    
        Case 192885
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192886)

    End Select

    Exit Sub
    
End Sub

Private Sub Calcula_SaldoInvest()

Dim lErro As Long
Dim dTotal As Double
Dim iLinha As Integer

On Error GoTo Erro_Calcula_SaldoInvest

    For iLinha = 1 To objGridC.iLinhasExistentes
        dTotal = dTotal + gobjAporte.colPagtoCondicionados.Item(iLinha).dValor
    Next
    For iLinha = 1 To objGridD.iLinhasExistentes
        dTotal = dTotal + StrParaDbl(GridD.TextMatrix(iLinha, iGrid_ValorD_Col))
    Next
    For iLinha = 1 To objGridSF.iLinhasExistentes
        dTotal = dTotal + StrParaDbl(GridSF.TextMatrix(iLinha, iGrid_ValorSF_Col))
    Next
    For iLinha = 1 To objGridSFC.iLinhasExistentes
        dTotal = dTotal + StrParaDbl(GridSFC.TextMatrix(iLinha, iGrid_ValorSFC_Col))
    Next
    
    SaldoInvest.Caption = Format(dTotal, "STANDARD")

    Exit Sub

Erro_Calcula_SaldoInvest:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194429)

    End Select

    Exit Sub

End Sub

Private Sub BotaoVerCred_Click(Index As Integer)

Dim colSelecao As New Collection
Dim objAporte As New ClassTRPAportes

On Error GoTo Erro_BotaoHistSFC_Click

    If GridD.Row = 0 Then gError 196780
    
    Set gobjAportePagtoDireto = gobjAporte.colPagtoDireto.Item(GridD.Row)
    
    If gobjAportePagtoDireto.lNumIntDocDestino <> 0 Then
    
        If gobjAportePagtoDireto.iFormaPagto <> FORMAPAGTO_TRP_OCR_CRED Then gError 196781
         
        colSelecao.Add gobjAportePagtoDireto.lNumIntDocDestino
         
        Call Chama_Tela("TRPCreditosUtilizadosLista", colSelecao, objAporte, Nothing, "NumIntDocCredito = ? ")

    End If

    Exit Sub

Erro_BotaoHistSFC_Click:

    Select Case gErr
    
        Case 196780
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 196781
            Call Rotina_Erro(vbOKOnly, "ERRO_FORMA_PAGTO_NAO_CREDITO", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196782)

    End Select

    Exit Sub
    
End Sub

Public Function Formata_CNPJ() As Long

Dim lErro As Long
Dim sFormato As String

On Error GoTo Erro_Formata_CNPJ
    
    If Len(Trim(CNPJ.Caption)) <> 0 Then
    
        Select Case Len(Trim(CNPJ.Caption))
    
            Case STRING_CPF
                
                'Critica Cpf
                lErro = Cpf_Critica(CNPJ.Caption)
                If lErro <> SUCESSO Then gError 131846
                
                'Formata e coloca na Tela
                sFormato = "000\.000\.000-00; ; ; "
                CNPJ.Caption = Format(CNPJ.Caption, sFormato)
    
            Case STRING_CGC 'CGC
                
                'Critica CGC
                lErro = Cgc_Critica(CNPJ.Caption)
                If lErro <> SUCESSO Then gError 131847
                
                'Formata e Coloca na Tela
                sFormato = "00\.000\.000\/0000-00; ; ; "
                CNPJ.Caption = Format(CNPJ.Caption, sFormato)
    
            Case Else
                    
                gError 131848
    
        End Select
    
    End If
    
    Formata_CNPJ = SUCESSO
    
    Exit Function

Erro_Formata_CNPJ:

    Formata_CNPJ = gErr

    Select Case gErr

        Case 131846 To 131847

        Case 131848
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179275)

    End Select

    Exit Function

End Function

Private Sub BotaoAbrirCli_Click()

Dim objCliente As New ClassCliente

On Error GoTo Erro_BotaoAbrirCli_Click

    objCliente.lCodigo = StrParaDbl(CodigoCliente.Caption)

    Call Chama_Tela("Clientes", objCliente)

    Exit Sub

Erro_BotaoAbrirCli_Click:

    Select Case gErr
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192881)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoOutrosAportes_Click()

Dim lErro As Long
Dim objTRPAporte As New ClassTRPAportes
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoOutrosAportes_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then

        objTRPAporte.lCodigo = Codigo.Text

    End If
    
    colSelecao.Add StrParaDbl(CodigoCliente.Caption)

    Call Chama_Tela("TRPAportesLista", colSelecao, objTRPAporte, objEventoCodigo, "Cliente = ?")

    Exit Sub

Erro_BotaoOutrosAportes_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190630)

    End Select

    Exit Sub

End Sub
