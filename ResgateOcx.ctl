VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl ResgateOcx 
   ClientHeight    =   5625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9480
   KeyPreview      =   -1  'True
   ScaleHeight     =   5625
   ScaleWidth      =   9480
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4545
      Index           =   1
      Left            =   210
      TabIndex        =   0
      Top             =   795
      Width           =   9030
      Begin VB.Frame Frame4 
         Caption         =   "Dados Financeiros"
         Height          =   2850
         Left            =   120
         TabIndex        =   43
         Top             =   1560
         Width           =   8835
         Begin VB.ComboBox ComboTipoMov 
            Height          =   315
            ItemData        =   "ResgateOcx.ctx":0000
            Left            =   765
            List            =   "ResgateOcx.ctx":000D
            TabIndex        =   106
            Text            =   "Combo1"
            Top             =   555
            Width           =   1935
         End
         Begin MSMask.MaskEdBox Descontos 
            Height          =   300
            Left            =   4605
            TabIndex        =   7
            Top             =   2295
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Irrf 
            Height          =   300
            Left            =   2310
            TabIndex        =   6
            Top             =   2310
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Rendimentos 
            Height          =   300
            Left            =   1980
            TabIndex        =   4
            Top             =   1380
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox ValorResgate 
            Height          =   300
            Left            =   5175
            TabIndex        =   5
            Top             =   1395
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox AplicacaoAdicional 
            Height          =   300
            Left            =   3450
            TabIndex        =   100
            Top             =   1395
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   150
            TabIndex        =   105
            Top             =   615
            Width           =   450
         End
         Begin VB.Label Label4 
            Caption         =   "Rendimentos"
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
            Left            =   2010
            TabIndex        =   102
            Top             =   1125
            Width           =   1140
         End
         Begin VB.Label Label2 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3225
            TabIndex        =   101
            Top             =   1425
            Width           =   150
         End
         Begin VB.Label Label33 
            Caption         =   "="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   6450
            TabIndex        =   51
            Top             =   2295
            Width           =   330
         End
         Begin VB.Label Label32 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4230
            TabIndex        =   52
            Top             =   2295
            Width           =   300
         End
         Begin VB.Label Label25 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2010
            TabIndex        =   53
            Top             =   2280
            Width           =   300
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Valor do Resgate"
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
            Left            =   255
            TabIndex        =   54
            Top             =   2040
            Width           =   1485
         End
         Begin VB.Label ValorResgateLabel 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   180
            TabIndex        =   55
            Top             =   2295
            Width           =   1695
         End
         Begin VB.Label Label31 
            Caption         =   "="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   6870
            TabIndex        =   56
            Top             =   1425
            Width           =   195
         End
         Begin VB.Label LabelSinalValorResgate 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5010
            TabIndex        =   57
            Top             =   1425
            Width           =   120
         End
         Begin VB.Label Label21 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1695
            TabIndex        =   58
            Top             =   1410
            Width           =   300
         End
         Begin VB.Label Label10 
            Caption         =   "Saldo Atual"
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
            Left            =   7290
            TabIndex        =   59
            Top             =   1155
            Width           =   1095
         End
         Begin VB.Label SaldoAtual 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7110
            TabIndex        =   60
            Top             =   1410
            Width           =   1455
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Anterior"
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
            Left            =   360
            TabIndex        =   61
            Top             =   1170
            Width           =   1215
         End
         Begin VB.Label SaldoAnterior 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   150
            TabIndex        =   62
            Top             =   1410
            Width           =   1485
         End
         Begin VB.Label LabelValorResgate 
            AutoSize        =   -1  'True
            Caption         =   "Valor do Resgate"
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
            Left            =   5220
            TabIndex        =   63
            Top             =   1125
            Width           =   1485
         End
         Begin VB.Label LabelAplicacao 
            Caption         =   "Aplicação"
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
            Left            =   3765
            TabIndex        =   64
            Top             =   1110
            Width           =   930
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Valor Creditado"
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
            Left            =   7020
            TabIndex        =   65
            Top             =   2055
            Width           =   1320
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Descontos"
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
            Left            =   4980
            TabIndex        =   66
            Top             =   2040
            Width           =   915
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "IRRF"
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
            Left            =   2910
            TabIndex        =   67
            Top             =   2040
            Width           =   450
         End
         Begin VB.Label ValorCreditadoLabel 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6840
            TabIndex        =   68
            Top             =   2250
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Dados Principais"
         Height          =   1125
         Left            =   90
         TabIndex        =   42
         Top             =   165
         Width           =   8835
         Begin VB.ComboBox CodResgate 
            Height          =   315
            ItemData        =   "ResgateOcx.ctx":003B
            Left            =   4365
            List            =   "ResgateOcx.ctx":003D
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   450
            Width           =   930
         End
         Begin MSMask.MaskEdBox CodAplicacao 
            Height          =   300
            Left            =   1290
            TabIndex        =   1
            Top             =   450
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   300
            Left            =   7815
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   435
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataResgate 
            Height          =   300
            Left            =   6645
            TabIndex        =   3
            Top             =   435
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelCodResgate 
            AutoSize        =   -1  'True
            Caption         =   "Movimento:"
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
            Left            =   3270
            TabIndex        =   69
            Top             =   495
            Width           =   990
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   6075
            TabIndex        =   70
            Top             =   465
            Width           =   480
         End
         Begin VB.Label LabelCodAplicacao 
            AutoSize        =   -1  'True
            Caption         =   "Aplicação:"
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
            Left            =   285
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   71
            Top             =   495
            Width           =   915
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   4545
      Index           =   2
      Left            =   210
      TabIndex        =   8
      Top             =   795
      Visible         =   0   'False
      Width           =   9030
      Begin VB.Frame Frame2 
         Caption         =   "Reaplicação do Saldo Atual"
         Height          =   1050
         Left            =   210
         TabIndex        =   46
         Top             =   2310
         Width           =   8625
         Begin MSMask.MaskEdBox TaxaPrevista 
            Height          =   300
            Left            =   4185
            TabIndex        =   14
            Top             =   495
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "##0.#0\"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorResgatePrevisto 
            Height          =   300
            Left            =   6645
            TabIndex        =   15
            Top             =   495
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   300
            Left            =   2520
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   480
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataResgatePrevista 
            Height          =   300
            Left            =   1470
            TabIndex        =   13
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label3 
            Caption         =   "Data Prevista:"
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
            Left            =   165
            TabIndex        =   92
            Top             =   495
            Width           =   1335
         End
         Begin VB.Label Label17 
            Caption         =   "Valor Previsto:"
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
            Left            =   5190
            TabIndex        =   93
            Top             =   540
            Width           =   1335
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Taxa (%):"
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
            Left            =   3285
            TabIndex        =   94
            Top             =   525
            Width           =   810
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Complemento"
         Height          =   1860
         Left            =   225
         TabIndex        =   49
         Top             =   300
         Width           =   8625
         Begin VB.ComboBox CodContaCorrente 
            Height          =   315
            Left            =   2430
            TabIndex        =   103
            Top             =   1020
            Width           =   1695
         End
         Begin VB.ComboBox Historico 
            Height          =   315
            Left            =   1200
            TabIndex        =   9
            Top             =   510
            Width           =   4350
         End
         Begin VB.ComboBox TipoMeioPagto 
            Height          =   315
            ItemData        =   "ResgateOcx.ctx":003F
            Left            =   2430
            List            =   "ResgateOcx.ctx":0041
            TabIndex        =   11
            Top             =   1440
            Width           =   1695
         End
         Begin MSMask.MaskEdBox NumRefExterna 
            Height          =   300
            Left            =   7080
            TabIndex        =   10
            Top             =   510
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Numero 
            Height          =   300
            Left            =   7095
            TabIndex        =   12
            Top             =   1125
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin VB.Label LabelCtaCorrente 
            AutoSize        =   -1  'True
            Caption         =   "Conta Corrente:"
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
            Left            =   1005
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   104
            Top             =   1050
            Width           =   1350
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   285
            TabIndex        =   95
            Top             =   570
            Width           =   825
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Doc. Externo:"
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
            Left            =   5805
            TabIndex        =   96
            Top             =   555
            Width           =   1185
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Forma de Recebimento:"
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
            Left            =   300
            TabIndex        =   97
            Top             =   1470
            Width           =   2025
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
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
            Left            =   6270
            TabIndex        =   98
            Top             =   1170
            Width           =   720
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4545
      Index           =   3
      Left            =   240
      TabIndex        =   16
      Top             =   810
      Visible         =   0   'False
      Width           =   9030
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4920
         TabIndex        =   99
         Tag             =   "1"
         Top             =   1560
         Width           =   870
      End
      Begin VB.CommandButton CTBBotaoModeloPadrao 
         Caption         =   "Modelo Padrão"
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
         Left            =   6270
         TabIndex        =   22
         Top             =   390
         Width           =   2700
      End
      Begin VB.CommandButton CTBBotaoLimparGrid 
         Caption         =   "Limpar Grid"
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
         Left            =   6270
         TabIndex        =   20
         Top             =   30
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6270
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   870
         Width           =   2700
      End
      Begin VB.CommandButton CTBBotaoImprimir 
         Caption         =   "Imprimir"
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
         Left            =   7710
         TabIndex        =   21
         Top             =   30
         Width           =   1245
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4680
         TabIndex        =   30
         Top             =   1920
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   32
         Top             =   2565
         Width           =   870
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   31
         Top             =   2175
         Width           =   1770
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2985
         Left            =   6345
         TabIndex        =   34
         Top             =   1500
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   210
         TabIndex        =   44
         Top             =   3435
         Width           =   5895
         Begin VB.Label CTBCclLabel 
            AutoSize        =   -1  'True
            Caption         =   "Centro de Custo:"
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
            Left            =   375
            TabIndex        =   72
            Top             =   675
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label CTBLabel7 
            AutoSize        =   -1  'True
            Caption         =   "Conta:"
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
            Left            =   1260
            TabIndex        =   73
            Top             =   345
            Width           =   570
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1935
            TabIndex        =   74
            Top             =   315
            Width           =   3720
         End
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1935
            TabIndex        =   75
            Top             =   645
            Visible         =   0   'False
            Width           =   3720
         End
      End
      Begin VB.CheckBox CTBLancAutomatico 
         Caption         =   "Recalcula Automaticamente"
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
         Left            =   3495
         TabIndex        =   25
         Top             =   915
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin MSMask.MaskEdBox CTBConta 
         Height          =   225
         Left            =   525
         TabIndex        =   26
         Top             =   1860
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBDebito 
         Height          =   225
         Left            =   3435
         TabIndex        =   29
         Top             =   1890
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
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
      Begin MSMask.MaskEdBox CTBCredito 
         Height          =   225
         Left            =   2295
         TabIndex        =   28
         Top             =   1830
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
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
      Begin MSMask.MaskEdBox CTBCcl 
         Height          =   225
         Left            =   1545
         TabIndex        =   27
         Top             =   1875
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
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
      Begin MSComCtl2.UpDown CTBUpDown 
         Height          =   300
         Left            =   1650
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   540
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil 
         Height          =   300
         Left            =   570
         TabIndex        =   19
         Top             =   540
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBLote 
         Height          =   300
         Left            =   5595
         TabIndex        =   18
         Top             =   165
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBDocumento 
         Height          =   300
         Left            =   3795
         TabIndex        =   17
         Top             =   120
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid CTBGridContabil 
         Height          =   1860
         Left            =   15
         TabIndex        =   33
         Top             =   1185
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   3281
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin MSComctlLib.TreeView CTBTvwCcls 
         Height          =   2985
         Left            =   6345
         TabIndex        =   35
         Top             =   1500
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5265
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView CTBTvwContas 
         Height          =   2985
         Left            =   6345
         TabIndex        =   36
         Top             =   1500
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5265
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.Label CTBLabel1 
         AutoSize        =   -1  'True
         Caption         =   "Modelo:"
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
         Left            =   6300
         TabIndex        =   23
         Top             =   660
         Width           =   690
      End
      Begin VB.Label CTBLabel21 
         Caption         =   "Origem:"
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
         Left            =   45
         TabIndex        =   76
         Top             =   180
         Width           =   720
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   840
         TabIndex        =   77
         Top             =   135
         Width           =   1530
      End
      Begin VB.Label CTBLabel14 
         Caption         =   "Período:"
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
         Left            =   4230
         TabIndex        =   78
         Top             =   630
         Width           =   735
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5025
         TabIndex        =   79
         Top             =   585
         Width           =   1185
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   80
         Top             =   570
         Width           =   1185
      End
      Begin VB.Label CTBLabel13 
         Caption         =   "Exercício:"
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
         Left            =   1995
         TabIndex        =   81
         Top             =   600
         Width           =   870
      End
      Begin VB.Label CTBLabel5 
         AutoSize        =   -1  'True
         Caption         =   "Lançamentos"
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
         Left            =   45
         TabIndex        =   82
         Top             =   960
         Width           =   1140
      End
      Begin VB.Label CTBLabelHistoricos 
         Caption         =   "Históricos"
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
         Left            =   6345
         TabIndex        =   83
         Top             =   1275
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label CTBLabelContas 
         Caption         =   "Plano de Contas"
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
         Left            =   6345
         TabIndex        =   84
         Top             =   1275
         Width           =   2340
      End
      Begin VB.Label CTBLabelCcl 
         Caption         =   "Centros de Custo / Lucro"
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
         Left            =   6345
         TabIndex        =   85
         Top             =   1275
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.Label CTBLabelTotais 
         Caption         =   "Totais:"
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
         Height          =   225
         Left            =   1800
         TabIndex        =   86
         Top             =   3060
         Width           =   615
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   87
         Top             =   3045
         Width           =   1155
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2475
         TabIndex        =   88
         Top             =   3045
         Width           =   1155
      End
      Begin VB.Label CTBLabel8 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   45
         TabIndex        =   89
         Top             =   570
         Width           =   480
      End
      Begin VB.Label CTBLabelDoc 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
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
         Left            =   2700
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   90
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label CTBLabelLote 
         AutoSize        =   -1  'True
         Caption         =   "Lote:"
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
         Left            =   5100
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   91
         Top             =   180
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7155
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ResgateOcx.ctx":0043
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ResgateOcx.ctx":019D
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "ResgateOcx.ctx":0327
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ResgateOcx.ctx":0859
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5040
      Left            =   120
      TabIndex        =   50
      Top             =   420
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   8890
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Complemento"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contabilização"
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
Attribute VB_Name = "ResgateOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Início da Contabilidade
Dim objGrid1 As AdmGrid
Dim objContabil As New ClassContabil

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1

Private Const APLICACAO1 As String = "Aplicacao_Codigo"
Private Const RESGATE1 As String = "Resgate_Codigo"
Private Const CONTACORRENTE1 As String = "Conta_Corrente"
Private Const HISTORICO1 As String = "Historico"
Private Const FORMA1 As String = "Tipo_Meio_Pagto"
Private Const VALORRESGATE1 As String = "Valor_Resgate"
Private Const VALORCREDITADO1 As String = "Valor_Creditado"
Private Const RENDIMENTOS1 As String = "Rendimentos"
Private Const IRRF1 As String = "IRRF"
Private Const DESCONTOS1 As String = "Descontos"
Private Const VALORPREVISTO1 As String = "Valor_Resg_Prev"
Private Const CTACONTACORRENTE As String = "Cta_Conta_Corrente"
Private Const CTATIPOAPLICACAO As String = "Cta_Tipo_Aplicacao"
Private Const CTARECEITAAPLICACAO As String = "Cta_Receita_Aplic"
Private Const TIPO_APLICACAO As String = "Tipo_Aplicacao"

Private WithEvents objEventoCodAplicacao As AdmEvento
Attribute objEventoCodAplicacao.VB_VarHelpID = -1
Private WithEvents objEventoCtaCorrente As AdmEvento
Attribute objEventoCtaCorrente.VB_VarHelpID = -1

Public iAlterado As Integer
Dim lCodAplicacao As Long
Dim iFrameAtual As Integer

Dim iRendimentos_Validate As Integer
Dim iValorResgate_Validate As Integer
Dim iIRRF_Validate As Integer
Dim iDescontos_Validate As Integer
Dim iAplicacaoAdicional_Validate As Integer

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Complemento = 2
Private Const TAB_Contabilizacao = 3

Private Sub BotaoExcluir_Click()

Dim lErro  As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objResgate As New ClassResgate

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se codigo da aplicacao foi informado
    If Len(Trim(CodAplicacao.Text)) = 0 Then Error 17599

    objResgate.lCodigoAplicacao = CLng(CodAplicacao.Text)

    'Verifica se codigo do resgate foi informado
    If Len(Trim(CodResgate.Text)) = 0 Then Error 17636

    objResgate.iSeqResgate = CInt(CodResgate.Text)
    objResgate.dSaldoAnterior = CDbl(SaldoAnterior.Caption)
    
    'Pede confirmacao da exclusao
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_RESGATE", objResgate.iSeqResgate, objResgate.lCodigoAplicacao)

    If vbMsgRes = vbYes Then

        'Chama a rotina de exclusao
        lErro = CF("Resgate_Exclui", objResgate, objContabil)
        If lErro <> SUCESSO Then Error 17600

        Call Limpa_Tela_Resgate2
                
        lCodAplicacao = 0

        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 17599
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_APLICACA0_NAO_PREENCHIDO", Err)
                   
        Case 17600
        
        Case 17636
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_RESGATE_NAO_PREENCHIDO", Err)
                   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174141)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama a rotina de gravacao
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 17542

    Call Limpa_Tela_Resgate

    CodResgate.Clear
    
    lCodAplicacao = 0
    
    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 17542

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174142)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Confirma o pedido de limpeza da tela
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 17633

    'Limpa a tela
    Call Limpa_Tela_Resgate
    
    CodResgate.Clear
    
    lCodAplicacao = 0
    
    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 17633

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174143)

    End Select

    Exit Sub

End Sub

Private Sub CodAplicacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodAplicacao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodAplicacao, iAlterado)

End Sub

Private Sub CodAplicacao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim lCodAplic As Long
Dim objResgate As New ClassResgate
Dim objAplicacao As New ClassAplicacao

On Error GoTo Erro_CodAplicacao_Validate

    'Verifica se CodAplicacao esta preenchido
    If Len(Trim(CodAplicacao.Text)) = 0 Then Exit Sub

    lErro = Long_Critica(CodAplicacao.Text)
    If lErro <> SUCESSO Then Error 17517

    lCodAplic = CLng(CodAplicacao.Text)

    If lCodAplicacao <> lCodAplic Then

        Call Limpa_Tela_Resgate
        
        CodResgate.Clear
        
        objResgate.lCodigoAplicacao = lCodAplic

        lErro = Carrega_Resgate(objResgate)
        If lErro <> SUCESSO Then Error 17518
        
        objAplicacao.lCodigo = lCodAplic
        lErro = CF("Aplicacao_Le", objAplicacao)
        If lErro <> SUCESSO And lErro <> 17241 Then Error 17519

        If lErro = 17241 Then Error 17520

        'Insere na combobox CodResgate
        CodResgate.AddItem CStr(objAplicacao.iProxSeqResgate)
        CodResgate.ItemData(CodResgate.NewIndex) = objAplicacao.iProxSeqResgate
                
        'Mostra o Saldo Anterior
        SaldoAnterior.Caption = Format(objAplicacao.dSaldoAplicado, "Standard")

        'Mostra o Saldo Atual
        SaldoAtual.Caption = Format(objAplicacao.dSaldoAplicado, "Standard")
        
        'Mostra o codigo da aplicacao
        CodAplicacao.PromptInclude = False
        CodAplicacao.Text = CStr(objAplicacao.lCodigo)
        CodAplicacao.PromptInclude = True
    
        lCodAplicacao = lCodAplic

    End If

    Exit Sub

Erro_CodAplicacao_Validate:

    Cancel = True


    Select Case Err

        Case 17517

        Case 17518, 17519

        Case 17520
            lErro = Rotina_Erro(vbOKOnly, "ERRO_APLICACAO_INEXISTENTE", Err, objAplicacao.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174144)

    End Select

    Exit Sub

End Sub

Private Sub CodContaCorrente_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodContaCorrente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_CodContaCorrente_Validate

    If Len(Trim(CodContaCorrente.Text)) = 0 Then Exit Sub

    'Verifica se esta preenchida com o item selecionado na ComboBox CodContacOrrente
    If CodContaCorrente.Text = CodContaCorrente.List(CodContaCorrente.ListIndex) Then Exit Sub

    lErro = Combo_Seleciona(CodContaCorrente, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 17528
    
    objContaCorrenteInt.iCodigo = iCodigo
    
    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
        If lErro <> SUCESSO And lErro <> 11807 Then Error 17529

        'Não encontrou a Conta Corrente no BD
        If lErro = 11807 Then Error 17530

        'Se alguma Filial tiver sido selecionada
        If giFilialEmpresa <> EMPRESA_TODA Then

            If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then Error 43535
        
        End If
        
        'Encontrou a Conta Corrente no BD, coloca no Text da Combo
        CodContaCorrente.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 17531

    Exit Sub

Erro_CodContaCorrente_Validate:

    Cancel = True


    Select Case Err

        Case 17528, 17529

        Case 17530
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CODCONTACORRENTE_INEXISTENTE", objContaCorrenteInt.iCodigo)
            
            If vbMsgRes = vbYes Then
                'Chama a tela de Contas Correntes
                Call Chama_Tela("CtaCorrenteInt", objContaCorrenteInt)
            Else
            End If

        Case 17531
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE1", Err, CodContaCorrente.Text)

        Case 43535
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL", Err, CodContaCorrente.Text, giFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174145)

    End Select

    Exit Sub

End Sub

Private Sub CodResgate_Click()

Dim lErro As Long
Dim lCodAplic As Long
Dim objResgate As New ClassResgate
Dim objAplicacao As New ClassAplicacao

On Error GoTo Erro_CodResgate_Click

    iAlterado = REGISTRO_ALTERADO

    'Verifica se CodAplicacao esta preenchido
    If Len(Trim(CodAplicacao.Text)) <> 0 Then
        
        lCodAplic = CLng(CodAplicacao.Text)
        
        Call Limpa_Tela_Resgate2

        objResgate.lCodigoAplicacao = lCodAplic
        objResgate.iSeqResgate = CInt(CodResgate.Text)
        
        'Lê o Resgate
        lErro = CF("Resgate_Le", objResgate)
        If lErro <> SUCESSO And lErro <> 17499 Then Error 17521

        If lErro = 17499 Then

            objAplicacao.lCodigo = objResgate.lCodigoAplicacao
            
            'Lê a Aplicação
            lErro = CF("Aplicacao_Le", objAplicacao)
            If lErro <> SUCESSO And lErro <> 17241 Then Error 17523
        
            If lErro = 17241 Then Error 17524
                        
            'Mostra o saldo anterior
            SaldoAnterior.Caption = Format(objAplicacao.dSaldoAplicado, "Standard")

            'Mostra o saldo atual
            SaldoAtual.Caption = Format(objAplicacao.dSaldoAplicado, "Standard")

            'Mostra o codigo da aplicacao
            CodAplicacao.PromptInclude = False
            CodAplicacao.Text = CStr(objAplicacao.lCodigo)
            CodAplicacao.PromptInclude = True
            
        Else
        
            If objResgate.iStatus = RESGATE_EXCLUIDO Then Error 55974
        
            lErro = Traz_Resgate_Tela(objResgate)
            If lErro <> SUCESSO Then Error 17525
        
        End If
        
    End If

    Exit Sub

Erro_CodResgate_Click:

    Select Case Err

        Case 17521, 17523, 17525

        Case 17524
            lErro = Rotina_Erro(vbOKOnly, "ERRO_APLICACAO_INEXISTENTE", Err, objAplicacao.lCodigo)

        Case 55974
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RESGATE_EXCLUIDO", Err, objResgate.iSeqResgate, objResgate.lCodigoAplicacao)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174146)

    End Select

    Exit Sub

End Sub

Private Sub CodResgate_Validate(Cancel As Boolean)

On Error GoTo Erro_CodResgate_Validate

    'Verifica preenchimento do sequencial
    If Len(Trim(CodResgate.Text)) > 0 Then

        'Verifica se o sequencial é numérico
        If Not IsNumeric(CodResgate.Text) Then Error 55962

        'Verifica se codigo é menor que um
        If CInt(CodResgate.Text) < 1 Then Error 55963

    End If

    Exit Sub

Erro_CodResgate_Validate:

    Cancel = True

    Select Case Err

        Case 55962, 55963
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INVALIDO1", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174147)

    End Select

    Exit Sub

End Sub

Private Sub ComboTipoMov_Click()

    Select Case ComboTipoMov.ListIndex
    
        Case 0 'resgate
            LabelValorResgate.Enabled = True
            ValorResgate.Enabled = True
            AplicacaoAdicional.Text = ""
            AplicacaoAdicional.Enabled = False
            LabelAplicacao.Enabled = False
            
        Case 1 'aplicacao adicional
            AplicacaoAdicional.Enabled = True
            LabelAplicacao.Enabled = True
            LabelValorResgate.Enabled = False
            ValorResgate.Text = ""
            ValorResgate.Enabled = False
        
        Case 2 'rendimento
            AplicacaoAdicional.Text = ""
            AplicacaoAdicional.Enabled = False
            ValorResgate.Text = ""
            ValorResgate.Enabled = False
            LabelAplicacao.Enabled = False
            LabelValorResgate.Enabled = False
        
    End Select
    
End Sub

Private Sub DataResgate_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataResgate_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataResgate, iAlterado)

End Sub

Private Sub DataResgate_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataResgate_Validate

    'Verifica se a data de resgate está preenchida
    If Len(DataResgate.ClipText) <> 0 Then

        'Verifica se a data final é válida
        lErro = Data_Critica(DataResgate.Text)
        If lErro <> SUCESSO Then Error 17527

    End If

    Exit Sub

Erro_DataResgate_Validate:

    Cancel = True


    Select Case Err

        Case 17527

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174148)

    End Select

    Exit Sub

End Sub

Private Sub DataResgatePrevista_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataResgatePrevista_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataResgatePrevista, iAlterado)

End Sub

Private Sub DataResgatePrevista_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataResgatePrevista_Validate

    'Verifica se a data de resgate prevista está preenchida
    If Len(DataResgatePrevista.ClipText) <> 0 Then

        'Verifica se a data final é válida
        lErro = Data_Critica(DataResgatePrevista.Text)
        If lErro <> SUCESSO Then Error 17526

    End If

    Exit Sub

Erro_DataResgatePrevista_Validate:

    Cancel = True


    Select Case Err

        Case 17526

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174149)

    End Select

    Exit Sub

End Sub

Private Sub Descontos_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Descontos_GotFocus()
    iDescontos_Validate = 0
End Sub

Private Sub Descontos_Validate(Cancel As Boolean)

Dim curTeste As Currency
Dim lErro As Long
Dim dValResg As Double
Dim dValCred As Double
Dim dDesc As Double
Dim dIrrf As Double

On Error GoTo Erro_Descontos_Validate

    If iValorResgate_Validate = 1 Or iRendimentos_Validate = 1 Or iIRRF_Validate = 1 Or iAplicacaoAdicional_Validate = 1 Then Exit Sub

    'Verifica se descontos foi preenchido
    If Len(Trim(Descontos.Text)) > 0 Then
        lErro = Valor_NaoNegativo_Critica(Descontos.Text)
        If lErro <> SUCESSO Then Error 17596

        dDesc = CDbl(Descontos.Text)

        Descontos.Text = Format(Descontos.Text, "Fixed")
        
    End If
    

    'Verifica se valor resgate label esta preenchido
    If Len(ValorResgateLabel.Caption) > 0 Then dValResg = CDbl(ValorResgateLabel.Caption)

    'Verifica se irrf esta preenchido
    If Len(Irrf.Text) > 0 Then dIrrf = CDbl(Irrf.Text)

    'Calcula o valor creditado
    dValCred = dValResg - dIrrf - dDesc

    'Mostra na tela o valor creditado
    ValorCreditadoLabel.Caption = Format(dValCred, "Standard")

    Exit Sub

Erro_Descontos_Validate:

    Cancel = True


    Select Case Err

        Case 17596
            iDescontos_Validate = 1

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174150)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
   
End Sub

Public Sub Form_UnLoad(Cancel As Integer)

    Set objEventoCodAplicacao = Nothing
    Set objEventoCtaCorrente = Nothing
    
    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing
    
    Set objGrid1 = Nothing
    Set objContabil = Nothing

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)
    
End Sub

Private Sub Historico_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Historico_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iTamanho As Integer
Dim iCodigo As Integer
Dim iIndice As Integer
Dim sDescricao As Long
Dim objHistMovCta As New ClassHistMovCta

On Error GoTo Erro_Historico_Validate

    'Verifica o tamanho do texto em historico
    iTamanho = Len(Trim(Historico.Text))

    If iTamanho = 0 Then Exit Sub

    'Verifica se é maior que o tamanho maximo
    If iTamanho > 50 Then Error 17538

    'Verifica se o que foi digitado é numerico
    If IsNumeric(Trim(Historico.Text)) Then
        
        'Verifica se é inteiro
        lErro = Valor_Inteiro_Critica(Trim(Historico.Text))
        If lErro <> SUCESSO Then Error 40738
        
        'peenche o objeto
        objHistMovCta.iCodigo = CInt(Trim(Historico.Text))
                
        'verifica se existe hitorico relacionado com codigo passado
        lErro = CF("HistMovCta_Le", objHistMovCta)
        If lErro <> SUCESSO And lErro <> 15011 Then Error 40739
        
        'se nao existir -----> Erro
        If lErro = 15011 Then Error 40740
        
        Historico.Text = objHistMovCta.sDescricao
        
    
    End If

    Exit Sub

Erro_Historico_Validate:

    Cancel = True


    Select Case Err

        Case 17538
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_HISTORICOMOVCONTA", Err)
        
        Case 40738
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INTEIRO", Err, Historico.Text)
                
        Case 40739
        
        Case 40740
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTMOVCTA_NAO_CADASTRADO", Err, objHistMovCta.iCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174151)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Form_Load

    If giTipoVersao = VERSAO_LIGHT Then
        Opcao.Tabs.Remove (TAB_Contabilizacao)
    End If
    
    iFrameAtual = 1

    Set objEventoCodAplicacao = New AdmEvento
    Set objEventoCtaCorrente = New AdmEvento

    iAlterado = 0

    lCodAplicacao = 0

    ComboTipoMov.ListIndex = 0 'colocar "resgate" como default.

    'Lê as contas correntes  com codigo e o nome reduzido existentes no BD e carrega na ComboBox
    lErro = Carrega_CodContaCorrente()
    If lErro <> SUCESSO Then Error 17491

    'Lê os tipos de pagamentos ativos existentes no BD
    lErro = Carrega_TipoMeioPagto()
    If lErro <> SUCESSO Then Error 17492

    'Lê os historicos com codigo e o nome existentes no BD e carrega na ComboBox
    lErro = Carrega_Historico()
    If lErro <> SUCESSO Then Error 17493
    
    'inicializacao da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_TESOURARIA)
    If lErro <> SUCESSO Then Error 39562
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 17491, 17492, 17493, 39562

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174152)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 39564

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 39565

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 39564
        
        Case 39565
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Sub Irrf_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Irrf_GotFocus()
    iIRRF_Validate = 0
End Sub

Private Sub Irrf_Validate(Cancel As Boolean)

Dim curTeste As Currency
Dim lErro As Long
Dim dValResg As Double
Dim dValCred As Double
Dim dDesc As Double
Dim dIrrf As Double

On Error GoTo Erro_Irrf_Validate

    If iValorResgate_Validate = 1 Or iRendimentos_Validate = 1 Or iIRRF_Validate = 1 Or iAplicacaoAdicional_Validate = 1 Then Exit Sub

    'Verifica se irrf foi preenchido
    If Len(Trim(Irrf.Text)) > 0 Then
        lErro = Valor_NaoNegativo_Critica(Irrf.Text)
        If lErro <> SUCESSO Then Error 17595

        dIrrf = CDbl(Irrf.Text)

        Irrf.Text = Format(Irrf.Text, "Fixed")
    
    End If
    

    'Verifica se valor resgate label esta preenchido
    If Len(ValorResgateLabel.Caption) > 0 Then dValResg = CDbl(ValorResgateLabel.Caption)

    'Verifica se descontos esta preenchido
    If Len(Descontos.Text) > 0 Then dDesc = CDbl(Descontos.Text)

    'Calcula o valor creditado
    dValCred = dValResg - dIrrf - dDesc

    'Mostra na tela o valor creditado
    ValorCreditadoLabel.Caption = Format(dValCred, "Standard")
    
    Exit Sub

Erro_Irrf_Validate:

    Cancel = True


    Select Case Err

        Case 17595
            iIRRF_Validate = 1

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174153)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodAplicacao_Click()

Dim objAplicacao As New ClassAplicacao
Dim colSelecao As Collection

    If Len(Trim(CodAplicacao.Text)) = 0 Then
        objAplicacao.lCodigo = 0
    Else
        objAplicacao.lCodigo = CLng(CodAplicacao.Text)
        
    End If
    
    Call Chama_Tela("AplicacaoLista", colSelecao, objAplicacao, objEventoCodAplicacao)

End Sub

Private Sub LabelCtaCorrente_Click()

Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim colSelecao As Collection

    If Len(CodContaCorrente.Text) = 0 Then
        objContaCorrenteInt.iCodigo = 0
    Else
        If CodContaCorrente.ListIndex <> -1 Then objContaCorrenteInt.iCodigo = CodContaCorrente.ItemData(CodContaCorrente.ListIndex)
    End If

    Call Chama_Tela("CtaCorrenteLista", colSelecao, objContaCorrenteInt, objEventoCtaCorrente)
    
End Sub

Private Sub Numero_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Numero_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Numero, iAlterado)

End Sub

Private Sub NumRefExterna_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoCodAplicacao_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objAplicacao As ClassAplicacao
Dim objResgate As New ClassResgate

On Error GoTo Erro_objEventoCodAplicacao_evSelecao

    Set objAplicacao = obj1
    
    CodResgate.Clear
    
    objResgate.lCodigoAplicacao = objAplicacao.lCodigo

    lErro = Carrega_Resgate(objResgate)
    If lErro <> SUCESSO Then Error 40714
    
    lErro = CF("Aplicacao_Le", objAplicacao)
    If lErro <> SUCESSO And lErro <> 17241 Then Error 40715
    
    If lErro = 17241 Then Error 40716

    'Insere na combobox CodResgate
    CodResgate.AddItem CStr(objAplicacao.iProxSeqResgate)
    CodResgate.ItemData(CodResgate.NewIndex) = objAplicacao.iProxSeqResgate
            
    'Mostra o Saldo Anterior
    SaldoAnterior.Caption = Format(objAplicacao.dSaldoAplicado, "Standard")

    'Mostra o Saldo Atual
    SaldoAtual.Caption = Format(objAplicacao.dSaldoAplicado, "Standard")
    
    'Mostra o codigo da aplicacao
    CodAplicacao.PromptInclude = False
    CodAplicacao.Text = CStr(objAplicacao.lCodigo)
    CodAplicacao.PromptInclude = True
      
    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoCodAplicacao_evSelecao:

    Select Case Err

        Case 40714, 40715

        Case 40716
            lErro = Rotina_Erro(vbOKOnly, "ERRO_APLICACAO_INEXISTENTE", Err, objAplicacao.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174154)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCtaCorrente_evSelecao(obj1 As Object)

Dim objContaCorrenteInt As ClassContasCorrentesInternas

    Set objContaCorrenteInt = obj1

    CodContaCorrente.Text = CStr(objContaCorrenteInt.iCodigo)

    Call CodContaCorrente_Validate(bSGECancelDummy)

    iAlterado = 0

    Me.Show

End Sub

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
        
        'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        If Opcao.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then Call objContabil.Contabil_Carga_Modelo_Padrao
        
        Select Case iFrameAtual
        
            Case TAB_Identificacao
                Parent.HelpContextID = IDH_RESGATE_ID
                
            Case TAB_Complemento
                Parent.HelpContextID = IDH_RESGATE_COMPLEMENTO
                
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_RESGATE_CONTABILIZACAO
                        
        End Select
        
    End If

End Sub

Private Function Carrega_CodContaCorrente() As Long
'Carrega as contas correntes na combo de contas correntes

Dim lErro As Long
Dim colCodigoNomeConta As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Carrega_CodContaCorrente

    'Le o nome e o codigo de todas a contas correntes
    lErro = CF("ContasCorrentes_Bancarias_Le_CodigosNomesRed", colCodigoNomeConta)
    If lErro <> SUCESSO Then Error 17494

    For Each objCodigoNome In colCodigoNomeConta

        'Insere na combo de contas correntes
        CodContaCorrente.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        CodContaCorrente.ItemData(CodContaCorrente.NewIndex) = objCodigoNome.iCodigo

    Next

    Carrega_CodContaCorrente = SUCESSO

    Exit Function

Erro_Carrega_CodContaCorrente:

    Carrega_CodContaCorrente = Err

    Select Case Err

        Case 17494

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174155)

    End Select

    Exit Function

End Function

Private Function Carrega_TipoMeioPagto() As Long
'Carrega na Combo TipoMeioPagto os tipo de meio de pagamento ativos

Dim lErro As Long
Dim colTipoMeioPagto As New Collection
Dim objTipoMeioPagto As ClassTipoMeioPagto

On Error GoTo Erro_Carrega_TipoMeioPagto

    'Le todos os tipo de pagamento
    lErro = CF("TipoMeioPagto_Le_Todos", colTipoMeioPagto)
    If lErro <> SUCESSO Then Error 17495

    For Each objTipoMeioPagto In colTipoMeioPagto

        'Verifica se estao ativos
        If objTipoMeioPagto.iInativo = TIPOMEIOPAGTO_ATIVO Then

            'Insere na Combo de TipoMeioPagto
            TipoMeioPagto.AddItem CStr(objTipoMeioPagto.iTipo) & SEPARADOR & objTipoMeioPagto.sDescricao
            TipoMeioPagto.ItemData(TipoMeioPagto.NewIndex) = objTipoMeioPagto.iTipo

        End If

    Next

    Carrega_TipoMeioPagto = SUCESSO

    Exit Function

Erro_Carrega_TipoMeioPagto:

    Carrega_TipoMeioPagto = Err

    Select Case Err

        Case 17495

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174156)

    End Select

    Exit Function

End Function

Private Function Carrega_Historico() As Long
'Carrega a combo de historicos com os historicos da tabela "HistPadraMovConta"

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Carrega_Historico

    'Le o Codigo e a descricao de todos os historicos
    lErro = CF("Cod_Nomes_Le", "HistPadraoMovConta", "Codigo", "Descricao", STRING_NOME, colCodigoNome)
    If lErro <> SUCESSO Then Error 17496

    For Each objCodigoNome In colCodigoNome

        'Insere na Combo de historicos
        Historico.AddItem objCodigoNome.sNome
        Historico.ItemData(Historico.NewIndex) = objCodigoNome.iCodigo

    Next

    Carrega_Historico = SUCESSO

    Exit Function

Erro_Carrega_Historico:

    Carrega_Historico = Err

    Select Case Err

        Case 17496

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174157)

    End Select
    
    Exit Function

End Function

Private Function Carrega_Resgate(objResgate As ClassResgate) As Long
'Carrega os resgates relativos ao CodAplicacao na combo de CodResgate

Dim lErro As Long
Dim colCodigo As New Collection
Dim objCodigo As ClassResgate

On Error GoTo Erro_Carrega_Resgate

    'Le todos os resgates da aplicacao em questao
    lErro = CF("Resgate_Le_Todos", objResgate, colCodigo)
    If lErro <> SUCESSO Then Error 17512
    
    CodResgate.Clear
    
    'Preenche a ComboBox CodResgate com os objetos da colecao colCodigo
     For Each objCodigo In colCodigo

        'Insere na combo de CodResgate
        CodResgate.AddItem CStr(objCodigo.iSeqResgate)
        CodResgate.ItemData(CodResgate.NewIndex) = objCodigo.iSeqResgate

    Next

    If colCodigo.Count = 0 Then
        objResgate.iSeqResgate = 0
    End If

    Carrega_Resgate = SUCESSO

    Exit Function

Erro_Carrega_Resgate:

    Carrega_Resgate = Err

    Select Case Err

        Case 17512

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174158)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objResgate As ClassResgate) As Long

Dim lErro As Long
Dim lCodigoAplicacao As Long

On Error GoTo Erro_Trata_Parametros

    'Se há uma aplicacao selecionada, exibir seus dados
    If Not (objResgate Is Nothing) Then

        'Verifica se tipo de aplicacao existe
        lErro = CF("Resgate_Le", objResgate)
        If lErro <> SUCESSO And lErro <> 17499 Then Error 17500

        'Se não encontrou o resgate em questão
        If lErro = 17499 Then Error 17501

        'Se encontrou verifica se o resgate esta ativo
        If objResgate.iStatus = RESGATE_EXCLUIDO Then Error 17502

        lCodigoAplicacao = objResgate.lCodigoAplicacao

        lErro = Traz_Resgate_Tela(objResgate)
        If lErro <> SUCESSO Then Error 17503
        'Resgate esta cadastrado

    Else

        DataResgate.Text = Format(gdtDataAtual, "dd/mm/yy")

    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 17500

        Case 17501
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RESGATE_INEXISTENTE", Err, objResgate.lCodigoAplicacao)

        Case 17502
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RESGATE_EXCLUIDO", Err, Err, objResgate.iSeqResgate, objResgate.lCodigoAplicacao)

        Case 17503
            Call Limpa_Tela_Resgate

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174159)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objResgate As New ClassResgate

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à tela
    sTabela = "Resgates"

    If Len(Trim(CodAplicacao.ClipText)) > 0 Then
        objResgate.lCodigoAplicacao = StrParaLong(CodAplicacao.ClipText)
    Else
        objResgate.lCodigoAplicacao = 0
    End If

    If Len(Trim(CodResgate.Text)) > 0 Then
        objResgate.iSeqResgate = Codigo_Extrai(CodResgate.Text)
    Else
        objResgate.iSeqResgate = 0
    End If

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    
    colCampoValor.Add "CodigoAplicacao", objResgate.lCodigoAplicacao, 0, "CodigoAplicacao"
    colCampoValor.Add "SeqResgate", objResgate.iSeqResgate, 0, "SeqResgate"
    colCampoValor.Add "ValorResgatado", objResgate.dValorResgatado, 0, "ValorResgatado"
    colCampoValor.Add "Rendimentos", objResgate.dRendimentos, 0, "Rendimentos"
    colCampoValor.Add "ValorIRRF", objResgate.dValorIRRF, 0, "ValorIRRF"
    colCampoValor.Add "Descontos", objResgate.dDescontos, 0, "Descontos"
    colCampoValor.Add "SaldoAnterior", objResgate.dSaldoAnterior, 0, "SaldoAnterior"
    
    colSelecao.Add "Status", OP_DIFERENTE, RESGATE_EXCLUIDO
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174160)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objResgate As New ClassResgate
Dim iIndice As Integer

On Error GoTo Erro_Tela_Preenche

    objResgate.lCodigoAplicacao = colCampoValor.Item("CodigoAplicacao").vValor
    objResgate.iSeqResgate = colCampoValor.Item("SeqResgate").vValor
    
    If objResgate.lCodigoAplicacao > 0 And objResgate.iSeqResgate > 0 Then
            
        lErro = CF("Resgate_Le", objResgate)
        If lErro <> SUCESSO And lErro <> 17499 Then Error 54802

        For iIndice = 0 To CodResgate.ListCount - 1
            If CodResgate.List(iIndice) = CStr(objResgate.iSeqResgate) Then
                CodResgate.ListIndex = iIndice
            End If
        Next
        
        lErro = Traz_Resgate_Tela(objResgate)
        If lErro <> SUCESSO Then Error 54803

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 54802, 54803

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174161)

    End Select

    Exit Sub

End Sub

Private Function Traz_Resgate_Tela(objResgate As ClassResgate) As Long
'Coloca na Tela os dados do Resgate passado como parametro

Dim lErro As Long
Dim objMovContaCorrente As New ClassMovContaCorrente
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim objTipoMeioPagto As New ClassTipoMeioPagto
Dim dSaldoAnt As Double
Dim dRend As Double
Dim dValResg As Double
Dim dSaldoAtual As Double
Dim dIrrf As Double
Dim dDescontos As Double
Dim dValCredit As Double

On Error GoTo Erro_Traz_Resgate_Tela

    objMovContaCorrente.lNumMovto = objResgate.lNumMovto

    'Carrega os dados do movimento relativos ao resgate a partir da chave
    lErro = CF("MovContaCorrente_Le", objMovContaCorrente)
    If lErro <> SUCESSO And lErro <> 11893 Then Error 17504

    'Se o movimento não estiver cadastrado ==> Erro
    If lErro = 11893 Then Error 17505

    If objMovContaCorrente.iExcluido = EXCLUIDO Then Error 17506
        
    If objMovContaCorrente.iTipo <> MOVCCI_RESGATE Then Error 17507

    'Passa os dados para a tela
    CodAplicacao.PromptInclude = False
    CodAplicacao.Text = CStr(objResgate.lCodigoAplicacao)
    CodAplicacao.PromptInclude = True
    
    DataResgate.Text = Format(objMovContaCorrente.dtDataMovimento, "dd/MM/yy")
    SaldoAnterior.Caption = Format(objResgate.dSaldoAnterior, "Standard")
    Rendimentos.Text = CStr(objResgate.dRendimentos)
    Irrf.Text = CStr(objResgate.dValorIRRF)
    Descontos.Text = CStr(objResgate.dDescontos)
        
    'Calcula Saldo Atual
    dValResg = objResgate.dValorResgatado

    'só rendimento
    If dValResg = 0 Then
            
            ComboTipoMov.ListIndex = 2
            AplicacaoAdicional.Text = 0
    ElseIf dValResg > 0 Then 'resgate
            ComboTipoMov.ListIndex = 0
            AplicacaoAdicional.Text = 0
    Else        'aplicacao adicional
            ComboTipoMov.ListIndex = 1
            AplicacaoAdicional.Text = CStr(-dValResg)
            dValResg = 0
    End If
    
    ValorResgate.Text = CStr(dValResg)
    ValorResgateLabel.Caption = Format(dValResg, "Standard")
    
    Call SaldoAtualAtualiza
    
    'Calcula o valor creditado
    dIrrf = objResgate.dValorIRRF
    dDescontos = objResgate.dDescontos
    dValCredit = dValResg - dIrrf - dDescontos

    'Mostra na tela o valor creditado
    ValorCreditadoLabel.Caption = Format(dValCredit, "Standard")
    
    Historico.Text = objMovContaCorrente.sHistorico
    NumRefExterna.Text = objMovContaCorrente.sNumRefExterna
    
    If objMovContaCorrente.lNumero <> 0 Then
        Numero.Text = CStr(objMovContaCorrente.lNumero)
    Else
        Numero.Text = ""
    End If

    'Verifica se a conta corrente existe
    lErro = CF("ContaCorrenteInt_Le", objMovContaCorrente.iCodConta, objContaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 17508

    If lErro = 11807 Then Error 17509

    CodContaCorrente.Text = CStr(objMovContaCorrente.iCodConta) & SEPARADOR & objContaCorrenteInt.sNomeReduzido

    'Verifica se o TiPoMeioPagto existe
    objTipoMeioPagto.iTipo = objMovContaCorrente.iTipoMeioPagto

    lErro = CF("TipoMeioPagto_Le", objTipoMeioPagto)
    If lErro <> SUCESSO And lErro <> 11909 Then Error 17510

    If lErro = 11909 Then Error 17511

    TipoMeioPagto.Text = CStr(objMovContaCorrente.iTipoMeioPagto) & SEPARADOR & objTipoMeioPagto.sDescricao
    
    If CodResgate.ListCount = 0 Then
    
        'Carrega todos os resgates referentes a aplicacao em questao
        lErro = Carrega_Resgate(objResgate)
        If lErro <> SUCESSO Then Error 55969
    
    End If
    
    'traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objResgate.lNumMovto)
    If lErro <> SUCESSO And lErro <> 36326 Then Error 39563
    
    iAlterado = 0
    
    Traz_Resgate_Tela = SUCESSO

    Exit Function

Erro_Traz_Resgate_Tela:

    Traz_Resgate_Tela = Err

    Select Case Err

        Case 17504, 17508, 17510, 17511, 39563, 55969

        Case 17505
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_CADASTRADO3", Err, objMovContaCorrente.lNumMovto)

        Case 17506
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVCONTACORRENTE_EXCLUIDO", Err, objMovContaCorrente.iCodConta, objMovContaCorrente.lSequencial)

        Case 17507
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_RESGATE", Err, objMovContaCorrente.lSequencial)

        Case 17509
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, objMovContaCorrente.iCodConta)

        Case 17511
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INEXISTENTE", Err, objMovContaCorrente.iTipoMeioPagto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174162)

    End Select

    Exit Function

End Function

Private Sub Rendimentos_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Rendimentos_GotFocus()
    iRendimentos_Validate = 0
End Sub

Private Sub Rendimentos_Validate(Cancel As Boolean)

Dim curTeste As Currency
Dim lErro As Long

On Error GoTo Erro_Rendimentos_Validate


    If iValorResgate_Validate = 1 Or iDescontos_Validate = 1 Or iIRRF_Validate = 1 Then Exit Sub

    'Verifica se rendimentos foi preenchido
    If Len(Trim(Rendimentos.Text)) > 0 Then
        
        lErro = Valor_NaoNegativo_Critica(Rendimentos.Text)
        If lErro <> SUCESSO Then Error 17533

        Rendimentos.Text = Format(Rendimentos.Text, "Fixed")
    
    End If
    
    Call SaldoAtualAtualiza
    
    Exit Sub

Erro_Rendimentos_Validate:

    Cancel = True


    Select Case Err

        Case 17533
            iRendimentos_Validate = 1

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174163)

    End Select

    Exit Sub

End Sub

Private Sub SaldoAnterior_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TaxaPrevista_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TaxaPrevista_Validate(Cancel As Boolean)

Dim curTeste As Currency
Dim lErro As Long
Dim dSalAtual As Double
Dim dValResgPrev As Double
Dim dPercentual As Double

On Error GoTo Erro_TaxaPrevista_Validate

    'Verifica se taxa prevista está preenchida
    If Len(TaxaPrevista.Text) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TaxaPrevista.Text)
        If lErro <> SUCESSO Then Error 17541

        'Verifica se valor aplicado está preenchido
        If Len(Trim(SaldoAtual.Caption)) > 0 Then

           dPercentual = CDbl(TaxaPrevista.Text) / 100

           dSalAtual = CDbl(SaldoAtual.Caption)

           dValResgPrev = (1 + dPercentual) * dSalAtual

           'Coloca o valor resgate previsto na tela
           ValorResgatePrevisto.Text = Format(dValResgPrev, "Fixed")

        End If

    End If

    Exit Sub

Erro_TaxaPrevista_Validate:

    Cancel = True


    Select Case Err

        Case 17541

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174164)

    End Select

    Exit Sub

End Sub

Private Sub TipoMeioPagto_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoMeioPagto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTipoMeioPagto As New ClassTipoMeioPagto
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_TipoMeioPagto_Validate

    'verifica se foi preenchido o TipoMeioPagto
    If Len(Trim(TipoMeioPagto.Text)) = 0 Then Exit Sub

    'verifica se esta preenchida com o item selecionado na ComboBox TipoMeioPagto
    If TipoMeioPagto.Text = TipoMeioPagto.List(TipoMeioPagto.ListIndex) Then Exit Sub

    'Tenta selecionar o TipoMeioPagto com o codigo digitado
    lErro = Combo_Seleciona(TipoMeioPagto, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 17534

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTipoMeioPagto.iTipo = iCodigo

        'Pesquisa no BD a existencia do tipo passado por parametro
        lErro = CF("TipoMeioPagto_Le", objTipoMeioPagto)
        If lErro <> SUCESSO And lErro <> 11909 Then Error 17535

        'Se não existir o tipomeiopagto com o codigo ==> Erro
        If lErro = 11909 Then Error 17536

        TipoMeioPagto.Text = CStr(objTipoMeioPagto.iTipo) & SEPARADOR & objTipoMeioPagto.sDescricao

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 17537

    Exit Sub

Erro_TipoMeioPagto_Validate:

    Cancel = True


    Select Case Err

        Case 17534, 17535

        Case 17536
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INEXISTENTE", Err, objTipoMeioPagto.iTipo)
            
        Case 17537
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INEXISTENTE1", Err, TipoMeioPagto.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174165)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_DownClick

    DataResgate.SetFocus

    If Len(DataResgate.ClipText) > 0 Then

        sData = DataResgate.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 17649

        DataResgate.Text = sData

    End If

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 17649

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174166)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_UpClick

    DataResgate.SetFocus

    If Len(DataResgate.ClipText) > 0 Then

        sData = DataResgate.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 17650

        DataResgate.Text = sData

    End If

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 17650

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174167)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown2_DownClick

    DataResgatePrevista.SetFocus

    If Len(DataResgatePrevista.ClipText) > 0 Then

        sData = DataResgatePrevista.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 17651

        DataResgatePrevista.Text = sData

    End If

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case Err

        Case 17651

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174168)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown2_UpClick

    DataResgatePrevista.SetFocus

    If Len(DataResgatePrevista.ClipText) > 0 Then

        sData = DataResgatePrevista.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 17652

        DataResgatePrevista.Text = sData

    End If

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case Err

        Case 17652

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174169)

    End Select

    Exit Sub

End Sub

Private Sub ValorCreditadoLabel_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorResgate_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorResgate_GotFocus()
    iValorResgate_Validate = 0
End Sub

Private Sub ValorResgate_Validate(Cancel As Boolean)

Dim curTeste As Currency
Dim lErro As Long
Dim dRend As Double
Dim dSalAnt As Double
Dim dValResg As Double
Dim dSalAtual As Double
Dim dDesc As Double
Dim dIrrf As Double
Dim dValCredLab As Double

On Error GoTo Erro_ValorResgate_Validate

    If iValorResgate_Validate = 1 Or iRendimentos_Validate = 1 Or iIRRF_Validate = 1 Or iAplicacaoAdicional_Validate = 1 Then Exit Sub

    'Verifica se valor resgate foi preenchido
    If Len(Trim(ValorResgate.Text)) > 0 Then
        lErro = Valor_NaoNegativo_Critica(ValorResgate.Text)
        If lErro <> SUCESSO Then Error 17594

        dValResg = CDbl(ValorResgate.Text)

        ValorResgate.Text = Format(ValorResgate.Text, "Fixed")
    
    End If
    
    Call SaldoAtualAtualiza

    'Mostra na tela o valor do resgate
    ValorResgateLabel.Caption = Format(dValResg, "Standard")

    'Verifica se valor Irrf esta preenchido
    If Len(Irrf.Text) > 0 Then dIrrf = CDbl(Irrf.Text)

    'Verifica se descontos esta preenchido
    If Len(Descontos.Text) > 0 Then dDesc = CDbl(Descontos.Text)

    'Calcula o valor creditado
    dValCredLab = dValResg - dIrrf - dDesc

    'Mostra na tela o valor creditado
    ValorCreditadoLabel.Caption = Format(dValCredLab, "Standard")
    
    Exit Sub

Erro_ValorResgate_Validate:

    Cancel = True


    Select Case Err

        Case 17594
            iValorResgate_Validate = 1

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174170)

    End Select

    Exit Sub

End Sub

Private Sub ValorResgateLabel_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorResgatePrevisto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Limpa_Tela_Resgate()

    Call Limpa_Tela(Me)

    CodContaCorrente = ""
    DataResgate.Text = Format(gdtDataAtual, "dd/mm/yy")
    SaldoAnterior = ""
    SaldoAtual.Caption = ""
    ValorResgateLabel = ""
    ValorCreditadoLabel = ""
    Historico.Text = ""
    TipoMeioPagto.Text = ""
    
    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

End Sub

Private Sub Limpa_Tela_Resgate2()

    CodContaCorrente = ""
    DataResgate.Text = Format(gdtDataAtual, "dd/mm/yy")
    SaldoAnterior = ""
    SaldoAtual.Caption = ""
    ValorResgateLabel = ""
    ValorCreditadoLabel = ""
    Historico.Text = ""
    TipoMeioPagto.Text = ""
    ValorResgate.Text = ""
    Rendimentos.Text = ""
    Descontos.Text = ""
    Irrf.Text = ""
    Numero.Text = ""
    NumRefExterna.Text = ""
    
    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

End Sub

Private Sub ValorResgatePrevisto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim curTeste As Currency
Dim dSalAtual As Double
Dim dValResgPrev As Double
Dim dPercentual As Double
Dim dValAplic As Double

On Error GoTo Erro_ValorResgatePrevisto_Validate

    'Verifica se valor do resgate previsto está preenchido
    If Len(ValorResgatePrevisto.Text) > 0 Then

        lErro = Valor_NaoNegativo_Critica(ValorResgatePrevisto.Text)
        If lErro <> SUCESSO Then Error 17539

        dValResgPrev = CDbl(ValorResgatePrevisto.Text)

        ValorResgatePrevisto.Text = Format(ValorResgatePrevisto.Text, "Fixed")

        'Verifica se saldo atual está preenchido
        If Len(Trim(SaldoAtual.Caption)) > 0 Then

            dSalAtual = CDbl(SaldoAtual.Caption)
           
            If dSalAtual <> 0 Then

                'Calcula a taxa prevista
                dPercentual = (dValResgPrev - dSalAtual) / dSalAtual * 100
                
                'Coloca a taxa na tela
                TaxaPrevista.Text = Format(dPercentual, "Fixed")
            
            End If
        End If

    End If

    Exit Sub

Erro_ValorResgatePrevisto_Validate:

    Cancel = True


    Select Case Err

        Case 17539

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174171)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim dSalAtual As Double
Dim dValCred As Double
Dim dValCredLab As Double
Dim dValResgPrev As Double
Dim dtDataResg As Date
Dim dtDataResgPrev As Date
Dim objResgate As New ClassResgate
Dim objAplicacao As New ClassAplicacao
Dim objMovContaCorrente As New ClassMovContaCorrente

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos essencias da tela foram preenchidos
    If Len(Trim(CodAplicacao.Text)) = 0 Then Error 17543

    If Len(Trim(CodResgate.Text)) = 0 Then Error 17544
    
    If Len(Trim(SaldoAtual.Caption)) <> 0 Then

        dSalAtual = CDbl(SaldoAtual.Caption)
        If dSalAtual < 0 Then Error 17550
    End If
    
    'If StrParaDbl(SaldoAnterior.Caption) = 0 Then Error 20594
    
    If Len(Trim(DataResgate.ClipText)) = 0 Then Error 17545
    
    If StrParaDbl(ValorResgate) < 0 Then Error 20593

    dtDataResg = CDate(DataResgate.Text)

    dValCred = StrParaDbl(ValorCreditadoLabel.Caption)

    If dValCred < 0 Then Error 17547

    If Len(Trim(CodContaCorrente.Text)) = 0 Then Error 17548

    If Len(Trim(TipoMeioPagto.Text)) = 0 Then Error 17549

    If Len(Trim(DataResgatePrevista.ClipText)) <> 0 Then

        dtDataResgPrev = CDate(DataResgatePrevista.Text)

       If dtDataResgPrev < dtDataResg Then Error 17551

    End If

    If Len(Trim(ValorResgatePrevisto.Text)) <> 0 Then

       dValResgPrev = CDbl(ValorResgatePrevisto.Text)

       If dValResgPrev < dSalAtual Then Error 17552

    End If

    'Verifica se valor creditado label esta preenchido
    If Len((Trim(ValorCreditadoLabel.Caption))) <> 0 Then

        dValCredLab = CDbl(ValorCreditadoLabel.Caption)
        
    End If

    'Passa os dados da Tela para objResgate e objMovContaCorrente
    lErro = Move_Tela_Memoria(objResgate, objMovContaCorrente, objAplicacao)
    If lErro <> SUCESSO Then Error 17554

    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(DataResgate.Text))
    If lErro <> SUCESSO Then Error 20834

    'Rotina encarregada de gravar o resgate
    lErro = CF("Resgate_Grava", objResgate, objMovContaCorrente, objAplicacao, objContabil)
    If lErro <> SUCESSO Then Error 17349

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
       
        Case 17349, 17554, 20834

        Case 17543
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_APLICACA0_NAO_PREENCHIDO", Err)

        Case 17544
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_RESGATE_NAO_PREENCHIDO", Err)

        Case 17545
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_RESGATE_NAO_PREENCHIDA", Err)

        Case 17547
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_POSITIVO", Err, ValorCreditadoLabel.Caption)

        Case 17548
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PREENCHIDA", Err)

        Case 17549
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_NAO_INFORMADO", Err)

        Case 17550
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_ATUAL_NEGATIVO", Err, SaldoAtual.Caption)

        Case 17551
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATARESGPREV_MENOR_DATARESG", Err, dtDataResgPrev, dtDataResg)

        Case 17552
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALRESGATE_MENOR_SALATUAL", Err, dValResgPrev, dSalAtual)

        Case 20593
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_RESGATE_NAO_PREENCHIDO", Err)
        
        Case 20594
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_ZERO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174172)

    End Select

    Exit Function

End Function

Function Move_Tela_Memoria(objResgate As ClassResgate, objMovContaCorrente As ClassMovContaCorrente, objAplicacao As ClassAplicacao) As Long
'Move os dados da tela para memoria.

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'Move os dados da tela para objResgate e objMovContaCorrente
    objMovContaCorrente.iCodConta = Codigo_Extrai(CodContaCorrente.Text)
    objResgate.lCodigoAplicacao = CLng(CodAplicacao.Text)
    objResgate.iSeqResgate = Codigo_Extrai(CodResgate.Text)
    objMovContaCorrente.dtDataMovimento = CDate(DataResgate.Text)
    objResgate.dValorCreditado = StrParaDbl(ValorCreditadoLabel.Caption)
    objMovContaCorrente.iFilialEmpresa = giFilialEmpresa
    If Len(Trim(SaldoAnterior.Caption)) > 0 Then
        objResgate.dSaldoAnterior = CDbl(SaldoAnterior.Caption)
    Else
        objResgate.dSaldoAnterior = 0
    End If

    If Len(Trim(Rendimentos.Text)) > 0 Then
        objResgate.dRendimentos = CDbl(Rendimentos.Text)
    Else
        objResgate.dRendimentos = 0
    End If

    objResgate.dValorResgatado = 0
    
    If Len(Trim(ValorResgate.Text)) > 0 Then
        objResgate.dValorResgatado = CDbl(ValorResgate.Text)
    End If

    If Len(Trim(AplicacaoAdicional.Text)) > 0 Then
        objResgate.dValorResgatado = -CDbl(AplicacaoAdicional.Text)
    End If

    If Len(Trim(Irrf.Text)) > 0 Then
        objResgate.dValorIRRF = CDbl(Irrf.Text)
    Else
        objResgate.dValorIRRF = 0
    End If

    If Len(Trim(Descontos.Text)) > 0 Then
        objResgate.dDescontos = CDbl(Descontos.Text)
    Else
        objResgate.dDescontos = 0
    End If

    If Len(Trim(Numero.Text)) > 0 Then
        objMovContaCorrente.lNumero = CLng(Numero.Text)
    Else
        objMovContaCorrente.lNumero = 0
    End If

    If Len(Trim(Historico.Text)) <> 0 Then objMovContaCorrente.sHistorico = Historico.Text

    If Len(Trim(NumRefExterna.Text)) <> 0 Then objMovContaCorrente.sNumRefExterna = NumRefExterna.Text

    objMovContaCorrente.iTipoMeioPagto = Codigo_Extrai(TipoMeioPagto.Text)

    'Move os dados da tela para objAplicacao
    If Len(Trim(DataResgatePrevista.ClipText)) > 0 Then
        objAplicacao.dtDataResgatePrevista = CDate(DataResgatePrevista.Text)
    Else
        objAplicacao.dtDataResgatePrevista = DATA_NULA
    End If

    If Len(Trim(ValorResgatePrevisto.Text)) > 0 Then
        objAplicacao.dValorResgatePrevisto = CDbl(ValorResgatePrevisto.Text)
    Else
        objAplicacao.dValorResgatePrevisto = 0
    End If

    If Len(Trim(TaxaPrevista.Text)) > 0 Then
        objAplicacao.dTaxaPrevista = CDbl(TaxaPrevista.Text)
    Else
        objAplicacao.dTaxaPrevista = 0
    End If

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174173)

    End Select

    Exit Function

End Function

'inicio contabilidade

Private Sub CTBBotaoModeloPadrao_Click()

    Call objContabil.Contabil_BotaoModeloPadrao_Click

End Sub

Private Sub CTBModelo_Click()

    Call objContabil.Contabil_Modelo_Click

End Sub

Private Sub CTBGridContabil_Click()

    Call objContabil.Contabil_GridContabil_Click

    If giTipoVersao = VERSAO_LIGHT Then
        Call objContabil.Contabil_GridContabil_Consulta_Click
    End If

End Sub

Private Sub CTBGridContabil_EnterCell()

    Call objContabil.Contabil_GridContabil_EnterCell

End Sub

Private Sub CTBGridContabil_GotFocus()

    Call objContabil.Contabil_GridContabil_GotFocus

End Sub

Private Sub CTBGridContabil_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_GridContabil_KeyPress(KeyAscii)

End Sub

Private Sub CTBGridContabil_KeyDown(KeyCode As Integer, Shift As Integer)

    Call objContabil.Contabil_GridContabil_KeyDown(KeyCode)
    
End Sub

Private Sub CTBGridContabil_LeaveCell()

        Call objContabil.Contabil_GridContabil_LeaveCell

End Sub

Private Sub CTBGridContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_GridContabil_Validate(Cancel)

End Sub

Private Sub CTBGridContabil_RowColChange()

    Call objContabil.Contabil_GridContabil_RowColChange

End Sub

Private Sub CTBGridContabil_Scroll()

    Call objContabil.Contabil_GridContabil_Scroll

End Sub

Private Sub CTBConta_Change()

    Call objContabil.Contabil_Conta_Change

End Sub

Private Sub CTBConta_GotFocus()

    Call objContabil.Contabil_Conta_GotFocus

End Sub

Private Sub CTBConta_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Conta_KeyPress(KeyAscii)

End Sub

Private Sub CTBConta_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Conta_Validate(Cancel)

End Sub

Private Sub CTBCcl_Change()

    Call objContabil.Contabil_Ccl_Change

End Sub

Private Sub CTBCcl_GotFocus()

    Call objContabil.Contabil_Ccl_GotFocus

End Sub

Private Sub CTBCcl_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Ccl_KeyPress(KeyAscii)

End Sub

Private Sub CTBCcl_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Ccl_Validate(Cancel)

End Sub

Private Sub CTBCredito_Change()

    Call objContabil.Contabil_Credito_Change

End Sub

Private Sub CTBCredito_GotFocus()

    Call objContabil.Contabil_Credito_GotFocus

End Sub

Private Sub CTBCredito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Credito_KeyPress(KeyAscii)

End Sub

Private Sub CTBCredito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Credito_Validate(Cancel)

End Sub

Private Sub CTBDebito_Change()

    Call objContabil.Contabil_Debito_Change

End Sub

Private Sub CTBDebito_GotFocus()

    Call objContabil.Contabil_Debito_GotFocus

End Sub

Private Sub CTBDebito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Debito_KeyPress(KeyAscii)

End Sub

Private Sub CTBDebito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Debito_Validate(Cancel)

End Sub

Private Sub CTBSeqContraPartida_Change()

    Call objContabil.Contabil_SeqContraPartida_Change

End Sub

Private Sub CTBSeqContraPartida_GotFocus()

    Call objContabil.Contabil_SeqContraPartida_GotFocus

End Sub

Private Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_SeqContraPartida_KeyPress(KeyAscii)

End Sub

Private Sub CTBSeqContraPartida_Validate(Cancel As Boolean)

    Call objContabil.Contabil_SeqContraPartida_Validate(Cancel)

End Sub

Private Sub CTBHistorico_Change()

    Call objContabil.Contabil_Historico_Change

End Sub

Private Sub CTBHistorico_GotFocus()

    Call objContabil.Contabil_Historico_GotFocus

End Sub

Private Sub CTBHistorico_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Historico_KeyPress(KeyAscii)

End Sub

Private Sub CTBHistorico_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Historico_Validate(Cancel)

End Sub

Private Sub CTBLancAutomatico_Click()

    Call objContabil.Contabil_LancAutomatico_Click

End Sub

Private Sub CTBAglutina_Click()
    
    Call objContabil.Contabil_Aglutina_Click

End Sub

Private Sub CTBAglutina_GotFocus()

    Call objContabil.Contabil_Aglutina_GotFocus

End Sub

Private Sub CTBAglutina_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Aglutina_KeyPress(KeyAscii)

End Sub

Private Sub CTBAglutina_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Aglutina_Validate(Cancel)

End Sub

Private Sub CTBTvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_NodeClick(Node)

End Sub

Private Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_Expand(Node, CTBTvwContas.Nodes)

End Sub

Private Sub CTBTvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwCcls_NodeClick(Node)

End Sub

Private Sub CTBListHistoricos_DblClick()

    Call objContabil.Contabil_ListHistoricos_DblClick

End Sub

Private Sub CTBBotaoLimparGrid_Click()

    Call objContabil.Contabil_Limpa_GridContabil

End Sub

Private Sub CTBLote_Change()

    Call objContabil.Contabil_Lote_Change

End Sub

Private Sub CTBLote_GotFocus()

    Call objContabil.Contabil_Lote_GotFocus

End Sub

Private Sub CTBLote_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Lote_Validate(Cancel, Parent)

End Sub

Private Sub CTBDataContabil_Change()

    Call objContabil.Contabil_DataContabil_Change

End Sub

Private Sub CTBDataContabil_GotFocus()

    Call objContabil.Contabil_DataContabil_GotFocus

End Sub

Private Sub CTBDataContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_DataContabil_Validate(Cancel, Parent)

End Sub

Private Sub objEventoLote_evSelecao(obj1 As Object)
'traz o lote selecionado para a tela

    Call objContabil.Contabil_objEventoLote_evSelecao(obj1)

End Sub

Private Sub objEventoDoc_evSelecao(obj1 As Object)

    Call objContabil.Contabil_objEventoDoc_evSelecao(obj1)

End Sub

Private Sub CTBDocumento_Change()

    Call objContabil.Contabil_Documento_Change

End Sub

Private Sub CTBDocumento_GotFocus()

    Call objContabil.Contabil_Documento_GotFocus

End Sub

Private Sub CTBBotaoImprimir_Click()
    
    Call objContabil.Contabil_BotaoImprimir_Click

End Sub

Private Sub CTBUpDown_DownClick()

    Call objContabil.Contabil_UpDown_DownClick
    
End Sub

Private Sub CTBUpDown_UpClick()

    Call objContabil.Contabil_UpDown_UpClick

End Sub

Private Sub CTBLabelDoc_Click()

    Call objContabil.Contabil_LabelDoc_Click
    
End Sub

Private Sub CTBLabelLote_Click()

    Call objContabil.Contabil_LabelLote_Click
    
End Sub

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long
Dim objTiposDeAplicacao As New ClassTiposDeAplicacao
Dim objAplicacao As New ClassAplicacao
Dim objContasCorrentesInternas As New ClassContasCorrentesInternas
Dim sContaTela As String

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case TIPO_APLICACAO
            
            If Len(Trim(CodAplicacao.Text)) > 0 Then
                
                objAplicacao.lCodigo = CLng(CodAplicacao.Text)
                
                'Le a Aplicacao
                lErro = CF("Aplicacao_Le", objAplicacao)
                If lErro <> SUCESSO And lErro <> 17241 Then gError 64418
                
                If lErro = 17241 Then gError 64419
                
                objTiposDeAplicacao.iCodigo = objAplicacao.iTipoAplicacao
                
                'Le  a conta no BD
                lErro = CF("TiposDeAplicacao_Le", objTiposDeAplicacao)
                If lErro <> SUCESSO And lErro <> 15068 Then gError 64420
                
                'Se não encontrou ---> Erro
                If lErro = 15068 Then gError 64421
                
                objMnemonicoValor.colValor.Add objTiposDeAplicacao.sDescricao
                
            Else
                objMnemonicoValor.colValor.Add ""
            End If
        
        Case CTATIPOAPLICACAO
            
            If Len(Trim(CodAplicacao.Text)) > 0 Then
                
                objAplicacao.lCodigo = CLng(CodAplicacao.Text)
                
                'Le a Aplicacao
                lErro = CF("Aplicacao_Le", objAplicacao)
                If lErro <> SUCESSO And lErro <> 17241 Then gError 64422
                
                If lErro = 17241 Then gError 64423
                
                objTiposDeAplicacao.iCodigo = objAplicacao.iTipoAplicacao
                
                'Le  a conta no BD
                lErro = CF("TiposDeAplicacao_Le", objTiposDeAplicacao)
                If lErro <> SUCESSO And lErro <> 15068 Then gError 64424
                
                'Se não encontrou ---> Erro
                If lErro = 15068 Then gError 64425
                                
                If objTiposDeAplicacao.sContaAplicacao <> "" Then
                                
                    lErro = Mascara_RetornaContaTela(objTiposDeAplicacao.sContaAplicacao, sContaTela)
                    If lErro <> SUCESSO Then gError 64449
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                
            Else
                objMnemonicoValor.colValor.Add ""
            End If

        Case CTARECEITAAPLICACAO
        
            If Len(Trim(CodAplicacao.Text)) > 0 Then
                
                objAplicacao.lCodigo = CLng(CodAplicacao.Text)
                
                'Le a Aplicacao
                lErro = CF("Aplicacao_Le", objAplicacao)
                If lErro <> SUCESSO Then gError 64426
                
                If lErro = 17241 Then gError 64427
                
                objTiposDeAplicacao.iCodigo = objAplicacao.iTipoAplicacao
                
                'Le  a conta no BD
                lErro = CF("TiposDeAplicacao_Le", objTiposDeAplicacao)
                If lErro <> SUCESSO And lErro <> 15068 Then gError 64428
                
                'Se não encontrou ---> Erro
                If lErro = 15068 Then gError 64429
                
                If objTiposDeAplicacao.sContaReceita <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objTiposDeAplicacao.sContaReceita, sContaTela)
                    If lErro <> SUCESSO Then gError 64450
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                
            Else
                objMnemonicoValor.colValor.Add ""
            End If

        Case CTACONTACORRENTE
            If Len(CodContaCorrente.Text) > 0 Then
                
                objContasCorrentesInternas.iCodigo = Codigo_Extrai(CodContaCorrente.Text)
                
                'Procura a conta no BD
                lErro = CF("ContaCorrenteInt_Le", objContasCorrentesInternas.iCodigo, objContasCorrentesInternas)
                If lErro <> SUCESSO And lErro <> 11807 Then gError 64430
            
                'Se nao estiver cadastrada --> Erro
                If lErro = 11807 Then gError 64431
                
                If objContasCorrentesInternas.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objContasCorrentesInternas.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 64451
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                
            Else
                objMnemonicoValor.colValor.Add ""
            End If
        
        Case RESGATE1
            If Len(CodResgate.Text) > 0 Then
                objMnemonicoValor.colValor.Add CInt(CodResgate.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case APLICACAO1
            If Len(CodAplicacao.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CLng(CodAplicacao.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case CONTACORRENTE1
            If Len(Trim(CodContaCorrente.Text)) > 0 Then
            
                objContasCorrentesInternas.iCodigo = Codigo_Extrai(CodContaCorrente.Text)
                
                'Procura a conta no BD
                lErro = CF("ContaCorrenteInt_Le", objContasCorrentesInternas.iCodigo, objContasCorrentesInternas)
                If lErro <> SUCESSO And lErro <> 11807 Then gError 64432
            
                'Se nao estiver cadastrada --> Erro
                If lErro = 11807 Then gError 64433
                
                objMnemonicoValor.colValor.Add objContasCorrentesInternas.sNomeReduzido

            Else
                objMnemonicoValor.colValor.Add ""
            End If
                
        Case HISTORICO1
            If Len(Historico.Text) > 0 Then
                objMnemonicoValor.colValor.Add Historico.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case FORMA1
            If Len(TipoMeioPagto.Text) > 0 Then
                objMnemonicoValor.colValor.Add TipoMeioPagto.ItemData(TipoMeioPagto.ListIndex)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case VALORRESGATE1
            If Len(ValorResgate.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorResgate.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        Case VALORCREDITADO1
            If Len(ValorCreditadoLabel.Caption) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorCreditadoLabel.Caption)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case RENDIMENTOS1
            If Len(Rendimentos.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(Rendimentos.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case IRRF1
            If Len(Irrf.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(Irrf.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        Case DESCONTOS1
            If Len(Descontos.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(Descontos.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case VALORPREVISTO1
            If Len(ValorResgatePrevisto.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorResgatePrevisto.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        Case Else
            gError 39566

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 39566
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case 64418, 64420, 64422, 64424, 64426, 64428, 64430, 64432, 64449, 64450, 64451
        
        Case 64419, 64422, 64427
            lErro = Rotina_Erro(vbOKOnly, "ERRO_APLICACAO_INEXISTENTE", gErr, objAplicacao.lCodigo)
        
        Case 64421, 64425, 64429
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOAPLICACAO_INEXISTENTE1", gErr, objTiposDeAplicacao.iCodigo)
        
        Case 64431, 64433
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", gErr, objContasCorrentesInternas.iCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174174)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RESGATE_ID
    Set Form_Load_Ocx = Me
    Caption = "Resgate"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Resgate"
    
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

'***** fim do trecho a ser copiado ******

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is CodAplicacao Then
            Call LabelCodAplicacao_Click
        ElseIf Me.ActiveControl Is CodContaCorrente Then
            Call LabelCtaCorrente_Click
        End If
    
    End If
    
End Sub


Private Sub Label33_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label33, Source, X, Y)
End Sub

Private Sub Label33_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label33, Button, Shift, X, Y)
End Sub

Private Sub Label32_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label32, Source, X, Y)
End Sub

Private Sub Label32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label32, Button, Shift, X, Y)
End Sub

Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label25, Source, X, Y)
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label25, Button, Shift, X, Y)
End Sub

Private Sub Label24_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label24, Source, X, Y)
End Sub

Private Sub Label24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label24, Button, Shift, X, Y)
End Sub

Private Sub ValorResgateLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorResgateLabel, Source, X, Y)
End Sub

Private Sub ValorResgateLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorResgateLabel, Button, Shift, X, Y)
End Sub

Private Sub Label31_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label31, Source, X, Y)
End Sub

Private Sub Label31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label31, Button, Shift, X, Y)
End Sub

Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub SaldoAtual_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SaldoAtual, Source, X, Y)
End Sub

Private Sub SaldoAtual_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SaldoAtual, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub SaldoAnterior_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SaldoAnterior, Source, X, Y)
End Sub

Private Sub SaldoAnterior_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SaldoAnterior, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub

Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
End Sub

Private Sub ValorCreditadoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorCreditadoLabel, Source, X, Y)
End Sub

Private Sub ValorCreditadoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorCreditadoLabel, Button, Shift, X, Y)
End Sub

Private Sub LabelCodResgate_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodResgate, Source, X, Y)
End Sub

Private Sub LabelCodResgate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodResgate, Button, Shift, X, Y)
End Sub

Private Sub Label30_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label30, Source, X, Y)
End Sub

Private Sub Label30_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label30, Button, Shift, X, Y)
End Sub

Private Sub LabelCodAplicacao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodAplicacao, Source, X, Y)
End Sub

Private Sub LabelCodAplicacao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodAplicacao, Button, Shift, X, Y)
End Sub

Private Sub LabelCtaCorrente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCtaCorrente, Source, X, Y)
End Sub

Private Sub LabelCtaCorrente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCtaCorrente, Button, Shift, X, Y)
End Sub

Private Sub CTBCclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclLabel, Source, X, Y)
End Sub

Private Sub CTBCclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclLabel, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel7, Source, X, Y)
End Sub

Private Sub CTBLabel7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel7, Button, Shift, X, Y)
End Sub

Private Sub CTBContaDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBContaDescricao, Source, X, Y)
End Sub

Private Sub CTBContaDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBContaDescricao, Button, Shift, X, Y)
End Sub

Private Sub CTBCclDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclDescricao, Source, X, Y)
End Sub

Private Sub CTBCclDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclDescricao, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel21, Source, X, Y)
End Sub

Private Sub CTBLabel21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel21, Button, Shift, X, Y)
End Sub

Private Sub CTBOrigem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBOrigem, Source, X, Y)
End Sub

Private Sub CTBOrigem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBOrigem, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel14, Source, X, Y)
End Sub

Private Sub CTBLabel14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel14, Button, Shift, X, Y)
End Sub

Private Sub CTBPeriodo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBPeriodo, Source, X, Y)
End Sub

Private Sub CTBPeriodo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBPeriodo, Button, Shift, X, Y)
End Sub

Private Sub CTBExercicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBExercicio, Source, X, Y)
End Sub

Private Sub CTBExercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBExercicio, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel13, Source, X, Y)
End Sub

Private Sub CTBLabel13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel13, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel5, Source, X, Y)
End Sub

Private Sub CTBLabel5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel5, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelHistoricos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelHistoricos, Source, X, Y)
End Sub

Private Sub CTBLabelHistoricos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelHistoricos, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelContas, Source, X, Y)
End Sub

Private Sub CTBLabelContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelContas, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelCcl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelCcl, Source, X, Y)
End Sub

Private Sub CTBLabelCcl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelCcl, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel1, Source, X, Y)
End Sub

Private Sub CTBLabel1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel1, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelTotais, Source, X, Y)
End Sub

Private Sub CTBLabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelTotais, Button, Shift, X, Y)
End Sub

Private Sub CTBTotalDebito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalDebito, Source, X, Y)
End Sub

Private Sub CTBTotalDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalDebito, Button, Shift, X, Y)
End Sub

Private Sub CTBTotalCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalCredito, Source, X, Y)
End Sub

Private Sub CTBTotalCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalCredito, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel8, Source, X, Y)
End Sub

Private Sub CTBLabel8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel8, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelDoc_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelDoc, Source, X, Y)
End Sub

Private Sub CTBLabelDoc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelDoc, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelLote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelLote, Source, X, Y)
End Sub

Private Sub CTBLabelLote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelLote, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub

Private Sub Label18_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label18, Source, X, Y)
End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label18, Button, Shift, X, Y)
End Sub

Private Sub Label36_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label36, Source, X, Y)
End Sub

Private Sub Label36_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label36, Button, Shift, X, Y)
End Sub

Private Sub Label37_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label37, Source, X, Y)
End Sub

Private Sub Label37_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label37, Button, Shift, X, Y)
End Sub

Private Sub Label34_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label34, Source, X, Y)
End Sub

Private Sub Label34_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label34, Button, Shift, X, Y)
End Sub

Private Sub Label35_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label35, Source, X, Y)
End Sub

Private Sub Label35_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label35, Button, Shift, X, Y)
End Sub


Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub
Private Sub CTBGerencial_Click()
    
    Call objContabil.Contabil_Gerencial_Click

End Sub

Private Sub CTBGerencial_GotFocus()

    Call objContabil.Contabil_Gerencial_GotFocus

End Sub

Private Sub CTBGerencial_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Gerencial_KeyPress(KeyAscii)

End Sub

Private Sub CTBGerencial_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Gerencial_Validate(Cancel)

End Sub

Private Sub SaldoAtualAtualiza()

Dim dRend As Double
Dim dSalAnt As Double
Dim dValResg As Double
Dim dSalAtual As Double
Dim dAplicAdicional As Double

    'Verifica se saldo anterior esta preenchido
    If Len(SaldoAnterior.Caption) > 0 Then
    
        dSalAnt = StrParaDbl(SaldoAnterior.Caption)
        dValResg = StrParaDbl(ValorResgate.Text)
        dRend = StrParaDbl(Rendimentos.Text)
        dAplicAdicional = StrParaDbl(AplicacaoAdicional.Text)
        
        'Calcula o saldo atual
        dSalAtual = dSalAnt + dRend + dAplicAdicional - dValResg
        
        'Mostra na tela o saldo atual
        SaldoAtual.Caption = Format(dSalAtual, "Standard")

    End If

End Sub

Private Sub AplicacaoAdicional_Validate(Cancel As Boolean)

Dim curTeste As Currency
Dim lErro As Long
Dim dRend As Double
Dim dSalAnt As Double
Dim dValResg As Double
Dim dSalAtual As Double
Dim dDesc As Double
Dim dIrrf As Double
Dim dValCredLab As Double

On Error GoTo Erro_AplicacaoAdicional_Validate

    If iValorResgate_Validate = 1 Or iRendimentos_Validate = 1 Or iIRRF_Validate = 1 Or iAplicacaoAdicional_Validate = 1 Then Exit Sub

    'Verifica se valor foi preenchido
    If Len(Trim(AplicacaoAdicional.Text)) > 0 Then
        
        lErro = Valor_NaoNegativo_Critica(AplicacaoAdicional.Text)
        If lErro <> SUCESSO Then Error 17594

        AplicacaoAdicional.Text = Format(AplicacaoAdicional.Text, "Fixed")
    
    End If
    
    Call SaldoAtualAtualiza

    Exit Sub

Erro_AplicacaoAdicional_Validate:

    Cancel = True


    Select Case Err

        Case 17594
            iAplicacaoAdicional_Validate = 1
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174170)

    End Select

    Exit Sub

End Sub

Private Sub AplicacaoAdicional_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub AplicacaoAdicional_GotFocus()
    iAplicacaoAdicional_Validate = 0
End Sub


