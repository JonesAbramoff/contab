VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ClassificacaoABC 
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   4785
      Index           =   1
      Left            =   285
      TabIndex        =   0
      Top             =   765
      Width           =   9015
      Begin VB.Frame Frame3 
         Caption         =   "Faixas de Classificação"
         Height          =   765
         Left            =   210
         TabIndex        =   34
         Top             =   3960
         Width           =   5475
         Begin MSMask.MaskEdBox FaixaA 
            Height          =   285
            Left            =   1155
            TabIndex        =   11
            Top             =   330
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   3
            Format          =   "0\%"
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FaixaB 
            Height          =   285
            Left            =   2925
            TabIndex        =   12
            Top             =   315
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   3
            Format          =   "0\%"
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Faixa B:"
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
            Left            =   2160
            TabIndex        =   38
            Top             =   360
            Width           =   705
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Faixa A:"
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
            Left            =   375
            TabIndex        =   37
            Top             =   360
            Width           =   705
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Faixa C:"
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
            Left            =   3885
            TabIndex        =   36
            Top             =   360
            Width           =   705
         End
         Begin VB.Label FaixaC 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4650
            TabIndex        =   35
            Top             =   330
            Width           =   555
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo de Produto"
         Height          =   1185
         Left            =   210
         TabIndex        =   30
         Top             =   1500
         Width           =   5475
         Begin VB.CheckBox TodosTipos 
            Caption         =   "Todos os tipos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   210
            TabIndex        =   5
            Top             =   300
            Value           =   1  'Checked
            Width           =   1605
         End
         Begin MSMask.MaskEdBox Tipo 
            Height          =   315
            Left            =   2580
            TabIndex        =   6
            Top             =   315
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label TipoLabel 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2070
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   33
            Top             =   330
            Width           =   450
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1620
            TabIndex        =   32
            Top             =   765
            Width           =   930
         End
         Begin VB.Label TipoDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2610
            TabIndex        =   31
            Top             =   720
            Width           =   2715
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Faixa de Tempo"
         Height          =   1065
         Left            =   210
         TabIndex        =   25
         Top             =   2760
         Width           =   5475
         Begin MSMask.MaskEdBox MesInicial 
            Height          =   285
            Left            =   1785
            TabIndex        =   7
            Top             =   285
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   2
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MesFinal 
            Height          =   285
            Left            =   1785
            TabIndex        =   8
            Top             =   675
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   2
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AnoInicial 
            Height          =   285
            Left            =   3855
            TabIndex        =   9
            Top             =   270
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AnoFinal 
            Height          =   285
            Left            =   3855
            TabIndex        =   10
            Top             =   660
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Ano Final:"
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
            Left            =   2910
            TabIndex        =   29
            Top             =   705
            Width           =   870
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Ano Inicial:"
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
            Left            =   2805
            TabIndex        =   28
            Top             =   330
            Width           =   975
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Mês Final:"
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
            Left            =   840
            TabIndex        =   27
            Top             =   705
            Width           =   885
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Mês Inicial:"
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
            Left            =   735
            TabIndex        =   26
            Top             =   330
            Width           =   990
         End
      End
      Begin VB.ListBox Classificacoes 
         Height          =   4350
         Left            =   6015
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   315
         Width           =   2775
      End
      Begin VB.Frame Frame4 
         Caption         =   "Identificação"
         Height          =   1395
         Left            =   210
         TabIndex        =   20
         Top             =   15
         Width           =   5475
         Begin VB.CheckBox AtualizaProdutosFilial 
            Caption         =   "Atualiza Classe ABC na tabela de Produtos"
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
            Left            =   420
            TabIndex        =   4
            Top             =   1035
            Width           =   4005
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   285
            Left            =   1380
            TabIndex        =   1
            Top             =   225
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   15
            PromptChar      =   "_"
         End
         Begin MSComCtl2.UpDown DataUpDown 
            Height          =   285
            Left            =   4635
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   210
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   285
            Left            =   3570
            TabIndex        =   2
            Top             =   210
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Descricao 
            Height          =   285
            Left            =   1365
            TabIndex        =   3
            Top             =   645
            Width           =   3510
            _ExtentX        =   6191
            _ExtentY        =   503
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   50
            PromptChar      =   " "
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   630
            TabIndex        =   24
            Top             =   285
            Width           =   660
         End
         Begin VB.Label Label3 
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
            Left            =   3030
            TabIndex        =   23
            Top             =   255
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   22
            Top             =   690
            Width           =   930
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Classificações"
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
         Left            =   6075
         TabIndex        =   39
         Top             =   90
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   4845
      Index           =   2
      Left            =   210
      TabIndex        =   14
      Top             =   750
      Visible         =   0   'False
      Width           =   9015
      Begin VB.PictureBox PictureABC 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   4755
         Left            =   240
         ScaleHeight     =   -127.554
         ScaleLeft       =   -30
         ScaleMode       =   0  'User
         ScaleTop        =   115
         ScaleWidth      =   155
         TabIndex        =   40
         Top             =   45
         Width           =   8565
         Begin VB.Label XLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   1
            Left            =   1755
            TabIndex        =   80
            Top             =   4050
            Width           =   150
         End
         Begin VB.Label XLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   2
            Left            =   2280
            TabIndex        =   79
            Top             =   4050
            Width           =   150
         End
         Begin VB.Label XLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   3
            Left            =   2850
            TabIndex        =   78
            Top             =   4050
            Width           =   150
         End
         Begin VB.Label XLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   4
            Left            =   3360
            TabIndex        =   77
            Top             =   4050
            Width           =   150
         End
         Begin VB.Label XLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   5
            Left            =   3885
            TabIndex        =   76
            Top             =   4050
            Width           =   150
         End
         Begin VB.Label XLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   6
            Left            =   4410
            TabIndex        =   75
            Top             =   4050
            Width           =   150
         End
         Begin VB.Label XLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   7
            Left            =   4980
            TabIndex        =   74
            Top             =   4050
            Width           =   150
         End
         Begin VB.Label XLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   8
            Left            =   5490
            TabIndex        =   73
            Top             =   4050
            Width           =   150
         End
         Begin VB.Label XLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   9
            Left            =   6030
            TabIndex        =   72
            Top             =   4050
            Width           =   150
         End
         Begin VB.Label XLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   10
            Left            =   6585
            TabIndex        =   71
            Top             =   4050
            Width           =   150
         End
         Begin VB.Label YLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   1
            Left            =   1005
            TabIndex        =   70
            Top             =   3510
            Width           =   150
         End
         Begin VB.Label YLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   2
            Left            =   1005
            TabIndex        =   69
            Top             =   3135
            Width           =   150
         End
         Begin VB.Label YLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   3
            Left            =   1005
            TabIndex        =   68
            Top             =   2820
            Width           =   150
         End
         Begin VB.Label YLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   4
            Left            =   1005
            TabIndex        =   67
            Top             =   2430
            Width           =   150
         End
         Begin VB.Label YLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   5
            Left            =   1005
            TabIndex        =   66
            Top             =   2070
            Width           =   150
         End
         Begin VB.Label YLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   6
            Left            =   1005
            TabIndex        =   65
            Top             =   1710
            Width           =   150
         End
         Begin VB.Label YLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   7
            Left            =   1005
            TabIndex        =   64
            Top             =   1350
            Width           =   150
         End
         Begin VB.Label YLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   8
            Left            =   1005
            TabIndex        =   63
            Top             =   1005
            Width           =   150
         End
         Begin VB.Label YLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   9
            Left            =   1005
            TabIndex        =   62
            Top             =   675
            Width           =   150
         End
         Begin VB.Label YLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   10
            Left            =   1005
            TabIndex        =   61
            Top             =   360
            Width           =   150
         End
         Begin VB.Label NumItensA 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10 ITENS"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   1545
            TabIndex        =   60
            Top             =   4530
            Width           =   585
         End
         Begin VB.Label NumItensB 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10 ITENS"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   2940
            TabIndex        =   59
            Top             =   4500
            Width           =   585
         End
         Begin VB.Label NumItensC 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10 ITENS"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   4695
            TabIndex        =   58
            Top             =   4515
            Width           =   585
         End
         Begin VB.Label EixoXLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "% ITENS"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   7725
            TabIndex        =   57
            Top             =   3825
            Width           =   540
         End
         Begin VB.Label EixoYLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "% R$"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   1470
            TabIndex        =   56
            Top             =   225
            Width           =   330
         End
         Begin VB.Label LabelA 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   1845
            TabIndex        =   55
            Top             =   3195
            Width           =   135
         End
         Begin VB.Label LabelB 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "B"
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
            Left            =   2745
            TabIndex        =   54
            Top             =   2640
            Width           =   135
         End
         Begin VB.Label LabelC 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "C"
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
            Left            =   3990
            TabIndex        =   53
            Top             =   2160
            Width           =   135
         End
         Begin VB.Line EixoX 
            X1              =   -12.504
            X2              =   106.138
            Y1              =   7.415
            Y2              =   7.415
         End
         Begin VB.Line EixoY 
            X1              =   -7.037
            X2              =   -7.037
            Y1              =   -2.773
            Y2              =   105.22
         End
         Begin VB.Line LinhaItens 
            X1              =   -7.857
            X2              =   86.728
            Y1              =   -6.034
            Y2              =   -6.034
         End
         Begin VB.Line TickLIA 
            X1              =   16.746
            X2              =   14.832
            Y1              =   -4.404
            Y2              =   -7.664
         End
         Begin VB.Line TickLIB 
            X1              =   44.63
            X2              =   42.716
            Y1              =   -4.404
            Y2              =   -7.664
         End
         Begin VB.Line TickLIC 
            X1              =   85.635
            X2              =   83.721
            Y1              =   -4.404
            Y2              =   -7.664
         End
         Begin VB.Label PercItensA 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10% "
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   1725
            TabIndex        =   52
            Top             =   4245
            Width           =   285
         End
         Begin VB.Label PercItensB 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10%"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   3075
            TabIndex        =   51
            Top             =   4245
            Width           =   255
         End
         Begin VB.Label PercItensC 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10%"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   4950
            TabIndex        =   50
            Top             =   4245
            Width           =   255
         End
         Begin VB.Line LinhaDemandas 
            X1              =   -20.979
            X2              =   -20.979
            Y1              =   108.887
            Y2              =   5.784
         End
         Begin VB.Line TickLDC 
            X1              =   -22.346
            X2              =   -19.339
            Y1              =   109.295
            Y2              =   109.295
         End
         Begin VB.Line TickLDB 
            X1              =   -22.346
            X2              =   -19.339
            Y1              =   83.213
            Y2              =   83.213
         End
         Begin VB.Line TickLDA 
            X1              =   -22.346
            X2              =   -19.339
            Y1              =   54.279
            Y2              =   54.279
         End
         Begin VB.Line TickLDBegin 
            X1              =   -22.892
            X2              =   -19.885
            Y1              =   5.377
            Y2              =   5.377
         End
         Begin VB.Label PercDemandaC 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10% "
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   315
            TabIndex        =   49
            Top             =   615
            Width           =   285
         End
         Begin VB.Label PercDemandaB 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10%"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   345
            TabIndex        =   48
            Top             =   1530
            Width           =   255
         End
         Begin VB.Label PercDemandaA 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10%"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   435
            TabIndex        =   47
            Top             =   2985
            Width           =   255
         End
         Begin VB.Image XArrow 
            Height          =   135
            Left            =   7470
            Picture         =   "ClassificacaoABC.ctx":0000
            Stretch         =   -1  'True
            Top             =   3900
            Width           =   165
         End
         Begin VB.Image YArrow 
            Height          =   120
            Left            =   1215
            Picture         =   "ClassificacaoABC.ctx":13042
            Stretch         =   -1  'True
            Top             =   255
            Width           =   120
         End
         Begin VB.Label DemandaC 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10000"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   345
            TabIndex        =   46
            Top             =   900
            Width           =   375
         End
         Begin VB.Label DemandaB 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10000"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   360
            TabIndex        =   45
            Top             =   1815
            Width           =   375
         End
         Begin VB.Label DemandaA 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "10000"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   330
            TabIndex        =   44
            Top             =   3255
            Width           =   375
         End
         Begin VB.Line TickLIBegin 
            X1              =   -5.67
            X2              =   -7.584
            Y1              =   -4.811
            Y2              =   -8.071
         End
         Begin VB.Label ZeroLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   1095
            TabIndex        =   43
            Top             =   4035
            Width           =   75
         End
         Begin VB.Line YTick 
            Index           =   1
            X1              =   -7.857
            X2              =   -5.944
            Y1              =   17.195
            Y2              =   17.195
         End
         Begin VB.Line YTick 
            Index           =   2
            X1              =   -7.857
            X2              =   -5.944
            Y1              =   27.791
            Y2              =   27.791
         End
         Begin VB.Line YTick 
            Index           =   3
            X1              =   -8.404
            X2              =   -6.49
            Y1              =   35.941
            Y2              =   35.941
         End
         Begin VB.Line YTick 
            Index           =   4
            X1              =   -7.857
            X2              =   -5.944
            Y1              =   46.944
            Y2              =   46.944
         End
         Begin VB.Line YTick 
            Index           =   5
            X1              =   -8.131
            X2              =   -6.217
            Y1              =   56.725
            Y2              =   56.725
         End
         Begin VB.Line YTick 
            Index           =   6
            X1              =   -7.584
            X2              =   -5.67
            Y1              =   66.505
            Y2              =   66.505
         End
         Begin VB.Line YTick 
            Index           =   7
            X1              =   -7.857
            X2              =   -5.944
            Y1              =   76.286
            Y2              =   76.286
         End
         Begin VB.Line YTick 
            Index           =   8
            X1              =   -8.131
            X2              =   -6.217
            Y1              =   85.659
            Y2              =   85.659
         End
         Begin VB.Line YTick 
            Index           =   9
            X1              =   -7.584
            X2              =   -5.67
            Y1              =   94.624
            Y2              =   94.624
         End
         Begin VB.Line YTick 
            Index           =   10
            X1              =   -7.857
            X2              =   -5.944
            Y1              =   103.182
            Y2              =   103.182
         End
         Begin VB.Line XTick 
            Index           =   1
            X1              =   3.351
            X2              =   3.351
            Y1              =   5.676
            Y2              =   8.637
         End
         Begin VB.Line XTick 
            Index           =   2
            X1              =   13.192
            X2              =   13.192
            Y1              =   5.676
            Y2              =   8.637
         End
         Begin VB.Line XTick 
            Index           =   3
            X1              =   23.034
            X2              =   23.034
            Y1              =   5.784
            Y2              =   9.045
         End
         Begin VB.Line XTick 
            Index           =   4
            X1              =   32.328
            X2              =   32.328
            Y1              =   5.676
            Y2              =   8.637
         End
         Begin VB.Line XTick 
            Index           =   5
            X1              =   42.169
            X2              =   42.169
            Y1              =   5.676
            Y2              =   8.637
         End
         Begin VB.Line XTick 
            Index           =   6
            X1              =   52.011
            X2              =   52.011
            Y1              =   5.676
            Y2              =   8.637
         End
         Begin VB.Line XTick 
            Index           =   7
            X1              =   61.852
            X2              =   61.852
            Y1              =   5.676
            Y2              =   8.637
         End
         Begin VB.Line XTick 
            Index           =   8
            X1              =   71.42
            X2              =   71.42
            Y1              =   5.676
            Y2              =   8.637
         End
         Begin VB.Line XTick 
            Index           =   9
            X1              =   81.261
            X2              =   81.261
            Y1              =   5.676
            Y2              =   8.637
         End
         Begin VB.Line XTick 
            Index           =   10
            X1              =   91.102
            X2              =   91.102
            Y1              =   5.676
            Y2              =   8.637
         End
         Begin VB.Label Itens 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "ITENS"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   705
            TabIndex        =   42
            Top             =   4260
            Width           =   405
         End
         Begin VB.Label Demanda 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "DEMANDA"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   165
            TabIndex        =   41
            Top             =   30
            Width           =   750
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7185
      ScaleHeight     =   495
      ScaleWidth      =   2145
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   30
      Width           =   2205
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ClassificacaoABC.ctx":26084
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ClassificacaoABC.ctx":261DE
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "ClassificacaoABC.ctx":26368
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ClassificacaoABC.ctx":2689A
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5325
      Left            =   135
      TabIndex        =   81
      Top             =   360
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   9393
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Classificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Curva ABC"
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
Attribute VB_Name = "ClassificacaoABC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iFrameAtual As Integer

Private WithEvents objEventoTipoProduto As AdmEvento
Attribute objEventoTipoProduto.VB_VarHelpID = -1

'Const LOCAIS. NÃO devem subir
Const FOLGA_GRAFICO_PERC As Double = 0.15 'Percentagem de folga do grafico em relacao ao PICTURE
Const FOLGA_CURVA_PERC As Double = 0.1 'Percentagem de folga da curva em relacao aos eixos
Const NUM_TICKS = 10  'Número de ticks - se alterar exclui/inclui ticks no form
Const PERC_MAX = 100 'Percentagem máxima - não deve ser alterado
Const TAMANHO_TICKX = 3
Const TAMANHO_TICKY = 2

'Constantes públicas dos tabs
Private Const TAB_Classificacao = 1
Private Const TAB_CurvaABC = 2

Function CurvaABC_Exibe(objCurvaABC As ClassCurvaABC) As Long
'Exibe curva ABC na tela
    
Dim lErro As Long
Dim dFolgaGraficoX As Double
Dim dFolgaGraficoY As Double
Dim dFolgaCurvaX As Double
Dim dFolgaCurvaY As Double

On Error GoTo Erro_CurvaABC_Exibe
    
    'Limpa gráfico ABC gerado anteriormente
    Call Limpa_Tela_GraficoABC
       
    'Se não tem produtos nas 3 classes não exibe curva
    If objCurvaABC.objPontoClasseA.dX = 0 Then Error 25394
    If objCurvaABC.objPontoClasseB.dX = 0 Then Error 25395
    
    'Define folgas em unidade de pontos
    dFolgaGraficoY = PERC_MAX * FOLGA_GRAFICO_PERC
    dFolgaCurvaY = PERC_MAX * FOLGA_CURVA_PERC
    dFolgaGraficoX = (PictureABC.Height / PictureABC.Width) * 1.5 * dFolgaGraficoY
    dFolgaCurvaX = (PictureABC.Height / PictureABC.Width) * dFolgaCurvaY
    
    'Define escalas de modo a ter projeção da curva nos eixos com 100 unidades
    PictureABC.ScaleHeight = -PERC_MAX - dFolgaCurvaY - 3 * dFolgaGraficoY
    PictureABC.ScaleWidth = PERC_MAX + dFolgaCurvaX + 3 * dFolgaGraficoX
    
    'Origem (0,0) fica sendo a interseção dos eixos
    PictureABC.ScaleTop = -PictureABC.ScaleHeight - 2 * dFolgaGraficoY
    PictureABC.ScaleLeft = -2 * dFolgaGraficoX
    
    'Exibe eixos e origem
    lErro = CurvaABC_ExibeEixos(dFolgaGraficoX, dFolgaGraficoY, dFolgaCurvaX, dFolgaCurvaY)
    If lErro <> SUCESSO Then Error 25396
    
    'Exibe linha de demandas
    lErro = CurvaABC_ExibeLinhaDemandas(objCurvaABC, dFolgaGraficoX)
    If lErro <> SUCESSO Then Error 25397
    
    'Exibe linha de Itens
    lErro = CurvaABC_ExibeLinhaItens(objCurvaABC, dFolgaGraficoY)
    If lErro <> SUCESSO Then Error 25398

    'Exibe curva e labels A, B, C
    lErro = CurvaABC_ExibeCurva(objCurvaABC)
    If lErro <> SUCESSO Then Error 25399
           
    CurvaABC_Exibe = SUCESSO
    
    Exit Function

Erro_CurvaABC_Exibe:
 
    CurvaABC_Exibe = Err
    
    Select Case Err
       
        Case 25394
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_CLASSEB", Err, objCurvaABC.lClassifABC)
       
        Case 25395
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_CLASSEC", Err, objCurvaABC.lClassifABC)
            
        Case 25396, 25397, 25398, 25399  'tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150782)

    End Select

    Exit Function

End Function

Function CurvaABC_ExibeCurva(objCurvaABC As ClassCurvaABC) As Long
'Exibe linha de itens do gráfico da curva ABC na tela
'CHAMADA POR CurvaABC_Exibe

Dim lErro As Long
Dim lCorCurva As Long 'cor da curva
Dim iIndice As Integer
Dim iPontoMeioClasseA As Integer 'ponto (em pontos da curva) correspondente ao meio da classe A
Dim iPontoMeioClasseB As Integer 'ponto (em pontos da curva) correspondente ao meio da classe B
Dim iPontoMeioClasseC As Integer 'ponto (em pontos da curva) correspondente ao meio da classe C

On Error GoTo Erro_CurvaABC_ExibeCurva
    
    'Exibe verticais A, B, C
    PictureABC.DrawStyle = vbDot
    PictureABC.DrawWidth = 1
    
    PictureABC.Line (objCurvaABC.objPontoClasseA.dX, 0)-(objCurvaABC.objPontoClasseA.dX, objCurvaABC.objPontoClasseA.dY)
    PictureABC.Line (objCurvaABC.objPontoClasseB.dX, 0)-(objCurvaABC.objPontoClasseB.dX, objCurvaABC.objPontoClasseB.dY)
    PictureABC.Line (100, 0)-(100, 100)
    
    'Exibe horizontais A, B, C
    PictureABC.Line (0, objCurvaABC.objPontoClasseA.dY)-(objCurvaABC.objPontoClasseA.dX, objCurvaABC.objPontoClasseA.dY)
    PictureABC.Line (0, objCurvaABC.objPontoClasseB.dY)-(objCurvaABC.objPontoClasseB.dX, objCurvaABC.objPontoClasseB.dY)
    PictureABC.Line (0, 100)-(100, 100)
    
    'Exibe pontos da curva com cor dada por QBColor
    lCorCurva = QBColor(1)
    PictureABC.DrawStyle = vbSolid
    PictureABC.DrawWidth = 2
    
    For iIndice = 1 To objCurvaABC.ColPontos.Count - 1
            
        PictureABC.Line (objCurvaABC.ColPontos(iIndice).dX, objCurvaABC.ColPontos(iIndice).dY)-(objCurvaABC.ColPontos(iIndice + 1).dX, objCurvaABC.ColPontos(iIndice + 1).dY), lCorCurva
    
    Next
    
    'Exibe labels A, B, C
    
    'Visibilidade
    LabelA.Visible = True
    LabelB.Visible = True
    LabelC.Visible = True
    
    With objCurvaABC
        
        'perc de produtos classe A divide por 2 (e por 100 para dar percent decimal) e multiplica por número de pontos menos dois extremos, soma 1 (primeiro extremo)
        iPontoMeioClasseA = 1 + Int((.objPontoClasseA.dX / 2) / 100 * (.ColPontos.Count - 1))
        LabelA.Top = 0.5 * .ColPontos(iPontoMeioClasseA).dY
        LabelA.Top = LabelA.Top + LabelA.Height / 2
        
        If (LabelA.Top - LabelA.Height) < 0 Then
            LabelA.Top = 1.2 * LabelA.Height
        End If
        
        LabelA.Left = .objPontoClasseA.dX / 2
        LabelA.Left = LabelA.Left - LabelA.Width / 2
        
        '[perc de produtos classe A + metade (perc produtos classe B)] vezes núm pontos menos dois extremos, soma 1 (primeiro extremo)
        iPontoMeioClasseB = 1 + Int((.objPontoClasseA.dX + (.objPontoClasseB.dX - .objPontoClasseA.dX) / 2) / 100 * (.ColPontos.Count - 1))
        LabelB.Top = 0.5 * .ColPontos(iPontoMeioClasseB).dY
        LabelB.Top = LabelB.Top + LabelB.Height / 2
        
        If (LabelB.Top - LabelB.Height) < 0 Then
            LabelB.Top = LabelB.Height
        End If
        
        LabelB.Left = .objPontoClasseA.dX + (.objPontoClasseB.dX - .objPontoClasseA.dX) / 2
        LabelB.Left = LabelB.Left - LabelB.Width / 2
    
        '[perc produtos classe A + perc produtos classe B + metade (perc produtos classe C)] vezes núm pontos menos dois extremos, soma 1 (primeiro extremo)
        iPontoMeioClasseC = 1 + Int((.objPontoClasseB.dX + (PERC_MAX - .objPontoClasseB.dX) / 2) / 100 * (.ColPontos.Count - 1))
        LabelC.Top = 0.5 * .ColPontos(iPontoMeioClasseC).dY
        LabelC.Top = LabelC.Top + LabelC.Height / 2
        LabelC.Left = .objPontoClasseB.dX + (PERC_MAX - .objPontoClasseB.dX) / 2
        LabelC.Left = LabelC.Left - LabelC.Width / 2
    
    End With

    CurvaABC_ExibeCurva = SUCESSO
    
    Exit Function

Erro_CurvaABC_ExibeCurva:
 
    CurvaABC_ExibeCurva = Err
    
    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150783)

    End Select

    Exit Function

End Function



Function CurvaABC_ExibeLinhaItens(objCurvaABC As ClassCurvaABC, dFolgaGraficoY As Double) As Long
'Exibe linha de itens do gráfico da curva ABC na tela
'CHAMADA POR CurvaABC_Exibe

Dim lErro As Long

On Error GoTo Erro_CurvaABC_ExibeLinhaItens
    
    'Visibilidade
    LinhaItens.Visible = True
    TickLIBegin.Visible = True
    TickLIA.Visible = True
    TickLIB.Visible = True
    TickLIC.Visible = True
    PercItensA.Visible = True
    PercItensB.Visible = True
    PercItensC.Visible = True
    NumItensA.Visible = True
    NumItensB.Visible = True
    NumItensC.Visible = True
    Itens.Visible = True
    
    'Linha de Itens
    With LinhaItens
        .X1 = 0
        .Y1 = -dFolgaGraficoY
        .X2 = PERC_MAX
        .Y2 = .Y1
    End With
   
    'Label Itens
    Itens.Left = -Itens.Width - 0.5 * Itens.Height
    Itens.Top = LinhaItens.Y1 - Itens.Height / 2
   
    'Ticks e Labels de LinhaItens (LI)
    TickLIBegin.X1 = 0
    TickLIBegin.X2 = TickLIBegin.X1
    TickLIBegin.Y1 = LinhaItens.Y1 - TAMANHO_TICKX
    TickLIBegin.Y2 = LinhaItens.Y1 + TAMANHO_TICKX
    
    TickLIA.X1 = objCurvaABC.objPontoClasseA.dX
    TickLIA.X2 = TickLIA.X1
    TickLIA.Y1 = LinhaItens.Y1 - TAMANHO_TICKX
    TickLIA.Y2 = LinhaItens.Y1 + TAMANHO_TICKX
    
    PercItensA.Caption = CStr(Int(TickLIA.X1)) & "%"
    PercItensA.Left = (TickLIA.X1 - PercItensA.Width) / 2
    PercItensA.Top = LinhaItens.Y1 + PercItensA.Height
    
    NumItensA.Caption = "(" & CStr(objCurvaABC.lItensA) & ")"
    NumItensA.Left = (TickLIA.X1 - NumItensA.Width) / 2
    NumItensA.Top = LinhaItens.Y1 - NumItensA.Height / 2
    
    TickLIB.X1 = objCurvaABC.objPontoClasseB.dX
    TickLIB.X2 = TickLIB.X1
    TickLIB.Y1 = LinhaItens.Y1 - TAMANHO_TICKX
    TickLIB.Y2 = LinhaItens.Y1 + TAMANHO_TICKX
    
    PercItensB.Caption = CStr(Int(TickLIB.X1 - TickLIA.X1)) & "%"
    PercItensB.Left = TickLIA.X1 + ((TickLIB.X1 - TickLIA.X1) - PercItensB.Width) / 2
    PercItensB.Top = LinhaItens.Y1 + PercItensB.Height
    
    NumItensB.Caption = "(" & CStr(objCurvaABC.lItensB) & ")"
    NumItensB.Top = LinhaItens.Y1 - NumItensB.Height / 2
    NumItensB.Left = TickLIA.X1 + ((TickLIB.X1 - TickLIA.X1) - NumItensB.Width) / 2
    
    TickLIC.X1 = 100
    TickLIC.X2 = TickLIC.X1
    TickLIC.Y1 = LinhaItens.Y1 - TAMANHO_TICKX
    TickLIC.Y2 = LinhaItens.Y1 + TAMANHO_TICKX
    
    PercItensC.Caption = CStr(100 - Int(TickLIA.X1) - Int(TickLIB.X1 - TickLIA.X1)) & "%"
    PercItensC.Left = TickLIB.X1 + ((TickLIC.X1 - TickLIB.X1) - PercItensC.Width) / 2
    PercItensC.Top = LinhaItens.Y1 + PercItensC.Height
    
    NumItensC.Caption = "(" & CStr(objCurvaABC.lItensC) & ")"
    NumItensC.Top = LinhaItens.Y1 - NumItensC.Height / 2
    NumItensC.Left = TickLIB.X1 + ((TickLIC.X1 - TickLIB.X1) - NumItensC.Width) / 2

    CurvaABC_ExibeLinhaItens = SUCESSO
    
    Exit Function

Erro_CurvaABC_ExibeLinhaItens:
 
    CurvaABC_ExibeLinhaItens = Err
    
    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150784)

    End Select

    Exit Function

End Function

Function CurvaABC_ExibeLinhaDemandas(objCurvaABC As ClassCurvaABC, dFolgaGraficoX As Double) As Long
'Exibe linha de demanda do gráfico da curva ABC na tela
'CHAMADA POR CurvaABC_Exibe
    
Dim lErro As Long

On Error GoTo Erro_CurvaABC_ExibeLinhaDemandas
    
    'Visibilidade
    LinhaDemandas.Visible = True
    TickLDA.Visible = True
    TickLDB.Visible = True
    TickLDC.Visible = True
    TickLDBegin.Visible = True
    PercDemandaA.Visible = True
    PercDemandaB.Visible = True
    PercDemandaC.Visible = True
    DemandaA.Visible = True
    DemandaB.Visible = True
    DemandaC.Visible = True
    Demanda.Visible = True
    
    'Linha Demandas
    With LinhaDemandas
        .X1 = -dFolgaGraficoX
        .Y1 = 0
        .X2 = .X1
        .Y2 = PERC_MAX
    End With
    
    'Label Demanda
    Demanda.Left = LinhaDemandas.X1 - Demanda.Width / 2
    Demanda.Top = LinhaDemandas.Y2 + 1.5 * Demanda.Height
    
    'Ticks e Labels de LinhaDemandas (LD)
    TickLDBegin.X1 = LinhaDemandas.X1 - TAMANHO_TICKY
    TickLDBegin.X2 = LinhaDemandas.X1 + TAMANHO_TICKY
    TickLDBegin.Y1 = 0
    TickLDBegin.Y2 = TickLDBegin.Y1
    
    TickLDA.X1 = LinhaDemandas.X1 - TAMANHO_TICKY
    TickLDA.X2 = LinhaDemandas.X1 + TAMANHO_TICKY
    TickLDA.Y1 = objCurvaABC.objPontoClasseA.dY
    TickLDA.Y2 = TickLDA.Y1
    
    PercDemandaA.Caption = CStr(objCurvaABC.iFaixaA) & "%"
    PercDemandaA.Left = LinhaDemandas.X1 - 1.2 * PercDemandaA.Width
    PercDemandaA.Top = (TickLDA.Y1 + PercDemandaA.Height) / 2
    
    DemandaA.Caption = "(" & Format((objCurvaABC.objPontoClasseA.dY / 100 * objCurvaABC.dDemandaTotal), "Standard") & ")"
    
    If PercDemandaA.Height + 4 * DemandaA.Height > TickLDA.Y1 Then
        
        DemandaA.Visible = False
    
    Else
            
        DemandaA.Visible = True
        DemandaA.Top = PercDemandaA.Top - PercDemandaA.Height - DemandaA.Height / 2
        DemandaA.Left = LinhaDemandas.X1 - DemandaA.Width / 2
    
    End If
    
    TickLDB.X1 = LinhaDemandas.X1 - TAMANHO_TICKY
    TickLDB.X2 = LinhaDemandas.X1 + TAMANHO_TICKY
    TickLDB.Y1 = objCurvaABC.objPontoClasseB.dY
    TickLDB.Y2 = TickLDB.Y1
    
    PercDemandaB.Caption = CStr(objCurvaABC.iFaixaB) & "%"
    PercDemandaB.Left = LinhaDemandas.X1 - 1.2 * PercDemandaB.Width
    PercDemandaB.Top = TickLDA.Y1 + ((TickLDB.Y1 - TickLDA.Y1) + PercDemandaB.Height) / 2
    
    DemandaB.Caption = "(" & Format((objCurvaABC.objPontoClasseB.dY - objCurvaABC.objPontoClasseA.dY) / 100 * objCurvaABC.dDemandaTotal, "Standard") & ")"
    
    If PercDemandaB.Height + 4 * DemandaB.Height > TickLDB.Y1 - TickLDA.Y1 Then
        
        DemandaB.Visible = False
    
    Else
            
        DemandaB.Visible = True
        DemandaB.Top = PercDemandaB.Top - PercDemandaB.Height - DemandaB.Height / 2
        DemandaB.Left = LinhaDemandas.X1 - DemandaB.Width / 2
    
    End If
    
    TickLDC.X1 = LinhaDemandas.X1 - TAMANHO_TICKY
    TickLDC.X2 = LinhaDemandas.X1 + TAMANHO_TICKY
    TickLDC.Y1 = 100
    TickLDC.Y2 = TickLDC.Y1
    
    PercDemandaC.Caption = CStr(100 - objCurvaABC.iFaixaA - objCurvaABC.iFaixaB) & "%"
    PercDemandaC.Left = LinhaDemandas.X1 - 1.2 * PercDemandaC.Width
    PercDemandaC.Top = TickLDB.Y1 + ((TickLDC.Y1 - TickLDB.Y1) + PercDemandaC.Height) / 2
    
    DemandaC.Caption = "(" & Format((1 - (objCurvaABC.objPontoClasseB.dY / 100)) * objCurvaABC.dDemandaTotal, "Standard") & ")"
    
    If PercDemandaC.Height + 4 * DemandaC.Height > TickLDC.Y1 - TickLDB.Y1 Then
        
        DemandaC.Visible = False
    
    Else
            
        DemandaC.Visible = True
        DemandaC.Top = PercDemandaC.Top - PercDemandaC.Height - DemandaC.Height / 2
        DemandaC.Left = LinhaDemandas.X1 - DemandaC.Width / 2
    
    End If
    
    'Alinhamento de percentuais de demanda
    If PercDemandaA.Left <= PercDemandaB.Left And PercDemandaA.Left <= PercDemandaC.Left Then
        PercDemandaB.Left = PercDemandaA.Left
        PercDemandaC.Left = PercDemandaA.Left
    ElseIf PercDemandaB.Left <= PercDemandaA.Left And PercDemandaB.Left <= PercDemandaC.Left Then
        PercDemandaA.Left = PercDemandaB.Left
        PercDemandaC.Left = PercDemandaB.Left
    Else
        PercDemandaA.Left = PercDemandaC.Left
        PercDemandaB.Left = PercDemandaC.Left
    End If

    CurvaABC_ExibeLinhaDemandas = SUCESSO
    
    Exit Function

Erro_CurvaABC_ExibeLinhaDemandas:
 
    CurvaABC_ExibeLinhaDemandas = Err
    
    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150785)

    End Select

    Exit Function

End Function


Function CurvaABC_ExibeEixos(dFolgaGraficoX As Double, dFolgaGraficoY As Double, dFolgaCurvaX As Double, dFolgaCurvaY As Double) As Long
'Exibe eixos do gráfico da curva ABC na tela e origem
'CHAMADA POR CurvaABC_Exibe
    
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_CurvaABC_ExibeEixos
    
    'Visibilidade de X e Y ticks e seus labels
    For iIndice = 1 To NUM_TICKS
        XTick(iIndice).Visible = True
        YTick(iIndice).Visible = True
        XLabel(iIndice).Visible = True
        YLabel(iIndice).Visible = True
    Next
    
    'Eixo X
    EixoX.X1 = -0.5 * dFolgaGraficoX
    EixoX.Y1 = 0
    EixoX.X2 = PERC_MAX + dFolgaCurvaX
    EixoX.Y2 = 0
    
    'Eixo Y
    EixoY.X1 = 0
    EixoY.Y1 = -0.5 * dFolgaGraficoY
    EixoY.X2 = 0
    EixoY.Y2 = PERC_MAX + dFolgaCurvaY
    
    'Origem
    ZeroLabel.Top = -ZeroLabel.Height / 3
    ZeroLabel.Left = -2.5 * ZeroLabel.Width
    
    'Setas X e Y
    XArrow.Top = EixoX.Y1 + (XArrow.Height / 2) - 0.5
    XArrow.Left = EixoX.X2
    YArrow.Left = EixoY.X1 - (YArrow.Width / 2) + 0.2
    YArrow.Top = EixoY.Y2 + YArrow.Height - 0.4
            
    'Labels dos Eixos X e Y
    EixoXLabel.Left = (XArrow.Left + XArrow.Width) + (dFolgaGraficoX - XArrow.Width - EixoXLabel.Width) / 2
    EixoXLabel.Top = EixoX.Y1 + EixoXLabel.Height / 2
    EixoYLabel.Left = EixoY.X1 - EixoYLabel.Width / 2
    EixoYLabel.Top = YArrow.Top + EixoYLabel.Height + (dFolgaGraficoY - YArrow.Height - EixoYLabel.Height) / 2
    
    'X e Y ticks e seus labels
    For iIndice = 1 To NUM_TICKS
        XTick(iIndice).X1 = CInt(iIndice * (100 / NUM_TICKS))
        XTick(iIndice).X2 = XTick(iIndice).X1
        XTick(iIndice).Y1 = -TAMANHO_TICKX / 2
        XTick(iIndice).Y2 = TAMANHO_TICKX / 2
        
        YTick(iIndice).X1 = -TAMANHO_TICKY / 2
        YTick(iIndice).X2 = TAMANHO_TICKY / 2
        YTick(iIndice).Y1 = CInt(iIndice * (100 / NUM_TICKS))
        YTick(iIndice).Y2 = YTick(iIndice).Y1
        
        XLabel(iIndice).Caption = CStr(CInt(XTick(iIndice).X1))
        XLabel(iIndice).Left = XTick(iIndice).X1 - XLabel(iIndice).Width / 2
        XLabel(iIndice).Top = XTick(iIndice).Y1
        
        YLabel(iIndice).Caption = CStr(CInt(YTick(iIndice).Y1))
        YLabel(iIndice).Left = YTick(iIndice).X1 - 1.2 * YLabel(iIndice).Width
        YLabel(iIndice).Top = YTick(iIndice).Y1 + YLabel(iIndice).Height / 2
    
    Next
    
    CurvaABC_ExibeEixos = SUCESSO
    
    Exit Function

Erro_CurvaABC_ExibeEixos:
 
    CurvaABC_ExibeEixos = Err
    
    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150786)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long
'Calcula e grava classif ABC

Dim lErro As Long
Dim objClassABC As New ClassClassificacaoABC
Dim objCurvaABC As New ClassCurvaABC

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Critica preenchimento dos campos
    If Len(Trim(Codigo.Text)) = 0 Then Error 25337
    If Len(Trim(Data.ClipText)) = 0 Then Error 25342
    
    If TodosTipos.Value = vbUnchecked And Len(Trim(Tipo.Text)) = 0 Then Error 25325
    
    If Len(Trim(MesInicial.Text)) = 0 Then Error 25326
    If Len(Trim(MesFinal.Text)) = 0 Then Error 25327
    If Len(Trim(AnoInicial.Text)) = 0 Then Error 25328
    If Len(Trim(AnoFinal.Text)) = 0 Then Error 25329
    
    'Verifica se mês/ano inicial <= mês/ano final
    If CInt(AnoInicial.Text) > CInt(AnoFinal.Text) Then Error 25400
    If CInt(AnoInicial.Text) = CInt(AnoFinal.Text) And CInt(MesInicial.Text) > CInt(MesFinal.Text) Then Error 25401
    
    'Verifica se
    If Len(Trim(FaixaA.Text)) = 0 Then Error 25330
    If Len(Trim(FaixaB.Text)) = 0 Then Error 25331
    
    'Passa dados para objClassABC
    objClassABC.iFilialEmpresa = giFilialEmpresa
    objClassABC.sCodigo = Trim(Codigo.Text)
    objClassABC.dtData = CDate(Data.Text)
    objClassABC.sDescricao = Trim(Descricao.Text)
    objClassABC.iAtualizaProdutosFilial = AtualizaProdutosFilial.Value
    
    If TodosTipos.Value = vbChecked Then
        objClassABC.iTipoProduto = TODOS_TIPOS
    Else
        objClassABC.iTipoProduto = CInt(Tipo.Text)
    End If
    
    objClassABC.iMesInicial = CInt(MesInicial.Text)
    objClassABC.iMesFinal = CInt(MesFinal.Text)
    objClassABC.iAnoInicial = CInt(AnoInicial.Text)
    objClassABC.iAnoFinal = CInt(AnoFinal.Text)
    objClassABC.iFaixaA = CInt(FaixaA.Text)
    objClassABC.iFaixaB = CInt(FaixaB.Text)

    'Calcula e grava no BD a classif ABC
    lErro = CF("ClassificacaoABC_Grava",objClassABC)
    If lErro <> SUCESSO Then Error 25332
    
    'Exibe as Faixas ajustadas depois da classif ABC
    FaixaA.PromptInclude = False
    FaixaA.Text = CStr(objClassABC.iFaixaA)
    FaixaA.PromptInclude = True
    FaixaB.PromptInclude = False
    FaixaB.Text = CStr(objClassABC.iFaixaB)
    FaixaB.PromptInclude = True
    FaixaC.Caption = CStr(100 - objClassABC.iFaixaA - objClassABC.iFaixaB) & "%"
    
    'Atualiza Lista de Classificações
    Call Classificacoes_Remove(objClassABC)
    Call Classificacoes_Adiciona(objClassABC)
        
    'Le os pontos da classificacao ABC
    lErro = CF("CurvaABC_LePontos",objClassABC, objCurvaABC)
    If lErro <> SUCESSO Then Error 25379
    
    'Exibe a curva da classificação ABC
    lErro = CurvaABC_Exibe(objCurvaABC)
    If lErro <> SUCESSO Then Error 25386
        
    'Esconde o frame atual, mostra o Frame da Curva ABC
    Frame1(iFrameAtual).Visible = False
    iFrameAtual = TabStrip1.Tabs.Count
    TabStrip1.Tabs(iFrameAtual).Selected = True
    Frame1(iFrameAtual).Visible = True
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function

Erro_Gravar_Registro:
 
    Gravar_Registro = Err
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
            
        Case 25325
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_TIPO_PRODUTO", Err)
            
        Case 25326
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_MES_INICIAL", Err)
                        
        Case 25327
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_MES_FINAL", Err)
        
        Case 25328
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_ANO_INICIAL", Err)
        
        Case 25329
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_ANO_FINAL", Err)
        
        Case 25330
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_FAIXA_A", Err)
        
        Case 25331
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_FAIXA_B", Err)
            
        Case 25332
        
        Case 25379, 25386 'Tratado na rotina chamada
            Call Limpa_Tela_GraficoABC
            
        Case 25337
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_CODIGO_CLASSIFABC", Err)
        
        Case 25342
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_DATA_CLASSIFABC", Err)
        
        Case 25400
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ANOINIC_MAIOR_ANOFINAL", Err)
        
        Case 25401
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MESINIC_MAIOR_MESFINAL", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150787)

    End Select

    Exit Function

End Function

Private Sub Classificacoes_Remove(objClassABC As ClassClassificacaoABC)
'Percorre a Lista de Classificações para remover a Classificação caso ela exista

Dim iIndice As Integer

For iIndice = 0 To Classificacoes.ListCount - 1

    If Classificacoes.ItemData(iIndice) = objClassABC.lNumInt Then
        Classificacoes.RemoveItem iIndice
        Exit For

    End If

Next

End Sub

Private Sub Classificacoes_Adiciona(objClassABC As ClassClassificacaoABC)
'Inclui Classificação na Lista

    Classificacoes.AddItem objClassABC.sCodigo & SEPARADOR & objClassABC.sDescricao
    Classificacoes.ItemData(Classificacoes.NewIndex) = objClassABC.lNumInt

End Sub

Private Sub AnoFinal_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub AnoFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(AnoFinal, iAlterado)

End Sub

Private Sub AnoFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iAno As Integer

On Error GoTo Erro_AnoFinal_Validate

    'Verifica se o Ano Final foi preenchido
    If Len(Trim(AnoFinal.ClipText)) = 0 Then Exit Sub
    
    iAno = CInt(AnoFinal.Text)
    
    If iAno < ANO_VALIDO Then Error 43480
    
    Exit Sub

Erro_AnoFinal_Validate:

    Cancel = True


    Select Case Err

        Case 43480
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ANO_INVALIDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150788)

    End Select

    Exit Sub
    
End Sub

Private Sub AnoInicial_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub AnoInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(AnoInicial, iAlterado)

End Sub

Private Sub AnoInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iAno As Integer

On Error GoTo Erro_AnoInicial_Validate

    'Verifica se o Ano Inicial foi preenchido
    If Len(Trim(AnoInicial.ClipText)) = 0 Then Exit Sub
    
    iAno = CInt(AnoInicial.Text)
    
    If iAno < ANO_VALIDO Then Error 43479
    
    Exit Sub

Erro_AnoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 43479
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ANO_INVALIDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150789)

    End Select

    Exit Sub

End Sub

Private Sub AtualizaProdutosFilial_Click()

    iAlterado = REGISTRO_ALTERADO
     
End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objClassABC As New ClassClassificacaoABC
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se Código da ClassifABC foi preenchido
    If Len(Trim(Codigo)) = 0 Then Error 25402

    'Preenche objClassABC
    objClassABC.iFilialEmpresa = giFilialEmpresa
    objClassABC.sCodigo = Trim(Codigo.Text)
    
    'Lê a ClassificacaoABC
    lErro = CF("ClassificacaoABC_Le",objClassABC)
    If lErro <> SUCESSO And lErro <> 43465 Then Error 25404
    
    'Se não encontrou a ClassificacaoABC --> Erro
    If lErro <> SUCESSO Then Error 25405

    'Envia aviso perguntando se realmente deseja excluir ClasseUM
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_CLASSEABC", objClassABC.sCodigo)

    If vbMsgRes = vbYes Then

        'Exclui a Classificação ABC
        lErro = CF("ClassificacaoABC_Exclui",objClassABC)
        If lErro <> SUCESSO Then Error 25403

        'Limpa a Tela
        Call Limpa_Tela_ClassificacaoABCTotal

        'Atualiza Lista de Classificações
        Call Classificacoes_Remove(objClassABC)
        
    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 25402
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_CODIGO_CLASSIFABC", Err)
            
        Case 25403, 25404 'Erro tratado na rotina chamada

        Case 25405
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLASSIFICACAOABC_INEXISTENTE1", Err, objClassABC.sCodigo, objClassABC.iFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150790)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 43485
    
    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 43485

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 150791)

    End Select

    Exit Sub

End Sub


Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 43486

    'Limpa a Tela
    Call Limpa_Tela_ClassificacaoABCTotal
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 43486

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150792)

    End Select

    Exit Sub

End Sub

Private Sub Classificacoes_DblClick()

Dim lErro As Long
Dim objClassABC As New ClassClassificacaoABC

On Error GoTo Erro_Classificacoes_DblClick

    'Guarda o valor do código da Classificação selecionada na ListBox Classificacoes
    objClassABC.lNumInt = Classificacoes.ItemData(Classificacoes.ListIndex)
    
    'Lê a ClassificacaoABC no BD
    lErro = CF("ClassificacaoABC_Le_NumInt",objClassABC)
    If lErro <> SUCESSO And lErro <> 43500 Then Error 43470

    'Se não encontrou a ClassificacaoABC --> Erro
    If lErro <> SUCESSO Then Error 43471

    'Exibe os dados da ClassificacaoABC
    lErro = Traz_ClassificacaoABC_Tela(objClassABC)
    If lErro <> SUCESSO Then Error 43472

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_Classificacoes_DblClick:

    Select Case Err
            
        Case 43470
            
        Case 43471
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLASSIFICACAOABC_INEXISTENTE2", Err, objClassABC.lNumInt)
            Classificacoes.RemoveItem (Classificacoes.ListIndex)
    
        Case 43472
            Call Limpa_Tela_GraficoABC
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150793)
    
    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'Verifica se Data foi digitada
    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(Data.Text)
    If lErro <> SUCESSO Then Error 43472

    Exit Sub

Erro_Data_Validate:

    Cancel = True


    Select Case Err

        Case 43472

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150794)

    End Select

    Exit Sub

End Sub

Private Sub DataUpDown_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntrada_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro Then Error 43473

    Exit Sub

Erro_UpDownEntrada_DownClick:

    Select Case Err

        Case 43473

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150795)

    End Select

    Exit Sub

End Sub

Private Sub DataUpDown_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_DataUpDown_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 43474

    Exit Sub

Erro_DataUpDown_UpClick:

    Select Case Err

        Case 43474

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150796)

    End Select

    Exit Sub

End Sub

Private Sub Descricao_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub FaixaA_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub FaixaA_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(FaixaA, iAlterado)

End Sub

Private Sub FaixaA_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iFaixa As Integer
Dim iFaixaA As Integer
Dim iFaixaB As Integer

On Error GoTo Erro_FaixaA_Validate
    
    If FaixaB.Tag <> "FaixaB" Then
    
        'Verifica se a Faixa A foi digitada
        If Len(Trim(FaixaA.Text)) = 0 Then Exit Sub
        
        iFaixaA = CInt(FaixaA.Text)
        
        If iFaixaA < 0 Or iFaixaA > 100 Then Error 43481
        
        'Verifica se FaixaB foi preenchida
        If Len(Trim(FaixaB.Text)) > 0 Then
            iFaixaB = CInt(FaixaB.Text)
            iFaixa = iFaixaA + iFaixaB
            If iFaixa > 100 Then Error 43482
            FaixaC.Caption = 100 - iFaixaA - iFaixaB
            FaixaC.Caption = FaixaC.Caption & "%"
        End If
    
    End If
    
    FaixaB.Tag = ""
    
    Exit Sub
    
Erro_FaixaA_Validate:

    Cancel = True


    Select Case Err

        Case 43481
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FAIXA_INVALIDA2", Err)
            FaixaA.Tag = "FaixaA"
            
            
        Case 43482
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FAIXA_MAXIMA2", Err)
            FaixaA.Tag = "FaixaA"
            FaixaC.Caption = ""
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150797)

    End Select

    Exit Sub

End Sub

Private Sub FaixaB_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub FaixaB_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(FaixaB, iAlterado)

End Sub

Private Sub FaixaB_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iFaixa As Integer
Dim iFaixaA As Integer
Dim iFaixaB As Integer

On Error GoTo Erro_FaixaB_Validate
    
    If FaixaA.Tag <> "FaixaA" Then
    
        'Verifica se a Faixa B foi digitada
        If Len(Trim(FaixaB.Text)) = 0 Then Exit Sub
        
        iFaixaB = CInt(FaixaB.Text)
        
        If iFaixaB < 0 Or iFaixaB > 100 Then Error 43483
        
        'Verifica se FaixaA foi preenchida
        If Len(Trim(FaixaA.Text)) > 0 Then
            iFaixaA = CInt(FaixaA.Text)
            iFaixa = iFaixaA + iFaixaB
            If iFaixa > 100 Then Error 43484
            FaixaC.Caption = 100 - iFaixaA - iFaixaB
            FaixaC.Caption = FaixaC.Caption & "%"
        End If
    
    End If
    
    FaixaA.Tag = ""
    
    Exit Sub
    
Erro_FaixaB_Validate:

    Cancel = True


    Select Case Err

        Case 43483
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FAIXA_INVALIDA2", Err)
            FaixaB.Tag = "FaixaB"
            
        Case 43484
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FAIXA_MAXIMA2", Err)
            FaixaB.Tag = "FaixaB"
            FaixaC.Caption = ""
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150798)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoTipoProduto = New AdmEvento

    iFrameAtual = 1
    
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    'Carrega todas as Classificações existentes para a Lista
    lErro = Carrega_ClassificacaoABC()
    If lErro <> SUCESSO Then Error 43454

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 43454

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150799)

    End Select

    iAlterado = 0
    
    Exit Sub
    
End Sub

Private Function Carrega_ClassificacaoABC() As Long
'Carrega a Lista com as Classificações existentes no BD

Dim lErro As Long
Dim iIndice As Integer
Dim colNumIntCodigo As New Collection
Dim objClassABC As ClassClassificacaoABC

On Error GoTo Erro_Carrega_ClassificacaoABC

    'Lê todas as Classificações existentes para esta FilialEmpresa
    lErro = CF("ClassificacoesABC_Le",colNumIntCodigo)
    If lErro <> SUCESSO Then Error 43455

    'Preenche a Lista das Classificações com os objetos da coleção colCodigos
    For Each objClassABC In colNumIntCodigo
       
        Classificacoes.AddItem objClassABC.sCodigo & SEPARADOR & objClassABC.sDescricao
        Classificacoes.ItemData(Classificacoes.NewIndex) = objClassABC.lNumInt
    
    Next
        
    Carrega_ClassificacaoABC = SUCESSO

    Exit Function

Erro_Carrega_ClassificacaoABC:

    Carrega_ClassificacaoABC = Err

    Select Case Err

        Case 43455

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150800)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoTipoProduto = Nothing

   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
   
End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objClassABC As New ClassClassificacaoABC

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "ClassificacaoABC"

    'Lê os dados da Tela ClassificacaoABC
    lErro = Move_Tela_Memoria(objClassABC)
    If lErro <> SUCESSO Then Error 43460

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumInt", CLng(0), 0, "NumInt"
    colCampoValor.Add "FilialEmpresa", objClassABC.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "Codigo", objClassABC.sCodigo, STRING_CLASSABC_CODIGO, "Codigo"
    colCampoValor.Add "Descricao", objClassABC.sDescricao, STRING_CLASSABC_DESCRICAO, "Descricao"
    colCampoValor.Add "Data", objClassABC.dtData, 0, "Data"
    colCampoValor.Add "MesInicial", objClassABC.iMesInicial, 0, "MesInicial"
    colCampoValor.Add "AnoInicial", objClassABC.iAnoInicial, 0, "AnoInicial"
    colCampoValor.Add "MesFinal", objClassABC.iMesFinal, 0, "MesFinal"
    colCampoValor.Add "AnoFinal", objClassABC.iAnoFinal, 0, "AnoFinal"
    colCampoValor.Add "FaixaA", objClassABC.iFaixaA, 0, "FaixaA"
    colCampoValor.Add "FaixaB", objClassABC.iFaixaB, 0, "FaixaB"
    colCampoValor.Add "TipoProduto", objClassABC.iTipoProduto, 0, "TipoProduto"
    colCampoValor.Add "DemandaTotal", objClassABC.dDemandaTotal, 0, "DesmandaTotal"
    colCampoValor.Add "AtualizaProdutosFilial", objClassABC.iAtualizaProdutosFilial, 0, "AtualizaProdutosFilial"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, objClassABC.iFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 43460

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150801)

    End Select

    Exit Sub

End Sub

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objClassABC As New ClassClassificacaoABC

On Error GoTo Erro_Tela_Preenche

    objClassABC.lNumInt = colCampoValor.Item("NumInt").vValor

    If objClassABC.lNumInt <> 0 Then

        'Carrega objClassABC com os dados passados em colCampoValor
        objClassABC.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
        objClassABC.sCodigo = colCampoValor.Item("Codigo").vValor
        objClassABC.sDescricao = colCampoValor.Item("Descricao").vValor
        objClassABC.dtData = colCampoValor.Item("Data").vValor
        objClassABC.iMesInicial = colCampoValor.Item("MesInicial").vValor
        objClassABC.iAnoInicial = colCampoValor.Item("AnoInicial").vValor
        objClassABC.iMesFinal = colCampoValor.Item("MesFinal").vValor
        objClassABC.iAnoFinal = colCampoValor.Item("AnoFinal").vValor
        objClassABC.iFaixaA = colCampoValor.Item("FaixaA").vValor
        objClassABC.iFaixaB = colCampoValor.Item("FaixaB").vValor
        objClassABC.iTipoProduto = colCampoValor.Item("TipoProduto").vValor
        objClassABC.dDemandaTotal = colCampoValor.Item("DemandaTotal").vValor
        objClassABC.iAtualizaProdutosFilial = colCampoValor.Item("AtualizaProdutosFilial").vValor

        'Traz dados do Almoxarifado para a Tela
        lErro = Traz_ClassificacaoABC_Tela(objClassABC)
        If lErro <> SUCESSO Then Error 43461

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 43461
            Call Limpa_Tela_GraficoABC

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150802)

    End Select

    Exit Sub

End Sub

Private Function Move_Tela_Memoria(objClassABC As ClassClassificacaoABC) As Long

Dim lErro As Long
Dim objTipoProduto As New ClassTipoDeProduto

On Error GoTo Erro_Move_Tela_Memoria:

    objClassABC.iFilialEmpresa = giFilialEmpresa

    If Len(Trim(Codigo.Text)) <> 0 Then
        objClassABC.sCodigo = Codigo.Text
    End If
    
    If Len(Trim(Data.ClipText)) <> 0 Then
        objClassABC.dtData = CDate(Data.Text)
    End If
    
    objClassABC.sDescricao = Descricao.Text
    
    If AtualizaProdutosFilial.Value = CLASSABC_ATUALIZA_PRODFILIAL Then
        objClassABC.iAtualizaProdutosFilial = CLASSABC_ATUALIZA_PRODFILIAL
    Else
        objClassABC.iAtualizaProdutosFilial = vbUnchecked
    End If
    
    'Verifica se algum Tipo de Produto foi selecionado
    If Len(Trim(Tipo.Text)) <> 0 Then
        'Preenche objTipoProduto
        objTipoProduto.iTipo = CInt(Tipo.Text)
        'Lê o Tipo de Produto
        lErro = CF("TipoDeProduto_Le",objTipoProduto)
        If lErro <> SUCESSO And lErro <> 1 Then Error 43466
        
        'Se não encontrou o Tipo de Produto --> Erro
        If lErro <> SUCESSO Then Error 43467
        
        objClassABC.iTipoProduto = objTipoProduto.iTipo
        
    End If
    
    If Len(Trim(MesInicial.Text)) <> 0 Then objClassABC.iMesInicial = CInt(MesInicial.Text)
    If Len(Trim(AnoInicial.Text)) <> 0 Then objClassABC.iAnoInicial = CInt(AnoInicial.Text)
    If Len(Trim(MesFinal.Text)) <> 0 Then objClassABC.iMesFinal = CInt(MesFinal.Text)
    If Len(Trim(AnoFinal.Text)) <> 0 Then objClassABC.iAnoFinal = CInt(AnoFinal.Text)
    If Len(Trim(FaixaA.Text)) <> 0 Then objClassABC.iFaixaA = CInt(FaixaA.Text)
    If Len(Trim(FaixaB.Text)) <> 0 Then objClassABC.iFaixaB = CInt(FaixaB.Text)
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err


        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150803)

    End Select

    Exit Function

End Function

Private Function Traz_ClassificacaoABC_Tela(objClassABC As ClassClassificacaoABC) As Long
'Traz os dados da ClassificacaoABC passada em objClassABC

Dim lErro As Long
Dim objCurvaABC As New ClassCurvaABC

On Error GoTo Erro_Traz_ClassificacaoABC_Tela

    'Limpa a tela ClassificaoABC
    Call Limpa_Tela_ClassificacaoABC
    
    Codigo.Text = objClassABC.sCodigo
    
    Data.Text = Format(objClassABC.dtData, "dd/mm/yy")
    
    Descricao.Text = objClassABC.sDescricao
    
    If objClassABC.iAtualizaProdutosFilial = CLASSABC_ATUALIZA_PRODFILIAL Then
        AtualizaProdutosFilial.Value = CLASSABC_ATUALIZA_PRODFILIAL
    Else
        AtualizaProdutosFilial.Value = vbUnchecked
    End If
    
    If objClassABC.iTipoProduto <> 0 Then
        TodosTipos.Value = vbUnchecked
        Tipo.Text = CStr(objClassABC.iTipoProduto)
        Call Tipo_Validate(bSGECancelDummy)
    Else
        TodosTipos.Value = vbChecked
    End If
    
    MesInicial.PromptInclude = False
    MesInicial.Text = CStr(objClassABC.iMesInicial)
    MesInicial.PromptInclude = True
    MesFinal.PromptInclude = False
    MesFinal.Text = CStr(objClassABC.iMesFinal)
    MesFinal.PromptInclude = True
    AnoInicial.Text = CStr(objClassABC.iAnoInicial)
    AnoFinal.Text = CStr(objClassABC.iAnoFinal)
    FaixaA.PromptInclude = False
    FaixaA.Text = CStr(objClassABC.iFaixaA)
    FaixaA.PromptInclude = True
    FaixaB.PromptInclude = False
    FaixaB.Text = CStr(objClassABC.iFaixaB)
    FaixaB.PromptInclude = True
    FaixaC.Caption = CStr(100 - objClassABC.iFaixaA - objClassABC.iFaixaB) & "%"
    
    'Chama CurvaABC_LePontos
    lErro = CF("CurvaABC_LePontos",objClassABC, objCurvaABC)
    If lErro <> SUCESSO Then Error 43468
    
    'Chama CurvaABC_Exibe
    lErro = CurvaABC_Exibe(objCurvaABC)
    If lErro <> SUCESSO Then Error 43469

    iAlterado = 0

    Traz_ClassificacaoABC_Tela = SUCESSO

    Exit Function

Erro_Traz_ClassificacaoABC_Tela:

    Traz_ClassificacaoABC_Tela = Err

    iAlterado = 0
    
    Select Case Err
    
        Case 43468, 43469
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150804)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_ClassificacaoABC()
'Limpa a Tela ClassificacaoABC
    
    'Chama o Limpa tela
    Call Limpa_Tela(Me)

    'Limpa os campos que não são limpos pelo Limpa_Tela
    Codigo.Text = ""
    Descricao.Text = ""
    AtualizaProdutosFilial.Value = vbUnchecked
    TodosTipos.Value = vbChecked
    TipoDescricao.Caption = ""
    FaixaC.Caption = ""

End Sub

Private Sub Limpa_Tela_ClassificacaoABCTotal()
'Limpa a Tela ClassificacaoABC
    
Dim lErro As Long
    
    'Limpa os campos da tela
    Call Limpa_Tela_ClassificacaoABC
    
    'limpa o grafico
    Call Limpa_Tela_GraficoABC
    
    'Fecha o Comando de Setas
    lErro = ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0

End Sub

Private Sub Limpa_Tela_GraficoABC()
'Limpa o Grafico ABC
    
Dim iIndice As Integer
Dim lErro As Long
    
    'X e Y ticks e seus labels
    For iIndice = 1 To NUM_TICKS
        XTick(iIndice).Visible = False
        YTick(iIndice).Visible = False
        XLabel(iIndice).Visible = False
        YLabel(iIndice).Visible = False
    Next
    
    'Demandas
    LinhaDemandas.Visible = False
    TickLDA.Visible = False
    TickLDB.Visible = False
    TickLDC.Visible = False
    TickLDBegin.Visible = False
    PercDemandaA.Visible = False
    PercDemandaB.Visible = False
    PercDemandaC.Visible = False
    DemandaA.Visible = False
    DemandaB.Visible = False
    DemandaC.Visible = False
    Demanda.Visible = False
    
    'Itens
    LinhaItens.Visible = False
    TickLIBegin.Visible = False
    TickLIA.Visible = False
    TickLIB.Visible = False
    TickLIC.Visible = False
    PercItensA.Visible = False
    PercItensB.Visible = False
    PercItensC.Visible = False
    NumItensA.Visible = False
    NumItensB.Visible = False
    NumItensC.Visible = False
    Itens.Visible = False
    
    'Labels A, B, C
    LabelA.Visible = False
    LabelB.Visible = False
    LabelC.Visible = False
    
    'Limpa curva gerada anteriormente
    PictureABC.Cls
    
End Sub

Private Sub MesFinal_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub MesFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(MesFinal, iAlterado)

End Sub

Private Sub MesFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iMes As Integer

On Error GoTo Erro_MesFinal_Validate

    'Verifica se o Mês Final foi digitado
    If Len(Trim(MesFinal.ClipText)) = 0 Then Exit Sub
    
    iMes = CInt(MesFinal.Text)
    
    If iMes < 1 Or iMes > 12 Then Error 43478
    
    Exit Sub

Erro_MesFinal_Validate:

    Cancel = True


    Select Case Err

        Case 43478
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MES_INVALIDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150805)

    End Select

    Exit Sub

End Sub

Private Sub MesInicial_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub MesInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(MesInicial, iAlterado)

End Sub

Private Sub MesInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iMes As Integer

On Error GoTo Erro_MesInicial_Validate

    'Verifica se o Mês Inicial foi digitado
    If Len(Trim(MesInicial.ClipText)) = 0 Then Exit Sub
    
    iMes = CInt(MesInicial.ClipText)
    
    If iMes < 1 Or iMes > 12 Then Error 43477
    
    Exit Sub

Erro_MesInicial_Validate:

    Cancel = True


    Select Case Err

        Case 43477
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MES_INVALIDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150806)

    End Select

    Exit Sub

End Sub

Private Sub objEventoTipoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoProduto As ClassTipoDeProduto

On Error GoTo Erro_objEventoTipoProduto_evSelecao

    Set objTipoProduto = obj1

    If objTipoProduto.iTipo <> 0 Then
        'Mostra o Tipo de Produto na tela
        Tipo.Text = objTipoProduto.iTipo
        TipoDescricao.Caption = objTipoProduto.sDescricao
            
        'Limpa Todos os Tipos
        TodosTipos.Value = vbUnchecked
    End If
    
    'Fecha o Comando de Setas
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoTipoProduto_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150807)

    End Select

    Exit Sub
    
End Sub

Private Sub TabStrip1_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
            
            'Esconde o frame atual, mostra o novo
            Frame1(TabStrip1.SelectedItem.Index).Visible = True
            Frame1(iFrameAtual).Visible = False
            'Armazena novo valor de iFrameAtual
            iFrameAtual = TabStrip1.SelectedItem.Index
    
            Select Case iFrameAtual
            
                Case TAB_Classificacao
                    Parent.HelpContextID = IDH_CLASSIFICACAO_ABC_CLASSIFICACAO
                    
                Case TAB_CurvaABC
                    Parent.HelpContextID = IDH_CLASSIFICACAO_ABC_CURVA_ABC
                            
            End Select
    
    End If
    
    Exit Sub
    
Erro_TabStrip1_Click:
    
    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150808)

    End Select

    Exit Sub

End Sub

Private Sub Tipo_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Tipo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Tipo, iAlterado)

End Sub

Private Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTipoProduto As New ClassTipoDeProduto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Tipo_Validate
   
    If Len(Trim(Tipo.ClipText)) = 0 Then
        TipoDescricao.Caption = ""
        TodosTipos.Value = vbChecked
        Exit Sub
    End If
    
    TodosTipos.Value = vbUnchecked
    
    objTipoProduto.iTipo = CInt(Tipo.Text)
    'Lê o Tipo de Produto
    lErro = CF("TipoDeProduto_Le",objTipoProduto)
    If lErro <> SUCESSO And lErro <> 22531 Then Error 43475
    
    'Se não encontrou o Tipo de Produto --> Erro
    If lErro <> SUCESSO Then Error 43476
    
    TipoDescricao.Caption = objTipoProduto.sDescricao
    
    Exit Sub

Erro_Tipo_Validate:

    Cancel = True


    Select Case Err

        Case 43475

        Case 43476
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", Err, objTipoProduto.iTipo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150809)

    End Select

    Exit Sub

End Sub

Private Sub TipoLabel_Click()

Dim lErro As Long
Dim colSelecao As Collection
Dim objTipoProduto As New ClassTipoDeProduto

    If Len(Trim(Tipo.Text)) <> 0 Then objTipoProduto.iTipo = CInt(Tipo.Text)

    'Chama a tela de Browe
    Call Chama_Tela("TipoProdutoLista", colSelecao, objTipoProduto, objEventoTipoProduto)


End Sub

Public Function Trata_Parametros(Optional objClassABC As ClassClassificacaoABC) As Long

    iAlterado = 0

    Trata_Parametros = SUCESSO

End Function

Private Sub TodosTipos_Click()

    iAlterado = REGISTRO_ALTERADO
    
    If TodosTipos.Value = vbChecked Then
        'Limpa e desabilita o Tipo
        Tipo.Text = ""
        TipoDescricao.Caption = ""
    End If
    
End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CLASSIFICACAO_ABC_CLASSIFICACAO
    Set Form_Load_Ocx = Me
    Caption = "Classificação ABC"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ClassificacaoABC"
    
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

'**** fim do trecho a ser copiado *****

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Tipo Then
            Call TipoLabel_Click
        End If
    End If

End Sub




Private Sub XLabel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(XLabel(Index), Source, X, Y)
End Sub

Private Sub XLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(XLabel(Index), Button, Shift, X, Y)
End Sub

Private Sub YLabel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(YLabel(Index), Source, X, Y)
End Sub

Private Sub YLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(YLabel(Index), Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label2(Index), Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2(Index), Button, Shift, X, Y)
End Sub


Private Sub NumItensA_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumItensA, Source, X, Y)
End Sub

Private Sub NumItensA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumItensA, Button, Shift, X, Y)
End Sub

Private Sub NumItensB_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumItensB, Source, X, Y)
End Sub

Private Sub NumItensB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumItensB, Button, Shift, X, Y)
End Sub

Private Sub NumItensC_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumItensC, Source, X, Y)
End Sub

Private Sub NumItensC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumItensC, Button, Shift, X, Y)
End Sub

Private Sub EixoXLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(EixoXLabel, Source, X, Y)
End Sub

Private Sub EixoXLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(EixoXLabel, Button, Shift, X, Y)
End Sub

Private Sub EixoYLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(EixoYLabel, Source, X, Y)
End Sub

Private Sub EixoYLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(EixoYLabel, Button, Shift, X, Y)
End Sub

Private Sub LabelA_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelA, Source, X, Y)
End Sub

Private Sub LabelA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelA, Button, Shift, X, Y)
End Sub

Private Sub LabelB_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelB, Source, X, Y)
End Sub

Private Sub LabelB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelB, Button, Shift, X, Y)
End Sub

Private Sub LabelC_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelC, Source, X, Y)
End Sub

Private Sub LabelC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelC, Button, Shift, X, Y)
End Sub

Private Sub PercItensA_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PercItensA, Source, X, Y)
End Sub

Private Sub PercItensA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PercItensA, Button, Shift, X, Y)
End Sub

Private Sub PercItensB_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PercItensB, Source, X, Y)
End Sub

Private Sub PercItensB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PercItensB, Button, Shift, X, Y)
End Sub

Private Sub PercItensC_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PercItensC, Source, X, Y)
End Sub

Private Sub PercItensC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PercItensC, Button, Shift, X, Y)
End Sub

Private Sub PercDemandaC_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PercDemandaC, Source, X, Y)
End Sub

Private Sub PercDemandaC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PercDemandaC, Button, Shift, X, Y)
End Sub

Private Sub PercDemandaB_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PercDemandaB, Source, X, Y)
End Sub

Private Sub PercDemandaB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PercDemandaB, Button, Shift, X, Y)
End Sub

Private Sub PercDemandaA_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PercDemandaA, Source, X, Y)
End Sub

Private Sub PercDemandaA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PercDemandaA, Button, Shift, X, Y)
End Sub

Private Sub DemandaC_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DemandaC, Source, X, Y)
End Sub

Private Sub DemandaC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DemandaC, Button, Shift, X, Y)
End Sub

Private Sub DemandaB_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DemandaB, Source, X, Y)
End Sub

Private Sub DemandaB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DemandaB, Button, Shift, X, Y)
End Sub

Private Sub DemandaA_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DemandaA, Source, X, Y)
End Sub

Private Sub DemandaA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DemandaA, Button, Shift, X, Y)
End Sub

Private Sub ZeroLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ZeroLabel, Source, X, Y)
End Sub

Private Sub ZeroLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ZeroLabel, Button, Shift, X, Y)
End Sub

Private Sub Itens_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Itens, Source, X, Y)
End Sub

Private Sub Itens_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Itens, Button, Shift, X, Y)
End Sub

Private Sub Demanda_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Demanda, Source, X, Y)
End Sub

Private Sub Demanda_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Demanda, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub FaixaC_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FaixaC, Source, X, Y)
End Sub

Private Sub FaixaC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FaixaC, Button, Shift, X, Y)
End Sub

Private Sub TipoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoLabel, Source, X, Y)
End Sub

Private Sub TipoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub TipoDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoDescricao, Source, X, Y)
End Sub

Private Sub TipoDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoDescricao, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TabStrip1_BeforeClick

    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
    
    If Cancel = False Then
    
        'Se o Tab selecionado for o da CurvaABC e se houve alteração na tela
        If TabStrip1.SelectedItem.Index <> TAB_CurvaABC And iAlterado = REGISTRO_ALTERADO Then
                    
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CLASSIFICACAOABC_ALTERADA")

            If vbMsgRes = vbYes Then
                        
                'Grava os Dados no BD
                lErro = Gravar_Registro()
                If lErro <> SUCESSO And lErro <> 25379 And lErro <> 25386 Then Error 52869
                
                iAlterado = 0
                
            Else
                Error 59359
                    
            End If
        
        End If
    
    End If
    
    Exit Sub
    
Erro_TabStrip1_BeforeClick:
    
    Cancel = True
    
    Select Case Err
        
        Case 52869, 59359
            'volta ao tab inicial
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150810)

    End Select

    Exit Sub

End Sub

