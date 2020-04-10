VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWizardEmpresa 
   Appearance      =   0  'Flat
   Caption         =   "Configuração"
   ClientHeight    =   5445
   ClientLeft      =   555
   ClientTop       =   915
   ClientWidth     =   8580
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "WizardEmpresa2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   8580
   Tag             =   "10"
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Termino da Instalação"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4830
      Index           =   15
      Left            =   0
      TabIndex        =   139
      Tag             =   "3000"
      Top             =   0
      Width           =   8535
      Begin VB.Label Label10 
         Caption         =   "Pressione o botão ""Terminar"" para que suas configurações sejam gravadas."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   780
         TabIndex        =   76
         Top             =   2655
         Width           =   4275
      End
      Begin VB.Label lblStep 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "A Configuração da Empresa está encerrada. "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   5
         Left            =   780
         TabIndex        =   77
         Tag             =   "3001"
         Top             =   630
         Width           =   3960
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   3075
         Index           =   5
         Left            =   5655
         Picture         =   "WizardEmpresa2.frx":014A
         Stretch         =   -1  'True
         Top             =   210
         Width           =   2430
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4830
      Index           =   14
      Left            =   0
      TabIndex        =   171
      Tag             =   "2000"
      Top             =   0
      Width           =   8655
      Begin VB.Frame Frame16 
         Caption         =   "Truncamento / Arredondamento"
         Height          =   780
         Left            =   195
         TabIndex        =   175
         Top             =   2280
         Width           =   4050
         Begin VB.OptionButton Truncamento 
            Caption         =   "Truncamento"
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
            Left            =   120
            TabIndex        =   177
            Top             =   360
            Value           =   -1  'True
            Width           =   1545
         End
         Begin VB.OptionButton Arredondamento 
            Caption         =   "Arredondamento"
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
            Left            =   2040
            TabIndex        =   176
            Top             =   360
            Width           =   1770
         End
      End
      Begin MSMask.MaskEdBox SimboloMoeda 
         Height          =   315
         Left            =   1920
         TabIndex        =   174
         ToolTipText     =   "Símbolo da moeda utilizada"
         Top             =   1560
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   2
         PromptChar      =   " "
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   2205
         Left            =   5760
         Picture         =   "WizardEmpresa2.frx":832C
         Top             =   120
         Width           =   2505
      End
      Begin VB.Label Label30 
         Caption         =   "Símbolo de Moeda:"
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
         Left            =   195
         TabIndex        =   178
         Top             =   1620
         Width           =   1665
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Módulo - Loja"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   195
         TabIndex        =   173
         Top             =   135
         Width           =   1905
      End
      Begin VB.Label Label8 
         Caption         =   "Permite que você defina as configurações de Loja"
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
         Index           =   12
         Left            =   480
         TabIndex        =   172
         Top             =   600
         Width           =   4290
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Passo 5"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4830
      Index           =   9
      Left            =   0
      TabIndex        =   55
      Tag             =   "2006"
      Top             =   -15
      Width           =   8535
      Begin VB.Frame FrameRegPedidoVenda 
         Caption         =   "Registro de Pedidos de Venda"
         Height          =   1650
         Left            =   1080
         TabIndex        =   57
         Top             =   2535
         Width           =   4320
         Begin VB.Frame Frame8 
            Caption         =   "Reserva de Produtos"
            Height          =   645
            Left            =   360
            TabIndex        =   58
            Top             =   705
            Width           =   3540
            Begin VB.OptionButton ReservaManPV 
               Caption         =   "Manual"
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
               Left            =   2430
               TabIndex        =   59
               Top             =   300
               Width           =   945
            End
            Begin VB.OptionButton ReservaAutoPV 
               Caption         =   "Automática"
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
               Left            =   435
               TabIndex        =   60
               Top             =   300
               Value           =   -1  'True
               Width           =   1905
            End
         End
         Begin VB.CheckBox CheckEditaComissoesPV 
            Caption         =   "Permite Editar Comissões"
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
            Left            =   375
            TabIndex        =   61
            Top             =   390
            Width           =   2460
         End
      End
      Begin VB.ListBox ListaConfiguraFAT 
         Height          =   510
         ItemData        =   "WizardEmpresa2.frx":E30A
         Left            =   1110
         List            =   "WizardEmpresa2.frx":E314
         Style           =   1  'Checkbox
         TabIndex        =   68
         Top             =   1485
         Width           =   4320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Módulo - Faturamento"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   195
         TabIndex        =   83
         Top             =   135
         Width           =   3045
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Configurações"
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
         Left            =   1110
         TabIndex        =   84
         Top             =   1245
         Width           =   1230
      End
      Begin VB.Label Label8 
         Caption         =   "Permite que você defina as configurações do faturamento"
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
         Index           =   7
         Left            =   750
         TabIndex        =   85
         Top             =   600
         Width           =   5265
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   2565
         Index           =   9
         Left            =   6120
         Picture         =   "WizardEmpresa2.frx":E378
         Top             =   60
         Width           =   2130
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Height          =   4890
      Index           =   13
      Left            =   0
      TabIndex        =   127
      Top             =   0
      Width           =   9255
      Begin VB.Frame Frame13 
         Caption         =   "Recebimento"
         Height          =   3075
         Left            =   4440
         TabIndex        =   159
         Top             =   1740
         Width           =   3900
         Begin VB.Frame Frame15 
            Caption         =   "Faixa de recebimento"
            Height          =   1095
            Left            =   195
            TabIndex        =   161
            Top             =   660
            Width           =   3525
            Begin MSMask.MaskEdBox PercentMaisReceb 
               Height          =   315
               Left            =   2310
               TabIndex        =   134
               Top             =   255
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox PercentMenosReceb 
               Height          =   315
               Left            =   2310
               TabIndex        =   135
               Top             =   675
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               Caption         =   "Percentagem a menos:"
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
               TabIndex        =   163
               Top             =   735
               Width           =   1950
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               Caption         =   "Percentagem a mais:"
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
               Left            =   450
               TabIndex        =   162
               Top             =   330
               Width           =   1785
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Recebimento fora da faixa"
            Height          =   1035
            Left            =   195
            TabIndex        =   160
            Top             =   1860
            Width           =   3540
            Begin VB.OptionButton RecebForaFaixa 
               Caption         =   "Não aceita recebimento"
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
               Index           =   0
               Left            =   330
               TabIndex        =   136
               Top             =   300
               Value           =   -1  'True
               Width           =   2415
            End
            Begin VB.OptionButton RecebForaFaixa 
               Caption         =   "Avisa e aceita recebimento"
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
               Index           =   1
               Left            =   315
               TabIndex        =   137
               Top             =   660
               Width           =   2655
            End
         End
         Begin VB.CheckBox NaoTemFaixaReceb 
            Caption         =   "Aceita qualquer quantidade sem aviso"
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
            Left            =   225
            TabIndex        =   133
            Top             =   285
            Width           =   3585
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Cotações / Concorrências"
         Height          =   3960
         Left            =   165
         TabIndex        =   153
         Top             =   855
         Width           =   4110
         Begin VB.Frame Frame7 
            Caption         =   "Cotações Anteriores"
            Height          =   2040
            Left            =   135
            TabIndex        =   154
            Top             =   1695
            Width           =   3795
            Begin VB.Frame Frame12 
               Caption         =   "Limites percentuais de quantidade para uso"
               Height          =   1200
               Left            =   210
               TabIndex        =   155
               Top             =   660
               Width           =   3405
               Begin MSMask.MaskEdBox PercentMaisQuantCotacaoAnterior 
                  Height          =   315
                  Left            =   2220
                  TabIndex        =   131
                  Top             =   315
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   6
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox PercentMenosQuantCotacaoAnterior 
                  Height          =   315
                  Left            =   2235
                  TabIndex        =   132
                  Top             =   735
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   6
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin VB.Label Label26 
                  AutoSize        =   -1  'True
                  Caption         =   "Percentagem a menos:"
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
                  TabIndex        =   157
                  Top             =   795
                  Width           =   1950
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Percentagem a mais:"
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
                  Left            =   360
                  TabIndex        =   156
                  Top             =   375
                  Width           =   1785
               End
            End
            Begin VB.CheckBox NaoConsideraQuantCotacaoAnterior 
               Caption         =   "Usa independente de quantidade"
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
               Left            =   255
               TabIndex        =   130
               Top             =   315
               Width           =   3165
            End
         End
         Begin VB.CheckBox CompradorAumentaQuant 
            Caption         =   "Comprador pode aumentar quantidades requisitadas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   195
            TabIndex        =   128
            Top             =   480
            Width           =   3660
         End
         Begin MSMask.MaskEdBox TaxaFinanceiraEmpresa 
            Height          =   315
            Left            =   3075
            TabIndex        =   129
            Top             =   1155
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Taxa Financeira p/ Concorrência:"
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
            TabIndex        =   158
            Top             =   1215
            Width           =   2880
         End
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   1230
         Index           =   13
         Left            =   5685
         Picture         =   "WizardEmpresa2.frx":1F31A
         Stretch         =   -1  'True
         Top             =   225
         Width           =   2280
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Continuação da configuração de compras"
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
         Left            =   195
         TabIndex        =   165
         Top             =   525
         Width           =   3555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Módulo - Compras"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   180
         TabIndex        =   164
         Top             =   15
         Width           =   2580
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Height          =   4875
      Index           =   12
      Left            =   60
      TabIndex        =   117
      Top             =   60
      Width           =   8415
      Begin VB.ComboBox FilialCompra 
         Height          =   315
         Left            =   2310
         TabIndex        =   118
         Top             =   975
         Width           =   2745
      End
      Begin VB.Frame Frame11 
         Caption         =   "Alçadas / Resíduo / Atraso / Ressuprimento"
         Height          =   1215
         Left            =   105
         TabIndex        =   147
         Top             =   1650
         Width           =   8250
         Begin VB.CheckBox ControleAlcada 
            Caption         =   "Controle de alçadas"
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
            Left            =   915
            TabIndex        =   119
            Top             =   360
            Width           =   2025
         End
         Begin MSMask.MaskEdBox Residuo 
            Height          =   315
            Left            =   2100
            TabIndex        =   120
            Top             =   780
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumComprasMediaAtraso 
            Height          =   315
            Left            =   7155
            TabIndex        =   121
            Top             =   330
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NumComprasTempoRessup 
            Height          =   315
            Left            =   7140
            TabIndex        =   122
            Top             =   765
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Nº Compras para Tempo Ressup:"
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
            Left            =   4230
            TabIndex        =   150
            Top             =   825
            Width           =   2820
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Nº Compras para Média de Atraso:"
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
            Left            =   4110
            TabIndex        =   149
            Top             =   390
            Width           =   2940
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Resíduo (%):"
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
            Index           =   10
            Left            =   885
            TabIndex        =   148
            Top             =   825
            Width           =   1110
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Ponto de Pedido"
         Height          =   1785
         Left            =   90
         TabIndex        =   138
         Top             =   2925
         Width           =   8250
         Begin VB.Frame Frame4 
            Caption         =   "Estoque de Segurança"
            Height          =   1320
            Left            =   4200
            TabIndex        =   140
            Top             =   225
            Width           =   3825
            Begin MSMask.MaskEdBox ConsumoMedioMax 
               Height          =   315
               Left            =   2970
               TabIndex        =   125
               Top             =   375
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   3
               Format          =   "#\%"
               Mask            =   "###"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox TempoRessupMax 
               Height          =   315
               Left            =   2970
               TabIndex        =   126
               Top             =   825
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   3
               Format          =   "#\%"
               Mask            =   "###"
               PromptChar      =   " "
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "% a Mais de Tempo de Ressup:"
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
               TabIndex        =   142
               Top             =   885
               Width           =   2685
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "% a Mais de Consumo Médio:"
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
               Left            =   390
               TabIndex        =   141
               Top             =   435
               Width           =   2490
            End
         End
         Begin MSMask.MaskEdBox MesesMediaTempoRessup 
            Height          =   315
            Left            =   2790
            TabIndex        =   124
            Top             =   1005
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MesesConsumoMedio 
            Height          =   315
            Left            =   2790
            TabIndex        =   123
            Top             =   630
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "meses"
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
            Left            =   3300
            TabIndex        =   146
            Top             =   1065
            Width           =   540
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "meses"
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
            Left            =   3300
            TabIndex        =   145
            Top             =   675
            Width           =   540
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Faixa cálculo Consumo Médio:"
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
            TabIndex        =   144
            Top             =   675
            Width           =   2595
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Faixa cálculo Tempo Ressup:"
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
            Left            =   225
            TabIndex        =   143
            Top             =   1065
            Width           =   2520
         End
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   1230
         Index           =   12
         Left            =   5535
         Picture         =   "WizardEmpresa2.frx":20ABC
         Stretch         =   -1  'True
         Top             =   105
         Width           =   2280
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Módulo - Compras"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   240
         TabIndex        =   152
         Top             =   210
         Width           =   2580
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filial de Compra Default:"
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
         Index           =   2
         Left            =   180
         TabIndex        =   151
         Top             =   1050
         Width           =   2100
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Passo 5"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4830
      Index           =   11
      Left            =   -15
      TabIndex        =   69
      Tag             =   "2006"
      Top             =   0
      Width           =   8550
      Begin VB.ListBox ListaConfiguraEST 
         Height          =   510
         ItemData        =   "WizardEmpresa2.frx":2225E
         Left            =   1110
         List            =   "WizardEmpresa2.frx":22268
         Style           =   1  'Checkbox
         TabIndex        =   70
         Top             =   1560
         Width           =   4320
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   1770
         Index           =   10
         Left            =   5640
         Picture         =   "WizardEmpresa2.frx":222CC
         Top             =   240
         Width           =   2640
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Módulo - Estoque"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   195
         TabIndex        =   78
         Top             =   135
         Width           =   2355
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Configurações"
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
         Left            =   1110
         TabIndex        =   79
         Top             =   1290
         Width           =   1230
      End
      Begin VB.Label Label8 
         Caption         =   "Permite que você defina as configurações do estoque"
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
         Index           =   8
         Left            =   600
         TabIndex        =   80
         Top             =   600
         Width           =   4740
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   4830
      Index           =   10
      Left            =   15
      TabIndex        =   16
      Top             =   0
      Width           =   8550
      Begin VB.Frame SSFrame1 
         Caption         =   "Descontos por Adiantamento de Pagamento"
         Height          =   1920
         Left            =   735
         TabIndex        =   166
         Top             =   2880
         Width           =   5025
         Begin VB.ComboBox TipoDesconto 
            Height          =   315
            Left            =   1020
            TabIndex        =   167
            Top             =   390
            Width           =   1890
         End
         Begin MSMask.MaskEdBox Dias 
            Height          =   225
            Left            =   3015
            TabIndex        =   168
            Top             =   405
            Width           =   645
            _ExtentX        =   1138
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
            Format          =   "0"
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercentualDesc 
            Height          =   225
            Left            =   3810
            TabIndex        =   169
            Top             =   405
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridDescontos 
            Height          =   1215
            Left            =   780
            TabIndex        =   170
            Top             =   420
            Width           =   4080
            _ExtentX        =   7197
            _ExtentY        =   2143
            _Version        =   393216
            Rows            =   4
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Registro de Notas Fiscais"
         Height          =   1695
         Left            =   735
         TabIndex        =   17
         Top             =   990
         Width           =   5025
         Begin VB.CheckBox CheckEditaComissoesNF 
            Caption         =   "Permite Editar Comissões"
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
            Left            =   630
            TabIndex        =   18
            Top             =   450
            Width           =   2535
         End
         Begin VB.Frame Frame6 
            Caption         =   "Alocação de Produtos"
            Height          =   645
            Left            =   660
            TabIndex        =   26
            Top             =   870
            Width           =   3660
            Begin VB.OptionButton AlocacaoAutoNF 
               Caption         =   "Automática"
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
               TabIndex        =   30
               Top             =   300
               Value           =   -1  'True
               Width           =   1905
            End
            Begin VB.OptionButton AlocacaoManNF 
               Caption         =   "Manual"
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
               Left            =   2415
               TabIndex        =   31
               Top             =   300
               Width           =   1215
            End
         End
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   2565
         Index           =   11
         Left            =   6210
         Picture         =   "WizardEmpresa2.frx":308D6
         Top             =   60
         Width           =   2130
      End
      Begin VB.Label Label8 
         Caption         =   "Continuação das configurações do faturamento"
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
         Index           =   9
         Left            =   750
         TabIndex        =   81
         Top             =   600
         Width           =   5340
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Módulo - Faturamento"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   195
         TabIndex        =   82
         Top             =   135
         Width           =   3045
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Passo 5"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4830
      Index           =   8
      Left            =   120
      TabIndex        =   51
      Tag             =   "2006"
      Top             =   0
      Width           =   8415
      Begin VB.ListBox ListaConfiguraCR 
         Height          =   510
         ItemData        =   "WizardEmpresa2.frx":41878
         Left            =   1110
         List            =   "WizardEmpresa2.frx":41882
         Style           =   1  'Checkbox
         TabIndex        =   52
         Top             =   1560
         Width           =   4320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Módulo - Contas a Receber"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   195
         TabIndex        =   86
         Top             =   135
         Width           =   3645
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Configurações"
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
         Left            =   1110
         TabIndex        =   87
         Top             =   1290
         Width           =   1230
      End
      Begin VB.Label Label8 
         Caption         =   "Permite que você defina as configurações de Contas a Receber"
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
         Index           =   6
         Left            =   360
         TabIndex        =   88
         Top             =   600
         Width           =   5490
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   2505
         Index           =   8
         Left            =   6120
         Picture         =   "WizardEmpresa2.frx":418E6
         Top             =   60
         Width           =   2055
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Passo 5"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4830
      Index           =   7
      Left            =   0
      TabIndex        =   53
      Tag             =   "2006"
      Top             =   0
      Width           =   8535
      Begin VB.ListBox ListaConfiguraCP 
         Height          =   510
         ItemData        =   "WizardEmpresa2.frx":517D8
         Left            =   1140
         List            =   "WizardEmpresa2.frx":517E2
         Style           =   1  'Checkbox
         TabIndex        =   54
         Top             =   1560
         Width           =   4320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Módulo - Contas a Pagar"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   195
         TabIndex        =   89
         Top             =   135
         Width           =   3360
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Configurações"
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
         Left            =   1110
         TabIndex        =   90
         Top             =   1290
         Width           =   1230
      End
      Begin VB.Label Label8 
         Caption         =   "Permite que você defina as configurações de Contas a Pagar"
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
         Index           =   5
         Left            =   240
         TabIndex        =   91
         Top             =   615
         Width           =   5340
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   2265
         Index           =   7
         Left            =   5640
         Picture         =   "WizardEmpresa2.frx":51846
         Top             =   60
         Width           =   2550
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Passo 5"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4830
      Index           =   6
      Left            =   0
      TabIndex        =   9
      Tag             =   "2006"
      Top             =   0
      Width           =   8535
      Begin VB.ListBox ListaConfigura 
         Height          =   510
         ItemData        =   "WizardEmpresa2.frx":637A4
         Left            =   1110
         List            =   "WizardEmpresa2.frx":637AE
         Style           =   1  'Checkbox
         TabIndex        =   56
         Top             =   1560
         Width           =   4320
      End
      Begin VB.Label Label8 
         Caption         =   "Permite que você defina as configurações da tesouraria"
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
         Index           =   2
         Left            =   750
         TabIndex        =   92
         Top             =   600
         Width           =   4785
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Configurações"
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
         Left            =   1110
         TabIndex        =   93
         Top             =   1290
         Width           =   1230
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Módulo - Tesouraria"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   195
         TabIndex        =   94
         Top             =   135
         Width           =   2715
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   1680
         Index           =   6
         Left            =   5640
         Picture         =   "WizardEmpresa2.frx":63812
         Top             =   360
         Width           =   2550
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Passo 4"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4830
      Index           =   5
      Left            =   0
      TabIndex        =   8
      Tag             =   "2004"
      Top             =   15
      Width           =   8535
      Begin VB.TextBox NomeExterno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4440
         MaxLength       =   20
         TabIndex        =   38
         Top             =   1215
         Width           =   1305
      End
      Begin MSMask.MaskEdBox NomePeriodo 
         Height          =   285
         Left            =   2760
         TabIndex        =   44
         Top             =   3120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
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
      Begin MSMask.MaskEdBox DataInicioExercicio 
         Height          =   315
         Left            =   1575
         TabIndex        =   45
         Top             =   1635
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataFimExercicio 
         Height          =   315
         Left            =   4440
         TabIndex        =   46
         Top             =   1635
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataInicioPeriodo 
         Height          =   285
         Left            =   4320
         TabIndex        =   47
         Tag             =   "1"
         Top             =   3120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridPeriodos 
         Height          =   1890
         Left            =   1950
         TabIndex        =   48
         Top             =   2880
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   3334
         _Version        =   393216
         Rows            =   7
         Cols            =   3
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin MSComCtl2.UpDown SpinDataInicio 
         Height          =   330
         Left            =   2775
         TabIndex        =   49
         Top             =   1620
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown SpinDataFim 
         Height          =   330
         Left            =   5640
         TabIndex        =   50
         Top             =   1635
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Frame Frame1 
         Caption         =   "Geração  Automática  de  Periodos"
         Height          =   705
         Index           =   3
         Left            =   360
         TabIndex        =   39
         Top             =   2160
         Width           =   7545
         Begin VB.CommandButton BotaoGeraPeriodos 
            DisabledPicture =   "WizardEmpresa2.frx":70B44
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
            Height          =   495
            Left            =   5880
            Picture         =   "WizardEmpresa2.frx":72786
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Gerar Períodos"
            Top             =   135
            Width           =   1305
         End
         Begin VB.ComboBox Periodicidade 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "WizardEmpresa2.frx":743C8
            Left            =   1440
            List            =   "WizardEmpresa2.frx":743E1
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   225
            Width           =   1455
         End
         Begin MSMask.MaskEdBox NumPeriodos 
            Height          =   315
            Left            =   4695
            TabIndex        =   42
            Top             =   240
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
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
         Begin MSComCtl2.UpDown SpinNumPeriodos 
            Height          =   330
            Left            =   5070
            TabIndex        =   43
            Top             =   225
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label LabelNumPeriodos 
            AutoSize        =   -1  'True
            Caption         =   "Num. de Periodos:"
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
            Left            =   3120
            TabIndex        =   95
            Top             =   270
            Width           =   1575
         End
         Begin VB.Label LabelPeriodicidade 
            AutoSize        =   -1  'True
            Caption         =   "Periodicidade:"
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
            Left            =   135
            TabIndex        =   96
            Top             =   270
            Width           =   1230
         End
      End
      Begin VB.Label Label3 
         Caption         =   "Exercicio:"
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
         Height          =   255
         Index           =   2
         Left            =   660
         TabIndex        =   99
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Permite que você defina o primeiro exercício contábil"
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
         Index           =   0
         Left            =   570
         TabIndex        =   97
         Top             =   615
         Width           =   5310
      End
      Begin VB.Label Exercicio 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1605
         TabIndex        =   98
         Top             =   1215
         Width           =   315
      End
      Begin VB.Label LabelDataInicio 
         AutoSize        =   -1  'True
         Caption         =   "Data Inicio:"
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
         Left            =   480
         TabIndex        =   100
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label LabelDataFim 
         AutoSize        =   -1  'True
         Caption         =   "Data Fim:"
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
         Left            =   3510
         TabIndex        =   101
         Top             =   1680
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
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
         Index           =   1
         Left            =   3780
         TabIndex        =   102
         Top             =   1260
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Módulo - Contabilidade"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   195
         TabIndex        =   103
         Top             =   135
         Width           =   3135
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   2190
         Index           =   4
         Left            =   6000
         Picture         =   "WizardEmpresa2.frx":7442B
         Top             =   0
         Width           =   2580
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Passo 3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4830
      Index           =   4
      Left            =   0
      TabIndex        =   7
      Tag             =   "2002"
      Top             =   0
      Width           =   8655
      Begin VB.ComboBox Tipo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "WizardEmpresa2.frx":85BFD
         Left            =   1155
         List            =   "WizardEmpresa2.frx":85BFF
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2115
         Width           =   1400
      End
      Begin VB.ComboBox Delimitador 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "WizardEmpresa2.frx":85C01
         Left            =   3495
         List            =   "WizardEmpresa2.frx":85C0E
         TabIndex        =   12
         Top             =   2115
         Width           =   1065
      End
      Begin VB.ComboBox Preenchimento 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "WizardEmpresa2.frx":85C1B
         Left            =   4575
         List            =   "WizardEmpresa2.frx":85C1D
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2520
         Width           =   3210
      End
      Begin MSMask.MaskEdBox Tamanho 
         Height          =   315
         Left            =   2535
         TabIndex        =   14
         Top             =   2190
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
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
         Mask            =   "99"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridSegmentos 
         Height          =   2580
         Left            =   135
         TabIndex        =   15
         Top             =   2295
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   4551
         _Version        =   393216
         Rows            =   10
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label Label8 
         Caption         =   "Permite que você defina o formato das contas contábeis"
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
         Index           =   1
         Left            =   735
         TabIndex        =   104
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Módulo - Contabilidade"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   195
         TabIndex        =   105
         Top             =   135
         Width           =   3135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Segmentos"
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
         Left            =   165
         TabIndex        =   106
         Top             =   2100
         Width           =   945
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   2190
         Index           =   3
         Left            =   5640
         Picture         =   "WizardEmpresa2.frx":85C1F
         Top             =   75
         Width           =   2580
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Passo 2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4830
      Index           =   3
      Left            =   0
      TabIndex        =   6
      Tag             =   "2000"
      Top             =   0
      Width           =   8655
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   2700
         Index           =   2
         Left            =   300
         TabIndex        =   27
         Top             =   1770
         Visible         =   0   'False
         Width           =   5355
         Begin VB.ComboBox Natureza 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "WizardEmpresa2.frx":973F1
            Left            =   2220
            List            =   "WizardEmpresa2.frx":973F3
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1725
            Width           =   2000
         End
         Begin VB.ComboBox TipoConta 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "WizardEmpresa2.frx":973F5
            Left            =   2220
            List            =   "WizardEmpresa2.frx":973F7
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   990
            Width           =   2000
         End
         Begin VB.Label Nat 
            AutoSize        =   -1  'True
            Caption         =   "Natureza:"
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
            Left            =   1275
            TabIndex        =   107
            Top             =   1800
            Width           =   840
         End
         Begin VB.Label TipoDaConta 
            AutoSize        =   -1  'True
            Caption         =   "Tipo da Conta:"
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
            Left            =   840
            TabIndex        =   108
            Top             =   1050
            Width           =   1275
         End
         Begin VB.Label Label7 
            Caption         =   "Valores Iniciais dos Campos nas Telas em que aparecem:"
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
            Left            =   120
            TabIndex        =   109
            Top             =   420
            Width           =   5310
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1(1)"
         Height          =   2700
         Index           =   1
         Left            =   300
         TabIndex        =   32
         Top             =   1770
         Visible         =   0   'False
         Width           =   5355
         Begin VB.Frame FrameCCL 
            Caption         =   "Utilização de Centro de Custo/Centro de Lucro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1755
            Left            =   0
            TabIndex        =   33
            Top             =   300
            Width           =   5235
            Begin VB.OptionButton CclExtra 
               Caption         =   "Utiliza Centro de Custo/Centro de Lucro Extra Contábil"
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
               Left            =   90
               TabIndex        =   36
               Top             =   1320
               Width           =   5115
            End
            Begin VB.OptionButton CclContabil 
               Caption         =   "Utiliza Centro de Custo/Centro de Lucro Contábil"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   90
               TabIndex        =   35
               Top             =   825
               Width           =   4515
            End
            Begin VB.OptionButton SemCcl 
               Caption         =   "Não utiliza Centro de Custo/Centro de Lucro"
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
               Left            =   90
               TabIndex        =   34
               Top             =   465
               Value           =   -1  'True
               Width           =   4245
            End
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   2700
         Index           =   0
         Left            =   300
         TabIndex        =   19
         Top             =   1770
         Width           =   5235
         Begin VB.Frame Frame2 
            Caption         =   "Lote"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1100
            Index           =   0
            Left            =   615
            TabIndex        =   23
            Top             =   1080
            Width           =   1935
            Begin VB.OptionButton LotePorPeriodo 
               Caption         =   "Por Período"
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
               Left            =   270
               TabIndex        =   25
               Top             =   270
               Width           =   1470
            End
            Begin VB.OptionButton LotePorExercicio 
               Caption         =   "Por Exercício"
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
               Left            =   240
               TabIndex        =   24
               Top             =   735
               Value           =   -1  'True
               Width           =   1530
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Documento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1100
            Left            =   3060
            TabIndex        =   20
            Top             =   1080
            Width           =   1890
            Begin VB.OptionButton DocPorPeriodo 
               Caption         =   "Por Período"
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
               Left            =   285
               TabIndex        =   22
               Top             =   300
               Width           =   1455
            End
            Begin VB.OptionButton DocPorExercicio 
               Caption         =   "Por Exercício"
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
               Left            =   270
               TabIndex        =   21
               Top             =   720
               Value           =   -1  'True
               Width           =   1515
            End
         End
         Begin VB.Label Label6 
            Caption         =   "Permite que você escolha como será feita a reinicialização da numeração dos seguintes campos:"
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
            Left            =   210
            TabIndex        =   110
            Top             =   255
            Width           =   5280
         End
      End
      Begin MSComctlLib.TabStrip Opcoes 
         Height          =   3405
         Left            =   120
         TabIndex        =   37
         Top             =   1200
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   6006
         MultiRow        =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Inicialização"
               Object.ToolTipText     =   "Indica como serão reinicializadas as numerações de alguns campos"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Centro de Custo/Lucro"
               Object.ToolTipText     =   "Utilização de centro de custo/centro de lucro"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Valores Iniciais"
               Object.ToolTipText     =   "Valores com que os campos serão inicializados"
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
      Begin VB.Label Label8 
         Caption         =   "Permite que você defina as configurações da contabilidade"
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
         Index           =   3
         Left            =   480
         TabIndex        =   111
         Top             =   600
         Width           =   5160
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Módulo - Contabilidade"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   195
         TabIndex        =   112
         Top             =   135
         Width           =   3135
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   2190
         Index           =   2
         Left            =   5760
         Picture         =   "WizardEmpresa2.frx":973F9
         Top             =   120
         Width           =   2580
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Passo 1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4830
      Index           =   2
      Left            =   0
      TabIndex        =   62
      Top             =   0
      Width           =   8535
      Begin VB.ComboBox Formato 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "WizardEmpresa2.frx":A8BCB
         Left            =   1290
         List            =   "WizardEmpresa2.frx":A8BCD
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   1410
         Width           =   2500
      End
      Begin VB.ComboBox Preenchimento1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "WizardEmpresa2.frx":A8BCF
         Left            =   4575
         List            =   "WizardEmpresa2.frx":A8BD1
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   3240
         Width           =   3210
      End
      Begin VB.ComboBox Delimitador1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "WizardEmpresa2.frx":A8BD3
         Left            =   3495
         List            =   "WizardEmpresa2.frx":A8BE0
         TabIndex        =   64
         Top             =   3360
         Width           =   1065
      End
      Begin VB.ComboBox Tipo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "WizardEmpresa2.frx":A8BED
         Left            =   1185
         List            =   "WizardEmpresa2.frx":A8BEF
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   3240
         Width           =   1400
      End
      Begin MSMask.MaskEdBox Tamanho1 
         Height          =   300
         Left            =   2550
         TabIndex        =   66
         Top             =   3120
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
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
         Mask            =   "99"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridSegmentos1 
         Height          =   2100
         Left            =   135
         TabIndex        =   67
         Top             =   2745
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   3704
         _Version        =   393216
         Rows            =   10
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label Label8 
         Caption         =   "Permite que você defina o formato dos centros de custo e dos produtos."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   4
         Left            =   480
         TabIndex        =   113
         Top             =   600
         Width           =   4950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Formato de:"
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
         Left            =   165
         TabIndex        =   114
         Top             =   1455
         Width           =   1020
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Segmentos"
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
         Left            =   150
         TabIndex        =   115
         Top             =   2520
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "SGE - Geral"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   195
         TabIndex        =   116
         Top             =   135
         Width           =   1635
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   2565
         Index           =   1
         Left            =   5640
         Picture         =   "WizardEmpresa2.frx":A8BF1
         Top             =   120
         Width           =   2565
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Passo 0"
      Enabled         =   0   'False
      Height          =   4830
      Index           =   1
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8535
      Begin VB.Label Label12 
         Caption         =   "As próximas telas permitirão que você configure o funcionamento do sistema de acordo com as opções escolhidas."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   3000
         TabIndex        =   73
         Top             =   1875
         Width           =   5055
      End
      Begin VB.Label Label11 
         Caption         =   "A Configuração da Empresa está sendo iniciada."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   3000
         TabIndex        =   74
         Top             =   375
         Width           =   5055
      End
      Begin VB.Label Label14 
         Caption         =   $"WizardEmpresa2.frx":BD4FB
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   150
         TabIndex        =   75
         Top             =   3105
         Width           =   7905
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   2175
         Index           =   0
         Left            =   210
         Picture         =   "WizardEmpresa2.frx":BD5CE
         Top             =   210
         Width           =   2565
      End
   End
   Begin VB.Frame fraStep 
      Caption         =   "Frame5"
      Height          =   1815
      Index           =   0
      Left            =   0
      TabIndex        =   72
      Top             =   375
      Width           =   2490
   End
   Begin VB.PictureBox picNav 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   8580
      TabIndex        =   0
      Top             =   4875
      Width           =   8580
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Finish"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   4
         Left            =   7140
         MaskColor       =   &H00000000&
         TabIndex        =   5
         Tag             =   "104"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   3
         Left            =   5745
         MaskColor       =   &H00000000&
         TabIndex        =   4
         Tag             =   "103"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   2
         Left            =   4635
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Tag             =   "102"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   1
         Left            =   3450
         MaskColor       =   &H00000000&
         TabIndex        =   2
         Tag             =   "101"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   0
         Left            =   108
         MaskColor       =   &H00000000&
         TabIndex        =   1
         Tag             =   "100"
         Top             =   120
         Width           =   1092
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   105
         X2              =   8254
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   120
         X2              =   8254
         Y1              =   30
         Y2              =   30
      End
   End
End
Attribute VB_Name = "frmWizardEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const NUM_STEPS = 16

Const MENSAGEM_TERMINO_CONFIG_EMPRESA1 = "A Configuração da Empresa "
Const MENSAGEM_TERMINO_CONFIG_EMPRESA2 = " está encerrada."
Const MENSAGEM_INICIO_CONFIG_EMPRESA1 = "A Configuração da Empresa "
Const MENSAGEM_INICIO_CONFIG_EMPRESA2 = " está sendo iniciada."

Const RES_ERROR_MSG = 30000

Const BTN_HELP = 0
Const BTN_CANCEL = 1
Const BTN_BACK = 2
Const BTN_NEXT = 3
Const BTN_FINISH = 4

Const STEP_INTRO = 0
Const STEP_1 = 1
Const STEP_2 = 2
Const STEP_3 = 3
Const STEP_4 = 4
Const STEP_5 = 5
Const STEP_6 = 6
Const STEP_7 = 7
Const STEP_8 = 8
Const STEP_9 = 9
Const STEP_10 = 10
Const STEP_11 = 11
Const STEP_12 = 12
Const STEP_13 = 13
Const STEP_14 = 14
Const STEP_FINISH = 15

Const DIR_NONE = 0
Const DIR_BACK = 1
Const DIR_NEXT = 2

Const FRM_TITLE = "Configuração da Empresa "
Const INTRO_KEY = "Tela de Introdução"
Const SHOW_INTRO = "Exibir Introdução"
Const TOPIC_TEXT = "<TOPIC_TEXT>"

'module level vars
Dim mnCurStep       As Integer
Dim mbHelpStarted   As Boolean
Dim mnVeStep        As Integer 'passo corrente que está visivel

Public VBInst       As VBIDE.VBE
Dim mbFinishOK      As Boolean

'DECLARACAO DE VARIAVEIS GLOBAIS
Dim iAlterado As Integer
Dim objConfiguraADM1 As ClassConfiguraADM

'=====================================================================
'DECLARACAO DE VARIAVEIS GLOBAIS PARA CONTABILIDADE TELA DE SEGMENTOS
'====================================================================
Dim CTB_Segmentos_objGrid1 As AdmGrid

'coluna referente ao tipo no grid da tela segmentos
Const COL_TIPO = 1
'coluna referente ao tamanho no grid da tela segmentos
Const COL_TAMANHO = 2
'coluna referente ao delimitador no grid da tela segmentos
Const COL_DELIMITADOR = 3
'coluna referente ao preenchimento no grid da tela segmentos
Const COL_PREENCHIMENTO = 4

'============================================================================
'FIM DA DECLARACAO DAS VARIAVEIS GLOBAIS PARA CONTABILIDADE TELA DE SEGMENTOS
'============================================================================

'========================================================================
'DECLARACAO DE VARIAVEIS GLOBAIS PARA CONTABILIDADE TELA DE CONFIGURACAO
'========================================================================

Dim CTB_Config_iFrameAtual As Integer

'========================================================================
'FIM DA DECLARACAO DAS VARIAVEIS GLOBAIS PARA CONTABILIDADE TELA DE CONFIGURACAO
'========================================================================

'========================================================================
'DECLARACAO DE VARIAVEIS GLOBAIS PARA CONTABILIDADE TELA DE EXERCICIO
'========================================================================

'Codigos de Periodicidade de Exercicio

Const PERIODICIDADE_ANUAL = 1
Const PERIODICIDADE_BIMENSAL = 2
Const PERIODICIDADE_LIVRE = 3
Const PERIODICIDADE_MENSAL = 4
Const PERIODICIDADE_QUADRIMESTRAL = 5
Const PERIODICIDADE_SEMESTRAL = 6
Const PERIODICIDADE_TRIMESTRAL = 7


Const GRID_NOME_COL = 1
Const GRID_DATAINI_COL = 2
Dim iExercicioMudou As Integer
Dim CTB_Exercicio_objGrid1 As AdmGrid
Dim iExercicio2 As Integer

'========================================================================
'FIM DA DECLARACAO DAS VARIAVEIS GLOBAIS PARA CONTABILIDADE TELA DE EXERCICIO
'========================================================================

'========================================================================
'DECLARACAO DE VARIAVEIS GLOBAIS PARA TESOURARIA TELA DE CONFIGURACAO
'========================================================================

Const TESCONFIG_AGLUTINA_LANCAM_POR_DIA = 0
Const TESCONFIG_GERA_LOTE_AUTOMATICO = 1

'========================================================================
'FIM DA DECLARACAO DAS VARIAVEIS GLOBAIS PARA TESOURARIA TELA DE CONFIGURACAO
'========================================================================

'========================================================================
'DECLARACAO DE VARIAVEIS GLOBAIS PARA CTAS A PAGAR TELA DE CONFIGURACAO
'========================================================================

Const CPCONFIG_AGLUTINA_LANCAM_POR_DIA = 0
Const CPCONFIG_GERA_LOTE_AUTOMATICO = 1

'==============================================================================
'FIM DA DECLARACAO DAS VARIAVEIS GLOBAIS PARA CTAS A PAGAR TELA DE CONFIGURACAO
'==============================================================================

'========================================================================
'DECLARACAO DE VARIAVEIS GLOBAIS PARA CTAS A RECEBER TELA DE CONFIGURACAO
'========================================================================

Const CRCONFIG_AGLUTINA_LANCAM_POR_DIA = 0
Const CRCONFIG_GERA_LOTE_AUTOMATICO = 1

'==============================================================================
'FIM DA DECLARACAO DAS VARIAVEIS GLOBAIS PARA CTAS A RECEBER TELA DE CONFIGURACAO
'==============================================================================

'========================================================================
'DECLARACAO DE VARIAVEIS GLOBAIS PARA FATURAMENTO TELA DE CONFIGURACAO
'========================================================================

Const FATCONFIG_AGLUTINA_LANCAM_POR_DIA = 0
Const FATCONFIG_GERA_LOTE_AUTOMATICO = 1
Const FATCONFIG_EDITA_COMISSOES_PV = 2

'GridDescontos
Dim objGridDescontos As AdmGrid
Dim iGrid_TipoDesconto_Col As Integer
Dim iGrid_Dias_Col As Integer
Dim iGrid_Percentual_Col As Integer

'==============================================================================
'FIM DA DECLARACAO DAS VARIAVEIS GLOBAIS PARA FATURAMENTO TELA DE CONFIGURACAO
'==============================================================================

'========================================================================
'DECLARACAO DE VARIAVEIS GLOBAIS PARA ESTOQUE TELA DE CONFIGURACAO
'========================================================================

Const ESTCONFIG_AGLUTINA_LANCAM_POR_DIA = 0
Const ESTCONFIG_GERA_LOTE_AUTOMATICO = 1

'==============================================================================
'FIM DA DECLARACAO DAS VARIAVEIS GLOBAIS PARA ESTOQUE TELA DE CONFIGURACAO
'==============================================================================

'=====================================================================
'DECLARACAO DE VARIAVEIS GLOBAIS PARA GERAL TELA DE SEGMENTOS
'====================================================================
Dim SGE_Segmentos_objGrid1 As AdmGrid
Dim SGE_Segmentos_sCodigo As String
Dim SGE_Segmentos_colSegmento As New Collection

Const FORMATO_PRODUTO = 1
Const FORMATO_CCL = 2

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Ccl = 2
Private Const TAB_Valores = 3



Private Sub Arredondamento_Click()
      iAlterado = REGISTRO_ALTERADO
End Sub

'============================================================================
'FIM DA DECLARACAO DAS VARIAVEIS GLOBAIS PARA GERAL TELA DE SEGMENTOS
'============================================================================

Private Sub cmdNav_Click(Index As Integer)
    
Dim nAltStep As Integer
Dim lHelpTopic As Long
Dim rc As Long
Dim lErro As Long
    
On Error GoTo Erro_cmdNav_Click

    Select Case Index
        Case BTN_HELP
                        
            SendKeys "{F1}", True
        
        Case BTN_CANCEL
            Unload Me
          
        Case BTN_BACK
            'place special cases here to jump
            'to alternate steps
LABEL_BTN_BACK:
            nAltStep = mnCurStep - 1
            lErro = SetStep(nAltStep, DIR_BACK)
            If lErro = 44862 Then
                mnCurStep = mnCurStep - 1
                GoTo LABEL_BTN_BACK
            End If
        Case BTN_NEXT
            'place special cases here to jump
            'to alternate steps
LABEL_BTN_NEXT:
            nAltStep = mnCurStep + 1
            lErro = SetStep(nAltStep, DIR_NEXT)
            If lErro = 44862 Then
                mnCurStep = mnCurStep + 1
                GoTo LABEL_BTN_NEXT
            End If
        Case BTN_FINISH
      
            lErro = Gravar_Registro()
            If lErro <> SUCESSO Then Error 44846
            
            objConfiguraADM1.iConfiguracaoOK = True
            
            Unload Me
            
'            If GetSetting(APP_CATEGORY, WIZARD_NAME, CONFIRM_KEY, vbNullString) = vbNullString Then
'                frmConfirm.Show vbModal
'            End If
        
    End Select
    
    Exit Sub
    
Erro_cmdNav_Click:

    Select Case Err

        Case 44846

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175813)

    End Select

    Exit Sub

End Sub

Private Sub ConsumoMedioMax_GotFocus()

    Call MaskEdBox_TrataGotFocus(ConsumoMedioMax, iAlterado)
    
End Sub

Private Sub DataFimExercicio_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFimExercicio)

End Sub

Private Sub DataInicioExercicio_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicioExercicio)

End Sub

Private Sub Form_Load()
    Dim i As Integer
    'init all vars
    mbFinishOK = False

    For i = STEP_1 To NUM_STEPS - 1
      fraStep(i).left = -10000
    Next
    
    'Load All string info for Form
    LoadResStrings Me
    
    'Determine 1st Step:
    If GetSetting(APP_CATEGORY, WIZARD_NAME, INTRO_KEY, vbNullString) = SHOW_INTRO Then
        Call SetStep(STEP_INTRO, DIR_NEXT)
    Else
        Call SetStep(STEP_1, DIR_NONE)
    End If
    
    If giTipoVersao = VERSAO_LIGHT Then
    
        FrameCCL.Enabled = False
        SemCcl.Enabled = False
        CclContabil.Enabled = False
        CclExtra.Enabled = False
        
        'Remove "Lote Automático" das Listas
        Call ListaConfigura.RemoveItem(1)
        ListaConfigura.Height = 285
        Call ListaConfiguraCP.RemoveItem(1)
        ListaConfiguraCP.Height = 285
        Call ListaConfiguraCR.RemoveItem(1)
        ListaConfiguraCR.Height = 285
        Call ListaConfiguraEST.RemoveItem(1)
        ListaConfiguraEST.Height = 285
        Call ListaConfiguraFAT.RemoveItem(1)
        ListaConfiguraFAT.Height = 285
        
        CheckEditaComissoesPV.Value = vbChecked
        FrameRegPedidoVenda.Visible = False
        
        CheckEditaComissoesNF.Value = vbChecked
        CheckEditaComissoesNF.Enabled = False
        
    End If
    
End Sub

Private Function SetStep(nStep As Integer, nDirection As Integer) As Long
  
Dim lErro As Long, iStep As Integer
  
On Error GoTo Erro_SetStep
  
    Select Case nStep
    
        Case STEP_INTRO
        
        Case STEP_1
            Me.HelpContextID = IDH_CONFIGURACAO_EMPRESA
            Label11.Caption = MENSAGEM_INICIO_CONFIG_EMPRESA1 & gsNomeEmpresa & MENSAGEM_INICIO_CONFIG_EMPRESA2
      
        Case STEP_2
            Me.HelpContextID = IDH_SGE_GERAL
            lErro = Valida_Step(SISTEMA_SGE)
            If lErro <> SUCESSO Then gError 44862
        
        Case STEP_3
            Me.HelpContextID = IDH_MODULO_CONTABILIDADE_ID
            lErro = SGE_Segmentos_Testa()
            If lErro <> SUCESSO Then gError 44804
        
            lErro = Valida_Step(MODULO_CONTABILIDADE)
            If lErro <> SUCESSO Then gError 44862
        
        Case STEP_4
            Me.HelpContextID = IDH_CONFIGURA_SEGMENTOS
            lErro = Valida_Step(MODULO_CONTABILIDADE)
            If lErro <> SUCESSO Then gError 44862
        
        Case STEP_5
            Me.HelpContextID = IDH_MODULO_CONTABILIDADE
            lErro = CTB_Segmentos_Testa()
            If lErro <> SUCESSO Then gError 44747
            
            lErro = Valida_Step(MODULO_CONTABILIDADE)
            If lErro <> SUCESSO Then gError 44862
            
        Case STEP_6
            Me.HelpContextID = IDH_MODULO_TESOURARIA
            lErro = CTB_Exercicio_Testa()
            If lErro <> SUCESSO Then gError 44655
      
            lErro = Valida_Step(MODULO_TESOURARIA)
            If lErro <> SUCESSO Then gError 44862
      
        Case STEP_7
            Me.HelpContextID = IDH_MODULO_CONTAS_PAGAR
            lErro = Valida_Step(MODULO_CONTASAPAGAR)
            If lErro <> SUCESSO Then gError 44862
      
        Case STEP_8
            Me.HelpContextID = IDH_MODULO_CONTAS_RECEBER
            lErro = Valida_Step(MODULO_CONTASARECEBER)
            If lErro <> SUCESSO Then gError 44862
      
        Case STEP_9
            Me.HelpContextID = IDH_MODULO_FATURAMENTO
            lErro = Valida_Step(MODULO_FATURAMENTO)
            If lErro <> SUCESSO Then gError 44862
      
        Case STEP_10
            Me.HelpContextID = IDH_MODULO_FATURAMENTO_CONTINUACAO
            lErro = Valida_Step(MODULO_FATURAMENTO)
            If lErro <> SUCESSO Then gError 44862
      
        Case STEP_11
            Me.HelpContextID = IDH_MODULO_ESTOQUE
            lErro = FAT_Parte2_Testa()
            If lErro <> SUCESSO Then gError 56703
            
            lErro = Valida_Step(MODULO_ESTOQUE)
            If lErro <> SUCESSO Then gError 44862
      
        Case STEP_12
            Me.HelpContextID = IDH_MODULO_COMPRAS
            lErro = Valida_Step(MODULO_COMPRAS)
            If lErro <> SUCESSO Then gError 44862
      
        Case STEP_13
            Me.HelpContextID = IDH_MODULO_COMPRAS
            
            lErro = COM_Parte1_Testa()
            If lErro <> SUCESSO Then gError 74937
            
            lErro = Valida_Step(MODULO_COMPRAS)
            If lErro <> SUCESSO Then gError 44862
            
        Case STEP_14
            '???Me.HelpContextID =IDH_MODULO_LOJA
            
            lErro = Valida_Step(MODULO_LOJA)
            If lErro <> SUCESSO Then gError 44862
            
        Case STEP_FINISH
            lErro = LJ_Parte1_Testa()
            If lErro <> SUCESSO Then gError 109339
            
            lblStep(5).Caption = MENSAGEM_TERMINO_CONFIG_EMPRESA1 & gsNomeEmpresa & MENSAGEM_TERMINO_CONFIG_EMPRESA2
            mbFinishOK = True
        
    End Select
    
    'move to new step
    fraStep(mnVeStep).Enabled = False
    fraStep(nStep).left = 0
    If nStep <> mnVeStep Then
        
        For iStep = STEP_INTRO To STEP_FINISH
        
            If iStep <> nStep Then
                fraStep(iStep).left = -10000
                fraStep(iStep).Enabled = False
            End If
    
        Next
    
    End If
    fraStep(nStep).Enabled = True
  
    SetCaption nStep
    SetNavBtns nStep
  
    SetStep = SUCESSO
  
    Exit Function

Erro_SetStep:

    SetStep = gErr

    Select Case gErr

        Case 44655, 44747, 44804, 44862, 56703, 74937, 109339, 109340

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175814)

    End Select

    Exit Function
  
End Function

Private Function LJ_Parte1_Testa() As Long

Dim lErro As Long

On Error GoTo Erro_LJ_Parte1_Testa

    lErro = Valida_Step(MODULO_LOJA)
    If lErro = SUCESSO Then
    
        'verifica se o símbolo moeda está preenchido
        If Len(Trim(SimboloMoeda.Text)) = 0 Then gError 109341

    End If
    
    LJ_Parte1_Testa = SUCESSO
    
    Exit Function
    
Erro_LJ_Parte1_Testa:
    
    LJ_Parte1_Testa = gErr
    
    Select Case gErr
    
        Case 109341
            Call Rotina_Erro(vbOKOnly, "ERRO_SIMBOLOMOEDA_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175815)
            
    End Select
    
    Exit Function

End Function

Private Sub SetNavBtns(nStep As Integer)
    mnVeStep = nStep
    mnCurStep = nStep
    
    If mnCurStep = STEP_1 Then
        cmdNav(BTN_BACK).Enabled = False
        cmdNav(BTN_NEXT).Enabled = True
    ElseIf mnCurStep = NUM_STEPS - 1 Then
        cmdNav(BTN_NEXT).Enabled = False
        cmdNav(BTN_BACK).Enabled = True
    Else
        cmdNav(BTN_BACK).Enabled = True
        cmdNav(BTN_NEXT).Enabled = True
    End If
    
    If mbFinishOK Then
        cmdNav(BTN_FINISH).Enabled = True
    Else
        cmdNav(BTN_FINISH).Enabled = False
    End If
End Sub

Private Sub SetCaption(nStep As Integer)
    On Error Resume Next

    Me.Caption = FRM_TITLE & gsNomeEmpresa
'    Me.Caption = FRM_TITLE & " - " & LoadResString(fraStep(nStep).Tag)

End Sub

'=========================================================
'this sub displays an error message when the user has
'not entered enough data to continue
'=========================================================
Sub IncompleteData(nIndex As Integer)
    On Error Resume Next
    Dim sTmp As String
      
    'get the base error message
    sTmp = LoadResString(RES_ERROR_MSG)
    'get the specific message
    sTmp = sTmp & vbCrLf & LoadResString(RES_ERROR_MSG + nIndex)
    Beep
    MsgBox sTmp, vbInformation
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim rc As Long
    'see if we need to save the settings
'    If chkSaveSettings(0).Value = vbChecked Then
      
'        SaveSetting APP_CATEGORY, WIZARD_NAME, "OptionName", Option Value
      
'    End If
  
    Set objConfiguraADM1 = Nothing
    Set CTB_Segmentos_objGrid1 = Nothing
    Set CTB_Exercicio_objGrid1 = Nothing
    Set objGridDescontos = Nothing
    Set SGE_Segmentos_objGrid1 = Nothing
    
''??? Jones: vamos aproveitar isto. 16/08/2000 If mbHelpStarted Then rc = WinHelp(Me.hwnd, HELP_FILE, HELP_QUIT, 0)
End Sub

Private Function Gravar_Registro() As Long

Dim lErro As Long
Dim lTransacao As Long
Dim lTransacaoDic As Long
Dim lConexao As Long

On Error GoTo Erro_Gravar_Registro
    
    lConexao = GL_lConexaoDic
    
    'Inicia a Transacao
    lTransacaoDic = Transacao_AbrirDic
    If lTransacaoDic = 0 Then gError 109341
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 109342
    
    lErro = SGE_Configuracao_Gravar_Registro()
    If lErro <> SUCESSO Then gError 109343
    
    lErro = SGE_Segmentos_Gravar_Registro()
    If lErro <> SUCESSO Then gError 109344
    
    lErro = CTB_Config_Gravar_Registro()
    If lErro <> SUCESSO Then gError 109345
    
    lErro = CTB_Segmentos_Gravar_Registro()
    If lErro <> SUCESSO Then gError 109346
    
    lErro = CTB_Exercicio_Gravar_Registro()
    If lErro <> SUCESSO Then gError 109347
    
    lErro = TES_Config_Gravar_Registro()
    If lErro <> SUCESSO Then gError 109348
    
    lErro = CP_Config_Gravar_Registro()
    If lErro <> SUCESSO Then gError 109349
    
    lErro = CR_Config_Gravar_Registro()
    If lErro <> SUCESSO Then gError 109350
    
    lErro = FAT_Config_Gravar_Registro()
    If lErro <> SUCESSO Then gError 109351
    
    lErro = EST_Config_Gravar_Registro()
    If lErro <> SUCESSO Then gError 109352
    
    lErro = COM_Config_Gravar_Registro()
    If lErro <> SUCESSO Then gError 109353
    
    lErro = LJ_Config_Gravar_Registro()
    If lErro <> SUCESSO Then gError 109354
    
    lErro = CF("ModuloFilEmp_Atualiza_Configurado", glEmpresa, EMPRESA_TODA, objConfiguraADM1.colModulosConfigurar)
    If lErro <> SUCESSO Then gError 109355
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 109356
    
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then gError 109357
    
    iAlterado = 0
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:
    
    Gravar_Registro = gErr
    
    Select Case gErr

        Case 109341, 109343 To 109355

        Case 109342
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 109356
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175816)

    End Select

    If gErr <> 109357 Then Call Transacao_Rollback
    Call Transacao_RollbackDic

    Exit Function
    
End Function

Private Sub LJ_Inicializa_Config()

Dim lErro As Long

On Error GoTo Erro_LJ_Inicializa_Config

    lErro = Valida_Step(MODULO_LOJA)
    
    If lErro = SUCESSO Then
    
        If gobjLoja.sTruncamentoArredondamento = LOJA_ARREDONDAMENTO Then
            Arredondamento.Value = True
        Else
            Truncamento.Value = True
        End If
        
        SimboloMoeda.Text = gobjLoja.sSimboloMoeda
    
    End If
        
    Exit Sub
    
    Exit Sub
    
Erro_LJ_Inicializa_Config:
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175817)
    
    End Select
    
    Exit Sub

End Sub

Private Function LJ_Config_Gravar_Registro() As Long

Dim lErro As Long
Dim colConfig As New ColLojaConfig

On Error GoTo Erro_LJ_Config_Gravar_Registro
    
    lErro = Valida_Step(MODULO_LOJA)
    
    If lErro = SUCESSO Then
    
        gobjLoja.sSimboloMoeda = SimboloMoeda.Text
        
        If Truncamento.Value = True Then
            gobjLoja.sTruncamentoArredondamento = LOJA_TRUNCAMENTO
        Else
            gobjLoja.sTruncamentoArredondamento = LOJA_ARREDONDAMENTO
        End If
        
        lErro = CF("ConfiguraLoja_MoverCampos_ColLojaConfig_EmpToda", gobjLoja, colConfig)
        If lErro <> SUCESSO Then gError 109388
        
        lErro = gobjLoja.Gravar_Trans(gobjLoja, colConfig)
        If lErro <> SUCESSO Then gError 109358
        
    End If

    LJ_Config_Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_LJ_Config_Gravar_Registro:
    
    LJ_Config_Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 109358, 109388
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175818)
    
    End Select
    
    Exit Function

End Function

Private Function Valida_Step(sModulo As String) As Long

Dim vModulo As Variant

    For Each vModulo In objConfiguraADM1.colModulosConfigurar

        If sModulo = vModulo Then
            Valida_Step = SUCESSO
            Exit Function
        End If
        
    Next
    
    Valida_Step = 44863

End Function

'=============================================================================
' TELA DE SEGMENTOS CONTABILIDADE
'=============================================================================

Private Function CTB_Segmentos_Testa() As Long
'verifica se os segmentos estão preenchidos.

Dim lErro As Long
Dim iTotalTamanho As Integer
Dim iLinha As Integer

On Error GoTo Erro_CTB_Segmentos_Testa

Dim colSegmento As Collection

    lErro = Valida_Step(MODULO_CONTABILIDADE)
    
    If lErro = SUCESSO Then

        If CTB_Segmentos_objGrid1.iLinhasExistentes = 0 Then Error 44746
    
        'percorre as linhas da coluna tamanho
        For iLinha = 1 To CTB_Segmentos_objGrid1.iLinhasExistentes
        
            'verifica se nao foi preenchido o tamanho
            If Len(Trim(GridSegmentos.TextMatrix(iLinha, COL_TAMANHO))) = 0 Then Error 44748
            'soma o valor total da coluna tamanho no grid
            iTotalTamanho = iTotalTamanho + CInt(GridSegmentos.TextMatrix(iLinha, COL_TAMANHO))
    
        Next
                  
        If iTotalTamanho > STRING_CONTA Then Error 44749

    End If
    
    CTB_Segmentos_Testa = SUCESSO
    
    Exit Function
    
Erro_CTB_Segmentos_Testa:

    CTB_Segmentos_Testa = Err

    Select Case Err
    
        Case 44649

        Case 44746
            Call Rotina_Erro(vbOKOnly, "ERRO_SEGMENTO_VAZIO", Err)
    
        Case 44748
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_TAMANHO_NAO_PREENCHIDO", Err)

        Case 44749
            Call Rotina_Erro(vbOKOnly, "ERRO_SEGMENTO_CONTA_MAIOR_PERMITIDO", Err, iTotalTamanho, STRING_CONTA)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175819)

    End Select

    Exit Function

End Function

Private Function CTB_Inicializa_Segmentos() As Long

Dim iIndice As Integer
Dim lErro As Long
Dim colSegmento As New Collection
Dim objSegmento As New ClassSegmento

On Error GoTo Erro_CTB_Inicializa_Segmentos

    lErro = Valida_Step(MODULO_CONTABILIDADE)
    
    If lErro = SUCESSO Then

        Set CTB_Segmentos_objGrid1 = New AdmGrid
               
        'inicializacao do grid
        Call Inicializa_Grid_Segmento
        
        'inicializar os tipos
        For iIndice = 1 To gobjColTipoSegmento.Count
            Tipo.AddItem gobjColTipoSegmento.Item(iIndice).sDescricao
        Next
    
        'inicializar os preenchimentos
        For iIndice = 1 To gobjColPreenchimento.Count
            Preenchimento.AddItem gobjColPreenchimento.Item(iIndice).sDescricao
        Next
    
        'preenche o obj com o formato corrente para usar em Segmento_Le_Codigo
        objSegmento.sCodigo = SEGMENTO_CONTA
    
        'preenche toda colecao(colSegmento) com o formato da conta
        lErro = CF("Segmento_Le_Codigo", objSegmento, colSegmento)
        If lErro <> SUCESSO Then Error 44741
    
        'preenche todo o grid da tabela segmento
        For Each objSegmento In colSegmento
    
            'coloca o tipo no grid da tela
            GridSegmentos.TextMatrix(objSegmento.iNivel, COL_TIPO) = gobjColTipoSegmento.Descricao(objSegmento.iTipo)
    
            'coloca o tamanho no grid da tela
            GridSegmentos.TextMatrix(objSegmento.iNivel, COL_TAMANHO) = objSegmento.iTamanho
    
            'coloca os delimitadores no grid da tela
            GridSegmentos.TextMatrix(objSegmento.iNivel, COL_DELIMITADOR) = objSegmento.sDelimitador
    
            'coloca o preenchimento no grid da tela
            GridSegmentos.TextMatrix(objSegmento.iNivel, COL_PREENCHIMENTO) = gobjColPreenchimento.Descricao(objSegmento.iPreenchimento)
    
            CTB_Segmentos_objGrid1.iLinhasExistentes = CTB_Segmentos_objGrid1.iLinhasExistentes + 1
    
        Next
    
        iAlterado = 0

    End If
    
    CTB_Inicializa_Segmentos = SUCESSO
    
    Exit Function
    
Erro_CTB_Inicializa_Segmentos:

    CTB_Inicializa_Segmentos = Err
    
    Select Case Err
    
        Case 44741, 44864
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175820)
    
    End Select
    
    Exit Function
    
End Function

Private Sub MesesConsumoMedio_GotFocus()

    Call MaskEdBox_TrataGotFocus(MesesConsumoMedio, iAlterado)
    
End Sub

Private Sub MesesMediaTempoRessup_GotFocus()

    Call MaskEdBox_TrataGotFocus(MesesMediaTempoRessup, iAlterado)
    
End Sub

Private Sub NumComprasMediaAtraso_GotFocus()

    Call MaskEdBox_TrataGotFocus(NumComprasMediaAtraso, iAlterado)
    
End Sub

Private Sub NumComprasTempoRessup_GotFocus()

    Call MaskEdBox_TrataGotFocus(NumComprasTempoRessup, iAlterado)
    
End Sub

Private Sub NumPeriodos_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumPeriodos)

End Sub

Private Sub PercentMaisQuantCotacaoAnterior_GotFocus()

    Call MaskEdBox_TrataGotFocus(PercentMaisQuantCotacaoAnterior, iAlterado)
    
End Sub

Private Sub PercentMaisReceb_GotFocus()

    Call MaskEdBox_TrataGotFocus(PercentMaisReceb, iAlterado)
    
End Sub

Private Sub PercentMenosQuantCotacaoAnterior_GotFocus()

    Call MaskEdBox_TrataGotFocus(PercentMenosQuantCotacaoAnterior, iAlterado)
    
End Sub

Private Sub PercentMenosReceb_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(PercentMenosReceb, iAlterado)
    
End Sub

Private Sub Residuo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Residuo, iAlterado)
    
End Sub

Private Sub SimboloMoeda_Change()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TaxaFinanceiraEmpresa_GotFocus()

    Call MaskEdBox_TrataGotFocus(TaxaFinanceiraEmpresa, iAlterado)
    
End Sub

Private Sub TempoRessupMax_GotFocus()

    Call MaskEdBox_TrataGotFocus(TempoRessupMax, iAlterado)
    
End Sub

Private Sub Tipo_Change()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tipo_Click()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tamanho_Change()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tamanho_Click()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Delimitador_Change()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Delimitador_Click()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Preenchimento_Change()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Preenchimento_Click()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tipo_GotFocus()

    Call Grid_Campo_Recebe_Foco(CTB_Segmentos_objGrid1)
    
End Sub

Private Sub Tipo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, CTB_Segmentos_objGrid1)

End Sub

Private Sub Tipo_LostFocus()

    Set CTB_Segmentos_objGrid1.objControle = Tipo
    
    Call Grid_Campo_Libera_Foco(CTB_Segmentos_objGrid1)

End Sub

Private Sub Tamanho_GotFocus()

    Call Grid_Campo_Recebe_Foco(CTB_Segmentos_objGrid1)
    
End Sub

Private Sub Tamanho_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, CTB_Segmentos_objGrid1)

End Sub

Private Sub Tamanho_LostFocus()

    Set CTB_Segmentos_objGrid1.objControle = Tamanho
    
    Call Grid_Campo_Libera_Foco(CTB_Segmentos_objGrid1)
    
End Sub

Private Sub Delimitador_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(CTB_Segmentos_objGrid1)

End Sub

Private Sub Delimitador_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, CTB_Segmentos_objGrid1)
    
End Sub

Private Sub Delimitador_LostFocus()

    Set CTB_Segmentos_objGrid1.objControle = Delimitador
    
    Call Grid_Campo_Libera_Foco(CTB_Segmentos_objGrid1)

End Sub

Private Sub Preenchimento_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(CTB_Segmentos_objGrid1)

End Sub

Private Sub Preenchimento_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, CTB_Segmentos_objGrid1)
    
End Sub

Private Sub Preenchimento_LostFocus()

    Set CTB_Segmentos_objGrid1.objControle = Preenchimento
    
    Call Grid_Campo_Libera_Foco(CTB_Segmentos_objGrid1)

End Sub

Private Sub GridSegmentos_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(CTB_Segmentos_objGrid1, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(CTB_Segmentos_objGrid1, iAlterado)
    End If
    

End Sub

Private Sub GridSegmentos_GotFocus()
    
    Call Grid_Recebe_Foco(CTB_Segmentos_objGrid1)

End Sub

Private Sub GridSegmentos_EnterCell()
    
    Call Grid_Entrada_Celula(CTB_Segmentos_objGrid1, iAlterado)

End Sub

Private Sub GridSegmentos_LeaveCell()
    
    Call Saida_Celula(CTB_Segmentos_objGrid1)
    
End Sub

Private Sub GridSegmentos_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, CTB_Segmentos_objGrid1)
    
End Sub

Private Sub GridSegmentos_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, CTB_Segmentos_objGrid1, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(CTB_Segmentos_objGrid1, iAlterado)
    End If

End Sub

Private Sub GridSegmentos_LostFocus()
    
    Call Grid_Libera_Foco(CTB_Segmentos_objGrid1)

End Sub

Private Sub GridSegmentos_RowColChange()

    Call Grid_RowColChange(CTB_Segmentos_objGrid1)
       
End Sub

Private Sub GridSegmentos_Scroll()

    Call Grid_Scroll(CTB_Segmentos_objGrid1)
    
End Sub

Private Function CTB_Segmentos_Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_CTB_Segmentos_Saida_Celula

   lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case objGridInt.objGrid.Col

            Case COL_TIPO
                
                lErro = Saida_Celula_Tipo(objGridInt)
                If lErro <> SUCESSO Then Error 44750

            Case COL_TAMANHO
                
                lErro = Saida_Celula_Tamanho(objGridInt)
                If lErro <> SUCESSO Then Error 44751

            Case COL_DELIMITADOR
            
                lErro = Saida_Celula_Delimitador(objGridInt)
                If lErro <> SUCESSO Then Error 44752
                
                
             Case COL_PREENCHIMENTO
             
                lErro = Saida_Celula_Preenchimento(objGridInt)
                If lErro <> SUCESSO Then Error 44753
                   

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 44754

    End If

    CTB_Segmentos_Saida_Celula = SUCESSO

    Exit Function

Erro_CTB_Segmentos_Saida_Celula:

    CTB_Segmentos_Saida_Celula = Err

    Select Case Err
        
        Case 44750, 44751, 44752, 44753
        
        Case 44754
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175821)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Delimitador(objGridInt As AdmGrid) As Long
'faz a critica da celula delimitador do grid que está deixando de ser a corrente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Delimitador

    Set objGridInt.objControle = Delimitador
    
    Delimitador.Text = Trim(Delimitador.Text)
    
    If Len(Delimitador.Text) > 0 And GridSegmentos.Row - GridSegmentos.FixedRows = objGridInt.iLinhasExistentes Then
       objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If
                
    If Len(Trim(Delimitador.Text)) > 1 Then Error 44755
    
    If Delimitador.Text = SEPARADOR Then Error 44756
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 44757

    Saida_Celula_Delimitador = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Delimitador:

    Saida_Celula_Delimitador = Err
    
    Select Case Err
    
        Case 44755
            Call Rotina_Erro(vbOKOnly, "ERRO_SAIDA_DELIMITADOR", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                 
        Case 44756
            Call Rotina_Erro(vbOKOnly, "ERRO_SAIDA_DELIMITADOR", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 44757
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175822)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Tamanho(objGridInt As AdmGrid) As Long
'faz a critica da celula Tamanho do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Tamanho

    Set objGridInt.objControle = Tamanho
    
    'verifica se foi preenchido o tamanho
    If Len(Trim(Tamanho.Text)) <> 0 Then
        
        'verifica se o tamanho é maior do que zero
        If CInt(Tamanho.Text) < 1 Then Error 44758
        
        If Len(Trim(Tamanho.Text)) > 0 And GridSegmentos.Row - GridSegmentos.FixedRows = objGridInt.iLinhasExistentes Then
           objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
               
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 44759

    Saida_Celula_Tamanho = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Tamanho:

    Saida_Celula_Tamanho = Err
    
    Select Case Err
    
        Case 44758
             Call Grid_Trata_Erro_Saida_Celula(objGridInt)
             Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_TAMANHO_INVALIDO", Err)
    
        Case 44759
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175823)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Tipo(objGridInt As AdmGrid) As Long
'faz a critica da celula tipo do grid que está deixando de ser a corrente
'se for preenchido, o numero de linhas existentes no grid aumenta uma unidade

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Tipo

    Set objGridInt.objControle = Tipo
    
    If Len(Trim(Tipo.Text)) > 0 And GridSegmentos.Row - GridSegmentos.FixedRows = objGridInt.iLinhasExistentes Then
       objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 44760

    Saida_Celula_Tipo = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Tipo:

    Saida_Celula_Tipo = Err
    
    Select Case Err
    
        Case 44760
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175824)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Preenchimento(objGridInt As AdmGrid) As Long
'faz a critica da celula preenchimento do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Preenchimento

    Set objGridInt.objControle = Preenchimento
                
    If Len(Trim(Preenchimento.Text)) > 0 And GridSegmentos.Row - GridSegmentos.FixedRows = objGridInt.iLinhasExistentes Then
       objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 44761

    Saida_Celula_Preenchimento = SUCESSO
    
Exit Function
    
Erro_Saida_Celula_Preenchimento:

    Saida_Celula_Preenchimento = Err
    
    Select Case Err
    
        Case 44761
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175825)
        
    End Select

    Exit Function

End Function

Function Inicializa_Grid_Segmento() As Long
   
    'tela em questão
    Set CTB_Segmentos_objGrid1.objForm = Me
    
    'titulos do grid
    CTB_Segmentos_objGrid1.colColuna.Add ("Segmento")
    CTB_Segmentos_objGrid1.colColuna.Add ("Tipo")
    CTB_Segmentos_objGrid1.colColuna.Add ("Tamanho")
    CTB_Segmentos_objGrid1.colColuna.Add ("Delimitador")
    CTB_Segmentos_objGrid1.colColuna.Add ("Preenchimento")
    
   'campos de edição do grid
    CTB_Segmentos_objGrid1.colCampo.Add (Tipo.Name)
    CTB_Segmentos_objGrid1.colCampo.Add (Tamanho.Name)
    CTB_Segmentos_objGrid1.colCampo.Add (Delimitador.Name)
    CTB_Segmentos_objGrid1.colCampo.Add (Preenchimento.Name)
    
    CTB_Segmentos_objGrid1.objGrid = GridSegmentos
   
    'todas as linhas do grid
    CTB_Segmentos_objGrid1.objGrid.Rows = 10
    
    'linhas visiveis do grid sem contar com as linhas fixas
    CTB_Segmentos_objGrid1.iLinhasVisiveis = 6
    
    CTB_Segmentos_objGrid1.objGrid.ColWidth(0) = 1000
    
    CTB_Segmentos_objGrid1.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    Call Grid_Inicializa(CTB_Segmentos_objGrid1)
    
    Inicializa_Grid_Segmento = SUCESSO
    
End Function

Function Grid_Segmentos(colSegmentos As Collection) As Long

Dim iIndice1 As Integer
Dim objSegmento As ClassSegmento
Dim lErro As Long

On Error GoTo Erro_Grid_Segmentos

    'percorre todas as linhas do grid
    For iIndice1 = 1 To CTB_Segmentos_objGrid1.iLinhasExistentes

        Set objSegmento = New ClassSegmento
                     
        'inclui o Formato Conta em objSegmento
        objSegmento.sCodigo = SEGMENTO_CONTA
            
        'inclui o nivel em objSegmento
        objSegmento.iNivel = iIndice1
        
        'verifica se foi preenchido o campo tipo
        If Len(Trim(GridSegmentos.TextMatrix(iIndice1, COL_TIPO))) = 0 Then Error 44743
        
        'inclui o tipo em objSegmento
        objSegmento.iTipo = gobjColTipoSegmento.TipoSegmento(GridSegmentos.TextMatrix(iIndice1, COL_TIPO))
         
        'verifica se foi preenchido o campo tamanho
        If Len(Trim(GridSegmentos.TextMatrix(iIndice1, COL_TAMANHO))) = 0 Then Error 44744
        
        'inclui o tamanho em objSegmento
        objSegmento.iTamanho = CInt(GridSegmentos.TextMatrix(iIndice1, COL_TAMANHO))
        
        'verifica se foi preenchido o campo delimitador
        If Len(Trim(GridSegmentos.TextMatrix(iIndice1, COL_DELIMITADOR))) = 0 Then Error 44745
        
        'inclui o delimitador em objSegmento
        objSegmento.sDelimitador = GridSegmentos.TextMatrix(iIndice1, COL_DELIMITADOR)
        
        'verifica se foi preenchido o campo preenchimento
        If Len(Trim(GridSegmentos.TextMatrix(iIndice1, COL_PREENCHIMENTO))) = 0 Then Error 44746
        
        'inclui o preenchimento em objSegmento
        objSegmento.iPreenchimento = gobjColPreenchimento.Preenchimento(GridSegmentos.TextMatrix(iIndice1, COL_PREENCHIMENTO))
        
        'Armazena o objeto objSegmento na coleção colSegmento
        colSegmentos.Add objSegmento

    Next

    Grid_Segmentos = SUCESSO

    Exit Function

Erro_Grid_Segmentos:

    Grid_Segmentos = Err

    Select Case Err

        Case 44743
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_TIPO_NAO_PREENCHIDO", Err)
            GridSegmentos.Row = iIndice1
            GridSegmentos.Col = COL_TIPO
            GridSegmentos.SetFocus
        
        Case 44744
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_TAMANHO_NAO_PREENCHIDO", Err)
            GridSegmentos.Row = iIndice1
            GridSegmentos.Col = COL_TAMANHO
            GridSegmentos.SetFocus

        Case 44745
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DELIMITADOR_NAO_PREENCHIDO", Err)
            GridSegmentos.Row = iIndice1
            GridSegmentos.Col = COL_DELIMITADOR
            GridSegmentos.SetFocus

        Case 44746
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_PREENCHIMENTO_NAO_PREENCHIDO", Err)
            GridSegmentos.Row = iIndice1
            GridSegmentos.Col = COL_PREENCHIMENTO
            GridSegmentos.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175826)

    End Select

    Exit Function

End Function

Private Function CTB_Segmentos_Gravar_Registro() As Long

Dim iTamanho As Integer
Dim iTotalTamanho As Integer
Dim iLinha As Integer
Dim lErro As Long
Dim colSegmento As New Collection
Dim objSegmento As ClassSegmento


On Error GoTo Erro_CTB_Segmentos_Gravar_Registro

    lErro = Valida_Step(MODULO_CONTABILIDADE)

    If lErro = SUCESSO Then
    
        'Preenche a colSegmentos com as informacoes contidas no Grid
        lErro = Grid_Segmentos(colSegmento)
        If lErro <> SUCESSO Then Error 44762
    
        'Grava os registros na tabela Segmentos com os dados de colSegmentos
        lErro = CF("Segmento_Grava_Conta_Trans", colSegmento)
        If lErro <> SUCESSO Then Error 44763
        
    End If
        
    CTB_Segmentos_Gravar_Registro = SUCESSO
       
    Exit Function
    
Erro_CTB_Segmentos_Gravar_Registro:
    
    CTB_Segmentos_Gravar_Registro = Err
    
    Select Case Err

        Case 44762, 44763
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175827)

    End Select

    Exit Function
    
End Function


'=============================================================================
' FIM TELA DE SEGMENTOS CONTABILIDADE
'=============================================================================

'=============================================================================
' TELA DE CONFIGURACAO CONTABILIDADE
'=============================================================================


Private Sub CclContabil_Click()
  
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CclExtra_Click()
  
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DocPorExercicio_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DocPorPeriodo_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CTB_Inicializa_Config()

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_CTB_Inicializa_Config
        
    lErro = Valida_Step(MODULO_CONTABILIDADE)
        
    If lErro = SUCESSO Then
    
        'inicializar os tipos de conta
        For iIndice = 1 To gobjColTipoConta.Count
            TipoConta.AddItem gobjColTipoConta.Item(iIndice).sDescricao
        Next
        
        TipoConta.ListIndex = 0
        
        'inicializar as naturezas de conta
        For iIndice = 1 To gobjColNaturezaConta.Count
            Natureza.AddItem gobjColNaturezaConta.Item(iIndice).sDescricao
        Next
        
        Natureza.ListIndex = 0
        
        iAlterado = 0
        
        CTB_Config_iFrameAtual = 0

    End If
    
    Exit Sub
    
Erro_CTB_Inicializa_Config:

    Select Case Err
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175828)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub LotePorExercicio_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LotePorPeriodo_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Opcoes_Click()

    If Opcoes.SelectedItem.Index - 1 <> CTB_Config_iFrameAtual Then
    
        If TabStrip_PodeTrocarTab(CTB_Config_iFrameAtual, Opcoes, Me, 0) <> SUCESSO Then Exit Sub
        
        Frame1(Opcoes.SelectedItem.Index - 1).Visible = True
        Frame1(CTB_Config_iFrameAtual).Visible = False
        CTB_Config_iFrameAtual = Opcoes.SelectedItem.Index - 1
        
        Select Case CTB_Config_iFrameAtual
        
            Case TAB_Identificacao
                Me.HelpContextID = IDH_MODULO_CONTABILIDADE_ID
                                
            Case TAB_Ccl
                Me.HelpContextID = IDH_MODULO_CONTABILIDADE_CENTRO_CUSTO_LUCRO
            
            Case TAB_Valores
                Me.HelpContextID = IDH_MODULO_CONTABILIDADE_VALORES_INICIAIS
            
        End Select
        
    End If

End Sub

Private Sub SemCcl_Click()
  
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoConta_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoConta_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Natureza_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Natureza_Click()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Function Leitura_Configuracao(objConfiguracao As ClassConfiguracao) As Long
'faz a leitura das marcacoes da tela de ConfiguracaoSetup

    'le a marcacao do Lote
    If LotePorPeriodo.Value Then
        objConfiguracao.iLotePorPeriodo = LOTE_INICIALIZADO_POR_PERIODO
    Else
        objConfiguracao.iLotePorPeriodo = LOTE_INICIALIZADO_POR_EXERCICIO
    End If
    
    'le a marcacao do Documento
    If DocPorPeriodo.Value Then
        objConfiguracao.iDocPorPeriodo = DOC_INICIALIZADO_POR_PERIODO
    Else
        objConfiguracao.iDocPorPeriodo = DOC_INICIALIZADO_POR_EXERCICIO
    End If
    
    'le a marcacao do Centro de Custo/Lucro
    If SemCcl.Value Then
        objConfiguracao.iUsoCcl = CCL_NAO_USA
    ElseIf CclContabil.Value Then
        objConfiguracao.iUsoCcl = CCL_USA_CONTABIL
    ElseIf CclExtra.Value Then
        objConfiguracao.iUsoCcl = CCL_USA_EXTRACONTABIL
    End If
    
    objConfiguracao.iTipoContaDefault = gobjColTipoConta.TipoConta(TipoConta.Text)
    objConfiguracao.iNaturezaDefault = gobjColNaturezaConta.NaturezaConta(Natureza.Text)

    Leitura_Configuracao = SUCESSO

End Function

Private Function CTB_Config_Gravar_Registro() As Long

Dim lErro As Long
Dim objConfiguracao As New ClassConfiguracao

On Error GoTo Erro_CTB_Config_Gravar_Registro

    lErro = Valida_Step(MODULO_CONTABILIDADE)

    If lErro = SUCESSO Then

        Call Leitura_Configuracao(objConfiguracao)
            
        'Grava os registros na tabela Configuracao com os dados de objConfiguracao
        lErro = CF("ConfiguracaoSetup_Altera_Trans", objConfiguracao)
        If lErro <> SUCESSO Then Error 44764
    
    End If
    
    CTB_Config_Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_CTB_Config_Gravar_Registro:
    
    CTB_Config_Gravar_Registro = Err
    
    Select Case Err

        Case 44764
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175829)

    End Select

    Exit Function
    
End Function

'=============================================================================
' FIM TELA DE CONFIGURACAO CONTABILIDADE
'=============================================================================

'=============================================================================
' TELA DE EXERCICIO CONTABILIDADE
'=============================================================================

Private Sub Exercicio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CTB_Inicializa_Exercicio()

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_CTB_Inicializa_Exercicio

    lErro = Valida_Step(MODULO_CONTABILIDADE)

    If lErro = SUCESSO Then
    
        iAlterado = 0
        
        Set CTB_Exercicio_objGrid1 = New AdmGrid
    
        lErro = Inicializa_Grid_ExercicioTela(CTB_Exercicio_objGrid1)
        If lErro <> SUCESSO Then Error 44766
    
        lErro = Criar_Exercicio()
        If lErro <> SUCESSO Then Error 44767

    End If
    
    Exit Sub

Erro_CTB_Inicializa_Exercicio:

    Select Case Err

        Case 44766, 44767

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175830)

    End Select

    Exit Sub

End Sub

Function Criar_Exercicio() As Long

Dim iUltimoExercicio
Dim objExercicio As New ClassExercicio
Dim dtData As Date
Dim iIndice As Integer

On Error GoTo Erro_Criar_Exercicio

    DataFimExercicio.Enabled = True
    SpinDataFim.Enabled = True
    Periodicidade.Enabled = True
    
    'Colocar a periodicidade = LIVRE
    For iIndice = 0 To Periodicidade.ListCount - 1
        If Periodicidade.ItemData(iIndice) = PERIODICIDADE_LIVRE Then
            Periodicidade.ListIndex = iIndice
            Exit For
        End If
    Next

    NumPeriodos.Enabled = True
    SpinNumPeriodos.Enabled = True
    BotaoGeraPeriodos.Enabled = True
    DataInicioPeriodo.Enabled = True
    CTB_Exercicio_objGrid1.iProibidoIncluir = 0
    CTB_Exercicio_objGrid1.iProibidoExcluir = 0

    NumPeriodos.PromptInclude = False
    NumPeriodos.Text = "1"
    NumPeriodos.PromptInclude = True

    NomeExterno.Text = Format(Date, "yyyy")

    'se não tem nenhum exercicio cadastrada, sua data inicial será 01/01 do ano corrente.
    dtData = CDate("01/01/" & Year(Date))
        
    DataInicioExercicio.Enabled = True
    SpinDataInicio.Enabled = True
        
    DataInicioExercicio.Text = Format(dtData, "dd/mm/yy")

    'data fim = 31/12 do ano da data início
    DataFimExercicio.Text = "31/12/" & Format(dtData, "yy")

    GridPeriodos.TopRow = 1

    Criar_Exercicio = SUCESSO

    Exit Function

Erro_Criar_Exercicio:

    Criar_Exercicio = Err
    
    Select Case Err

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175831)

    End Select

    Exit Function

End Function

Function Inicializa_Grid_ExercicioTela(objGridInt As AdmGrid) As Long
'inicializa o grid de períodos do form ExercicioTela /m


   Set CTB_Exercicio_objGrid1.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("Periodo")
    objGridInt.colColuna.Add ("Nome")
    objGridInt.colColuna.Add ("Data Inicio")

   'campos de edição do grid
    objGridInt.colCampo.Add (NomePeriodo.Name)
    objGridInt.colCampo.Add (DataInicioPeriodo.Name)

    objGridInt.objGrid = GridPeriodos

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 13

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 5

    GridPeriodos.ColWidth(0) = 1000

    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_ExercicioTela = SUCESSO

End Function

Sub GeracaoPeriodos(dtDataInicio As Date, dtDataFim As Date, iPeriodicidade As Integer, iNumPeriodo As Integer, colPeriodos As Collection)
'gera períodos entre as datas de entrada ( dtDataInicio e dtDataFim )
'os períodos são retornados na coleção colPeriodos

Dim iDuracaoPeriodo As Integer

    'determina a duração de cada período
    Select Case iPeriodicidade

        Case PERIODICIDADE_ANUAL
            iDuracaoPeriodo = 12
        Case PERIODICIDADE_BIMENSAL
            iDuracaoPeriodo = 2
        Case PERIODICIDADE_MENSAL
            iDuracaoPeriodo = 1
        Case PERIODICIDADE_QUADRIMESTRAL
            iDuracaoPeriodo = 4
        Case PERIODICIDADE_SEMESTRAL
            iDuracaoPeriodo = 6
        Case PERIODICIDADE_TRIMESTRAL
            iDuracaoPeriodo = 3
        Case PERIODICIDADE_LIVRE
            'Calcula a duracao de periodos de acordo com a quantidade de periodos passada como parametro
            Call CalculaPeriodos_Livre(dtDataInicio, dtDataFim, iNumPeriodo, colPeriodos)
            Exit Sub
    End Select

    'Traz em ColPeriodos todos os periodos calculados
    Call Calcula_Periodos(dtDataInicio, dtDataFim, iDuracaoPeriodo, colPeriodos)

    Exit Sub

End Sub

Sub HabilitaCampos()
'só pode gerar períodos se a periodicidade não for livre
'para periodicidade livre os períodos são determinados por Número de Períodos

Dim iHabBotaoGera As Integer
Dim iHabNumPer As Integer

    iHabBotaoGera = False
    iHabNumPer = False

    'Exercício selecionado e Data Início e Data Fim preenchidos
    If CTB_Exercicio_objGrid1.iProibidoIncluir = 0 And Len(DataInicioExercicio.ClipText) > 0 And Len(DataFimExercicio.ClipText) > 0 Then
    
        iHabBotaoGera = True

        'periodicidade livre permite selecionar o numero de períodos a serem gerados
        If Periodicidade.ItemData(Periodicidade.ListIndex) = PERIODICIDADE_LIVRE Then
            iHabNumPer = True
        End If
        
    End If
    
    'habilita ou desabilita os campos
    BotaoGeraPeriodos.Enabled = iHabBotaoGera
    DataInicioPeriodo.Enabled = iHabBotaoGera
    NumPeriodos.Enabled = iHabNumPer
    SpinNumPeriodos.Enabled = iHabNumPer

End Sub

Function MoveDadosTela_Variaveis(objExercicio As ClassExercicio, colPeriodos As Collection) As Long
'Move os dados do exercicio da tela para objExercicio e os
'dados referebtes ao período para colPeriodos

Dim iIndice As Integer
Dim objPeriodo As ClassPeriodo
Dim lErro As Long

On Error GoTo Erro_MoveDadosTela_Variaveis

    'dados do exercício
    objExercicio.sNomeExterno = NomeExterno.Text
    objExercicio.dtDataInicio = CDate(DataInicioExercicio.Text)
    objExercicio.dtDataFim = CDate(DataFimExercicio.Text)
    objExercicio.iNumPeriodos = CTB_Exercicio_objGrid1.iLinhasExistentes
    objExercicio.iExercicio = 1

    'dados dos períodos
    For iIndice = 1 To CTB_Exercicio_objGrid1.iLinhasExistentes

        Set objPeriodo = New ClassPeriodo
        
        objPeriodo.sNomeExterno = GridPeriodos.TextMatrix(iIndice, GRID_NOME_COL)
        If objPeriodo.sNomeExterno = "" Then Error 44768

        If GridPeriodos.TextMatrix(iIndice, GRID_DATAINI_COL) = "" Then Error 44769
        
        objPeriodo.dtDataInicio = CDate(GridPeriodos.TextMatrix(iIndice, GRID_DATAINI_COL))

        'se for o primeiro periodo, verifica se a data inicio coincide com a data inicio do exercicio.
        If iIndice = 1 And objPeriodo.dtDataInicio <> objExercicio.dtDataInicio Then Error 44770

        If objPeriodo.dtDataInicio > objExercicio.dtDataFim Then Error 44771
        
        If iIndice = CTB_Exercicio_objGrid1.iLinhasExistentes Then
            objPeriodo.dtDataFim = DataFimExercicio.Text
        Else
            If GridPeriodos.TextMatrix(iIndice + 1, GRID_DATAINI_COL) = "" Then Error 44772
            
            objPeriodo.dtDataFim = (CDate(GridPeriodos.TextMatrix(iIndice + 1, GRID_DATAINI_COL)) - 1)
        End If

        colPeriodos.Add objPeriodo

    Next

    MoveDadosTela_Variaveis = SUCESSO

    Exit Function

Erro_MoveDadosTela_Variaveis:

    MoveDadosTela_Variaveis = Err

    Select Case Err

        Case 44768
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_PERIODO_VAZIO", Err)

        Case 44769
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIO_PERIODO_VAZIA", Err, iIndice)

        Case 44770
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINI_PRIMEIRO_PERIODO", Err, CStr(objPeriodo.dtDataInicio), CStr(objExercicio.dtDataInicio))

        Case 44771
             Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FORA_EXERCICIO", Err, CStr(objPeriodo.dtDataInicio), CStr(objExercicio.dtDataInicio), CStr(objExercicio.dtDataFim))
             
        Case 44772
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIO_PERIODO_VAZIA", Err, iIndice + 1)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175832)

    End Select

    Exit Function

End Function

Private Sub DataFimExercicio_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub BotaoGeraPeriodos_Click()

Dim lErro As Long
Dim colPeriodos As New Collection
Dim objPeriodo As ClassPeriodo
Dim iConta As Integer
Dim dtDataFim As Date
Dim dtDataInicio As Date
Dim iPeriodicidade As Integer
Dim iNumPeriodo As Integer

On Error GoTo Erro_BotaoGeraPeriodos_Click

    dtDataInicio = CDate(DataInicioExercicio.Text)
    dtDataFim = CDate(DataFimExercicio.Text)
    
    iPeriodicidade = Periodicidade.ItemData(Periodicidade.ListIndex)

    If Trim(NumPeriodos.Text) = "" Then
        iNumPeriodo = 1
        NumPeriodos.PromptInclude = False
        NumPeriodos.Text = "1"
        NumPeriodos.PromptInclude = True
    Else
        iNumPeriodo = CInt(NumPeriodos.Text)
    End If

    'Chama a Rotina que gera todos os periodos
    Call GeracaoPeriodos(dtDataInicio, dtDataFim, iPeriodicidade, iNumPeriodo, colPeriodos)

    'preenche o grid com os períodos gerados
    lErro = PreencheGridPeriodos(colPeriodos)
    If lErro <> SUCESSO Then Error 44773

    GridPeriodos.TopRow = 1

    Exit Sub

Erro_BotaoGeraPeriodos_Click:

    Select Case Err

        Case 44773

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175833)

    End Select

    Exit Sub

End Sub

Private Sub DataFimExercicio_LostFocus()

Dim lErro As Long

On Error GoTo Erro_DataFimExercicio_LostFocus

    'verifica se a data final está vazia
    If Len(DataFimExercicio.ClipText) = 0 Then Error 44774

    'verifica se a data final é válida
    lErro = Data_Critica(DataFimExercicio.Text)
    If lErro <> SUCESSO Then Error 44775

    'data inicial não pode ser maior que a data final
    If CDate(DataInicioExercicio.Text) > CDate(DataFimExercicio.Text) Then Error 44776

    Call HabilitaCampos

    Exit Sub

Erro_DataFimExercicio_LostFocus:

    Select Case Err

        Case 44774
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FINAL_EXERCICIO_NAO_PREENCHIDA", Err)
            DataFimExercicio.SetFocus
             
        Case 44775
            DataFimExercicio.SetFocus

        Case 44776
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FINAL_EXERCICIO_MENOR", Err, DataFimExercicio.Text, DataInicioExercicio.Text)
            DataFimExercicio.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175834)

    End Select

    Exit Sub

End Sub

Private Sub DataInicioExercicio_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataInicioExercicio_LostFocus()

Dim lErro As Long

On Error GoTo Erro_DataInicioExercicio_LostFocus

    'verifica se a data inicial está vazia
    If Len(DataInicioExercicio.ClipText) = 0 Then Error 44777

    'verifica se a data inicial é válida
    lErro = Data_Critica(DataInicioExercicio.Text)
    If lErro <> SUCESSO Then Error 44778

    'data inicial não pode ser maior que a data final
    If CDate(DataInicioExercicio.Text) > CDate(DataFimExercicio.Text) Then Error 44779

    Call HabilitaCampos

    Exit Sub

Erro_DataInicioExercicio_LostFocus:

    Select Case Err

        Case 44777
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_EXERCICIO_NAO_PREENCHIDA", Err)
            DataInicioExercicio.SetFocus

        Case 44778
            DataInicioExercicio.SetFocus

        Case 44779
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_EXERCICIO_MAIOR", Err, DataInicioExercicio.Text, DataFimExercicio.Text)
            DataInicioExercicio.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175835)

    End Select

    Exit Sub

End Sub

Private Sub NomeExterno_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NomePeriodo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NomePeriodo_GotFocus()
    Call Grid_Campo_Recebe_Foco(CTB_Exercicio_objGrid1)
End Sub

Private Sub DataInicioPeriodo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, CTB_Exercicio_objGrid1)
End Sub

Private Sub DataInicioPeriodo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataInicioPeriodo_GotFocus()
        Call Grid_Campo_Recebe_Foco(CTB_Exercicio_objGrid1)
End Sub

Function PreencheGridPeriodos(colPeriodos As Collection) As Long
'preenche o grid com os períodos passados na coleção colPeriodos

Dim iIndice As Integer
Dim objPeriodo As ClassPeriodo

On Error GoTo Erro_PreencheGridPeriodos

    'Limpa o grid
    Call Grid_Limpa(CTB_Exercicio_objGrid1)

    CTB_Exercicio_objGrid1.iLinhasExistentes = colPeriodos.Count

    'preenche o grid com os dados retornados na coleção colPeriodos
    For iIndice = 1 To colPeriodos.Count

        Set objPeriodo = colPeriodos.Item(iIndice)
        
        GridPeriodos.TextMatrix(iIndice, GRID_NOME_COL) = objPeriodo.sNomeExterno
        GridPeriodos.TextMatrix(iIndice, GRID_DATAINI_COL) = Format(objPeriodo.dtDataInicio, "dd/mm/yyyy")

    Next

    PreencheGridPeriodos = SUCESSO

    Exit Function

Erro_PreencheGridPeriodos:

    PreencheGridPeriodos = Err

    Select Case Err

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175836)

    End Select

    Exit Function

End Function

Private Sub DataInicioPeriodo_LostFocus()

    Set CTB_Exercicio_objGrid1.objControle = DataInicioPeriodo

    Call Grid_Campo_Libera_Foco(CTB_Exercicio_objGrid1)

End Sub

Private Sub NomePeriodo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, CTB_Exercicio_objGrid1)
End Sub

Private Sub NomePeriodo_LostFocus()

    Set CTB_Exercicio_objGrid1.objControle = NomePeriodo
    Call Grid_Campo_Libera_Foco(CTB_Exercicio_objGrid1)

End Sub

Private Sub GridPeriodos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(CTB_Exercicio_objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(CTB_Exercicio_objGrid1, iAlterado)
    End If

End Sub

Private Sub GridPeriodos_GotFocus()
    Call Grid_Recebe_Foco(CTB_Exercicio_objGrid1)
End Sub

Private Sub GridPeriodos_EnterCell()
    Call Grid_Entrada_Celula(CTB_Exercicio_objGrid1, iAlterado)
End Sub

Private Sub GridPeriodos_LeaveCell()
    Call Saida_Celula(CTB_Exercicio_objGrid1)
End Sub

Private Sub GridPeriodos_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, CTB_Exercicio_objGrid1)
End Sub

Private Sub GridPeriodos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, CTB_Exercicio_objGrid1, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(CTB_Exercicio_objGrid1, iAlterado)
    End If

End Sub

Private Sub GridPeriodos_LostFocus()
    Call Grid_Libera_Foco(CTB_Exercicio_objGrid1)
End Sub

Private Sub GridPeriodos_RowColChange()
    Call Grid_RowColChange(CTB_Exercicio_objGrid1)
End Sub

Private Sub GridPeriodos_Scroll()
    Call Grid_Scroll(CTB_Exercicio_objGrid1)
End Sub

Function CTB_Exercicio_Testa() As Long

Dim lErro As Long
Dim objExercicio As New ClassExercicio
Dim colPeriodos As New Collection

On Error GoTo Erro_CTB_Exercicio_Testa

    lErro = Valida_Step(MODULO_CONTABILIDADE)
    
    If lErro = SUCESSO Then

        'verifica se o nome do exercício foi preenchido
        If NomeExterno.Text = "" Then Error 44652
    
        'Verifica se pelo menos um periodo foi gerado
        If CTB_Exercicio_objGrid1.iLinhasExistentes = 0 Then Error 44653
        
        'move os dados da tela para as variáveis
        lErro = MoveDadosTela_Variaveis(objExercicio, colPeriodos)
        If lErro <> SUCESSO Then Error 44654

    End If
    
    CTB_Exercicio_Testa = SUCESSO
    
    Exit Function

Erro_CTB_Exercicio_Testa:

    CTB_Exercicio_Testa = Err

    Select Case Err

        Case 44652
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_EXERCICIO_VAZIO", Err)
            NomeExterno.SetFocus
            
        Case 44653
            Call Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_SEM_PERIODO", Err, NomeExterno.Text)

        Case 44654

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175837)

    End Select

    Exit Function

End Function

Private Function CTB_Exercicio_Gravar_Registro() As Long
'grava os dados do exercicio em questão /m

Dim lErro As Long
Dim objExercicio As New ClassExercicio
Dim colPeriodos As New Collection

On Error GoTo Erro_CTB_Exercicio_Gravar_Registro

    lErro = Valida_Step(MODULO_CONTABILIDADE)

    If lErro = SUCESSO Then
    
        'verifica se o nome do exercício foi preenchido
        If NomeExterno.Text = "" Then Error 44779
    
        'Verifica se pelo menos um periodo foi gerado
        If CTB_Exercicio_objGrid1.iLinhasExistentes = 0 Then Error 44780
        
        'move os dados da tela para as variáveis
        lErro = MoveDadosTela_Variaveis(objExercicio, colPeriodos)
        If lErro <> SUCESSO Then Error 44781
    
        lErro = CF("Exercicio_Grava_Trans", objExercicio, colPeriodos)
        If lErro <> SUCESSO Then Error 44782
        
    End If
    
    CTB_Exercicio_Gravar_Registro = SUCESSO
    
    Exit Function

Erro_CTB_Exercicio_Gravar_Registro:

    CTB_Exercicio_Gravar_Registro = Err

    Select Case Err

        Case 44779
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_EXERCICIO_VAZIO", Err)
            NomeExterno.SetFocus
            
        Case 44780
            Call Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_SEM_PERIODO", Err, NomeExterno.Text)

        Case 44781, 44782

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175838)

    End Select

    Exit Function

End Function

Function Critica_Campo_DataInicioPeriodo(sData As String) As Long
'faz a crítica da data do início do período de acordo com o exercício e com o período anterior

Dim lErro As Long
Dim objExercicio As New ClassExercicio
Dim dtData As Date

On Error GoTo Erro_Critica_Campo_DataInicioPeriodo

    lErro = Data_Critica(sData)
    If lErro <> SUCESSO Then Error 44783
    
    dtData = CDate(sData)

    objExercicio.dtDataInicio = CDate(DataInicioExercicio.Text)
    objExercicio.dtDataFim = CDate(DataFimExercicio.Text)

    'verifica se data dtData está dentro do exercício
    If dtData < objExercicio.dtDataInicio Or dtData > objExercicio.dtDataFim Then Error 44784

    'verifica se o período é maior que 1
    If GridPeriodos.Row > 1 Then

        'se a data da linha anterior estiver preenchida
        If Len(Trim(GridPeriodos.TextMatrix(GridPeriodos.Row - 1, GRID_DATAINI_COL))) > 0 Then

            'data início deve ser maior que a data início do período anterior
            If dtData <= CDate(GridPeriodos.TextMatrix(GridPeriodos.Row - 1, GRID_DATAINI_COL)) Then Error 44785
            
        End If

    Else

        'data início = data início exercício, quando período = 1
        If dtData <> objExercicio.dtDataInicio Then

            GridPeriodos.TextMatrix(GridPeriodos.Row, GRID_DATAINI_COL) = Format(objExercicio.dtDataInicio, "dd/mm/yyyy")
            Error 44786

        End If

    End If
    
    If GridPeriodos.Row < CTB_Exercicio_objGrid1.iLinhasExistentes Then
    
        'se a data do periodo seguinte estiver preenchido
        If Len(Trim(GridPeriodos.TextMatrix(GridPeriodos.Row + 1, GRID_DATAINI_COL))) > 0 Then
    
            If dtData >= CDate(GridPeriodos.TextMatrix(GridPeriodos.Row + 1, GRID_DATAINI_COL)) Then Error 44787
            
        End If

    End If

    Critica_Campo_DataInicioPeriodo = SUCESSO

    Exit Function

Erro_Critica_Campo_DataInicioPeriodo:

    Critica_Campo_DataInicioPeriodo = Err

    Select Case Err

        Case 44783, 13639

        Case 44784
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FORA_EXERCICIO", Err, CStr(dtData), DataInicioExercicio.Text, DataFimExercicio.Text)

        Case 44785
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINI_PERIODO_MENOR_PERIODO_ANT", Err, CStr(dtData), GridPeriodos.TextMatrix(GridPeriodos.Row - 1, GRID_DATAINI_COL))

        Case 44786
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINI_PRIMEIRO_PERIODO", Err, CStr(dtData), DataInicioExercicio.Text)

        Case 44787
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINI_PERIODO_MAIOR_PERIODO_SEG", Err, CStr(dtData), GridPeriodos.TextMatrix(GridPeriodos.Row + 1, GRID_DATAINI_COL))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175839)

    End Select

    Exit Function

End Function

Private Function CTB_Exercicio_Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente /m

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_CTB_Exercicio_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case objGridInt.objGrid.Col

            Case GRID_NOME_COL
                lErro = Saida_Celula_Nome(objGridInt)
                If lErro <> SUCESSO Then Error 44788

            Case GRID_DATAINI_COL
                lErro = Saida_Celula_DataIni(objGridInt)
                If lErro <> SUCESSO Then Error 44789

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 44790

    End If

    CTB_Exercicio_Saida_Celula = SUCESSO

    Exit Function

Erro_CTB_Exercicio_Saida_Celula:

    CTB_Exercicio_Saida_Celula = Err

    Select Case Err

        Case 44788, 44789

        Case 44790
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175840)

    End Select

    Exit Function

End Function

Private Sub NumPeriodos_LostFocus()

Dim iNumPeriodos As Integer

On Error GoTo Erro_NumPeriodos_LostFocus

    If Len(Trim(NumPeriodos.Text)) = 0 Then Exit Sub

    iNumPeriodos = CInt(NumPeriodos.Text)

    'número de períodos não pode ser maior que NUM_MAX_PERIODOS
    If iNumPeriodos > NUM_MAX_PERIODOS Then

        NumPeriodos.Text = CStr(NUM_MAX_PERIODOS)
        Error 44791

    'número de períodos não pode ser menor que 1
    ElseIf iNumPeriodos < 1 Then

            NumPeriodos.PromptInclude = False
            NumPeriodos.Text = "1"
            NumPeriodos.PromptInclude = True
            Error 44792

    End If

    Exit Sub

Erro_NumPeriodos_LostFocus:

    Select Case Err

        Case 44791, 44792
            Call Rotina_Erro(vbOKOnly, "ERRO_NUM_PERIODO_INVALIDO", Err, NUM_MAX_PERIODOS)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175841)

    End Select

    Exit Sub

End Sub

Private Sub Periodicidade_click()
    iAlterado = REGISTRO_ALTERADO
    Call HabilitaCampos
End Sub

Private Sub SpinDataFim_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_SpinDataFim_UpClick

    DataFimExercicio.SetFocus

    If Len(Trim(DataFimExercicio.ClipText)) > 0 Then

        sData = DataFimExercicio.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 44792

        DataFimExercicio.PromptInclude = False
        DataFimExercicio.Text = sData
        DataFimExercicio.PromptInclude = True

    End If

    Exit Sub

Erro_SpinDataFim_UpClick:

    Select Case Err

        Case 44792

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175842)

    End Select

    Exit Sub

End Sub

Private Sub SpinDataFim_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_SpinDataFim_DownClick

    DataFimExercicio.SetFocus

    If Len(Trim(DataFimExercicio.ClipText)) > 0 Then

        sData = DataFimExercicio.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 44793

        DataFimExercicio.PromptInclude = False
        DataFimExercicio.Text = sData
        DataFimExercicio.PromptInclude = True
        
    End If

    Exit Sub

Erro_SpinDataFim_DownClick:

    Select Case Err

        Case 44793

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175843)

    End Select

    Exit Sub

End Sub

Private Sub SpinDataInicio_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_SpinDataInicio_UpClick

    DataInicioExercicio.SetFocus

    If Len(Trim(DataInicioExercicio.ClipText)) > 0 Then

        sData = DataInicioExercicio.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 44794

        DataInicioExercicio.Text = sData
    End If

    Exit Sub

Erro_SpinDataInicio_UpClick:

    Select Case Err

        Case 44794

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175844)

    End Select

    Exit Sub

End Sub

Private Sub SpinDataInicio_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_SpinDataInicio_DownClick

    DataInicioExercicio.SetFocus

    If Len(Trim(DataInicioExercicio.ClipText)) > 0 Then

        sData = DataInicioExercicio.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 44795

        DataInicioExercicio.Text = sData
    End If

    Exit Sub

Erro_SpinDataInicio_DownClick:

    Select Case Err

        Case 44795

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175845)

    End Select

    Exit Sub

End Sub

Private Sub SpinNumPeriodos_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub SpinNumPeriodos_DownClick()

Dim iNumPeriodos As Integer

On Error GoTo Erro_SpinNumPeriodos_DownClick

    NumPeriodos.SetFocus

    If Len(Trim(NumPeriodos.Text)) = 0 Then
        iNumPeriodos = 0
    Else
        iNumPeriodos = CInt(NumPeriodos.Text)
    End If

    'número de períodos não pode ser menor que 1
    If iNumPeriodos = 1 Then Error 44796

    NumPeriodos.PromptInclude = False
    NumPeriodos.Text = CStr(iNumPeriodos - 1)
    NumPeriodos.PromptInclude = True

    Exit Sub

Erro_SpinNumPeriodos_DownClick:

    Select Case Err

        Case 44796

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175846)

    End Select

    Exit Sub

End Sub

Private Sub SpinNumPeriodos_UpClick()

Dim iNumPeriodos As Integer

On Error GoTo Erro_SpinNumPeriodos_UpClick

    NumPeriodos.SetFocus

    If Len(Trim(NumPeriodos.Text)) = 0 Then
        iNumPeriodos = 0
    Else
        iNumPeriodos = CInt(NumPeriodos.Text)
    End If

    'número de períodos não pode ser maior que NUM_MAX_PERIODOS
    If iNumPeriodos = NUM_MAX_PERIODOS Then Error 44797

    NumPeriodos.PromptInclude = False
    NumPeriodos.Text = CStr(iNumPeriodos + 1)
    NumPeriodos.PromptInclude = True

    Exit Sub

Erro_SpinNumPeriodos_UpClick:

    Select Case Err

        Case 44797

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175847)

    End Select

    Exit Sub

End Sub

Private Function CalculaPeriodos_Livre(dtDataInicio As Date, dtDataFim As Date, iNumPeriodo As Integer, colPeriodos As Collection) As Long
' Calcula os Periodos quando a Periodicidade é livre

Dim iTotalDias As Integer
Dim iDuracaoPeriodo As Integer
Dim iPeriodo As Integer
Dim dtDataIniPer As Date
Dim dtDataFimPer As Date
Dim objPeriodo As ClassPeriodo

On Error GoTo Erro_CalculaPeriodos_Livre

    'Calcula o total de dias do Exercicio
    iTotalDias = (dtDataFim - dtDataInicio) + 1

    'Verifica se o numero de periodos requeridos é maior do que o total de dias do Exercicio
    If iTotalDias < iNumPeriodo Then Error 44798

    iDuracaoPeriodo = iTotalDias \ iNumPeriodo

    dtDataIniPer = dtDataInicio
    dtDataFimPer = dtDataInicio + iDuracaoPeriodo - 1

    For iPeriodo = 1 To iNumPeriodo

        Set objPeriodo = New ClassPeriodo

        objPeriodo.dtDataInicio = dtDataIniPer
        objPeriodo.dtDataFim = dtDataFimPer
        objPeriodo.sNomeExterno = "Periodo " & CStr(iPeriodo)

        colPeriodos.Add objPeriodo

        dtDataIniPer = dtDataFimPer + 1
        dtDataFimPer = dtDataIniPer + iDuracaoPeriodo - 1

    Next

    If iTotalDias Mod iNumPeriodo > 0 Then
        Call Rotina_Aviso(vbOKOnly, "AVISO_ULTIMO_PERIODO_MAIOR")
    End If

    CalculaPeriodos_Livre = SUCESSO

    Exit Function

Erro_CalculaPeriodos_Livre:

    CalculaPeriodos_Livre = Err

    Select Case Err

        Case 44798
            Call Rotina_Erro(vbOKOnly, "ERRO_TOTAL_PERIODOS_MAIOR_TOTAL_DIAS", Err, iNumPeriodo, iTotalDias)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175848)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Nome(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Nome

    Set objGridInt.objControle = NomePeriodo

    'testa se o nome do periodo está preenchido
    If Len(Trim(NomePeriodo.Text)) = 0 And GridPeriodos.Row - GridPeriodos.FixedRows < objGridInt.iLinhasExistentes Then Error 44799

    If Len(Trim(NomePeriodo.Text)) > 0 And GridPeriodos.Row - GridPeriodos.FixedRows = objGridInt.iLinhasExistentes Then
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If

    'verifica se já não existe outro período no exercício com o mesmo nome
    For iIndice = 1 To objGridInt.iLinhasExistentes

        If iIndice <> GridPeriodos.Row Then
            If Trim(NomePeriodo.Text) = Trim(GridPeriodos.TextMatrix(iIndice, GRID_NOME_COL)) Then Error 44800
        End If

    Next
    'critica da coluna 1 fim

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 44801

    Saida_Celula_Nome = SUCESSO

    Exit Function

Erro_Saida_Celula_Nome:

    Saida_Celula_Nome = Err

    Select Case Err

        Case 44799
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_PERIODO_VAZIO", Err, GridPeriodos.Row)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 44800
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_PERIODO_JA_USADO", Err, Trim(NomePeriodo.Text), iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 44801
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175849)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataIni(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataIni

    Set objGridInt.objControle = DataInicioPeriodo
    
    If Len(DataInicioPeriodo.ClipText) > 0 Then

        lErro = Critica_Campo_DataInicioPeriodo(DataInicioPeriodo.Text)
        If lErro <> SUCESSO Then Error 44802
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 44803

    Saida_Celula_DataIni = SUCESSO

    Exit Function

Erro_Saida_Celula_DataIni:

    Saida_Celula_DataIni = Err

    Select Case Err

        Case 44802, 44803
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175850)

    End Select

    Exit Function

End Function

Sub Calcula_Periodos(dtDataInicio As Date, dtDataFim As Date, iDuracaoPeriodo As Integer, colPeriodos As Collection)

Dim iDiaFinal As Integer
Dim iMesFinal As Integer
Dim iAnoFinal As Integer
Dim iTotPeriodo As Integer
Dim objPeriodo As ClassPeriodo
Dim iNumPeriodo As Integer
Dim dtDataFinal As Date

On Error GoTo Erro_CalCula_Periodos

    iTotPeriodo = 1
    iNumPeriodo = 1
    iMesFinal = Month(dtDataInicio)
    iAnoFinal = Year(dtDataInicio)
    
    Do While iMesFinal > iTotPeriodo * iDuracaoPeriodo
        iTotPeriodo = iTotPeriodo + 1
    Loop
    
    iMesFinal = iTotPeriodo * iDuracaoPeriodo
    
    iDiaFinal = Dias_Mes(iMesFinal, iAnoFinal)
    
    dtDataFinal = CDate(CStr(iDiaFinal) & "/" & CStr(iMesFinal) & "/" & CStr(iAnoFinal))

    If dtDataFim < dtDataFinal Then dtDataFinal = dtDataFim
    
    'adiciona um período ( nome, data inicial e final ) à coleção
    Set objPeriodo = New ClassPeriodo
    objPeriodo.sNomeExterno = "Periodo " & CStr(iNumPeriodo)
    objPeriodo.dtDataInicio = dtDataInicio
    objPeriodo.dtDataFim = dtDataFinal
    objPeriodo.iFechado = 0
    colPeriodos.Add objPeriodo
    
    Do While dtDataFim > dtDataFinal And iNumPeriodo < NUM_MAX_PERIODOS
    
        iNumPeriodo = iNumPeriodo + 1
    
        dtDataInicio = dtDataFinal + 1
        
        iMesFinal = Month(dtDataInicio)
        iAnoFinal = Year(dtDataInicio)
    
        iTotPeriodo = 1
    
        Do While iMesFinal > iTotPeriodo * iDuracaoPeriodo
            iTotPeriodo = iTotPeriodo + 1
        Loop

        iMesFinal = iTotPeriodo * iDuracaoPeriodo
        
        iDiaFinal = Dias_Mes(iMesFinal, iAnoFinal)
    
        dtDataFinal = CDate(CStr(iDiaFinal) & "/" & CStr(iMesFinal) & "/" & CStr(iAnoFinal))

        If dtDataFim < dtDataFinal Or iNumPeriodo = NUM_MAX_PERIODOS Then dtDataFinal = dtDataFim
    
        'adiciona um período ( nome, data inicial e final ) à coleção
        Set objPeriodo = New ClassPeriodo
        objPeriodo.sNomeExterno = "Periodo " & CStr(iNumPeriodo)
        objPeriodo.dtDataInicio = dtDataInicio
        objPeriodo.dtDataFim = dtDataFinal
        objPeriodo.iFechado = 0
        colPeriodos.Add objPeriodo

    Loop

    Exit Sub

Erro_CalCula_Periodos:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175851)

    End Select

    Exit Sub

End Sub
'=============================================================================
' FIM TELA DE EXERCICIO CONTABILIDADE
'=============================================================================

'=============================================================================
' TELA DE CONFIGURACAO TESOURARIA
'=============================================================================

Private Sub TES_Inicializa_Config()
           
Dim lErro As Long

On Error GoTo Erro_TES_Inicializa_Config
    
    lErro = Valida_Step(MODULO_TESOURARIA)
    
    If lErro = SUCESSO Then
    
        'Checa Aglutina lançamentos por dia
        If gobjTES.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA Then
            ListaConfigura.Selected(TESCONFIG_AGLUTINA_LANCAM_POR_DIA) = True
        Else
            ListaConfigura.Selected(TESCONFIG_AGLUTINA_LANCAM_POR_DIA) = False
        End If
        
        If giTipoVersao = VERSAO_FULL Then
            'Checa Exige Preenchimento Data de Saída
            If gobjTES.iGeraLoteAutomatico = GERA_LOTE_AUTOMATICO Then
                ListaConfigura.Selected(TESCONFIG_GERA_LOTE_AUTOMATICO) = True
            Else
                ListaConfigura.Selected(TESCONFIG_GERA_LOTE_AUTOMATICO) = False
            End If
        End If
    End If
    
    Exit Sub

Erro_TES_Inicializa_Config:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175852)

    End Select

    Exit Sub
    
End Sub
    
Private Function TES_Config_Gravar_Registro() As Long

Dim lErro As Long
    
On Error GoTo Erro_TES_Config_Gravar_Registro
    
    lErro = Valida_Step(MODULO_TESOURARIA)
    
    If lErro = SUCESSO Then
    
        If ListaConfigura.Selected(TESCONFIG_AGLUTINA_LANCAM_POR_DIA) = True Then
            gobjTES.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA
        Else
            gobjTES.iAglutinaLancamPorDia = NAO_AGLUTINA_LANCAM_POR_DIA
        End If
        
        If giTipoVersao = VERSAO_FULL Then

            If ListaConfigura.Selected(TESCONFIG_GERA_LOTE_AUTOMATICO) = True Then
                gobjTES.iGeraLoteAutomatico = GERA_LOTE_AUTOMATICO
            Else
                gobjTES.iGeraLoteAutomatico = NAO_GERA_LOTE_AUTOMATICO
            End If
        ElseIf giTipoVersao = VERSAO_LIGHT Then
            gobjTES.iGeraLoteAutomatico = NAO_GERA_LOTE_AUTOMATICO
        End If
        
        'chama gobjFAT.Gravar()
        lErro = gobjTES.Gravar_Trans()
        If lErro <> SUCESSO Then Error 44765
    
    End If
    
    TES_Config_Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_TES_Config_Gravar_Registro:

    TES_Config_Gravar_Registro = Err

    Select Case Err
    
        Case 44765
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175853)
            
    End Select

    Exit Function
    
End Function
'=============================================================================
' FIM TELA DE CONFIGURACAO TESOURARIA
'=============================================================================


'=============================================================================
' TELA DE SEGMENTOS GERAL
'=============================================================================

Private Function SGE_Segmentos_Testa() As Long
'verifica se os segmentos estão preenchidos.

Dim lErro As Long
Dim colSegmento As Collection


On Error GoTo Erro_SGE_Segmentos_Testa

    lErro = Valida_Step(SISTEMA_SGE)

    If lErro = SUCESSO Then

        'Salva o formato selecionado em SGE_Segmentos_colSegmento
        lErro = SGE_Segmentos_Salva_Formato()
        If lErro <> SUCESSO Then Error 44805

        For Each colSegmento In SGE_Segmentos_colSegmento
    
            If colSegmento.Count = 0 Then Error 44806
            
        Next
        
    End If

    SGE_Segmentos_Testa = SUCESSO
    
    Exit Function
    
Erro_SGE_Segmentos_Testa:

    SGE_Segmentos_Testa = Err

    Select Case Err
    
        Case 44805

        Case 44806
            Call Rotina_Erro(vbOKOnly, "ERRO_SEGMENTO_VAZIO", Err)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175854)

    End Select

    Exit Function

End Function

Private Function SGE_Inicializa_Segmentos() As Long

Dim iIndice As Integer
Dim colSegmento As Collection
Dim sDescricao As String
Dim lErro As Long
Dim objSegmento As New ClassSegmento

On Error GoTo Erro_SGE_Inicializa_Segmentos

    lErro = Valida_Step(SISTEMA_SGE)

    If lErro = SUCESSO Then
    
        Set SGE_Segmentos_objGrid1 = New AdmGrid
               
        'inicializacao do grid
        Call SGE_Inicializa_Grid_Segmento
        
        'inicializar os formatos
        For iIndice = 1 To gobjColCodigoSegmento.Count
            If gobjColCodigoSegmento.Item(iIndice).sCodigo = SEGMENTO_PRODUTO Or gobjColCodigoSegmento.Item(iIndice).sCodigo = SEGMENTO_CCL Then
                Formato.AddItem gobjColCodigoSegmento.Item(iIndice).sDescricao
            End If
        Next
        
        'inicializar os tipos
        For iIndice = 1 To gobjColTipoSegmento.Count
            Tipo1.AddItem gobjColTipoSegmento.Item(iIndice).sDescricao
        Next
    
        'inicializar os preenchimentos
        For iIndice = 1 To gobjColPreenchimento.Count
            Preenchimento1.AddItem gobjColPreenchimento.Item(iIndice).sDescricao
        Next
    
        'coloca a descricao referente ao centro de custo em sDescricao
        sDescricao = gobjColCodigoSegmento.Descricao(SEGMENTO_CCL)
    
        Set colSegmento = New Collection
    
        'preenche o obj com o formato corrente para usar em Segmento_Le_Codigo
        objSegmento.sCodigo = SEGMENTO_CCL
    
        'preenche toda colecao(colSegmento) com o formato do centro de custo
        lErro = CF("Segmento_Le_Codigo", objSegmento, colSegmento)
        If lErro <> SUCESSO Then Error 44833
    
        SGE_Segmentos_colSegmento.Add colSegmento, SEGMENTO_CCL
        
        Set colSegmento = New Collection
    
        'preenche o obj com o formato corrente para usar em Segmento_Le_Codigo
        objSegmento.sCodigo = SEGMENTO_PRODUTO
    
        'preenche toda colecao(colSegmento) com o formato do produto
        lErro = CF("Segmento_Le_Codigo", objSegmento, colSegmento)
        If lErro <> SUCESSO Then Error 44834
    
        SGE_Segmentos_colSegmento.Add colSegmento, SEGMENTO_PRODUTO
        
        'mostra o formato centro de custo como formato inicial
        For iIndice = 0 To Formato.ListCount - 1
            If Formato.List(iIndice) = sDescricao Then
                Formato.ListIndex = iIndice
                Exit For
            End If
        Next
        
        Call SGE_Carga_Grid
    
    End If
    
    SGE_Inicializa_Segmentos = SUCESSO
    
    Exit Function
    
Erro_SGE_Inicializa_Segmentos:

    SGE_Inicializa_Segmentos = Err
    
    Select Case Err
    
        Case 44833, 44834
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175855)
    
    End Select
    
    Exit Function
    
End Function

Private Sub Formato_Click()

Dim lErro As Long
Dim objSegmento As ClassSegmento
Dim colSegmento As Collection
Dim iIndice As Integer

On Error GoTo Erro_Formato_Click

    If Len(SGE_Segmentos_sCodigo) = 0 Then
        SGE_Segmentos_sCodigo = gobjColCodigoSegmento.Codigo(Formato.Text)
        Exit Sub
    End If

    'Se trocou o formato selecionado
    If Formato.Text <> gobjColCodigoSegmento.Descricao(SGE_Segmentos_sCodigo) Then

        'Salva o formato selecionado em CTB_Segmentos_colSegmento
        lErro = SGE_Segmentos_Salva_Formato()
        If lErro <> SUCESSO Then Error 44646
       
        Call SGE_Carga_Grid

    End If

    Exit Sub

Erro_Formato_Click:

    Select Case Err

        Case 44646
            For iIndice = 0 To Formato.ListCount - 1
                If Formato.List(iIndice) = gobjColCodigoSegmento.Descricao(SGE_Segmentos_sCodigo) Then
                    Formato.ListIndex = iIndice
                    Exit For
                End If
            Next

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175856)

    End Select

    Exit Sub

End Sub

Sub SGE_Carga_Grid()

Dim colSegmento As Collection
Dim objSegmento As ClassSegmento

    'Preenche o grid com o novo formato selecionado
    Set colSegmento = SGE_Segmentos_colSegmento.Item(gobjColCodigoSegmento.Codigo(Formato.Text))

    Call Grid_Limpa(SGE_Segmentos_objGrid1)

    SGE_Segmentos_objGrid1.iLinhasExistentes = 0

    'preenche todo o grid da tabela segmento
    For Each objSegmento In colSegmento

        'coloca o tipo no grid da tela
        GridSegmentos1.TextMatrix(objSegmento.iNivel, COL_TIPO) = gobjColTipoSegmento.Descricao(objSegmento.iTipo)

        'coloca o tamanho no grid da tela
        GridSegmentos1.TextMatrix(objSegmento.iNivel, COL_TAMANHO) = objSegmento.iTamanho

        'coloca os delimitadores no grid da tela
        GridSegmentos1.TextMatrix(objSegmento.iNivel, COL_DELIMITADOR) = objSegmento.sDelimitador

        'coloca o preenchimento no grid da tela
        GridSegmentos1.TextMatrix(objSegmento.iNivel, COL_PREENCHIMENTO) = gobjColPreenchimento.Descricao(objSegmento.iPreenchimento)

        SGE_Segmentos_objGrid1.iLinhasExistentes = SGE_Segmentos_objGrid1.iLinhasExistentes + 1

    Next

    SGE_Segmentos_sCodigo = gobjColCodigoSegmento.Codigo(Formato.Text)

End Sub

Function SGE_Segmentos_Salva_Formato() As Long

Dim lErro As Long
Dim iTamanho As Integer
Dim iTotalTamanho As Integer
Dim colSegmento As New Collection
Dim iLinha As Integer

On Error GoTo Erro_SGE_Segmentos_Salva_Formato

    If Len(SGE_Segmentos_sCodigo) > 0 Then

        'percorre as linhas da coluna tamanho
        For iLinha = 1 To SGE_Segmentos_objGrid1.iLinhasExistentes
        
            'verifica se nao foi preenchido o tamanho
            If Len(Trim(GridSegmentos1.TextMatrix(iLinha, COL_TAMANHO))) = 0 Then Error 44807
            'soma o valor total da coluna tamanho no grid
            iTotalTamanho = iTotalTamanho + CInt(GridSegmentos1.TextMatrix(iLinha, COL_TAMANHO))
    
        Next
                  
        'verifica se tamanho do produto ultrapassou tamanho pre_definido
        If SGE_Segmentos_sCodigo = SEGMENTO_PRODUTO And iTotalTamanho > STRING_PRODUTO Then
            Error 44808
                
        'verifica se tamanho ccl ultrapassou tamanho pre_definido
        ElseIf SGE_Segmentos_sCodigo = SEGMENTO_CCL And iTotalTamanho > STRING_CCL Then
            Error 44809
        End If
    
        'Preenche a colSegmentos com as informacoes contidas no Grid
        lErro = Grid_Segmentos1(colSegmento)
        If lErro <> SUCESSO Then Error 44810

        SGE_Segmentos_colSegmento.Remove (SGE_Segmentos_sCodigo)
    
        SGE_Segmentos_colSegmento.Add colSegmento, SGE_Segmentos_sCodigo
    
    End If

    SGE_Segmentos_Salva_Formato = SUCESSO

    Exit Function

Erro_SGE_Segmentos_Salva_Formato:

    SGE_Segmentos_Salva_Formato = Err

    Select Case Err

        Case 44807
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_TAMANHO_NAO_PREENCHIDO", Err)
            
        Case 44808
            Call Rotina_Erro(vbOKOnly, "ERRO_SEGMENTO_PRODUTO_MAIOR_PERMITIDO", Err, iTotalTamanho, STRING_PRODUTO)
        
        Case 44809
            Call Rotina_Erro(vbOKOnly, "ERRO_SEGMENTO_CCL_MAIOR_PERMITIDO", Err, iTotalTamanho, STRING_CCL)
        
        Case 44810
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175857)

    End Select

    Exit Function

End Function

Private Function SGE_Segmentos_Gravar_Registro() As Long

Dim lErro As Long
Dim colSegmentos As Collection

On Error GoTo Erro_SGE_Segmentos_Gravar_Registro

    lErro = Valida_Step(SISTEMA_SGE)

    If lErro = SUCESSO Then
        Set colSegmentos = SGE_Segmentos_colSegmento.Item(SEGMENTO_PRODUTO)
        
        'Grava os registros na tabela Segmentos com os dados de colSegmentos
        lErro = CF("Segmento_Grava_Produto_Trans", colSegmentos)
        If lErro <> SUCESSO Then Error 44812
            
        Set colSegmentos = SGE_Segmentos_colSegmento.Item(SEGMENTO_CCL)
            
        'Grava os registros na tabela Segmentos com os dados de colSegmentos
        lErro = CF("Segmento_Grava_Ccl_Trans", colSegmentos)
        If lErro <> SUCESSO Then Error 44813
        
    End If
    
    SGE_Segmentos_Gravar_Registro = SUCESSO
       
    Exit Function
    
Erro_SGE_Segmentos_Gravar_Registro:
    
    SGE_Segmentos_Gravar_Registro = Err
    
    Select Case Err

        Case 44812, 44813

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175858)

    End Select

    Exit Function
    
End Function

Private Sub Tipo1_Change()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tipo1_Click()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tamanho1_Change()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tamanho1_Click()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Delimitador1_Change()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Delimitador1_Click()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Preenchimento1_Change()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Preenchimento1_Click()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tipo1_GotFocus()

    Call Grid_Campo_Recebe_Foco(SGE_Segmentos_objGrid1)
    
End Sub

Private Sub Tipo1_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, SGE_Segmentos_objGrid1)

End Sub

Private Sub Tipo1_LostFocus()

    Set SGE_Segmentos_objGrid1.objControle = Tipo1
    
    Call Grid_Campo_Libera_Foco(SGE_Segmentos_objGrid1)

End Sub

Private Sub Tamanho1_GotFocus()

    Call Grid_Campo_Recebe_Foco(SGE_Segmentos_objGrid1)
    
End Sub

Private Sub Tamanho1_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, SGE_Segmentos_objGrid1)

End Sub

Private Sub Tamanho1_LostFocus()

    Set SGE_Segmentos_objGrid1.objControle = Tamanho1
    
    Call Grid_Campo_Libera_Foco(SGE_Segmentos_objGrid1)
    
End Sub

Private Sub Delimitador1_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(SGE_Segmentos_objGrid1)

End Sub

Private Sub Delimitador1_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, SGE_Segmentos_objGrid1)
    
End Sub

Private Sub Delimitador1_LostFocus()

    Set SGE_Segmentos_objGrid1.objControle = Delimitador1
    
    Call Grid_Campo_Libera_Foco(SGE_Segmentos_objGrid1)

End Sub

Private Sub Preenchimento1_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(SGE_Segmentos_objGrid1)

End Sub

Private Sub Preenchimento1_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, SGE_Segmentos_objGrid1)
    
End Sub

Private Sub Preenchimento1_LostFocus()

    Set SGE_Segmentos_objGrid1.objControle = Preenchimento1
    
    Call Grid_Campo_Libera_Foco(SGE_Segmentos_objGrid1)

End Sub

Private Sub GridSegmentos1_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(SGE_Segmentos_objGrid1, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(SGE_Segmentos_objGrid1, iAlterado)
    End If
    

End Sub

Private Sub GridSegmentos1_GotFocus()
    
    Call Grid_Recebe_Foco(SGE_Segmentos_objGrid1)

End Sub

Private Sub GridSegmentos1_EnterCell()
    
    Call Grid_Entrada_Celula(SGE_Segmentos_objGrid1, iAlterado)

End Sub

Private Sub GridSegmentos1_LeaveCell()
    
    Call Saida_Celula(SGE_Segmentos_objGrid1)
    
End Sub

Private Sub GridSegmentos1_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, SGE_Segmentos_objGrid1)
    
End Sub

Private Sub GridSegmentos1_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, SGE_Segmentos_objGrid1, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(SGE_Segmentos_objGrid1, iAlterado)
    End If

End Sub

Private Sub GridSegmentos1_LostFocus()
    
    Call Grid_Libera_Foco(SGE_Segmentos_objGrid1)

End Sub

Private Sub GridSegmentos1_RowColChange()

    Call Grid_RowColChange(SGE_Segmentos_objGrid1)
       
End Sub

Private Sub GridSegmentos1_Scroll()

    Call Grid_Scroll(SGE_Segmentos_objGrid1)
    
End Sub

Private Function SGE_Segmentos_Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_SGE_Segmentos_Saida_Celula

   lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case objGridInt.objGrid.Col

            Case COL_TIPO
                
                lErro = Saida_Celula_Tipo1(objGridInt)
                If lErro <> SUCESSO Then Error 44816

            Case COL_TAMANHO
                
                lErro = Saida_Celula_Tamanho1(objGridInt)
                If lErro <> SUCESSO Then Error 44817

            Case COL_DELIMITADOR
            
                lErro = Saida_Celula_Delimitador1(objGridInt)
                If lErro <> SUCESSO Then Error 44818
                
                
             Case COL_PREENCHIMENTO
             
                lErro = Saida_Celula_Preenchimento1(objGridInt)
                If lErro <> SUCESSO Then Error 44819

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 44820

    End If

    SGE_Segmentos_Saida_Celula = SUCESSO

    Exit Function

Erro_SGE_Segmentos_Saida_Celula:

    SGE_Segmentos_Saida_Celula = Err

    Select Case Err
        
        Case 44816, 44817, 44818, 44819
        
        Case 44820
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175859)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Delimitador1(objGridInt As AdmGrid) As Long
'faz a critica da celula delimitador do grid que está deixando de ser a corrente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Delimitador1

    Set objGridInt.objControle = Delimitador1
    
    Delimitador1.Text = Trim(Delimitador1.Text)
    
    If Len(Delimitador1.Text) > 0 And GridSegmentos1.Row - GridSegmentos1.FixedRows = objGridInt.iLinhasExistentes Then
       objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If
                
    If Len(Trim(Delimitador1.Text)) > 1 Then Error 44821
    
    If Delimitador1.Text = SEPARADOR Then Error 44822
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 44823

    Saida_Celula_Delimitador1 = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Delimitador1:

    Saida_Celula_Delimitador1 = Err
    
    Select Case Err
    
        Case 44821
            Call Rotina_Erro(vbOKOnly, "ERRO_SAIDA_DELIMITADOR", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                 
        Case 44822
            Call Rotina_Erro(vbOKOnly, "ERRO_SAIDA_DELIMITADOR", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 44823
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175860)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Tamanho1(objGridInt As AdmGrid) As Long
'faz a critica da celula Tamanho do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Tamanho1

    Set objGridInt.objControle = Tamanho1
    
    'verifica se foi preenchido o tamanho
    If Len(Trim(Tamanho1.Text)) <> 0 Then
        
        'verifica se o tamanho é maior do que zero
        If CInt(Tamanho1.Text) < 1 Then Error 44824
        
        If Len(Trim(Tamanho1.Text)) > 0 And GridSegmentos1.Row - GridSegmentos1.FixedRows = objGridInt.iLinhasExistentes Then
           objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
               
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 44825

    Saida_Celula_Tamanho1 = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Tamanho1:

    Saida_Celula_Tamanho1 = Err
    
    Select Case Err
    
        Case 44824
             Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_TAMANHO_INVALIDO", Err)
             Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case 44825
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175861)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Tipo1(objGridInt As AdmGrid) As Long
'faz a critica da celula tipo do grid que está deixando de ser a corrente
'se for preenchido, o numero de linhas existentes no grid aumenta uma unidade

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Tipo1

    Set objGridInt.objControle = Tipo1
    
    If Len(Trim(Tipo1.Text)) > 0 And GridSegmentos1.Row - GridSegmentos1.FixedRows = objGridInt.iLinhasExistentes Then
       objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If
                
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 44826

    Saida_Celula_Tipo1 = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Tipo1:

    Saida_Celula_Tipo1 = Err
    
    Select Case Err
    
        Case 44826
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175862)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_Preenchimento1(objGridInt As AdmGrid) As Long
'faz a critica da celula preenchimento do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Preenchimento1

    Set objGridInt.objControle = Preenchimento1
                
    If Len(Trim(Preenchimento1.Text)) > 0 And GridSegmentos1.Row - GridSegmentos1.FixedRows = objGridInt.iLinhasExistentes Then
       objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 44827

    Saida_Celula_Preenchimento1 = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Preenchimento1:

    Saida_Celula_Preenchimento1 = Err
    
    Select Case Err
    
        Case 44827
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175863)
        
    End Select

    Exit Function

End Function

Function SGE_Inicializa_Grid_Segmento() As Long
   
    'tela em questão
    Set SGE_Segmentos_objGrid1.objForm = Me
    
    'titulos do grid
    SGE_Segmentos_objGrid1.colColuna.Add ("Segmento")
    SGE_Segmentos_objGrid1.colColuna.Add ("Tipo")
    SGE_Segmentos_objGrid1.colColuna.Add ("Tamanho")
    SGE_Segmentos_objGrid1.colColuna.Add ("Delimitador")
    SGE_Segmentos_objGrid1.colColuna.Add ("Preenchimento")
    
   'campos de edição do grid
    SGE_Segmentos_objGrid1.colCampo.Add (Tipo1.Name)
    SGE_Segmentos_objGrid1.colCampo.Add (Tamanho1.Name)
    SGE_Segmentos_objGrid1.colCampo.Add (Delimitador1.Name)
    SGE_Segmentos_objGrid1.colCampo.Add (Preenchimento1.Name)
    
    SGE_Segmentos_objGrid1.objGrid = GridSegmentos1
   
    'todas as linhas do grid
    SGE_Segmentos_objGrid1.objGrid.Rows = 10
    
    'linhas visiveis do grid sem contar com as linhas fixas
    SGE_Segmentos_objGrid1.iLinhasVisiveis = 6
    
    SGE_Segmentos_objGrid1.objGrid.ColWidth(0) = 1000
    
    SGE_Segmentos_objGrid1.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    Call Grid_Inicializa(SGE_Segmentos_objGrid1)
    
    SGE_Inicializa_Grid_Segmento = SUCESSO
    
End Function

Function Grid_Segmentos1(colSegmentos As Collection) As Long

Dim iIndice1 As Integer
Dim objSegmento As ClassSegmento

On Error GoTo Erro_Grid_Segmentos1

    'percorre todas as linhas do grid
    For iIndice1 = 1 To SGE_Segmentos_objGrid1.iLinhasExistentes

        Set objSegmento = New ClassSegmento
                     
        'verifica se foi preenchido o campo formato
        If Len(Trim(Formato.Text)) = 0 Then Error 44828
        
        'inclui o Formato(codigo) em objSegmento
        objSegmento.sCodigo = SGE_Segmentos_sCodigo
              
        'inclui o nivel em objSegmento
        objSegmento.iNivel = iIndice1
        
        'verifica se foi preenchido o campo tipo
        If Len(Trim(GridSegmentos1.TextMatrix(iIndice1, COL_TIPO))) = 0 Then Error 44829
        
        'inclui o tipo em objSegmento
        objSegmento.iTipo = gobjColTipoSegmento.TipoSegmento(GridSegmentos1.TextMatrix(iIndice1, COL_TIPO))
         
        'verifica se foi preenchido o campo tamanho
        If Len(Trim(GridSegmentos1.TextMatrix(iIndice1, COL_TAMANHO))) = 0 Then Error 44830
        
        'inclui o tamanho em objSegmento
        objSegmento.iTamanho = CInt(GridSegmentos1.TextMatrix(iIndice1, COL_TAMANHO))
        
        'verifica se foi preenchido o campo delimitador
        If Len(Trim(GridSegmentos1.TextMatrix(iIndice1, COL_DELIMITADOR))) = 0 Then Error 44831
        
        'inclui o delimitador em objSegmento
        objSegmento.sDelimitador = GridSegmentos1.TextMatrix(iIndice1, COL_DELIMITADOR)
        
        'verifica se foi preenchido o campo preenchimento
        If Len(Trim(GridSegmentos1.TextMatrix(iIndice1, COL_PREENCHIMENTO))) = 0 Then Error 44832
        
        'inclui o preenchimento em objSegmento
        objSegmento.iPreenchimento = gobjColPreenchimento.Preenchimento(GridSegmentos1.TextMatrix(iIndice1, COL_PREENCHIMENTO))
        
        'Armazena o objeto objSegmento na coleção colSegmento
        colSegmentos.Add objSegmento

    Next

    Grid_Segmentos1 = SUCESSO

    Exit Function

Erro_Grid_Segmentos1:

    Grid_Segmentos1 = Err

    Select Case Err

        Case 44828
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_FORMATO_NAO_PREENCHIDO", Err)
            Formato.SetFocus
            
        Case 44829
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_TIPO_NAO_PREENCHIDO", Err)
            GridSegmentos1.Row = iIndice1
            GridSegmentos1.Col = COL_TIPO
            GridSegmentos1.SetFocus
        
        Case 44830
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_TAMANHO_NAO_PREENCHIDO", Err)
            GridSegmentos1.Row = iIndice1
            GridSegmentos1.Col = COL_TAMANHO
            GridSegmentos1.SetFocus

        Case 44831
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DELIMITADOR_NAO_PREENCHIDO", Err)
            GridSegmentos1.Row = iIndice1
            GridSegmentos1.Col = COL_DELIMITADOR
            GridSegmentos1.SetFocus

        Case 44832
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_PREENCHIMENTO_NAO_PREENCHIDO", Err)
            GridSegmentos1.Row = iIndice1
            GridSegmentos1.Col = COL_PREENCHIMENTO
            GridSegmentos1.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175864)

    End Select

    Exit Function

End Function
'=============================================================================
' FIM TELA DE SEGMENTOS GERAL
'=============================================================================

'=============================================================================
' TELA DE CONFIGURACAO CONTAS A PAGAR
'=============================================================================

Private Sub CP_Inicializa_Config()
           
Dim lErro As Long

On Error GoTo Erro_CP_Inicializa_Config
    
    lErro = Valida_Step(MODULO_CONTASAPAGAR)
    
    If lErro = SUCESSO Then
    
        'Checa Aglutina lançamentos por dia
        If gobjCP.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA Then
            ListaConfiguraCP.Selected(CPCONFIG_AGLUTINA_LANCAM_POR_DIA) = True
        Else
            ListaConfiguraCP.Selected(CPCONFIG_AGLUTINA_LANCAM_POR_DIA) = False
        End If
        
        If giTipoVersao = VERSAO_FULL Then
            'Checa Gera Lote Automatico
            If gobjCP.iGeraLoteAutomatico = GERA_LOTE_AUTOMATICO Then
                ListaConfiguraCP.Selected(CPCONFIG_GERA_LOTE_AUTOMATICO) = True
            Else
                ListaConfiguraCP.Selected(CPCONFIG_GERA_LOTE_AUTOMATICO) = False
            End If
        End If
    End If
    
    Exit Sub

Erro_CP_Inicializa_Config:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175865)

    End Select

    Exit Sub
    
End Sub
    
Private Function CP_Config_Gravar_Registro() As Long

Dim lErro As Long
    
On Error GoTo Erro_CP_Config_Gravar_Registro
    
    lErro = Valida_Step(MODULO_CONTASAPAGAR)
    
    If lErro = SUCESSO Then
    
        If ListaConfiguraCP.Selected(CPCONFIG_AGLUTINA_LANCAM_POR_DIA) = True Then
            gobjCP.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA
        Else
            gobjCP.iAglutinaLancamPorDia = NAO_AGLUTINA_LANCAM_POR_DIA
        End If
        
        If giTipoVersao = VERSAO_FULL Then
            If ListaConfiguraCP.Selected(CPCONFIG_GERA_LOTE_AUTOMATICO) = True Then
                gobjCP.iGeraLoteAutomatico = GERA_LOTE_AUTOMATICO
            Else
                gobjCP.iGeraLoteAutomatico = NAO_GERA_LOTE_AUTOMATICO
            End If
        
        ElseIf giTipoVersao = VERSAO_LIGHT Then
            gobjCP.iGeraLoteAutomatico = NAO_GERA_LOTE_AUTOMATICO
        End If
        
        lErro = gobjCP.Gravar_Trans()
        If lErro <> SUCESSO Then Error 41809
    
    End If
    
    CP_Config_Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_CP_Config_Gravar_Registro:

    CP_Config_Gravar_Registro = Err

    Select Case Err
    
        Case 41809
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175866)
            
    End Select

    Exit Function
    
End Function
'=============================================================================
' FIM TELA DE CONFIGURACAO CONTAS A PAGAR
'=============================================================================

'=============================================================================
' TELA DE CONFIGURACAO CONTAS A RECEBER
'=============================================================================

Private Sub CR_Inicializa_Config()
           
Dim lErro As Long

On Error GoTo Erro_CR_Inicializa_Config
    
    lErro = Valida_Step(MODULO_CONTASARECEBER)
    
    If lErro = SUCESSO Then
    
        'Checa Aglutina lançamentos por dia
        If gobjCR.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA Then
            ListaConfiguraCR.Selected(CRCONFIG_AGLUTINA_LANCAM_POR_DIA) = True
        Else
            ListaConfiguraCR.Selected(CRCONFIG_AGLUTINA_LANCAM_POR_DIA) = False
        End If
        
        If giTipoVersao = VERSAO_FULL Then
            'Checa Gera Lote Automatico
            If gobjCR.iGeraLoteAutomatico = GERA_LOTE_AUTOMATICO Then
                ListaConfiguraCR.Selected(CRCONFIG_GERA_LOTE_AUTOMATICO) = True
            Else
                ListaConfiguraCR.Selected(CRCONFIG_GERA_LOTE_AUTOMATICO) = False
            End If
        End If
        
    End If
    
    Exit Sub

Erro_CR_Inicializa_Config:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175867)

    End Select

    Exit Sub
    
End Sub
    
Private Function CR_Config_Gravar_Registro() As Long

Dim lErro As Long
    
On Error GoTo Erro_CR_Config_Gravar_Registro
    
    lErro = Valida_Step(MODULO_CONTASARECEBER)
    
    If lErro = SUCESSO Then
    
        If ListaConfiguraCR.Selected(CRCONFIG_AGLUTINA_LANCAM_POR_DIA) = True Then
            gobjCR.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA
        Else
            gobjCR.iAglutinaLancamPorDia = NAO_AGLUTINA_LANCAM_POR_DIA
        End If
        
        If giTipoVersao = VERSAO_FULL Then
            If ListaConfiguraCR.Selected(CRCONFIG_GERA_LOTE_AUTOMATICO) = True Then
                gobjCR.iGeraLoteAutomatico = GERA_LOTE_AUTOMATICO
            Else
                gobjCR.iGeraLoteAutomatico = NAO_GERA_LOTE_AUTOMATICO
            End If
        ElseIf giTipoVersao = VERSAO_LIGHT Then
            gobjCR.iGeraLoteAutomatico = NAO_GERA_LOTE_AUTOMATICO
        End If
        
        lErro = gobjCR.Gravar_Trans()
        If lErro <> SUCESSO Then Error 41810
    
    End If
    
    CR_Config_Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_CR_Config_Gravar_Registro:

    CR_Config_Gravar_Registro = Err

    Select Case Err
    
        Case 41810
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175868)
            
    End Select

    Exit Function
    
End Function
'=============================================================================
' FIM TELA DE CONFIGURACAO CONTAS A RECEBER
'=============================================================================

'=============================================================================
' TELA DE CONFIGURACAO FATURAMENTO
'=============================================================================

Private Sub FAT_Inicializa_Config()
           
Dim lErro As Long, iIndice As Integer

On Error GoTo Erro_FAT_Inicializa_Config
    
    lErro = Valida_Step(MODULO_FATURAMENTO)
    
    If lErro = SUCESSO Then
    
        'Checa Aglutina lançamentos por dia
        If gobjFAT.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA Then
            ListaConfiguraFAT.Selected(FATCONFIG_AGLUTINA_LANCAM_POR_DIA) = True
        Else
            ListaConfiguraFAT.Selected(FATCONFIG_AGLUTINA_LANCAM_POR_DIA) = False
        End If
        
        If giTipoVersao = VERSAO_FULL Then
            'Checa Gera Lote Automatico
            If gobjFAT.iGeraLoteAutomatico = GERA_LOTE_AUTOMATICO Then
                ListaConfiguraFAT.Selected(FATCONFIG_GERA_LOTE_AUTOMATICO) = True
            Else
                ListaConfiguraFAT.Selected(FATCONFIG_GERA_LOTE_AUTOMATICO) = False
            End If
        
            'Checa Edita Comissoes PV
            If gobjFAT.iPedidoVendaEditaComissao = PEDVENDA_EDITA_COMISSAO Then
                CheckEditaComissoesPV.Value = MARCADO
            Else
                CheckEditaComissoesPV.Value = DESMARCADO
            End If
        
            'Checa a Reserva Automática
            If gobjFAT.iPedidoReservaAutomatica = PEDVENDA_RESERVA_AUTOMATICA Then
                ReservaAutoPV.Value = True
            Else
                ReservaManPV.Value = True
            End If
        
            'checa se edita comissoes nf
            If gobjFAT.iNFiscalEditaComissao = NFISCAL_EDITA_COMISSAO Then
                CheckEditaComissoesNF.Value = MARCADO
            Else
                CheckEditaComissoesNF.Value = DESMARCADO
            End If
                    
        End If
        
        'Checa a alocação automática
        If gobjFAT.iNFiscalAlocacaoAutomatica = NFISCAL_ALOCA_AUTOMATICA Then
            AlocacaoAutoNF.Value = True
        Else
            AlocacaoManNF.Value = True
        End If
        
        'Carrega os Tipos de Desconto
        lErro = Carrega_TipoDesconto()
        If lErro <> SUCESSO Then Error 56681
        
        Set objGridDescontos = New AdmGrid
        
        lErro = Inicializa_Grid_Descontos(objGridDescontos)
        If lErro <> SUCESSO Then Error 56682
    
        For iIndice = 0 To TipoDesconto.ListCount - 1
            
            If TipoDesconto.ItemData(iIndice) = gobjCRFAT.iDescontoCodigo1 Then
                GridDescontos.TextMatrix(1, iGrid_TipoDesconto_Col) = TipoDesconto.List(iIndice)
                GridDescontos.TextMatrix(1, iGrid_Dias_Col) = gobjCRFAT.iDescontoDias1
                GridDescontos.TextMatrix(1, iGrid_Percentual_Col) = Format(gobjCRFAT.dDescontoPerc1, "Percent")
            End If
            
            If TipoDesconto.ItemData(iIndice) = gobjCRFAT.iDescontoCodigo2 Then
                GridDescontos.TextMatrix(2, iGrid_TipoDesconto_Col) = TipoDesconto.List(iIndice)
                GridDescontos.TextMatrix(2, iGrid_Dias_Col) = gobjCRFAT.iDescontoDias2
                GridDescontos.TextMatrix(2, iGrid_Percentual_Col) = Format(gobjCRFAT.dDescontoPerc2, "Percent")
            End If
            
            If TipoDesconto.ItemData(iIndice) = gobjCRFAT.iDescontoCodigo3 Then
                GridDescontos.TextMatrix(3, iGrid_TipoDesconto_Col) = TipoDesconto.List(iIndice)
                GridDescontos.TextMatrix(3, iGrid_Dias_Col) = gobjCRFAT.iDescontoDias3
                GridDescontos.TextMatrix(3, iGrid_Percentual_Col) = Format(gobjCRFAT.dDescontoPerc3, "Percent")
            End If
            
        Next

    End If
    
    Exit Sub

Erro_FAT_Inicializa_Config:

    Select Case Err

        Case 56681, 56682
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175869)

    End Select

    Exit Sub
    
End Sub
    
Private Function FAT_Parte2_Testa() As Long
'verifica se o frame 2 da configuracao de fat foi preenchido corretamente

Dim lErro As Long, iIndice As Integer, iMaiorLinhaPreenchida As Integer

On Error GoTo Erro_FAT_Parte2_Testa

    lErro = Valida_Step(MODULO_FATURAMENTO)
    
    If lErro = SUCESSO Then
    
        iMaiorLinhaPreenchida = 0
        
        For iIndice = 1 To objGridDescontos.iLinhasExistentes
         
            If Len(Trim(GridDescontos.TextMatrix(iIndice, iGrid_TipoDesconto_Col))) <> 0 _
                Or Len(Trim(GridDescontos.TextMatrix(iIndice, iGrid_Dias_Col))) <> 0 _
                Or Len(Trim(GridDescontos.TextMatrix(iIndice, iGrid_Percentual_Col))) <> 0 Then iMaiorLinhaPreenchida = iIndice
        
        Next
        
        'ou os tres campos de uma linha estao preenchidos ou nenhum deles está.
        'Nao pode haver "buraco" (ex.: preencher linha 2 sem preencher a um)
        For iIndice = 1 To iMaiorLinhaPreenchida
         
            If Len(Trim(GridDescontos.TextMatrix(iIndice, iGrid_TipoDesconto_Col))) = 0 Then Error 56684
            If Len(Trim(GridDescontos.TextMatrix(iIndice, iGrid_Dias_Col))) = 0 Then Error 56685
            If Len(Trim(GridDescontos.TextMatrix(iIndice, iGrid_Percentual_Col))) = 0 Then Error 56686
        
        Next
            
        'se a linha 3 está preenchida
        If Len(Trim(GridDescontos.TextMatrix(3, iGrid_TipoDesconto_Col))) <> 0 Then
        
            'a qtde de dias e o percentual tem que ser estritamente decrescentes no grid para que os descontos fiquem ordenados
            If PercentParaDbl(GridDescontos.TextMatrix(3, iGrid_Percentual_Col)) >= PercentParaDbl(GridDescontos.TextMatrix(2, iGrid_Percentual_Col)) Then Error 56687
            If StrParaInt(GridDescontos.TextMatrix(3, iGrid_Dias_Col)) >= StrParaInt(GridDescontos.TextMatrix(2, iGrid_Dias_Col)) Then Error 56688
            
        End If
        
        'se a linha 2 está preenchida
        If Len(Trim(GridDescontos.TextMatrix(2, iGrid_TipoDesconto_Col))) <> 0 Then
        
            'a qtde de dias e o percentual tem que ser estritamente decrescentes no grid para que os descontos fiquem ordenados
            If PercentParaDbl(GridDescontos.TextMatrix(2, iGrid_Percentual_Col)) >= PercentParaDbl(GridDescontos.TextMatrix(1, iGrid_Percentual_Col)) Then Error 56689
            If StrParaInt(GridDescontos.TextMatrix(2, iGrid_Dias_Col)) >= StrParaInt(GridDescontos.TextMatrix(1, iGrid_Dias_Col)) Then Error 56690
            
        End If
        
    End If
    
    FAT_Parte2_Testa = SUCESSO
     
    Exit Function
    
Erro_FAT_Parte2_Testa:

    FAT_Parte2_Testa = Err
     
    Select Case Err
          
        Case 56687, 56688, 56689, 56690
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_DESCONTO_NAO_ORDEM_DECRESCENTE", Err)
        
        Case 56684
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_DESCONTO_TIPODESCONTO_NAO_PRENCHIDO", Err, iIndice)
        
        Case 56685
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_DESCONTO_DIAS_NAO_PRENCHIDO", Err, iIndice)
        
        Case 56686
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_DESCONTO_PERCENTUAL_NAO_PRENCHIDO", Err, iIndice)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175870)
     
    End Select
     
    Exit Function

End Function

Private Function FAT_Config_Gravar_Registro() As Long

Dim lErro As Long, iIndice As Integer
    
On Error GoTo Erro_FAT_Config_Gravar_Registro
    
    lErro = Valida_Step(MODULO_FATURAMENTO)
    
    If lErro = SUCESSO Then
    
        If ListaConfiguraFAT.Selected(FATCONFIG_AGLUTINA_LANCAM_POR_DIA) = True Then
            gobjFAT.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA
        Else
            gobjFAT.iAglutinaLancamPorDia = NAO_AGLUTINA_LANCAM_POR_DIA
        End If
        
        If giTipoVersao = VERSAO_FULL Then
            
            If ListaConfiguraFAT.Selected(FATCONFIG_GERA_LOTE_AUTOMATICO) = True Then
                gobjFAT.iGeraLoteAutomatico = GERA_LOTE_AUTOMATICO
            Else
                gobjFAT.iGeraLoteAutomatico = NAO_GERA_LOTE_AUTOMATICO
            End If
        
            'Alteracao Daniel ( = TRUE por vbChecked )
            If CheckEditaComissoesPV.Value = vbChecked Then
                gobjFAT.iPedidoVendaEditaComissao = PEDVENDA_EDITA_COMISSAO
            Else
                gobjFAT.iPedidoVendaEditaComissao = PEDVENDA_NAO_EDITA_COMISSAO
            End If
        
            If ReservaAutoPV.Value = True Then
                gobjFAT.iPedidoReservaAutomatica = PEDVENDA_RESERVA_AUTOMATICA
            Else
                gobjFAT.iPedidoReservaAutomatica = PEDVENDA_NAO_RESERVA_AUTOMATICA
            End If

            'Alteracao Daniel ( = TRUE por vbChecked )
            If CheckEditaComissoesNF.Value = vbChecked Then
                gobjFAT.iNFiscalEditaComissao = NFISCAL_EDITA_COMISSAO
            Else
                gobjFAT.iNFiscalEditaComissao = NFISCAL_NAO_EDITA_COMISSAO
            End If
        
        ElseIf giTipoVersao = VERSAO_LIGHT Then
        
            gobjFAT.iGeraLoteAutomatico = NAO_GERA_LOTE_AUTOMATICO
            gobjFAT.iPedidoVendaEditaComissao = PEDVENDA_EDITA_COMISSAO
            gobjFAT.iNFiscalEditaComissao = NFISCAL_EDITA_COMISSAO
        
        End If
        
        
        If AlocacaoAutoNF.Value = True Then
            gobjFAT.iNFiscalAlocacaoAutomatica = NFISCAL_ALOCA_AUTOMATICA
        Else
            gobjFAT.iNFiscalAlocacaoAutomatica = NFISCAL_NAO_ALOCA_AUTOMATICA
        End If
            
        'Verifica se linha do Grid Descontos esta preenchida
        If Len(Trim(GridDescontos.TextMatrix(1, iGrid_Percentual_Col))) <> 0 And Len(Trim(GridDescontos.TextMatrix(1, iGrid_Dias_Col))) <> 0 And Len(Trim(GridDescontos.TextMatrix(1, iGrid_TipoDesconto_Col))) <> 0 Then
            gobjCRFAT.dDescontoPerc1 = PercentParaDbl(GridDescontos.TextMatrix(1, iGrid_Percentual_Col))
            gobjCRFAT.iDescontoDias1 = StrParaInt(GridDescontos.TextMatrix(1, iGrid_Dias_Col))
            gobjCRFAT.iDescontoCodigo1 = Codigo_Extrai(GridDescontos.TextMatrix(1, iGrid_TipoDesconto_Col))
            
            If Len(Trim(GridDescontos.TextMatrix(2, iGrid_Percentual_Col))) <> 0 And Len(Trim(GridDescontos.TextMatrix(2, iGrid_Dias_Col))) <> 0 And Len(Trim(GridDescontos.TextMatrix(2, iGrid_TipoDesconto_Col))) <> 0 Then
                gobjCRFAT.dDescontoPerc2 = PercentParaDbl(GridDescontos.TextMatrix(2, iGrid_Percentual_Col))
                gobjCRFAT.iDescontoDias2 = StrParaInt(GridDescontos.TextMatrix(2, iGrid_Dias_Col))
                gobjCRFAT.iDescontoCodigo2 = Codigo_Extrai(GridDescontos.TextMatrix(2, iGrid_TipoDesconto_Col))
            
                If Len(Trim(GridDescontos.TextMatrix(3, iGrid_Percentual_Col))) <> 0 And Len(Trim(GridDescontos.TextMatrix(3, iGrid_Dias_Col))) <> 0 And Len(Trim(GridDescontos.TextMatrix(3, iGrid_TipoDesconto_Col))) <> 0 Then
                    gobjCRFAT.dDescontoPerc3 = PercentParaDbl(GridDescontos.TextMatrix(3, iGrid_Percentual_Col))
                    gobjCRFAT.iDescontoDias3 = StrParaInt(GridDescontos.TextMatrix(3, iGrid_Dias_Col))
                    gobjCRFAT.iDescontoCodigo3 = Codigo_Extrai(GridDescontos.TextMatrix(3, iGrid_TipoDesconto_Col))
                
                Else 'Se linha 3 nao estiver preenchida
                    gobjCRFAT.dDescontoPerc3 = 0
                    gobjCRFAT.iDescontoDias3 = 0
                    gobjCRFAT.iDescontoCodigo3 = 0
                
                End If
        
            Else    'Se linha 2 nao estiver preenchida
                gobjCRFAT.dDescontoPerc2 = 0
                gobjCRFAT.iDescontoDias2 = 0
                gobjCRFAT.iDescontoCodigo2 = 0
                gobjCRFAT.dDescontoPerc3 = 0
                gobjCRFAT.iDescontoDias3 = 0
                gobjCRFAT.iDescontoCodigo3 = 0
                
            End If
        
        Else    'Se linha 1 nao estiver preenchida
            gobjCRFAT.dDescontoPerc1 = 0
            gobjCRFAT.iDescontoDias1 = 0
            gobjCRFAT.iDescontoCodigo1 = 0
            gobjCRFAT.dDescontoPerc2 = 0
            gobjCRFAT.iDescontoDias2 = 0
            gobjCRFAT.iDescontoCodigo2 = 0
            gobjCRFAT.dDescontoPerc3 = 0
            gobjCRFAT.iDescontoDias3 = 0
            gobjCRFAT.iDescontoCodigo3 = 0
                   
        End If
    
        lErro = gobjFAT.Gravar_Trans()
        If lErro <> SUCESSO Then Error 41811
    
        lErro = gobjCRFAT.Gravar_Trans()
        If lErro <> SUCESSO Then Error 56656
    
    End If
    
    FAT_Config_Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_FAT_Config_Gravar_Registro:

    FAT_Config_Gravar_Registro = Err

    Select Case Err
    
        Case 41811, 56656
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175871)
            
    End Select

    Exit Function
    
End Function

Private Function Inicializa_Grid_Descontos(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Descontos
    
    Set objGridInt.objForm = Me
    
    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Tipo Desconto")
    objGridInt.colColuna.Add ("Dias")
    objGridInt.colColuna.Add ("Percentual")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (TipoDesconto.Name)
    objGridInt.colCampo.Add (Dias.Name)
    objGridInt.colCampo.Add (PercentualDesc.Name)

    'Grid do GridInterno
    objGridInt.objGrid = GridDescontos

    'Colunas do Grid
    iGrid_TipoDesconto_Col = 1
    iGrid_Dias_Col = 2
    iGrid_Percentual_Col = 3

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_DESCONTOS + 1
    objGridInt.iLinhasExistentes = NUM_MAXIMO_DESCONTOS
    
    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridDescontos.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    objGridInt.iProibidoExcluir = PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Descontos = SUCESSO

    Exit Function

End Function

Private Function Carrega_TipoDesconto() As Long
'Carrega os Tipos de Desconto

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_TipoDesconto

    'Lê o código e a descrição de todos os Tipos de Desconto
    lErro = CF("Cod_Nomes_Le", "TiposDeDesconto", "Codigo", "DescReduzida", STRING_TIPOSDEDESCONTO_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 56683

    For Each objCodDescricao In colCodigoDescricao

        'se for desconto do tipo percentual
        If objCodDescricao.iCodigo = Percentual Or objCodDescricao.iCodigo = PERC_ANT_DIA Or objCodDescricao.iCodigo = PERC_ANT_DIA_UTIL Then
        
            'Adiciona o ítem na List da Combo TipoDesconto
            TipoDesconto.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
            TipoDesconto.ItemData(TipoDesconto.NewIndex) = objCodDescricao.iCodigo
            
        End If
        
    Next

    Carrega_TipoDesconto = SUCESSO

    Exit Function

Erro_Carrega_TipoDesconto:

    Carrega_TipoDesconto = Err

    Select Case Err

        Case 56683

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175872)

    End Select

    Exit Function

End Function

Private Sub Dias_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Dias_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDescontos)

End Sub

Private Sub Dias_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDescontos)

End Sub

Private Sub Dias_LostFocus()

    Set objGridDescontos.objControle = Dias
    Call Grid_Campo_Libera_Foco(objGridDescontos)

End Sub

Private Sub PercentualDesc_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub PercentualDesc_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDescontos)

End Sub

Private Sub PercentualDesc_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDescontos)

End Sub

Private Sub PercentualDesc_LostFocus()

    Set objGridDescontos.objControle = PercentualDesc
    Call Grid_Campo_Libera_Foco(objGridDescontos)

End Sub


Private Sub TipoDesconto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoDesconto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDescontos)

End Sub

Private Sub TipoDesconto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDescontos)

End Sub

Private Sub TipoDesconto_LostFocus()

    Set objGridDescontos.objControle = TipoDesconto
    Call Grid_Campo_Libera_Foco(objGridDescontos)

End Sub

Private Function Saida_Celula_GridDescontos(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridDescontos

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then
    
        'Verifica qual a coluna atual do Grid
        Select Case objGridInt.objGrid.Col
        
            Case iGrid_TipoDesconto_Col
                'Faz a crítica do tipo de desconto
                lErro = Saida_Celula_TipoDesconto(objGridInt)
                If lErro <> SUCESSO Then Error 56691
        
            Case iGrid_Dias_Col
                'Faz a crítica de Dias
                lErro = Saida_Celula_Dias(objGridInt)
                If lErro <> SUCESSO Then Error 56692
        
            Case iGrid_Percentual_Col
                'Faz a crítica do Percentual do desconto
                lErro = Saida_Celula_Percentual(objGridInt)
                If lErro <> SUCESSO Then Error 56693
        
        End Select
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 56694

    End If
    
    Saida_Celula_GridDescontos = SUCESSO

    Exit Function

Erro_Saida_Celula_GridDescontos:

    Saida_Celula_GridDescontos = Err

    Select Case Err

        Case 56691, 56692, 56693

        Case 56694
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175873)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TipoDesconto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Tipo Desconto do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_TipoDesconto

    Set objGridInt.objControle = TipoDesconto

    'Verifica se o Tipo foi preenchido
    If Len(Trim(TipoDesconto.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If TipoDesconto.ListIndex = -1 Then

            'Tenta selecioná-lo na combo
            lErro = Combo_Seleciona_Grid(TipoDesconto, iCodigo)
            If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then Error 56695
        
            'Não foi encontrado
            If lErro = 25085 Then Error 56696
            
            If lErro = 25086 Then Error 56697

        End If
            
        'Acrescenta uma linha no Grid se for o caso
        If GridDescontos.Row = objGridInt.iLinhasExistentes And objGridInt.iLinhasExistentes < NUM_MAXIMO_DESCONTOS Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    Else
        
        GridDescontos.TextMatrix(GridDescontos.Row, iGrid_Percentual_Col) = ""
        GridDescontos.TextMatrix(GridDescontos.Row, iGrid_Dias_Col) = ""
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56698

    Saida_Celula_TipoDesconto = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoDesconto:

    Saida_Celula_TipoDesconto = Err

    Select Case Err

        Case 56695
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56696
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO", Err, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56697
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO1", Err, TipoDesconto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56698
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175874)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Percentual(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Percentual

    Set objGridInt.objControle = PercentualDesc

    If Len(Trim(PercentualDesc.ClipText)) > 0 Then

        'Verifica se o Percentual foi preenchido
        lErro = Porcentagem_Critica(PercentualDesc.Text)
        If lErro <> SUCESSO Then Error 56699

        'Formata o Percentual
        PercentualDesc.Text = Format(PercentualDesc.Text, "Fixed")

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56700

    Saida_Celula_Percentual = SUCESSO

    Exit Function

Erro_Saida_Celula_Percentual:

    Saida_Celula_Percentual = Err

    Select Case Err

        Case 56699, 56700
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175875)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Dias(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Data do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Dias

    Set objGridInt.objControle = Dias

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 56702

    Saida_Celula_Dias = SUCESSO

    Exit Function

Erro_Saida_Celula_Dias:

    Saida_Celula_Dias = Err

    Select Case Err

        Case 56702
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175876)

    End Select

    Exit Function

End Function

Private Sub GridDescontos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridDescontos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDescontos, iAlterado)
    End If

End Sub

Private Sub GridDescontos_EnterCell()

    Call Grid_Entrada_Celula(objGridDescontos, iAlterado)

End Sub

Private Sub GridDescontos_GotFocus()

    Call Grid_Recebe_Foco(objGridDescontos)

End Sub

Private Sub GridDescontos_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridDescontos)
     
End Sub

Private Sub GridDescontos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridDescontos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDescontos, iAlterado)
    End If

End Sub

Private Sub GridDescontos_LeaveCell()

    Call Saida_Celula(objGridDescontos)

End Sub

Private Sub GridDescontos_LostFocus()

    Call Grid_Libera_Foco(objGridDescontos)

End Sub

Private Sub GridDescontos_RowColChange()

    Call Grid_RowColChange(objGridDescontos)

End Sub

Private Sub GridDescontos_Scroll()

    Call Grid_Scroll(objGridDescontos)

End Sub
'=============================================================================
' FIM TELA DE CONFIGURACAO FATURAMENTO
'=============================================================================

'=============================================================================
' TELA DE CONFIGURACAO ESTOQUE
'=============================================================================

Private Sub EST_Inicializa_Config()
           
Dim lErro As Long

On Error GoTo Erro_EST_Inicializa_Config
    
    lErro = Valida_Step(MODULO_ESTOQUE)
    
    If lErro = SUCESSO Then
    
        'Checa Aglutina lançamentos por dia
        If gobjEST.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA Then
            ListaConfiguraEST.Selected(ESTCONFIG_AGLUTINA_LANCAM_POR_DIA) = True
        Else
            ListaConfiguraEST.Selected(ESTCONFIG_AGLUTINA_LANCAM_POR_DIA) = False
        End If
        
        If giTipoVersao = VERSAO_FULL Then
            'Checa Gera Lote Automatico
            If gobjEST.iGeraLoteAutomatico = GERA_LOTE_AUTOMATICO Then
                ListaConfiguraEST.Selected(ESTCONFIG_GERA_LOTE_AUTOMATICO) = True
            Else
                ListaConfiguraEST.Selected(ESTCONFIG_GERA_LOTE_AUTOMATICO) = False
            End If
        End If
        
    End If
    
    Exit Sub

Erro_EST_Inicializa_Config:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175877)

    End Select

    Exit Sub
    
End Sub
    
Private Function EST_Config_Gravar_Registro() As Long

Dim lErro As Long
    
On Error GoTo Erro_EST_Config_Gravar_Registro
    
    lErro = Valida_Step(MODULO_ESTOQUE)
    
    If lErro = SUCESSO Then
        
        If ListaConfiguraEST.Selected(ESTCONFIG_AGLUTINA_LANCAM_POR_DIA) = True Then
            gobjEST.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA
        Else
            gobjEST.iAglutinaLancamPorDia = NAO_AGLUTINA_LANCAM_POR_DIA
        End If
        
        If giTipoVersao = VERSAO_FULL Then
            If ListaConfiguraEST.Selected(ESTCONFIG_GERA_LOTE_AUTOMATICO) = True Then
                gobjEST.iGeraLoteAutomatico = GERA_LOTE_AUTOMATICO
            Else
                gobjEST.iGeraLoteAutomatico = NAO_GERA_LOTE_AUTOMATICO
            End If
        ElseIf giTipoVersao = VERSAO_LIGHT Then
        
            gobjEST.iGeraLoteAutomatico = NAO_GERA_LOTE_AUTOMATICO
        End If
        
        lErro = gobjEST.Gravar_Trans()
        If lErro <> SUCESSO Then Error 41812
    
    End If
    
    EST_Config_Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_EST_Config_Gravar_Registro:

    EST_Config_Gravar_Registro = Err

    Select Case Err
    
        Case 41812
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175878)
            
    End Select

    Exit Function
    
End Function
'=============================================================================
' FIM TELA DE CONFIGURACAO ESTOQUE
'=============================================================================

'=============================================================================
' TELA DE CONFIGURACAO COMPRAS
'=============================================================================

Private Sub COM_Inicializa_Config()
           
Dim lErro As Long

On Error GoTo Erro_COM_Inicializa_Config
    
    lErro = Valida_Step(MODULO_COMPRAS)
    
    If lErro = SUCESSO Then
    
        lErro = Carrega_FilialCompra()
        If lErro <> SUCESSO Then gError 74899
    
    
        'Visibilidade para versão LIGHT
        If giTipoVersao = VERSAO_LIGHT Then
        
            FilialCompra.ListIndex = 0
            FilialCompra.left = POSICAO_FORA_TELA
            FilialCompra.TabStop = False
            Label1(2).left = POSICAO_FORA_TELA
            Label1(2).Visible = False
            
        End If
    
    End If
    Exit Sub

Erro_COM_Inicializa_Config:

    Select Case gErr

        Case 74899
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175879)

    End Select

    Exit Sub
    
End Sub
Private Function Carrega_FilialCompra() As Long
'Carrega a Combo FilialCompra com Codigo-Nome das filiais empresas

Dim lErro As Long
Dim colFiliais As New Collection
Dim objFiliais As ClassUsuarioEmpresa
Dim lCodEmpresa As Long

On Error GoTo Erro_Carrega_FilialCompra

    'Le todas as filiais e coloca todas as filiais em colFiliais
    lErro = FiliaisEmpresa_Le_Usuario(gsUsuario, glEmpresa, colFiliais)
    If lErro <> SUCESSO Then Error 49353

    'Carrega a Filial na combo FilialCompra
    For Each objFiliais In colFiliais
        If objFiliais.iCodFilial <> EMPRESA_TODA Then
            FilialCompra.AddItem objFiliais.iCodFilial & SEPARADOR & objFiliais.sNomeFilial
        End If
    Next

    Carrega_FilialCompra = SUCESSO

    Exit Function

Erro_Carrega_FilialCompra:

    Carrega_FilialCompra = Err

    Select Case Err

        Case 49353

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175880)

    End Select

    Exit Function

End Function

Private Sub MesesConsumoMedio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MesesConsumoMedio_Validate

    'Verifica se MesesConsumoMedio foi preenchido
    If Len(Trim(MesesConsumoMedio.Text)) > 0 Then
    
        'Critica o valor preenchido
        lErro = Valor_Positivo_Critica(MesesConsumoMedio.Text)
        If lErro <> SUCESSO Then Error 49365
    
    End If
    
    Exit Sub
    
Erro_MesesConsumoMedio_Validate:

    Cancel = True
    
    Select Case Err
    
        Case 49365
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175881)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub MesesMediaTempoRessup_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MesesMediaTempoRessup_Validate

    'Verifica se MesesMediaTempoRessup foi preenchido
    If Len(Trim(MesesMediaTempoRessup.Text)) > 0 Then
    
        'Critica o valor preenchido
        lErro = Valor_Positivo_Critica(MesesMediaTempoRessup.Text)
        If lErro <> SUCESSO Then Error 49366
    
    End If
    
    Exit Sub
    
Erro_MesesMediaTempoRessup_Validate:

    Cancel = True

    Select Case Err
    
        Case 49366
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175882)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub NaoConsideraQuantCotacaoAnterior_Click()

Dim lErro As Long

On Error GoTo Erro_NaoConsideraQuantCotacaoAnterior

    iAlterado = REGISTRO_ALTERADO

    If NaoConsideraQuantCotacaoAnterior.Value = vbChecked Then

        'Habilita os controles
        PercentMaisQuantCotacaoAnterior.Enabled = False
        PercentMenosQuantCotacaoAnterior.Enabled = False

    Else

        'Desabilita os controles
        PercentMaisQuantCotacaoAnterior.Enabled = True
        PercentMenosQuantCotacaoAnterior.Enabled = True

    End If

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_NaoConsideraQuantCotacaoAnterior:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175883)

    End Select

    Exit Sub

End Sub

Private Sub FilialCompra_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_FilialCompra_Validate

    'Verifica se FilialCompra foi preenchida
    If Len(Trim(FilialCompra.Text)) = 0 Then Exit Sub

    If FilialCompra.ListIndex <> -1 Then Exit Sub
    
    'Tenta selecionar  a filialempresa na combo
    lErro = Combo_Seleciona(FilialCompra, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 49331

    'Se não encontrou o CÓDIGO
    If lErro = 6730 Then

        objFilialEmpresa.lCodEmpresa = glEmpresa
        objFilialEmpresa.iCodFilial = iCodigo

        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then Error 49364
        If lErro = 27378 Then Error 49634
        
        FilialCompra.Text = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome

    End If

    'Não encontrou a STRING
    If lErro = 6731 Then Error 49363

    Exit Sub

Erro_FilialCompra_Validate:

    Cancel = True
    
    Select Case Err

        Case 49331

        Case 49363
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", Err, FilialCompra.Text)

        Case 49634
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err, objFilialEmpresa.iCodFilial)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175884)

    End Select

    Exit Sub

End Sub
Private Sub NaoTemFaixaReceb_Click()

Dim lErro As Long

On Error GoTo Erro_NaoTemFaixaReceb_Click

    'Alteracao Daniel ( = False por vbUnchecked )
    'Verifica valor na checkbox
    If NaoTemFaixaReceb.Value = vbUnchecked Then

        'Habilita os controles
        PercentMaisReceb.Enabled = True
        PercentMenosReceb.Enabled = True
        RecebForaFaixa(0).Enabled = True
        RecebForaFaixa(1).Enabled = True

    Else

        'Desabilita os controles
        PercentMaisReceb.Enabled = False
        PercentMenosReceb.Enabled = False
        RecebForaFaixa(0).Enabled = False
        RecebForaFaixa(1).Enabled = False

    End If
           
    Exit Sub

Erro_NaoTemFaixaReceb_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175885)

    End Select

    Exit Sub

End Sub

Private Sub PercentMaisQuantCotacaoAnterior_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dPercentMaisQuantCotacaoAnterior As Double

On Error GoTo Erro_PercentMaisQuantCotacaoAnterior_Validate

    'Verifica se PercentMaisQuantCotacaoAnterior foi preenchido
    If Len(Trim(PercentMaisQuantCotacaoAnterior.Text)) <> 0 Then

        'Verifica se é porcentagem
        lErro = Porcentagem_Critica(PercentMaisQuantCotacaoAnterior.Text)
        If lErro <> SUCESSO Then Error 49336

        dPercentMaisQuantCotacaoAnterior = StrParaDbl(PercentMaisQuantCotacaoAnterior.Text)

        'Verifica se a porcentagem é igual a 100%
        If dPercentMaisQuantCotacaoAnterior = 100 Then Error 49337

        'Coloca o valor no formato fixed da tela
        PercentMaisQuantCotacaoAnterior.Text = Format(dPercentMaisQuantCotacaoAnterior, "Fixed")

    End If

    Exit Sub

Erro_PercentMaisQuantCotacaoAnterior_Validate:

    Cancel = True

    Select Case Err

        Case 49336

        Case 49337
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_IGUAL_100", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175886)

    End Select

    Exit Sub

End Sub
Private Sub PercentMaisReceb_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dPercentMaisReceb As Double

On Error GoTo Erro_PercentMaisReceb_Validate

    'Verifica se PercentMaisRecebe foi preenchido
    If Len(Trim(PercentMaisReceb.Text)) <> 0 Then

        'Verifica se é porcentagem
        lErro = Porcentagem_Critica(PercentMaisReceb.Text)
        If lErro <> SUCESSO Then Error 49340

        dPercentMaisReceb = StrParaDbl(PercentMaisReceb.Text)

        'Coloca o valor no formato fixed da tela
        PercentMaisReceb.Text = Format(dPercentMaisReceb, "Fixed")

    End If

    Exit Sub

Erro_PercentMaisReceb_Validate:

    Cancel = True
    
    Select Case Err

        Case 49340

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175887)

    End Select

    Exit Sub

End Sub

Private Sub PercentMenosQuantCotacaoAnterior_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dPercentMenosQuantCotacaoAnterior As Double

On Error GoTo Erro_PercentMenosQuantCotacaoAnterior_Validate

    'Verifica se PercentMenosQuantCotacaoAnterior foi preenchido
    If Len(Trim(PercentMenosQuantCotacaoAnterior.Text)) <> 0 Then

        'Verifica se é porcentagem
        lErro = Porcentagem_Critica(PercentMenosQuantCotacaoAnterior.Text)
        If lErro <> SUCESSO Then Error 49338

        dPercentMenosQuantCotacaoAnterior = StrParaDbl(PercentMenosQuantCotacaoAnterior.Text)

        'Verifica se porcentagem igual a 100%
        If dPercentMenosQuantCotacaoAnterior = 100 Then Error 49339

        'Coloca o valor no formato fixed da tela
        PercentMenosQuantCotacaoAnterior.Text = Format(dPercentMenosQuantCotacaoAnterior, "Fixed")

    End If

    Exit Sub

Erro_PercentMenosQuantCotacaoAnterior_Validate:

    Cancel = True
    
    Select Case Err

        Case 49338

        Case 49339
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_IGUAL_100", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175888)

    End Select

    Exit Sub

End Sub
Private Sub PercentMenosReceb_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dPercentMenosReceb As Double

On Error GoTo Erro_PercentMenosReceb_Validate

    'Verifica se PercentMenosRecebe foi preenchido
    If Len(Trim(PercentMenosReceb.Text)) <> 0 Then

        'Verifica se é porcentagem
        lErro = Porcentagem_Critica(PercentMenosReceb.Text)
        If lErro <> SUCESSO Then Error 49342

        dPercentMenosReceb = StrParaDbl(PercentMenosReceb.Text)

        'Verifica se porcentagem igual a 100%
        If dPercentMenosReceb = 100 Then Error 49343

        'Coloca o valor no formato fixed da tela
        PercentMenosReceb.Text = Format(dPercentMenosReceb, "Fixed")

    End If

    Exit Sub

Erro_PercentMenosReceb_Validate:

    Cancel = True

    Select Case Err

        Case 49342

        Case 49343
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_IGUAL_100", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175889)

    End Select

    Exit Sub

End Sub
Private Sub Residuo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dResiduo As Double

On Error GoTo Erro_Residuo_Validate

    'Verifica se Residuo foi preenchido
    If Len(Trim(Residuo.Text)) <> 0 Then

        'Verifica se é porcentagem
        lErro = Porcentagem_Critica(Residuo.Text)
        If lErro <> SUCESSO Then Error 49332
    
        dResiduo = StrParaDbl(Residuo.Text)
    
        'Verifica se Residuo é igual a 100%
        If dResiduo = 100 Then Error 49333
    
        'Coloca o valor no formato Fixed da tela
        Residuo.Text = Format(dResiduo, "Fixed")

    End If
    
    Exit Sub

Erro_Residuo_Validate:

    Cancel = True
    
    Select Case Err

        Case 49332

        Case 49333
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_IGUAL_100", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175890)

    End Select

    Exit Sub

End Sub
Private Sub TaxaFinanceiraEmpresa_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dTaxaFinanceiraEmpresa As Double

On Error GoTo Erro_TaxaFinanceiraEmpresa_Validate

    'Verifica se TaxaFinanceiraEmpresa foi preenchido
    If Len(Trim(TaxaFinanceiraEmpresa.Text)) <> 0 Then

        'Verifica se é porcentagem
        lErro = Porcentagem_Critica(TaxaFinanceiraEmpresa.Text)
        If lErro <> SUCESSO Then Error 49334

        dTaxaFinanceiraEmpresa = StrParaDbl(TaxaFinanceiraEmpresa.Text)

        'Coloca o valor no formato fixed da tela
        TaxaFinanceiraEmpresa.Text = Format(dTaxaFinanceiraEmpresa, "Fixed")

    End If

    Exit Sub

Erro_TaxaFinanceiraEmpresa_Validate:

    Cancel = True
    
    Select Case Err

        Case 49334

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175891)

    End Select

    Exit Sub

End Sub
Function COM_Parte1_Testa()

Dim lErro As Long

On Error GoTo Erro_COM_Parte1_Testa
    
    lErro = Valida_Step(MODULO_COMPRAS)
    
    If lErro = SUCESSO Then

        If Len(Trim(Residuo.ClipText)) = 0 Then gError 74900
        If Len(Trim(MesesConsumoMedio.ClipText)) = 0 Then gError 74901
        If Len(Trim(MesesMediaTempoRessup.ClipText)) = 0 Then Error 74902
    
    End If

    COM_Parte1_Testa = SUCESSO
    
    Exit Function

Erro_COM_Parte1_Testa:

    COM_Parte1_Testa = gErr
    
    Select Case gErr

        Case 74900
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RESIDUO_NAO_PREENCHIDO", gErr)
    
        Case 74901
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MESESCONSUMOMEDIO_NAO_PREENCHIDO", gErr)

        Case 74902
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MESESMEDIATEMPORESSUP_NAO_PREENCHIDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175892)

    End Select

End Function

Function COM_Config_Gravar_Registro()

Dim lErro As Long

On Error GoTo Erro_COM_Config_Gravar_Registro

    'FilialCompra
    gobjCOM.iFilialCompra = Codigo_Extrai(FilialCompra.Text)
    gobjCOM.iCompradorAumentaQuant = CompradorAumentaQuant.Value
    gobjCOM.iControleAlcada = ControleAlcada.Value
    
    If Len(Trim(Residuo.Text)) <> 0 Then
        gobjCOM.dResiduo = StrParaDbl(Residuo.Text) / 100
    Else
        gobjCOM.dResiduo = 0
    End If
    
    gobjCOM.iNumComprasMediaAtraso = StrParaInt(NumComprasMediaAtraso.Text)
    gobjCOM.iNumComprasTempoRessup = StrParaInt(NumComprasTempoRessup.Text)
    gobjCOM.dTaxaFinanceiraEmpresa = StrParaDbl(TaxaFinanceiraEmpresa.Text) / 100
    gobjCOM.iMesesConsumoMedio = StrParaInt(MesesConsumoMedio.Text)
    gobjCOM.iMesesMediaTempoRessup = StrParaInt(MesesMediaTempoRessup.Text)

    'Estoque de Seguranca
    gobjCOM.dConsumoMedioMax = StrParaDbl(ConsumoMedioMax.Text) / 100
    gobjCOM.dTempoRessupMax = StrParaDbl(TempoRessupMax.Text) / 100
    
    'Cotacoes Anteriores
    gobjCOM.iConsideraQuantCotacaoAnterior = NaoConsideraQuantCotacaoAnterior.Value
    gobjCOM.dPercentMaisQuantCotacaoAnterior = StrParaDbl(PercentMaisQuantCotacaoAnterior.Text) / 100
    gobjCOM.dPercentMenosQuantCotacaoAnterior = StrParaDbl(PercentMenosQuantCotacaoAnterior.Text) / 100

    'Faixa de Recebimento
    gobjCOM.iTemFaixaReceb = NaoTemFaixaReceb.Value
    gobjCOM.dPercentMenosReceb = StrParaDbl(PercentMenosReceb.Text) / 100
    gobjCOM.dPercentMaisReceb = StrParaDbl(PercentMaisReceb.Text) / 100

    If RecebForaFaixa(0).Value Then
        gobjCOM.iRecebForaFaixa = REJEITA_RECEBIMENTO
    Else
        gobjCOM.iRecebForaFaixa = ACEITA_RECEBIMENTO
    End If

    lErro = gobjCOM.Gravar_Trans
    If lErro <> SUCESSO Then gError 74903

    COM_Config_Gravar_Registro = SUCESSO

    Exit Function

Erro_COM_Config_Gravar_Registro:

    COM_Config_Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 74903
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175893)
            
    End Select

    Exit Function

End Function
'=============================================================================
' FIM TELA DE CONFIGURACAO COMPRAS
'=============================================================================


Function Trata_Parametros(objConfiguraADM As ClassConfiguraADM) As Long

On Error GoTo Erro_Trata_Parametros

    Set objConfiguraADM1 = objConfiguraADM
    
    Call SGE_Inicializa_Segmentos
    Call CTB_Inicializa_Segmentos
    Call CTB_Inicializa_Config
    Call CTB_Inicializa_Exercicio
    Call TES_Inicializa_Config
    Call CP_Inicializa_Config
    Call CR_Inicializa_Config
    Call FAT_Inicializa_Config
    Call EST_Inicializa_Config
    Call COM_Inicializa_Config
    Call LJ_Inicializa_Config
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175894)
    
    End Select
    
    Exit Function

End Function

Private Function SGE_Configuracao_Gravar_Registro() As Long

Dim lErro As Long
Dim colSegmentos As Collection

On Error GoTo Erro_SGE_Configuracao_Gravar_Registro

    lErro = Valida_Step(SISTEMA_SGE)

    If lErro = SUCESSO Then
    
        lErro = CF("Configuracao_Altera_DataInstalacao")
        If lErro <> SUCESSO Then Error 55244
        
    End If
    
    SGE_Configuracao_Gravar_Registro = SUCESSO
       
    Exit Function
    
Erro_SGE_Configuracao_Gravar_Registro:
    
    SGE_Configuracao_Gravar_Registro = Err
    
    Select Case Err

        Case 55244

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175895)

    End Select

    Exit Function
    
End Function


'=============================================================================
' TRECHO COMUM A MAIS DE UMA TELA DE CONFIGURACAO
'=============================================================================

Sub Rotina_Grid_Enable(iLinha As Integer, objControle As Object, iCaminho As Integer)
   
Dim iTipo As Integer

    'Pesquisa a controle da coluna em questão
    Select Case objControle.Name
        
        '=============================================================================
        ' TELA DE CONFIGURACAO FATURAMENTO
        '=============================================================================
        Case PercentualDesc.Name
                
            If Len(Trim(GridDescontos.TextMatrix(GridDescontos.Row, iGrid_TipoDesconto_Col))) > 0 Then
                
                iTipo = Codigo_Extrai(GridDescontos.TextMatrix(GridDescontos.Row, iGrid_TipoDesconto_Col))
                
                If iTipo = Percentual Or iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Then
                    objControle.Enabled = True
                Else
                    objControle.Enabled = False
                End If
            Else
                objControle.Enabled = False
            End If
        
        Case Dias.Name
            
            If Len(Trim(GridDescontos.TextMatrix(GridDescontos.Row, iGrid_TipoDesconto_Col))) > 0 Then
                
                iTipo = Codigo_Extrai(GridDescontos.TextMatrix(GridDescontos.Row, iGrid_TipoDesconto_Col))
                
                If iTipo = Percentual Or iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Then
                    objControle.Enabled = True
                Else
                    objControle.Enabled = False
                End If
            Else
                objControle.Enabled = False
            End If
        
        '=============================================================================
        ' FIM DA TELA DE CONFIGURACAO FATURAMENTO
        '=============================================================================
    
    End Select
    
    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    If objGridInt Is CTB_Exercicio_objGrid1 Then
    
        lErro = CTB_Exercicio_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 44650
        
    ElseIf objGridInt Is CTB_Segmentos_objGrid1 Then
    
        lErro = CTB_Segmentos_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 44651
    
    ElseIf objGridInt Is SGE_Segmentos_objGrid1 Then
    
        lErro = SGE_Segmentos_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 44815
    
    ElseIf objGridInt Is objGridDescontos Then
    
        lErro = Saida_Celula_GridDescontos(objGridInt)
        If lErro <> SUCESSO Then Error 56701
    
    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err
        
        Case 44650, 44651, 44815, 56701
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175896)

    End Select

    Exit Function

End Function

'=============================================================================
' FIM DO TRECHO COMUM A MAIS DE UMA TELA DE CONFIGURACAO
'=============================================================================


Private Sub lblStep_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(lblStep(Index), Source, X, Y)
End Sub

Private Sub lblStep_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(lblStep(Index), Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label3(Index), Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3(Index), Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label8(Index), Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8(Index), Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub

Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub LabelNumPeriodos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNumPeriodos, Source, X, Y)
End Sub

Private Sub LabelNumPeriodos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNumPeriodos, Button, Shift, X, Y)
End Sub

Private Sub LabelPeriodicidade_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPeriodicidade, Source, X, Y)
End Sub

Private Sub LabelPeriodicidade_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPeriodicidade, Button, Shift, X, Y)
End Sub

Private Sub Exercicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Exercicio, Source, X, Y)
End Sub

Private Sub Exercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Exercicio, Button, Shift, X, Y)
End Sub

Private Sub LabelDataInicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataInicio, Source, X, Y)
End Sub

Private Sub LabelDataInicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataInicio, Button, Shift, X, Y)
End Sub

Private Sub LabelDataFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataFim, Source, X, Y)
End Sub

Private Sub LabelDataFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataFim, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Nat_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Nat, Source, X, Y)
End Sub

Private Sub Nat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Nat, Button, Shift, X, Y)
End Sub

Private Sub TipoDaConta_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoDaConta, Source, X, Y)
End Sub

Private Sub TipoDaConta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoDaConta, Button, Shift, X, Y)
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

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub


Private Sub Opcoes_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcoes)
End Sub
Private Sub Truncamento_Click()
      iAlterado = REGISTRO_ALTERADO
End Sub
