VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TRVOcrCasos 
   ClientHeight    =   6240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10515
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   10515
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5325
      Index           =   2
      Left            =   120
      TabIndex        =   140
      Top             =   885
      Visible         =   0   'False
      Width           =   10305
      Begin VB.Frame Frame0 
         Caption         =   "Autorização"
         Height          =   480
         Index           =   15
         Left            =   15
         TabIndex        =   186
         Top             =   -30
         Width           =   10230
         Begin VB.ComboBox CGAutorizadoPor 
            Height          =   315
            Left            =   7800
            TabIndex        =   46
            Text            =   "CGAutorizadoPor"
            Top             =   120
            Width           =   2340
         End
         Begin VB.ComboBox CGStatus 
            Height          =   315
            Left            =   3975
            TabIndex        =   45
            Text            =   "CGStatus"
            Top             =   120
            Width           =   2340
         End
         Begin VB.ComboBox CGAnalise 
            Height          =   315
            Left            =   1185
            TabIndex        =   44
            Text            =   "CGAnalise"
            Top             =   120
            Width           =   1905
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Autorizado Por:"
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
            Height          =   240
            Index           =   43
            Left            =   6285
            TabIndex        =   189
            Top             =   180
            Width           =   1500
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   42
            Left            =   3120
            TabIndex        =   188
            Top             =   180
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Análise:"
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
            Height          =   240
            Index           =   41
            Left            =   300
            TabIndex        =   187
            Top             =   180
            Width           =   810
         End
      End
      Begin VB.Frame Frame0 
         Caption         =   "Favorecido"
         Height          =   1875
         Index           =   11
         Left            =   15
         TabIndex        =   148
         Top             =   3435
         Width           =   10230
         Begin VB.CommandButton BotaoLimparFavorecido 
            Caption         =   "Limpar"
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
            Left            =   9060
            TabIndex        =   96
            Top             =   1545
            Width           =   1035
         End
         Begin VB.CommandButton BotaoAbrirFav 
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
            Height          =   285
            Left            =   9435
            TabIndex        =   56
            ToolTipText     =   "Abre a tela de cliente com o passageiro"
            Top             =   390
            Width           =   660
         End
         Begin VB.TextBox Logradouro 
            Height          =   285
            Left            =   1095
            MaxLength       =   40
            TabIndex        =   62
            Top             =   1275
            Width           =   9015
         End
         Begin VB.ComboBox Pais 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4920
            TabIndex        =   58
            Top             =   690
            Width           =   2535
         End
         Begin VB.ComboBox Estado 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   9495
            TabIndex        =   59
            Top             =   690
            Width           =   630
         End
         Begin VB.TextBox NomeFavorecido 
            Height          =   285
            Left            =   1095
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   55
            Top             =   405
            Width           =   6360
         End
         Begin MSMask.MaskEdBox ContaCorrente 
            Height          =   275
            Left            =   4905
            TabIndex        =   53
            Top             =   135
            Width           =   2550
            _ExtentX        =   4498
            _ExtentY        =   476
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Agencia 
            Height          =   275
            Left            =   3045
            TabIndex        =   52
            Top             =   135
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   476
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   7
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Banco 
            Height          =   275
            Left            =   1095
            TabIndex        =   51
            Top             =   135
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   476
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            Mask            =   "999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FavorecidoCGC 
            Height          =   275
            Left            =   8355
            TabIndex        =   54
            Top             =   135
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   476
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            Mask            =   "##############"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Bairro 
            Height          =   270
            Left            =   4920
            TabIndex        =   61
            Top             =   1005
            Width           =   5190
            _ExtentX        =   9155
            _ExtentY        =   476
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Cidade 
            Height          =   270
            Left            =   1095
            TabIndex        =   60
            Top             =   1005
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   476
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CEP 
            Height          =   315
            Left            =   1095
            TabIndex        =   57
            Top             =   690
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#####-###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Numero 
            Height          =   270
            Left            =   1095
            TabIndex        =   63
            Top             =   1560
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   476
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Complemento 
            Height          =   270
            Left            =   4920
            TabIndex        =   95
            Top             =   1560
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   476
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            Caption         =   "Forn:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   34
            Left            =   7845
            TabIndex        =   163
            Top             =   450
            Width           =   465
         End
         Begin VB.Label FornFavorecido 
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   8355
            TabIndex        =   162
            Top             =   405
            Width           =   1095
         End
         Begin VB.Label PaisLabel 
            AutoSize        =   -1  'True
            Caption         =   "País:"
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
            Left            =   4425
            TabIndex        =   161
            Top             =   780
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "CEP:"
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
            Index           =   33
            Left            =   645
            TabIndex        =   160
            Top             =   750
            Width           =   465
         End
         Begin VB.Label LabelCidade 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
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
            TabIndex        =   159
            Top             =   1050
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
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
            Index           =   32
            Left            =   9105
            TabIndex        =   158
            Top             =   735
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
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
            Index           =   31
            Left            =   4320
            TabIndex        =   157
            Top             =   1050
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Logradouro:"
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
            Index           =   30
            Left            =   60
            TabIndex        =   156
            Top             =   1320
            Width           =   1035
         End
         Begin VB.Label Label1 
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
            Index           =   27
            Left            =   360
            TabIndex        =   155
            Top             =   1605
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Complemento:"
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
            Index           =   26
            Left            =   3705
            TabIndex        =   154
            Top             =   1605
            Width           =   1200
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CPF:"
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
            Left            =   7890
            TabIndex        =   153
            Top             =   195
            Width           =   420
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
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
            Height          =   195
            Left            =   540
            TabIndex        =   152
            Top             =   465
            Width           =   555
         End
         Begin VB.Label Label77 
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
            Height          =   195
            Left            =   4290
            TabIndex        =   151
            Top             =   180
            Width           =   570
         End
         Begin VB.Label Label78 
            AutoSize        =   -1  'True
            Caption         =   "Agência:"
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
            Left            =   2235
            TabIndex        =   150
            Top             =   180
            Width           =   765
         End
         Begin VB.Label LabelBanco 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
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
            Left            =   465
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   149
            Top             =   180
            Width           =   615
         End
      End
      Begin VB.Frame Frame0 
         Caption         =   "Serviços"
         Height          =   3015
         Index           =   6
         Left            =   15
         TabIndex        =   144
         Top             =   420
         Width           =   10230
         Begin VB.Frame Frame0 
            Caption         =   "Responsabilidade do pagamento do seguro"
            Height          =   780
            Index           =   20
            Left            =   3960
            TabIndex        =   292
            Top             =   2190
            Width           =   4005
            Begin VB.CheckBox AntecPagto 
               Caption         =   "Antecipar"
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
               Left            =   2790
               TabIndex        =   49
               Top             =   525
               Width           =   1185
            End
            Begin MSMask.MaskEdBox SrvTotalSegTrvRS 
               Height          =   255
               Left            =   1545
               TabIndex        =   48
               Top             =   210
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   450
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
               Format          =   "#,##0.00##"
               PromptChar      =   " "
            End
            Begin VB.Label SrvTotalSegSegRS 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   1545
               TabIndex        =   295
               Top             =   480
               Width           =   1200
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Seguradora R$:"
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
               Index           =   40
               Left            =   105
               TabIndex        =   294
               Top             =   525
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Travel Ace R$:"
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
               Index           =   25
               Left            =   195
               TabIndex        =   293
               Top             =   240
               Width           =   1320
            End
         End
         Begin VB.Frame Frame0 
            Caption         =   "Autorizado Seguro"
            Height          =   780
            Index           =   19
            Left            =   2025
            TabIndex        =   287
            Top             =   2190
            Width           =   1890
            Begin VB.Label SrvTotalSegUS 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   480
               TabIndex        =   291
               Top             =   480
               Width           =   1200
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "U$:"
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
               Index           =   58
               Left            =   75
               TabIndex        =   290
               Top             =   525
               Width           =   390
            End
            Begin VB.Label SrvTotalSegRS 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   480
               TabIndex        =   289
               Top             =   210
               Width           =   1200
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "R$:"
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
               Index           =   24
               Left            =   60
               TabIndex        =   288
               Top             =   255
               Width           =   390
            End
         End
         Begin VB.Frame Frame0 
            Caption         =   "Autorizado Assistência"
            Height          =   780
            Index           =   14
            Left            =   90
            TabIndex        =   282
            Top             =   2190
            Width           =   1890
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "R$:"
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
               Index           =   57
               Left            =   60
               TabIndex        =   286
               Top             =   255
               Width           =   390
            End
            Begin VB.Label SrvTotalAssistRS 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   480
               TabIndex        =   285
               Top             =   210
               Width           =   1200
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "U$:"
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
               Index           =   39
               Left            =   75
               TabIndex        =   284
               Top             =   525
               Width           =   390
            End
            Begin VB.Label SrvTotalAssistUS 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   480
               TabIndex        =   283
               Top             =   480
               Width           =   1200
            End
         End
         Begin VB.TextBox SrvDescricaoDet 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   30
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   279
            Top             =   1860
            Width           =   10155
         End
         Begin VB.ComboBox IIMoeda 
            Height          =   315
            Left            =   6405
            TabIndex        =   275
            Text            =   "Combo1"
            Top             =   1755
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.CheckBox SrvSol 
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
            Left            =   600
            TabIndex        =   262
            Top             =   945
            Width           =   510
         End
         Begin VB.TextBox SrvTipo 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   3765
            MaxLength       =   250
            TabIndex        =   216
            Top             =   765
            Width           =   1050
         End
         Begin VB.TextBox SrvMoeda 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   2715
            MaxLength       =   250
            TabIndex        =   215
            Top             =   915
            Width           =   585
         End
         Begin MSMask.MaskEdBox SrvVlrLimite 
            Height          =   255
            Left            =   1515
            TabIndex        =   214
            Top             =   855
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
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
         Begin MSMask.MaskEdBox SrvVlrAutoUS 
            Height          =   255
            Left            =   6675
            TabIndex        =   213
            Top             =   450
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   450
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
         Begin MSMask.MaskEdBox SrvVlrAutoRS 
            Height          =   255
            Left            =   5805
            TabIndex        =   212
            Top             =   420
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   450
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
         Begin MSMask.MaskEdBox SrvVlrSolUS 
            Height          =   255
            Left            =   4875
            TabIndex        =   211
            Top             =   420
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   450
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
         Begin MSMask.MaskEdBox SrvVlrSolRS 
            Height          =   255
            Left            =   3930
            TabIndex        =   210
            Top             =   450
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   450
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
         Begin VB.TextBox SrvDescricao 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   960
            MaxLength       =   250
            TabIndex        =   208
            Top             =   435
            Width           =   4335
         End
         Begin VB.CheckBox SrvAuto 
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
            Left            =   510
            TabIndex        =   209
            Top             =   465
            Width           =   510
         End
         Begin MSFlexGridLib.MSFlexGrid GridSrv 
            Height          =   1545
            Left            =   15
            TabIndex        =   47
            Top             =   180
            Width           =   10155
            _ExtentX        =   17912
            _ExtentY        =   2725
            _Version        =   393216
            Rows            =   15
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox Cambio 
            Height          =   255
            Left            =   9060
            TabIndex        =   50
            Top             =   2520
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   450
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
            Format          =   "#,##0.00##"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Câmbio U$:"
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
            Height          =   210
            Index           =   21
            Left            =   7890
            TabIndex        =   185
            Top             =   2565
            Width           =   1110
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5295
      Index           =   1
      Left            =   75
      TabIndex        =   108
      Top             =   900
      Width           =   10320
      Begin VB.Frame Frame0 
         Caption         =   "Dados para Contato"
         Height          =   795
         Index           =   5
         Left            =   30
         TabIndex        =   121
         Top             =   2340
         Width           =   10215
         Begin VB.CommandButton BotaoLimparContato 
            Caption         =   "Limpar"
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
            Left            =   8970
            TabIndex        =   20
            Top             =   435
            Width           =   1080
         End
         Begin MSMask.MaskEdBox Contato 
            Height          =   285
            Left            =   1200
            TabIndex        =   21
            Top             =   180
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Email1 
            Height          =   285
            Left            =   6285
            TabIndex        =   22
            Top             =   180
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Telefone1 
            Height          =   285
            Left            =   1200
            TabIndex        =   23
            Top             =   465
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   40
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Telefone2 
            Height          =   285
            Left            =   6285
            TabIndex        =   24
            Top             =   465
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   40
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Telefone 2:"
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
            Left            =   5235
            TabIndex        =   128
            Top             =   510
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Telefone 1:"
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
            Left            =   180
            TabIndex        =   127
            Top             =   510
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "E-mail:"
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
            Index           =   13
            Left            =   5640
            TabIndex        =   126
            Top             =   210
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Contato:"
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
            Index           =   9
            Left            =   435
            TabIndex        =   125
            Top             =   210
            Width           =   750
         End
      End
      Begin VB.Frame Frame0 
         Caption         =   "Andamento"
         Height          =   2085
         Index           =   4
         Left            =   30
         TabIndex        =   119
         Top             =   3150
         Width           =   10215
         Begin VB.CommandButton BotaoFatReemb 
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
            Height          =   345
            Left            =   9495
            TabIndex        =   43
            Top             =   1695
            Width           =   555
         End
         Begin MSComCtl2.UpDown UpDown 
            Height          =   300
            Index           =   4
            Left            =   9450
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   765
            Width           =   225
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.CommandButton BotaoFatJur 
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
            Height          =   345
            Left            =   5820
            TabIndex        =   42
            Top             =   1695
            Width           =   555
         End
         Begin VB.CommandButton BotaoFatCobr 
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
            Height          =   345
            Left            =   2355
            TabIndex        =   41
            Top             =   1695
            Width           =   555
         End
         Begin VB.CheckBox Judicial 
            Caption         =   "Judicial"
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
            Left            =   1200
            TabIndex        =   40
            Top             =   1470
            Width           =   1170
         End
         Begin VB.CommandButton BotaoProg 
            Caption         =   "P"
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
            Index           =   4
            Left            =   9675
            TabIndex        =   33
            Top             =   765
            Width           =   360
         End
         Begin VB.CommandButton BotaoHoje 
            Caption         =   "H"
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
            Index           =   3
            Left            =   9675
            TabIndex        =   39
            Top             =   1080
            Width           =   360
         End
         Begin VB.CommandButton BotaoProg 
            Caption         =   "P"
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
            Index           =   5
            Left            =   6015
            TabIndex        =   36
            Top             =   1080
            Width           =   360
         End
         Begin VB.CommandButton BotaoHoje 
            Caption         =   "H"
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
            Index           =   2
            Left            =   9675
            TabIndex        =   30
            Top             =   135
            Width           =   360
         End
         Begin VB.CommandButton BotaoHoje 
            Caption         =   "H"
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
            Left            =   6015
            TabIndex        =   27
            Top             =   135
            Width           =   360
         End
         Begin MSComCtl2.UpDown UpDown 
            Height          =   300
            Index           =   1
            Left            =   5775
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   150
            Width           =   225
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataDocsRec 
            Height          =   300
            Left            =   4650
            TabIndex        =   25
            Top             =   150
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDown 
            Height          =   300
            Index           =   2
            Left            =   9435
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   150
            Width           =   225
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEnvioAnalise 
            Height          =   300
            Left            =   8310
            TabIndex        =   28
            Top             =   150
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDown 
            Height          =   300
            Index           =   5
            Left            =   5775
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   1095
            Width           =   225
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataProgrFin 
            Height          =   300
            Left            =   4650
            TabIndex        =   34
            Top             =   1095
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDown 
            Height          =   300
            Index           =   3
            Left            =   9435
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   1095
            Width           =   225
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataPagtoPax 
            Height          =   300
            Left            =   8310
            TabIndex        =   37
            Top             =   1095
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataLimite 
            Height          =   300
            Left            =   8310
            TabIndex        =   31
            Top             =   780
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label DataAbertura 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1200
            TabIndex        =   266
            Top             =   150
            Width           =   1740
         End
         Begin VB.Label Label1 
            Caption         =   "Fat.Reembolso:"
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
            Index           =   49
            Left            =   6945
            TabIndex        =   247
            Top             =   1770
            Width           =   1320
         End
         Begin VB.Label FatReembolso 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   8310
            TabIndex        =   246
            Top             =   1725
            Width           =   1185
         End
         Begin VB.Label FatJuridico 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4650
            TabIndex        =   245
            Top             =   1725
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "Fat.Jurídico:"
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
            Index           =   48
            Left            =   3495
            TabIndex        =   244
            Top             =   1770
            Width           =   1065
         End
         Begin VB.Label FatCobertura 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1185
            TabIndex        =   243
            Top             =   1725
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "Fat.Cobert:"
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
            Index           =   29
            Left            =   165
            TabIndex        =   242
            Top             =   1755
            Width           =   1125
         End
         Begin VB.Label DataEnvioFinanc 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1200
            TabIndex        =   193
            Top             =   1095
            Width           =   1740
         End
         Begin VB.Label AutorizadoPor 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   8310
            TabIndex        =   192
            Top             =   465
            Width           =   1740
         End
         Begin VB.Label Status 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4650
            TabIndex        =   191
            Top             =   465
            Width           =   1740
         End
         Begin VB.Label Analise 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1200
            TabIndex        =   190
            Top             =   465
            Width           =   1740
         End
         Begin VB.Label TotalAutoUS 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4650
            TabIndex        =   147
            Top             =   780
            Width           =   1740
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Autoriz. U$:"
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
            Index           =   16
            Left            =   3555
            TabIndex        =   146
            Top             =   840
            Width           =   1035
         End
         Begin VB.Label TotalAutoRS 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1200
            TabIndex        =   145
            Top             =   780
            Width           =   1740
         End
         Begin VB.Label DataPagtoReemb 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   8310
            TabIndex        =   139
            Top             =   1410
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Pagto Reemb:"
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
            Index           =   23
            Left            =   7065
            TabIndex        =   138
            Top             =   1455
            Width           =   1215
         End
         Begin VB.Label VlrCond 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4650
            TabIndex        =   137
            Top             =   1410
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Vlr Cond.:"
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
            Index           =   22
            Left            =   3735
            TabIndex        =   136
            Top             =   1455
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data Limite:"
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
            Index           =   20
            Left            =   7245
            TabIndex        =   135
            Top             =   840
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data Pagto:"
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
            Index           =   19
            Left            =   7245
            TabIndex        =   134
            Top             =   1140
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Progr Fin:"
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
            Index           =   18
            Left            =   3750
            TabIndex        =   133
            Top             =   1125
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Envio Finan:"
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
            Index           =   17
            Left            =   60
            TabIndex        =   132
            Top             =   1125
            Width           =   1080
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Autoriz. R$:"
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
            Index           =   15
            Left            =   90
            TabIndex        =   131
            Top             =   840
            Width           =   1035
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Autorizado Por:"
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
            Index           =   14
            Left            =   6795
            TabIndex        =   130
            Top             =   510
            Width           =   1500
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   12
            Left            =   3795
            TabIndex        =   129
            Top             =   495
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Análise:"
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
            Index           =   7
            Left            =   315
            TabIndex        =   124
            Top             =   495
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Envio Análise:"
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
            Index           =   6
            Left            =   7065
            TabIndex        =   123
            Top             =   210
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Docs Rec.:"
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
            Index           =   5
            Left            =   3630
            TabIndex        =   122
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Aberto em:"
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
            Left            =   210
            TabIndex        =   120
            Top             =   180
            Width           =   930
         End
      End
      Begin VB.Frame Frame0 
         Caption         =   "Dados do Voucher"
         Height          =   1125
         Index           =   3
         Left            =   30
         TabIndex        =   110
         Top             =   1200
         Width           =   10215
         Begin VB.TextBox VouPaxNome 
            Height          =   285
            Left            =   1215
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   9
            Top             =   210
            Width           =   3810
         End
         Begin VB.TextBox VouTitular 
            Height          =   285
            Left            =   6300
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   10
            Top             =   210
            Width           =   3810
         End
         Begin MSMask.MaskEdBox ClienteVou 
            Height          =   285
            Left            =   1215
            TabIndex        =   11
            Top             =   495
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   60
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   285
            Left            =   6300
            TabIndex        =   12
            Top             =   495
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDown 
            Height          =   300
            Index           =   8
            Left            =   2340
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   780
            Width           =   225
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmissao 
            Height          =   300
            Left            =   1215
            TabIndex        =   13
            Top             =   780
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDown 
            Height          =   300
            Index           =   9
            Left            =   4770
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   780
            Width           =   225
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataIda 
            Height          =   300
            Left            =   3645
            TabIndex        =   15
            Top             =   780
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDown 
            Height          =   300
            Index           =   10
            Left            =   7425
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   780
            Width           =   225
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataVolta 
            Height          =   300
            Left            =   6300
            TabIndex        =   17
            Top             =   780
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QtdePax 
            Height          =   285
            Left            =   9540
            TabIndex        =   19
            Top             =   780
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Passageiro:"
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
            Left            =   30
            TabIndex        =   118
            Top             =   255
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Titular:"
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
            Index           =   0
            Left            =   5295
            TabIndex        =   117
            Top             =   255
            Width           =   975
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   480
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   116
            Top             =   540
            Width           =   660
         End
         Begin VB.Label ProdutoLabel 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5535
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   115
            Top             =   540
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Emissão:"
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
            Left            =   375
            TabIndex        =   114
            Top             =   810
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ida:"
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
            Left            =   3225
            TabIndex        =   113
            Top             =   810
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Volta:"
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
            Left            =   5745
            TabIndex        =   112
            Top             =   810
            Width           =   510
         End
         Begin VB.Label Label1 
            Caption         =   "Qtde Pax:"
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
            Height          =   285
            Index           =   28
            Left            =   8565
            TabIndex        =   111
            Top             =   810
            Width           =   885
         End
      End
      Begin VB.Frame Frame0 
         Caption         =   "Eventos (Dados importados da Matriz)"
         Height          =   1170
         Index           =   2
         Left            =   30
         TabIndex        =   109
         Top             =   0
         Width           =   10215
         Begin VB.TextBox EvTelefone 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   7155
            MaxLength       =   250
            TabIndex        =   207
            Top             =   570
            Width           =   1555
         End
         Begin VB.TextBox EvCidade 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   6090
            MaxLength       =   250
            TabIndex        =   206
            Top             =   555
            Width           =   1355
         End
         Begin VB.TextBox EvTipo 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   4305
            MaxLength       =   250
            TabIndex        =   205
            Top             =   630
            Width           =   1810
         End
         Begin VB.TextBox EvEstado 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   2865
            MaxLength       =   250
            TabIndex        =   204
            Top             =   645
            Width           =   1610
         End
         Begin VB.TextBox EvPais 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   1455
            MaxLength       =   250
            TabIndex        =   203
            Top             =   675
            Width           =   1880
         End
         Begin VB.TextBox EvData 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   465
            MaxLength       =   250
            TabIndex        =   202
            Top             =   675
            Width           =   1045
         End
         Begin MSFlexGridLib.MSFlexGrid GridEv 
            Height          =   375
            Left            =   60
            TabIndex        =   8
            Top             =   195
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   661
            _Version        =   393216
            Rows            =   15
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
      Height          =   5265
      Index           =   7
      Left            =   105
      TabIndex        =   172
      Top             =   930
      Visible         =   0   'False
      Width           =   10260
      Begin VB.Frame FrameProc 
         Caption         =   "Processo"
         Height          =   2925
         Left            =   60
         TabIndex        =   180
         Top             =   -30
         Width           =   5400
         Begin VB.TextBox Comarca 
            Height          =   300
            Left            =   1395
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   82
            Top             =   855
            Width           =   3900
         End
         Begin VB.Frame Frame3 
            Caption         =   "Status"
            Height          =   870
            Left            =   60
            TabIndex        =   271
            Top             =   1635
            Width           =   5280
            Begin VB.OptionButton OptAcordo 
               Caption         =   "Acordo"
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
               TabIndex        =   87
               Top             =   255
               Width           =   1185
            End
            Begin VB.OptionButton OptCondenacao 
               Caption         =   "Condenação"
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
               Left            =   2325
               TabIndex        =   86
               Top             =   240
               Width           =   1680
            End
            Begin VB.CheckBox Condenado 
               Caption         =   "Perda de Causa"
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
               Left            =   135
               TabIndex        =   85
               Top             =   225
               Width           =   1845
            End
            Begin MSMask.MaskEdBox ValorCondenacao 
               Height          =   315
               Left            =   1350
               TabIndex        =   88
               Top             =   495
               Width           =   1215
               _ExtentX        =   2143
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
               Format          =   "#,##0.00##"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Valor a pagar:"
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
               Height          =   285
               Index           =   44
               Left            =   90
               TabIndex        =   272
               Top             =   540
               Width           =   1245
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Valor da Causa"
            Height          =   540
            Left            =   60
            TabIndex        =   268
            Top             =   1110
            Width           =   5280
            Begin MSMask.MaskEdBox DanoMaterial 
               Height          =   315
               Left            =   1335
               TabIndex        =   83
               Top             =   165
               Width           =   1215
               _ExtentX        =   2143
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
               Format          =   "#,##0.00##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox DanoMoral 
               Height          =   315
               Left            =   4005
               TabIndex        =   84
               Top             =   165
               Width           =   1215
               _ExtentX        =   2143
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
               Format          =   "#,##0.00##"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Dano Moral:"
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
               Height          =   285
               Index           =   55
               Left            =   2865
               TabIndex        =   270
               Top             =   210
               Width           =   1080
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Dano Material:"
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
               Height          =   285
               Index           =   54
               Left            =   60
               TabIndex        =   269
               Top             =   210
               Width           =   1260
            End
         End
         Begin VB.CheckBox Procon 
            Caption         =   "Procon"
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
            Left            =   4080
            TabIndex        =   80
            Top             =   225
            Width           =   1170
         End
         Begin VB.CommandButton BotaoHoje 
            Caption         =   "H"
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
            Index           =   11
            Left            =   2790
            TabIndex        =   79
            Top             =   135
            Width           =   360
         End
         Begin VB.CommandButton BotaoHoje 
            Caption         =   "H"
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
            Left            =   2820
            TabIndex        =   91
            Top             =   2520
            Width           =   360
         End
         Begin VB.CheckBox JudicialE 
            Caption         =   "Judicial"
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
            Left            =   4095
            TabIndex        =   76
            Top             =   2580
            Visible         =   0   'False
            Width           =   1170
         End
         Begin MSMask.MaskEdBox NumProcesso 
            Height          =   315
            Left            =   1395
            TabIndex        =   81
            Top             =   510
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataFimProcesso 
            Height          =   315
            Left            =   1410
            TabIndex        =   89
            Top             =   2550
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDown 
            Height          =   300
            Index           =   6
            Left            =   2580
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   2535
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataIniProcesso 
            Height          =   315
            Left            =   1395
            TabIndex        =   77
            Top             =   165
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDown 
            Height          =   300
            Index           =   11
            Left            =   2550
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   150
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Iniciado em:"
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
            Index           =   53
            Left            =   30
            TabIndex        =   267
            Top             =   225
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Finalizado em:"
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
            Index           =   37
            Left            =   45
            TabIndex        =   195
            Top             =   2580
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Comarca:"
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
            Index           =   45
            Left            =   480
            TabIndex        =   194
            Top             =   885
            Width           =   900
         End
         Begin VB.Label Label16 
            Caption         =   "Núm. Processo:"
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
            Left            =   30
            TabIndex        =   182
            Top             =   555
            Width           =   1950
         End
      End
      Begin VB.Frame FrameGAdv 
         Caption         =   "Gastos com advogado"
         Height          =   3135
         Left            =   5505
         TabIndex        =   248
         Top             =   2055
         Width           =   4740
         Begin VB.TextBox GADesc 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   1005
            MaxLength       =   250
            TabIndex        =   254
            Top             =   2025
            Width           =   1815
         End
         Begin MSMask.MaskEdBox GAValor 
            Height          =   255
            Left            =   1485
            TabIndex        =   253
            Top             =   840
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   450
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
            Format          =   "#,##0.00##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox GAData 
            Height          =   255
            Left            =   555
            TabIndex        =   252
            Top             =   795
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridGA 
            Height          =   1620
            Left            =   45
            TabIndex        =   249
            Top             =   210
            Width           =   4650
            _ExtentX        =   8202
            _ExtentY        =   2858
            _Version        =   393216
            Rows            =   15
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
         Begin VB.Label GATotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3435
            TabIndex        =   251
            Top             =   2730
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "Total:"
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
            Index           =   50
            Left            =   2745
            TabIndex        =   250
            Top             =   2760
            Width           =   630
         End
      End
      Begin VB.Frame Frame0 
         Caption         =   "Observação"
         Height          =   2310
         Index           =   17
         Left            =   60
         TabIndex        =   201
         Top             =   2895
         Width           =   5400
         Begin VB.TextBox Obs 
            Height          =   2010
            Left            =   75
            MultiLine       =   -1  'True
            TabIndex        =   94
            Top             =   240
            Width           =   5235
         End
      End
      Begin VB.Frame Frame0 
         Caption         =   "Fatura"
         Height          =   180
         Index           =   16
         Left            =   60
         TabIndex        =   196
         Top             =   2025
         Visible         =   0   'False
         Width           =   5400
         Begin VB.CommandButton BotaoAbrirFatProc 
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
            Height          =   345
            Left            =   2265
            TabIndex        =   93
            Top             =   150
            Width           =   645
         End
         Begin VB.Label Label1 
            Caption         =   "Emitida em:"
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
            Index           =   47
            Left            =   3015
            TabIndex        =   200
            Top             =   210
            Width           =   1110
         End
         Begin VB.Label DataEmiProc 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4155
            TabIndex        =   199
            Top             =   180
            Width           =   1185
         End
         Begin VB.Label Label1 
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
            Height          =   315
            Index           =   46
            Left            =   165
            TabIndex        =   198
            Top             =   210
            Width           =   750
         End
         Begin VB.Label NumeroFatProc 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   975
            TabIndex        =   197
            Top             =   180
            Width           =   1275
         End
      End
      Begin VB.Frame FramePC 
         Caption         =   "Pagamentos"
         Height          =   1950
         Left            =   5490
         TabIndex        =   181
         Top             =   -30
         Width           =   4755
         Begin MSMask.MaskEdBox PCPagamento 
            Height          =   255
            Left            =   2400
            TabIndex        =   241
            Top             =   510
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PCVencimento 
            Height          =   255
            Left            =   0
            TabIndex        =   239
            Top             =   225
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PCValor 
            Height          =   255
            Left            =   1140
            TabIndex        =   240
            Top             =   270
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   450
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
            Format          =   "#,##0.00##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridPC 
            Height          =   1620
            Left            =   45
            TabIndex        =   92
            Top             =   195
            Width           =   4665
            _ExtentX        =   8229
            _ExtentY        =   2858
            _Version        =   393216
            Rows            =   15
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
      Height          =   5265
      Index           =   3
      Left            =   105
      TabIndex        =   255
      Top             =   915
      Visible         =   0   'False
      Width           =   10275
      Begin VB.Frame Frame0 
         Caption         =   "Documentos necessários"
         Height          =   5265
         Index           =   18
         Left            =   15
         TabIndex        =   256
         Top             =   -15
         Width           =   10245
         Begin VB.TextBox DNObsDet 
            BackColor       =   &H8000000F&
            Height          =   690
            Left            =   795
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   281
            Top             =   4515
            Width           =   9405
         End
         Begin VB.TextBox DNDescDet 
            BackColor       =   &H8000000F&
            Height          =   690
            Left            =   795
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   280
            Top             =   3810
            Width           =   9405
         End
         Begin VB.TextBox DNObs 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   6165
            MaxLength       =   250
            TabIndex        =   263
            Top             =   2310
            Width           =   2640
         End
         Begin VB.TextBox DNDesc 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   1470
            TabIndex        =   261
            Top             =   2235
            Width           =   4440
         End
         Begin VB.CheckBox DNRecebido 
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
            Left            =   990
            TabIndex        =   260
            Top             =   3060
            Width           =   1050
         End
         Begin VB.CheckBox DNNU 
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
            Left            =   975
            TabIndex        =   259
            Top             =   2625
            Width           =   540
         End
         Begin VB.CheckBox DNNS 
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
            Left            =   1005
            TabIndex        =   258
            Top             =   2280
            Width           =   540
         End
         Begin MSFlexGridLib.MSFlexGrid GridDN 
            Height          =   1785
            Left            =   30
            TabIndex        =   257
            Top             =   195
            Width           =   10170
            _ExtentX        =   17939
            _ExtentY        =   3149
            _Version        =   393216
            Rows            =   15
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
         Begin VB.Label Label1 
            Caption         =   "Obs:"
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
            Index           =   52
            Left            =   270
            TabIndex        =   265
            Top             =   4740
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Desc:"
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
            Index           =   51
            Left            =   165
            TabIndex        =   264
            Top             =   4050
            Width           =   495
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5205
      Index           =   6
      Left            =   90
      TabIndex        =   141
      Top             =   930
      Visible         =   0   'False
      Width           =   10275
      Begin VB.Frame Frame0 
         Caption         =   "Nova Anotação"
         Height          =   1815
         Index           =   13
         Left            =   75
         TabIndex        =   178
         Top             =   15
         Width           =   10140
         Begin VB.TextBox TextoAnot 
            Height          =   1230
            Left            =   1155
            MaxLength       =   5000
            MultiLine       =   -1  'True
            TabIndex        =   73
            Top             =   255
            Width           =   8880
         End
         Begin VB.CommandButton BotaoGravarAnot 
            Caption         =   "Gravar"
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
            Left            =   1140
            TabIndex        =   74
            Top             =   1500
            Width           =   1095
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Texto:"
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
            Left            =   525
            TabIndex        =   179
            Top             =   315
            Width           =   555
         End
      End
      Begin VB.Frame Frame0 
         Caption         =   "Anotações"
         Height          =   3330
         Index           =   10
         Left            =   75
         TabIndex        =   177
         Top             =   1845
         Width           =   10140
         Begin VB.TextBox AnotTextoDet 
            BackColor       =   &H8000000F&
            Height          =   1275
            Left            =   90
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   278
            Top             =   1995
            Width           =   9915
         End
         Begin VB.TextBox AnotUsuario 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   5385
            MaxLength       =   250
            TabIndex        =   238
            Top             =   120
            Width           =   1305
         End
         Begin VB.TextBox AnotTexto 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   2475
            TabIndex        =   237
            Top             =   675
            Width           =   6135
         End
         Begin VB.TextBox AnotHora 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   1380
            MaxLength       =   250
            TabIndex        =   236
            Top             =   390
            Width           =   720
         End
         Begin VB.TextBox AnotData 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            MaxLength       =   250
            TabIndex        =   235
            Top             =   180
            Width           =   1020
         End
         Begin MSFlexGridLib.MSFlexGrid GridAnot 
            Height          =   1785
            Left            =   75
            TabIndex        =   75
            Top             =   195
            Width           =   9960
            _ExtentX        =   17568
            _ExtentY        =   3149
            _Version        =   393216
            Rows            =   15
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
      Height          =   5235
      Index           =   5
      Left            =   105
      TabIndex        =   143
      Top             =   930
      Visible         =   0   'False
      Width           =   10230
      Begin VB.Frame Frame0 
         Caption         =   "Novo Contato"
         Height          =   2070
         Index           =   12
         Left            =   75
         TabIndex        =   171
         Top             =   45
         Width           =   10095
         Begin VB.CommandButton BotaoGravarHist 
            Caption         =   "Gravar"
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
            Left            =   1140
            TabIndex        =   71
            Top             =   1785
            Width           =   1095
         End
         Begin VB.TextBox AssuntoHist 
            Height          =   1230
            Left            =   1155
            MaxLength       =   5000
            MultiLine       =   -1  'True
            TabIndex        =   70
            Top             =   540
            Width           =   8835
         End
         Begin VB.ComboBox OrigemHist 
            Height          =   315
            ItemData        =   "TRVOcrCasos.ctx":0000
            Left            =   1155
            List            =   "TRVOcrCasos.ctx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   66
            ToolTipText     =   "Selecione quem originou o relacionamento: o seu cliente ou a sua empresa."
            Top             =   195
            Width           =   1215
         End
         Begin MSComCtl2.UpDown UpDown 
            Height          =   300
            Index           =   7
            Left            =   4170
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   180
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataHist 
            Height          =   300
            Left            =   3165
            TabIndex        =   67
            ToolTipText     =   "Informe a data quando ocorreu o relacionamento. Em caso de agendamento, informe a data de quando ocorrerá."
            Top             =   195
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox HoraHist 
            Height          =   315
            Left            =   5430
            TabIndex        =   69
            Top             =   195
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelAssunto 
            AutoSize        =   -1  'True
            Caption         =   "Assunto:"
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
            TabIndex        =   176
            Top             =   600
            Width           =   930
         End
         Begin VB.Label LabelData 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2595
            TabIndex        =   175
            Top             =   240
            Width           =   480
         End
         Begin VB.Label LabelOrigem 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   465
            TabIndex        =   174
            Top             =   255
            Width           =   660
         End
         Begin VB.Label LabelHora 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
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
            Left            =   4890
            TabIndex        =   173
            Top             =   255
            Width           =   480
         End
      End
      Begin VB.Frame Frame0 
         Caption         =   "Histórico"
         Height          =   3015
         Index           =   9
         Left            =   75
         TabIndex        =   170
         Top             =   2145
         Width           =   10095
         Begin VB.TextBox HistAssuntoDet 
            BackColor       =   &H8000000F&
            Height          =   945
            Left            =   90
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   277
            Top             =   1995
            Width           =   9900
         End
         Begin VB.TextBox HistUsuario 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   6240
            MaxLength       =   250
            TabIndex        =   234
            Top             =   1185
            Width           =   1305
         End
         Begin VB.TextBox HistOrigem 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   6075
            MaxLength       =   250
            TabIndex        =   233
            Top             =   555
            Width           =   1020
         End
         Begin VB.TextBox HistAssunto 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   3165
            TabIndex        =   232
            Top             =   780
            Width           =   4980
         End
         Begin VB.TextBox HistHora 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   2055
            MaxLength       =   250
            TabIndex        =   231
            Top             =   825
            Width           =   720
         End
         Begin VB.TextBox HistData 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   255
            Left            =   675
            MaxLength       =   250
            TabIndex        =   230
            Top             =   615
            Width           =   1020
         End
         Begin MSFlexGridLib.MSFlexGrid GridHist 
            Height          =   1785
            Left            =   75
            TabIndex        =   72
            Top             =   195
            Width           =   9915
            _ExtentX        =   17489
            _ExtentY        =   3149
            _Version        =   393216
            Rows            =   15
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
      Height          =   5295
      Index           =   4
      Left            =   75
      TabIndex        =   142
      Top             =   900
      Visible         =   0   'False
      Width           =   10305
      Begin VB.Frame Frame0 
         Caption         =   "Faturas de Reembolso"
         Height          =   2610
         Index           =   8
         Left            =   60
         TabIndex        =   167
         Top             =   2640
         Width           =   10185
         Begin VB.TextBox OFDescricao 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   1620
            MaxLength       =   250
            TabIndex        =   229
            Top             =   1245
            Width           =   3750
         End
         Begin MSMask.MaskEdBox OFValorUS 
            Height          =   255
            Left            =   6660
            TabIndex        =   228
            Top             =   885
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   450
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
         Begin MSMask.MaskEdBox OFValorRS 
            Height          =   255
            Left            =   5700
            TabIndex        =   227
            Top             =   855
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   450
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
         Begin VB.TextBox OFNumero 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   4245
            MaxLength       =   20
            TabIndex        =   226
            Top             =   750
            Width           =   1110
         End
         Begin MSMask.MaskEdBox OFDataEmi 
            Height          =   255
            Left            =   3120
            TabIndex        =   225
            Top             =   735
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox OFDataRec 
            Height          =   255
            Left            =   1965
            TabIndex        =   224
            Top             =   540
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.CheckBox OFConsiderar 
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
            Left            =   1155
            TabIndex        =   223
            Top             =   555
            Width           =   435
         End
         Begin MSFlexGridLib.MSFlexGrid GridOF 
            Height          =   1785
            Left            =   75
            TabIndex        =   65
            Top             =   210
            Width           =   10035
            _ExtentX        =   17701
            _ExtentY        =   3149
            _Version        =   393216
            Rows            =   15
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
         Begin VB.Label OFTotalRS 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6420
            TabIndex        =   184
            Top             =   2310
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Total R$:"
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
            Index           =   38
            Left            =   5280
            TabIndex        =   183
            Top             =   2355
            Width           =   990
         End
         Begin VB.Label OFTotalUS 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   8865
            TabIndex        =   169
            Top             =   2310
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Total U$:"
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
            Index           =   36
            Left            =   7860
            TabIndex        =   168
            Top             =   2355
            Width           =   930
         End
      End
      Begin VB.Frame Frame0 
         Caption         =   "Faturas Internacionais"
         Height          =   2625
         Index           =   7
         Left            =   60
         TabIndex        =   164
         Top             =   -15
         Width           =   10185
         Begin VB.TextBox IINumero 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   3810
            MaxLength       =   20
            TabIndex        =   222
            Top             =   1635
            Width           =   1110
         End
         Begin VB.TextBox IIObs 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   840
            MaxLength       =   250
            TabIndex        =   221
            Top             =   1575
            Width           =   4200
         End
         Begin MSMask.MaskEdBox IIValorRS 
            Height          =   255
            Left            =   4395
            TabIndex        =   220
            Top             =   1155
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   450
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
         Begin MSMask.MaskEdBox IIValorMoeda 
            Height          =   255
            Left            =   3660
            TabIndex        =   219
            Top             =   600
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   450
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
         Begin MSMask.MaskEdBox IIDataEmi 
            Height          =   255
            Left            =   3255
            TabIndex        =   218
            Top             =   1110
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox IIDataRec 
            Height          =   255
            Left            =   1830
            TabIndex        =   217
            Top             =   975
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridII 
            Height          =   1965
            Left            =   75
            TabIndex        =   64
            Top             =   195
            Width           =   10035
            _ExtentX        =   17701
            _ExtentY        =   3466
            _Version        =   393216
            Rows            =   15
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Total U$:"
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
            Index           =   56
            Left            =   7890
            TabIndex        =   274
            Top             =   2370
            Width           =   930
         End
         Begin VB.Label IITotalUS 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   8895
            TabIndex        =   273
            Top             =   2325
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Total R$:"
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
            Height          =   210
            Index           =   35
            Left            =   4815
            TabIndex        =   166
            Top             =   2370
            Width           =   1590
         End
         Begin VB.Label IITotalRS 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6495
            TabIndex        =   165
            Top             =   2325
            Width           =   1245
         End
      End
   End
   Begin VB.CommandButton BotaoLibera 
      Height          =   525
      Left            =   8250
      Picture         =   "TRVOcrCasos.ctx":0020
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   15
      Width           =   555
   End
   Begin VB.Frame Frame0 
      Caption         =   "Busca por caso"
      Height          =   585
      Index           =   1
      Left            =   4575
      TabIndex        =   106
      Top             =   -30
      Width           =   3645
      Begin VB.CommandButton BotaoTrazerCaso 
         Height          =   315
         Left            =   3225
         Picture         =   "TRVOcrCasos.ctx":0462
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Trazer Dados"
         Top             =   210
         Width           =   345
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   840
         TabIndex        =   5
         Top             =   225
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin VB.Label LabelCodigo 
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
         Height          =   270
         Left            =   90
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   107
         Top             =   270
         Width           =   750
      End
   End
   Begin VB.Frame Frame0 
      Caption         =   "Busca por voucher"
      Height          =   585
      Index           =   0
      Left            =   30
      TabIndex        =   102
      Top             =   -30
      Width           =   4500
      Begin VB.CommandButton BotaoAtualizarVou 
         Height          =   315
         Left            =   3390
         Picture         =   "TRVOcrCasos.ctx":0834
         Style           =   1  'Graphical
         TabIndex        =   276
         ToolTipText     =   "Atualiza os dados da ocorrência de acordo com os dados do voucher"
         Top             =   210
         Width           =   345
      End
      Begin VB.CommandButton BotaoAbrirVou 
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
         Height          =   330
         Left            =   3765
         TabIndex        =   4
         Top             =   195
         Width           =   660
      End
      Begin VB.CommandButton BotaoTrazerVou 
         Height          =   315
         Left            =   3015
         Picture         =   "TRVOcrCasos.ctx":0C86
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Trazer Dados"
         Top             =   210
         Width           =   345
      End
      Begin MSMask.MaskEdBox TipVou 
         Height          =   315
         Left            =   330
         TabIndex        =   0
         Top             =   225
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   1
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Serie 
         Height          =   315
         Left            =   900
         TabIndex        =   1
         Top             =   225
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   1
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NumVou 
         Height          =   315
         Left            =   1710
         TabIndex        =   2
         Top             =   225
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   10
         Mask            =   "##########"
         PromptChar      =   " "
      End
      Begin VB.Label LabelNumVou 
         Caption         =   "Vou:"
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
         Left            =   1290
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   105
         Top             =   270
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "S:"
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
         Index           =   64
         Left            =   675
         TabIndex        =   104
         Top             =   270
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "T:"
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
         Index           =   65
         Left            =   120
         TabIndex        =   103
         Top             =   270
         Width           =   435
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   8880
      ScaleHeight     =   450
      ScaleWidth      =   1485
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   30
      Width           =   1545
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   45
         Picture         =   "TRVOcrCasos.ctx":1058
         Style           =   1  'Graphical
         TabIndex        =   101
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   525
         Picture         =   "TRVOcrCasos.ctx":11B2
         Style           =   1  'Graphical
         TabIndex        =   97
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1005
         Picture         =   "TRVOcrCasos.ctx":16E4
         Style           =   1  'Graphical
         TabIndex        =   98
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5685
      Left            =   30
      TabIndex        =   100
      Top             =   555
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   10028
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   7
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Solicitações e Autorização"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Documentos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Faturas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Histórico"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Anotações"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Jurídico"
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
Attribute VB_Name = "TRVOcrCasos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gbTrazendoDados As Boolean

Dim iAlterado As Integer
Dim iFrameAtual As Integer

Dim gsCEPAnt As String
Dim gbMudouCEP As Boolean
Dim giPaisAnt As Integer
Dim giIndexBrasil As Integer
Dim gdCambioAnt As Double
Dim gsProdAnt As String
Dim glNumIntDoc As Long
Dim glNumIntDocTitPagProc As Long
Dim glNumIntDocTitPagCober As Long
Dim glNumIntDocTitRecReembolso As Long

Dim iValorSolAlterado As Integer
Dim iValorAutoAlterado As Integer
Dim iTipoProc As Integer

Const TIPO_PROC_PERGUNTAR = 0
Const TIPO_PROC_NAO_PERGUNTAR_E_ALTERAR = 1
Const TIPO_PROC_NAO_PERGUNTAR_E_NAO_ALTERAR = 2

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoBanco As AdmEvento
Attribute objEventoBanco.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoCidade As AdmEvento
Attribute objEventoCidade.VB_VarHelpID = -1
Private WithEvents objEventoPais As AdmEvento
Attribute objEventoPais.VB_VarHelpID = -1

Dim objGridEv As AdmGrid
Dim iGrid_EvData_Col As Integer
Dim iGrid_EvPais_Col As Integer
Dim iGrid_EvEstado_Col As Integer
Dim iGrid_EvCidade_Col As Integer
Dim iGrid_EvTipo_Col As Integer
Dim iGrid_EvTelefone_Col As Integer

Dim objGridGA As AdmGrid
Dim iGrid_GAData_Col As Integer
Dim iGrid_GAValor_Col As Integer
Dim iGrid_GADesc_Col As Integer

Dim objGridDN As AdmGrid
Dim iGrid_DNNS_Col As Integer
Dim iGrid_DNNU_Col As Integer
Dim iGrid_DNRecebido_Col As Integer
Dim iGrid_DNDesc_Col As Integer
Dim iGrid_DNObs_Col As Integer

Dim objGridSrv As AdmGrid
Dim iGrid_SrvSol_Col As Integer
Dim iGrid_SrvAuto_Col As Integer
Dim iGrid_SrvDescricao_Col As Integer
Dim iGrid_SrvVlrSolRS_Col As Integer
Dim iGrid_SrvVlrSolUS_Col As Integer
Dim iGrid_SrvVlrAutoRS_Col As Integer
Dim iGrid_SrvVlrAutoUS_Col As Integer
Dim iGrid_SrvMoeda_Col As Integer
Dim iGrid_SrvVlrLimite_Col As Integer
Dim iGrid_SrvTipo_Col As Integer

Dim objGridII As AdmGrid
Dim iGrid_IIDataRec_Col As Integer
Dim iGrid_IIDataEmi_Col As Integer
Dim iGrid_IINumero_Col As Integer
'Dim iGrid_IIMoeda_Col As Integer
Dim iGrid_IIValorMoeda_Col As Integer
Dim iGrid_IIValorRS_Col As Integer
Dim iGrid_IIObs_Col As Integer

Dim objGridOF As AdmGrid
Dim iGrid_OFConsiderar_Col As Integer
Dim iGrid_OFDataRec_Col As Integer
Dim iGrid_OFDataEmi_Col As Integer
Dim iGrid_OFNumero_Col As Integer
Dim iGrid_OFValorUS_Col As Integer
Dim iGrid_OFValorRS_Col As Integer
Dim iGrid_OFDescricao_Col As Integer

Dim objGridHist As AdmGrid
Dim iGrid_HistData_Col As Integer
Dim iGrid_HistHora_Col As Integer
Dim iGrid_HistAssunto_Col As Integer
Dim iGrid_HistOrigem_Col As Integer
Dim iGrid_HistUsuario_Col As Integer

Dim objGridAnot As AdmGrid
Dim iGrid_AnotData_Col As Integer
Dim iGrid_AnotHora_Col As Integer
Dim iGrid_AnotTexto_Col As Integer
Dim iGrid_AnotUsuario_Col As Integer

Dim objGridPC As AdmGrid
Dim iGrid_PCValor_Col As Integer
Dim iGrid_PCVencimento_Col As Integer
Dim iGrid_PCPagamento_Col As Integer

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Ocorrências da Assistência"
    Call Form_Load

End Function

Public Function Name() As String
    Name = "TRVOcrCasos"
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

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
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
    Set objEventoBanco = Nothing
    Set objEventoCliente = Nothing
    Set objEventoProduto = Nothing
    Set objEventoCidade = Nothing
    Set objEventoPais = Nothing
    
    Set objGridEv = Nothing
    Set objGridSrv = Nothing
    Set objGridII = Nothing
    Set objGridOF = Nothing
    Set objGridHist = Nothing
    Set objGridAnot = Nothing
    Set objGridPC = Nothing
    Set objGridGA = Nothing
    Set objGridDN = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208555)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long, iIndice As Integer
Dim colCodigo As New Collection
Dim vCodigo As Variant
Dim objCodigoDescricao As AdmCodigoNome
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Form_Load

    Set objEventoCodigo = New AdmEvento
    Set objEventoBanco = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoCidade = New AdmEvento
    Set objEventoPais = New AdmEvento
    
    Set objGridEv = New AdmGrid
    Set objGridSrv = New AdmGrid
    Set objGridII = New AdmGrid
    Set objGridOF = New AdmGrid
    Set objGridHist = New AdmGrid
    Set objGridAnot = New AdmGrid
    Set objGridPC = New AdmGrid
    Set objGridGA = New AdmGrid
    Set objGridDN = New AdmGrid
    
    CGAnalise.Clear
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TRVOCRCASO_ANALISE, CGAnalise)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    CGAutorizadoPor.Clear
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TRVOCRCASO_AUTOPOR, CGAutorizadoPor)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    CGStatus.Clear
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TRVOCRCASO_STATUS, CGStatus)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Inicializa_Grid_EV(objGridEv)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Inicializa_Grid_SRV(objGridSrv)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Inicializa_Grid_II(objGridII)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Inicializa_Grid_OF(objGridOF)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Inicializa_Grid_Hist(objGridHist)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Inicializa_Grid_Anot(objGridAnot)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Inicializa_Grid_PC(objGridPC)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Inicializa_Grid_GA(objGridGA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Inicializa_Grid_DN(objGridDN)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Lê cada codigo da tabela Estados
    lErro = CF("Codigos_Le", "Estados", "Sigla", TIPO_STR, colCodigo, STRING_ESTADOS_SIGLA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Lê cada codigo e descricao da tabela Paises
    lErro = CF("Cod_Nomes_Le", "Paises", "Codigo", "Nome", STRING_PAISES_NOME, colCodigoDescricao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Estado.Clear
    
    'Preenche as ComboBox Estados com os objetos da colecao colCodigo
    For Each vCodigo In colCodigo
        Estado.AddItem vCodigo
    Next
    
    'Preenche cada ComboBox País com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        Pais.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        Pais.ItemData(Pais.NewIndex) = objCodigoDescricao.iCodigo
    Next
    
    Logradouro.MaxLength = STRING_ENDERECO
    Bairro.MaxLength = STRING_BAIRRO
    Cidade.MaxLength = STRING_CIDADE
    Telefone1.MaxLength = STRING_TELEFONE
    Telefone2.MaxLength = STRING_TELEFONE
    Email1.MaxLength = STRING_EMAIL
    Contato.MaxLength = STRING_CONTATO
    
   'Seleciona Brasil se existir
    For iIndice = 0 To Pais.ListCount - 1
        If right(Pais.Text, 6) = "Brasil" Then
            giIndexBrasil = iIndice
            Exit For
        End If
    Next
    
    'carrega a combo de Moedas
    lErro = Carrega_Moeda()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call Limpa_Tela_TRVOcrCasos

    iFrameAtual = 1
    iAlterado = 0
    
    iValorSolAlterado = 0
    iValorAutoAlterado = 0
    iTipoProc = TIPO_PROC_PERGUNTAR

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208556)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Function Inicializa_Grid_EV(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_EV

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("País")
    objGridInt.colColuna.Add ("Estado")
    objGridInt.colColuna.Add ("Cidade")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Telefone")

    'campos de edição do grid
    objGridInt.colCampo.Add (EvData.Name)
    objGridInt.colCampo.Add (EvPais.Name)
    objGridInt.colCampo.Add (EvEstado.Name)
    objGridInt.colCampo.Add (EvCidade.Name)
    objGridInt.colCampo.Add (EvTipo.Name)
    objGridInt.colCampo.Add (EvTelefone.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_EvData_Col = 1
    iGrid_EvPais_Col = 2
    iGrid_EvEstado_Col = 3
    iGrid_EvCidade_Col = 4
    iGrid_EvTipo_Col = 5
    iGrid_EvTelefone_Col = 6

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridEv

    'Linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 2
    
    'Largura da primeira coluna
    objGridInt.objGrid.ColWidth(0) = 300
    
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_EV = SUCESSO

    Exit Function

Erro_Inicializa_Grid_EV:

    Inicializa_Grid_EV = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208557)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_SRV(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_SRV

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Sol.")
    objGridInt.colColuna.Add ("Auto.")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Solic. U$")
    objGridInt.colColuna.Add ("Solic. R$")
    objGridInt.colColuna.Add ("Auto. U$")
    objGridInt.colColuna.Add ("Auto. R$")
    objGridInt.colColuna.Add ("Moeda")
    objGridInt.colColuna.Add ("Limite")
    objGridInt.colColuna.Add ("Tipo")

    'campos de edição do grid
    objGridInt.colCampo.Add (SrvSol.Name)
    objGridInt.colCampo.Add (SrvAuto.Name)
    objGridInt.colCampo.Add (SrvDescricao.Name)
    objGridInt.colCampo.Add (SrvVlrSolUS.Name)
    objGridInt.colCampo.Add (SrvVlrSolRS.Name)
    objGridInt.colCampo.Add (SrvVlrAutoUS.Name)
    objGridInt.colCampo.Add (SrvVlrAutoRS.Name)
    objGridInt.colCampo.Add (SrvMoeda.Name)
    objGridInt.colCampo.Add (SrvVlrLimite.Name)
    objGridInt.colCampo.Add (SrvTipo.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_SrvSol_Col = 1
    iGrid_SrvAuto_Col = 2
    iGrid_SrvDescricao_Col = 3
    iGrid_SrvVlrSolUS_Col = 4
    iGrid_SrvVlrSolRS_Col = 5
    iGrid_SrvVlrAutoUS_Col = 6
    iGrid_SrvVlrAutoRS_Col = 7
    iGrid_SrvMoeda_Col = 8
    iGrid_SrvVlrLimite_Col = 9
    iGrid_SrvTipo_Col = 10

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridSrv

    'Linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 4
    
    'Largura da primeira coluna
    objGridInt.objGrid.ColWidth(0) = 300
    
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_SRV = SUCESSO

    Exit Function

Erro_Inicializa_Grid_SRV:

    Inicializa_Grid_SRV = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208558)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_II(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_II

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Recebto")
    objGridInt.colColuna.Add ("Emissão")
    objGridInt.colColuna.Add ("Número")
'    objGridInt.colColuna.Add ("Moeda")
    objGridInt.colColuna.Add ("Valor U$")
    objGridInt.colColuna.Add ("Valor R$")
    objGridInt.colColuna.Add ("Observação")

    'campos de edição do grid
    objGridInt.colCampo.Add (IIDataRec.Name)
    objGridInt.colCampo.Add (IIDataEmi.Name)
    objGridInt.colCampo.Add (IINumero.Name)
'    objGridInt.colCampo.Add (IIMoeda.Name)
    objGridInt.colCampo.Add (IIValorMoeda.Name)
    objGridInt.colCampo.Add (IIValorRS.Name)
    objGridInt.colCampo.Add (IIObs.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_IIDataRec_Col = 1
    iGrid_IIDataEmi_Col = 2
    iGrid_IINumero_Col = 3
'    iGrid_IIMoeda_Col = 4
    iGrid_IIValorMoeda_Col = 4
    iGrid_IIValorRS_Col = 5
    iGrid_IIObs_Col = 6

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridII

    'Linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 5
    
    'Largura da primeira coluna
    objGridInt.objGrid.ColWidth(0) = 300

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_II = SUCESSO

    Exit Function

Erro_Inicializa_Grid_II:

    Inicializa_Grid_II = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208559)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_OF(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_OF

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("C")
    objGridInt.colColuna.Add ("Recebto")
    objGridInt.colColuna.Add ("Emissão")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Valor U$")
    objGridInt.colColuna.Add ("Valor R$")
    objGridInt.colColuna.Add ("Descrição")

    'campos de edição do grid
    objGridInt.colCampo.Add (OFConsiderar.Name)
    objGridInt.colCampo.Add (OFDataRec.Name)
    objGridInt.colCampo.Add (OFDataEmi.Name)
    objGridInt.colCampo.Add (OFNumero.Name)
    objGridInt.colCampo.Add (OFValorUS.Name)
    objGridInt.colCampo.Add (OFValorRS.Name)
    objGridInt.colCampo.Add (OFDescricao.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_OFConsiderar_Col = 1
    iGrid_OFDataRec_Col = 2
    iGrid_OFDataEmi_Col = 3
    iGrid_OFNumero_Col = 4
    iGrid_OFValorUS_Col = 5
    iGrid_OFValorRS_Col = 6
    iGrid_OFDescricao_Col = 7

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridOF

    'Linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 5
    
    'Largura da primeira coluna
    objGridInt.objGrid.ColWidth(0) = 300

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_OF = SUCESSO

    Exit Function

Erro_Inicializa_Grid_OF:

    Inicializa_Grid_OF = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208560)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Hist(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Hist

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Hora")
    objGridInt.colColuna.Add ("Assunto")
    objGridInt.colColuna.Add ("Origem")
    objGridInt.colColuna.Add ("Usuário")

    'campos de edição do grid
    objGridInt.colCampo.Add (HistData.Name)
    objGridInt.colCampo.Add (HistHora.Name)
    objGridInt.colCampo.Add (HistAssunto.Name)
    objGridInt.colCampo.Add (HistOrigem.Name)
    objGridInt.colCampo.Add (HistUsuario.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_HistData_Col = 1
    iGrid_HistHora_Col = 2
    iGrid_HistAssunto_Col = 3
    iGrid_HistOrigem_Col = 4
    iGrid_HistUsuario_Col = 5

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridHist

    'Linhas do grid
    objGridInt.objGrid.Rows = 500 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 5
    
    'Largura da primeira coluna
    objGridInt.objGrid.ColWidth(0) = 300

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Hist = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Hist:

    Inicializa_Grid_Hist = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208561)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Anot(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Anot

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Hora")
    objGridInt.colColuna.Add ("Assunto")
    objGridInt.colColuna.Add ("Usuário")

    'campos de edição do grid
    objGridInt.colCampo.Add (AnotData.Name)
    objGridInt.colCampo.Add (AnotHora.Name)
    objGridInt.colCampo.Add (AnotTexto.Name)
    objGridInt.colCampo.Add (AnotUsuario.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_AnotData_Col = 1
    iGrid_AnotHora_Col = 2
    iGrid_AnotTexto_Col = 3
    iGrid_AnotUsuario_Col = 4

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridAnot

    'Linhas do grid
    objGridInt.objGrid.Rows = 500 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 5
    
    'Largura da primeira coluna
    objGridInt.objGrid.ColWidth(0) = 300

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Anot = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Anot:

    Inicializa_Grid_Anot = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208562)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_PC(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_PC

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Pagto")

    'campos de edição do grid
    objGridInt.colCampo.Add (PCValor.Name)
    objGridInt.colCampo.Add (PCVencimento.Name)
    objGridInt.colCampo.Add (PCPagamento.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_PCValor_Col = 1
    iGrid_PCVencimento_Col = 2
    iGrid_PCPagamento_Col = 3

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridPC

    'Linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 5
    
    'Largura da primeira coluna
    objGridInt.objGrid.ColWidth(0) = 300

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_PC = SUCESSO

    Exit Function

Erro_Inicializa_Grid_PC:

    Inicializa_Grid_PC = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208563)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objOcrCasos Is Nothing) Then

        lErro = Traz_TRVOcrCasos_Tela(objOcrCasos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208564)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim sProduto As String, iPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria

    objOcrCasos.sCodigo = Trim(Codigo.Text)
    objOcrCasos.sTipVou = TipVou.Text
    objOcrCasos.sSerie = Serie.Text
    If Len(Trim(NumVou.Text)) < 10 Then objOcrCasos.lNumVou = StrParaLong(NumVou.Text)
    objOcrCasos.sNumVouTexto = NumVou.Text
    objOcrCasos.spaxnome = VouPaxNome.Text
    objOcrCasos.sTitularNome = VouTitular.Text
    objOcrCasos.lClienteVou = LCodigo_Extrai(ClienteVou.Text)
    objOcrCasos.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objOcrCasos.dtDataIda = StrParaDate(DataIda.Text)
    objOcrCasos.dtDataVolta = StrParaDate(DataVolta.Text)
    objOcrCasos.iQtdPax = StrParaInt(QtdePax.Text)
    
    lErro = CF("Produto_Formata", Produto.Text, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    objOcrCasos.sProduto = sProduto

    If Len(Trim(DataAbertura.Caption)) <> 0 Then
        objOcrCasos.dtDataAbertura = StrParaDate(DataAbertura.Caption)
    Else
        objOcrCasos.dtDataAbertura = gdtDataAtual
    End If
    
    objOcrCasos.dtDataDocsRec = StrParaDate(DataDocsRec.Text)
    objOcrCasos.dtDataEnvioAnalise = StrParaDate(DataEnvioAnalise.Text)
    objOcrCasos.lCGAnalise = LCodigo_Extrai(CGAnalise.Text)
    objOcrCasos.lCGStatus = LCodigo_Extrai(CGStatus.Text)
    objOcrCasos.lCGAutorizadoPor = LCodigo_Extrai(CGAutorizadoPor.Text)
    objOcrCasos.dValorAutorizadoTotalRS = StrParaDbl(TotalAutoRS.Caption)
    objOcrCasos.dValorAutorizadoTotalUS = StrParaDbl(TotalAutoUS.Caption)
    objOcrCasos.dtDataLimite = StrParaDate(DataLimite.Text)
    objOcrCasos.dtDataProgFinanc = StrParaDate(DataProgrFin.Text)
    objOcrCasos.dtDataPagtoPax = StrParaDate(DataPagtoPax.Text)
    
    If Judicial.Value = vbChecked Then
        objOcrCasos.iJudicial = MARCADO
    Else
        objOcrCasos.iJudicial = DESMARCADO
    End If
    
    If Condenado.Value = vbChecked Then
        objOcrCasos.iCondenado = MARCADO
        If OptCondenacao.Value Then
            objOcrCasos.iPerdaTipo = TRV_OCRCASOS_PERDA_CONDENACAO
        Else
            objOcrCasos.iPerdaTipo = TRV_OCRCASOS_PERDA_ACORDO
        End If
    Else
        objOcrCasos.iCondenado = DESMARCADO
        objOcrCasos.iPerdaTipo = 0
    End If
    
    objOcrCasos.sNumProcesso = NumProcesso.Text
    objOcrCasos.dValorCondenacao = StrParaDbl(ValorCondenacao.Text)
    objOcrCasos.dProcessoDanoMoral = StrParaDbl(DanoMoral.Text)
    objOcrCasos.dProcessoDanoMaterial = StrParaDbl(DanoMaterial.Text)
    objOcrCasos.sComarca = Comarca.Text
    objOcrCasos.dtDataFimProcesso = StrParaDate(DataFimProcesso.Text)
    objOcrCasos.dtDataIniProcesso = StrParaDate(DataIniProcesso.Text)
    If Procon.Value = vbChecked Then
        objOcrCasos.iProcon = MARCADO
    Else
        objOcrCasos.iProcon = DESMARCADO
    End If
    
    objOcrCasos.dValorAutorizadoSeguroRS = StrParaDbl(SrvTotalSegRS.Caption)
    objOcrCasos.dValorAutoSegRespTrvRS = StrParaDbl(SrvTotalSegTrvRS.Text)
    objOcrCasos.dValorAutorizadoSeguroUS = StrParaDbl(SrvTotalSegUS.Caption)
    objOcrCasos.dValorAutorizadoAssistRS = StrParaDbl(SrvTotalAssistRS.Caption)
    objOcrCasos.dValorAutorizadoAssistUS = StrParaDbl(SrvTotalAssistUS.Caption)
    
    objOcrCasos.dCambio = StrParaDbl(Cambio.Text)
    
    If AntecPagto.Value = vbChecked Then
        objOcrCasos.iAnteciparPagtoSeguro = MARCADO
    Else
        objOcrCasos.iAnteciparPagtoSeguro = DESMARCADO
    End If
    
    objOcrCasos.iBanco = StrParaInt(Banco.Text)
    objOcrCasos.sAgencia = Agencia.Text
    objOcrCasos.sContaCorrente = ContaCorrente.Text
    objOcrCasos.sNomeFavorecido = NomeFavorecido.Text
    
    objOcrCasos.sFavorecidoCGC = FavorecidoCGC.ClipText
    
    objOcrCasos.dValorInvoicesTotal = StrParaDbl(IITotalRS.Caption)
    objOcrCasos.dValorInvoicesTotalUS = StrParaDbl(IITotalUS.Caption)
    objOcrCasos.dValorDespesasTotalRS = StrParaDbl(OFTotalRS.Caption)
    objOcrCasos.dValorDespesasTotalUS = StrParaDbl(OFTotalUS.Caption)
    
    objOcrCasos.dValorGastosAdvRS = StrParaDbl(GATotal.Caption)
    
    objOcrCasos.sOBS = Obs.Text
    
    lErro = Move_Endereco_Memoria(objOcrCasos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Move_Srv_Memoria(objOcrCasos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Move_II_Memoria(objOcrCasos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Move_OF_Memoria(objOcrCasos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    lErro = Move_PC_Memoria(objOcrCasos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    lErro = Move_Docs_Memoria(objOcrCasos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    lErro = Move_GAdv_Memoria(objOcrCasos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208565)

    End Select

    Exit Function

End Function

Function Move_Endereco_Memoria(objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim objEndereco As ClassEndereco

On Error GoTo Erro_Move_Endereco_Memoria

    Set objEndereco = objOcrCasos.objEndereco

    objEndereco.sLogradouro = Trim(Logradouro.Text)
    objEndereco.sComplemento = Trim(Complemento.Text)
    objEndereco.lNumero = StrParaLong(Numero.Text)
    objEndereco.sBairro = Trim(Bairro.Text)
    objEndereco.sCidade = Trim(Cidade.Text)
    objEndereco.sCEP = Trim(CEP.Text)
    
    objEndereco.sTelNumero1 = Telefone1.Text
    objEndereco.sTelNumero2 = Telefone2.Text

    objEndereco.iCodigoPais = Codigo_Extrai(Pais.Text)
    objEndereco.sSiglaEstado = Trim(Estado.Text)
    If objEndereco.iCodigoPais = 0 Then objEndereco.iCodigoPais = PAIS_BRASIL
    If objEndereco.iCodigoPais = PAIS_BRASIL And (Estado.ListIndex = -1 Or Len(Trim(Estado.Text)) = 0) Then Estado.Text = "SP"
   
    objEndereco.sEmail = Trim(Email1.Text)
    objEndereco.sContato = Trim(Contato.Text)

    Move_Endereco_Memoria = SUCESSO

    Exit Function

Erro_Move_Endereco_Memoria:

    Move_Endereco_Memoria = gErr

    Select Case gErr
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208585)

    End Select

    Exit Function
    
End Function

Function Move_Srv_Memoria(objOcrCasos As ClassTRVOcrCasos, Optional bForca As Boolean = False) As Long

Dim lErro As Long
Dim iLinha As Integer, sDescGrid As String
Dim objOcrCasosSrv As ClassTRVOcrCasosSrv

On Error GoTo Erro_Move_Srv_Memoria

    For iLinha = 1 To objGridSrv.iLinhasExistentes
    
        Set objOcrCasosSrv = New ClassTRVOcrCasosSrv
    
        objOcrCasosSrv.dValorAutorizadoRS = StrParaDbl(GridSrv.TextMatrix(iLinha, iGrid_SrvVlrAutoRS_Col))
        objOcrCasosSrv.dValorAutorizadoUS = StrParaDbl(GridSrv.TextMatrix(iLinha, iGrid_SrvVlrAutoUS_Col))
        objOcrCasosSrv.dValorLimite = StrParaDbl(GridSrv.TextMatrix(iLinha, iGrid_SrvVlrLimite_Col))
        objOcrCasosSrv.dValorSolicitadoRS = StrParaDbl(GridSrv.TextMatrix(iLinha, iGrid_SrvVlrSolRS_Col))
        objOcrCasosSrv.dValorSolicitadoUS = StrParaDbl(GridSrv.TextMatrix(iLinha, iGrid_SrvVlrSolUS_Col))
        objOcrCasosSrv.iAutorizado = StrParaInt(GridSrv.TextMatrix(iLinha, iGrid_SrvAuto_Col))
        objOcrCasosSrv.iMoeda = Codigo_Extrai(GridSrv.TextMatrix(iLinha, iGrid_SrvMoeda_Col))
        objOcrCasosSrv.iTipo = Codigo_Extrai(GridSrv.TextMatrix(iLinha, iGrid_SrvTipo_Col))
        objOcrCasosSrv.iSolicitado = StrParaInt(GridSrv.TextMatrix(iLinha, iGrid_SrvSol_Col))
        
        sDescGrid = GridSrv.TextMatrix(iLinha, iGrid_SrvDescricao_Col)
        
        objOcrCasosSrv.sDescricao = Mid(sDescGrid, InStr(1, sDescGrid, "-") + 1)
        objOcrCasosSrv.lCodigoServ = LCodigo_Extrai(sDescGrid)
        
        objOcrCasosSrv.iSeq = iLinha
        
        If objOcrCasosSrv.iSolicitado = MARCADO Or bForca Then objOcrCasos.colCoberturas.Add objOcrCasosSrv

    Next

    Move_Srv_Memoria = SUCESSO

    Exit Function

Erro_Move_Srv_Memoria:

    Move_Srv_Memoria = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208586)

    End Select

    Exit Function
    
End Function

Function Move_II_Memoria(objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objOcrCasosInvoices As ClassTRvOcrCasosInvoices

On Error GoTo Erro_Move_II_Memoria

    For iLinha = 1 To objGridII.iLinhasExistentes
    
        Set objOcrCasosInvoices = New ClassTRvOcrCasosInvoices
    
        objOcrCasosInvoices.dtDataFatura = StrParaDate(GridII.TextMatrix(iLinha, iGrid_IIDataEmi_Col))
        objOcrCasosInvoices.dtDataRecepcao = StrParaDate(GridII.TextMatrix(iLinha, iGrid_IIDataRec_Col))
        objOcrCasosInvoices.dValorMoeda = StrParaDbl(GridII.TextMatrix(iLinha, iGrid_IIValorMoeda_Col))
        objOcrCasosInvoices.dValorRS = StrParaDbl(GridII.TextMatrix(iLinha, iGrid_IIValorRS_Col))
        objOcrCasosInvoices.iMoeda = MOEDA_DOLAR 'Codigo_Extrai(GridII.TextMatrix(iLinha, iGrid_IIMoeda_Col))
        objOcrCasosInvoices.sNumero = GridII.TextMatrix(iLinha, iGrid_IINumero_Col)
        objOcrCasosInvoices.sOBS = GridII.TextMatrix(iLinha, iGrid_IIObs_Col)
        
        objOcrCasosInvoices.iSeq = iLinha
        
        objOcrCasos.colInvoices.Add objOcrCasosInvoices

    Next

    Move_II_Memoria = SUCESSO

    Exit Function

Erro_Move_II_Memoria:

    Move_II_Memoria = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208588)

    End Select

    Exit Function
    
End Function

Function Move_OF_Memoria(objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objOcrCasosOF As ClassTRVOcrCasosOutrasFat

On Error GoTo Erro_Move_OF_Memoria

    For iLinha = 1 To objGridOF.iLinhasExistentes
    
        Set objOcrCasosOF = New ClassTRVOcrCasosOutrasFat
    
        objOcrCasosOF.dtDataFatura = StrParaDate(GridOF.TextMatrix(iLinha, iGrid_OFDataEmi_Col))
        objOcrCasosOF.dtDataRecepcao = StrParaDate(GridOF.TextMatrix(iLinha, iGrid_OFDataRec_Col))
        objOcrCasosOF.dValorUS = StrParaDbl(GridOF.TextMatrix(iLinha, iGrid_OFValorUS_Col))
        objOcrCasosOF.dValorRS = StrParaDbl(GridOF.TextMatrix(iLinha, iGrid_OFValorRS_Col))
        objOcrCasosOF.sNumero = GridOF.TextMatrix(iLinha, iGrid_OFNumero_Col)
        objOcrCasosOF.sDescricao = GridOF.TextMatrix(iLinha, iGrid_OFDescricao_Col)
        objOcrCasosOF.iConsiderar = StrParaInt(GridOF.TextMatrix(iLinha, iGrid_OFConsiderar_Col))
        
        objOcrCasosOF.iSeq = iLinha
        
        objOcrCasos.colOutrasFaturas.Add objOcrCasosOF

    Next

    Move_OF_Memoria = SUCESSO

    Exit Function

Erro_Move_OF_Memoria:

    Move_OF_Memoria = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208589)

    End Select

    Exit Function
    
End Function

Function Move_PC_Memoria(objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objOcrCasosPC As ClassTRVOcrCasosParcCond

On Error GoTo Erro_Move_PC_Memoria

    For iLinha = 1 To objGridPC.iLinhasExistentes
    
        Set objOcrCasosPC = New ClassTRVOcrCasosParcCond
    
        objOcrCasosPC.dtDataVencimento = StrParaDate(GridPC.TextMatrix(iLinha, iGrid_PCVencimento_Col))
        objOcrCasosPC.dtDataPagto = StrParaDate(GridPC.TextMatrix(iLinha, iGrid_PCPagamento_Col))
        objOcrCasosPC.dValor = StrParaDbl(GridPC.TextMatrix(iLinha, iGrid_PCValor_Col))
        
        objOcrCasosPC.iSeq = iLinha
        
        objOcrCasos.colParcProcesso.Add objOcrCasosPC

    Next

    Move_PC_Memoria = SUCESSO

    Exit Function

Erro_Move_PC_Memoria:

    Move_PC_Memoria = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208590)

    End Select

    Exit Function
    
End Function

Function Valida_Dados_Tela(objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim iIndice As Integer, dValorTotal As Double
Dim objOcrCasosBD As New ClassTRVOcrCasos
Dim objOcrCasosAux As ClassTRVOcrCasos
Dim objOcrCasosSrv As ClassTRVOcrCasosSrv
Dim objOcrCasosInvoices As ClassTRvOcrCasosInvoices
Dim objOcrCasosOF As ClassTRVOcrCasosOutrasFat
Dim objOcrCasosPC As ClassTRVOcrCasosParcCond
Dim objOcrCasosPCBD As ClassTRVOcrCasosParcCond
Dim objOcrCasosEv As ClassTRVOcrCasoImport
Dim vbResult As VbMsgBoxResult, colRetorno As New Collection
Dim vValor As Variant, objVou As ClassTRVVouchers

On Error GoTo Erro_Valida_Dados_Tela

    If glNumIntDoc = 0 Then gError 208650
    
    objOcrCasosBD.lNumIntDoc = glNumIntDoc

    lErro = CF("TRVOcrCasos_Le", objOcrCasosBD)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM

    If lErro = ERRO_LEITURA_SEM_DADOS Then gError 208651
    
    'Não pode alterar o código
    If UCase(objOcrCasosBD.sCodigo) <> UCase(objOcrCasos.sCodigo) Then gError 208652
    
'    'Não pode alterar o voucher
'    If objOcrCasosBD.lNumVou <> objOcrCasos.lNumVou Then gError 208653

    objOcrCasos.lNumIntDocTitPagProcesso = objOcrCasosBD.lNumIntDocTitPagProcesso
    objOcrCasos.lNumIntDocTitPagCobertura = objOcrCasosBD.lNumIntDocTitPagCobertura
    
    'Se já faturou a condenação judicial
    If objOcrCasosBD.lNumIntDocTitPagProcesso <> 0 Then
    
        'Não pode mexer no valor e nem nas parcelas
        If Abs(objOcrCasosBD.dValorCondenacao - objOcrCasos.dValorCondenacao) > DELTA_VALORMONETARIO Then gError 208654
        
        If objOcrCasosBD.colParcProcesso.Count <> objOcrCasos.colParcProcesso.Count Then gError 208655
        
        For iIndice = 1 To objOcrCasosBD.colParcProcesso.Count
            Set objOcrCasosPC = objOcrCasos.colParcProcesso.Item(iIndice)
            Set objOcrCasosPCBD = objOcrCasosBD.colParcProcesso.Item(iIndice)
        
            If objOcrCasosPC.dtDataVencimento <> objOcrCasosPCBD.dtDataVencimento Then gError 208656
            If Abs(objOcrCasosPC.dValor - objOcrCasosPCBD.dValor) > DELTA_VALORMONETARIO Then gError 208657
        Next
        
    End If
    
    'Se já faturou ao passageiro
    If objOcrCasosBD.lNumIntDocTitPagCobertura <> 0 Then
    
        'Não pode marcar ou desmarcar uma antecipação de pagamento
        If objOcrCasosBD.iAnteciparPagtoSeguro <> objOcrCasos.iAnteciparPagtoSeguro Then gError 208658
    
        'Não pode mexer no valor do que é assistência e nem do que é seguro com antecipação de pagamento
        If Abs(objOcrCasosBD.dValorAutorizadoAssistRS - objOcrCasos.dValorAutorizadoAssistRS) > DELTA_VALORMONETARIO Then gError 208659
        If objOcrCasosBD.iAnteciparPagtoSeguro = MARCADO And Abs(objOcrCasosBD.dValorAutorizadoSeguroRS - objOcrCasos.dValorAutorizadoSeguroRS) > DELTA_VALORMONETARIO Then gError 208660
        If Abs(objOcrCasosBD.dValorAutoSegRespTrvRS - objOcrCasos.dValorAutoSegRespTrvRS) > DELTA_VALORMONETARIO Then gError 208660
        
        If objOcrCasosBD.dtDataLimite <> objOcrCasos.dtDataLimite Then gError 209265
        If objOcrCasosBD.dtDataProgFinanc <> objOcrCasos.dtDataProgFinanc Then gError 209266
        
        If objOcrCasosBD.lCGAnalise <> objOcrCasos.lCGAnalise Then gError 208681
        If objOcrCasosBD.lCGAutorizadoPor <> objOcrCasos.lCGAutorizadoPor Then gError 208682
        If objOcrCasosBD.lCGStatus <> objOcrCasos.lCGStatus Then gError 208683
    End If

    iIndice = 0
    For Each objOcrCasosSrv In objOcrCasos.colCoberturas
    
        iIndice = iIndice + 1
    
        If objOcrCasosSrv.iAutorizado = MARCADO And objOcrCasosSrv.dValorAutorizadoRS = 0 Then gError 208661
    
        'Verifica se o valor autorizado está dentro do limite
        If objOcrCasosSrv.iMoeda = MOEDA_DOLAR And objOcrCasosSrv.dValorLimite > 0 Then
            If objOcrCasosSrv.dValorAutorizadoUS > objOcrCasosSrv.dValorLimite + DELTA_VALORMONETARIO Then gError 208662
        ElseIf objOcrCasosSrv.iMoeda = MOEDA_REAL And objOcrCasosSrv.dValorLimite > 0 Then
            If objOcrCasosSrv.dValorAutorizadoRS > objOcrCasosSrv.dValorLimite + DELTA_VALORMONETARIO Then gError 208663
        End If
    Next
    
    iIndice = 0
    For Each objOcrCasosInvoices In objOcrCasos.colInvoices
    
        iIndice = iIndice + 1
        
        'Informa caso existam faturas sem valor ou data
        If objOcrCasosInvoices.dtDataRecepcao = DATA_NULA Then gError 208664
        If objOcrCasosInvoices.dValorMoeda = 0 Then gError 208665
    Next
    
    iIndice = 0
    vbResult = vbNo
    For Each objOcrCasosOF In objOcrCasos.colOutrasFaturas
    
        iIndice = iIndice + 1
        
        'Informa caso existam faturas sem valor ou data
        If objOcrCasosOF.dtDataRecepcao = DATA_NULA Then gError 208666
        If objOcrCasosOF.dValorRS = 0 Then gError 208667
        
        'Se ainda não disse que era para continuar
        If vbResult = vbNo Then
            'Se está considerando a fatura como despesa e ele tem data de emissão
            If objOcrCasosOF.iConsiderar = MARCADO And objOcrCasosOF.dtDataFatura <> DATA_NULA Then
                'Se a vigência do voucher está preenchida
                If objOcrCasos.dtDataIda <> DATA_NULA And objOcrCasos.dtDataVolta <> DATA_NULA Then
                    'Se a data de fatura está fora do período de vigência
                    If objOcrCasosOF.dtDataFatura < objOcrCasos.dtDataIda Or objOcrCasosOF.dtDataFatura > objOcrCasos.dtDataVolta Then
                        vbResult = Rotina_Aviso(vbYesNo, "AVISO_TRVOCRCASOS_FATURA_FORA_VIGENCIA")
                        If vbResult = vbNo Then gError ERRO_SEM_MENSAGEM
                        Exit For
                    End If
                End If
            End If
        End If
        
    Next
    
    iIndice = 0
    dValorTotal = 0
    For Each objOcrCasosPC In objOcrCasos.colParcProcesso
    
        iIndice = iIndice + 1
        
        'Informa caso existam parcelas sem valor ou data
        If objOcrCasosPC.dtDataVencimento = DATA_NULA Then gError 208668
        If objOcrCasosPC.dValor = 0 Then gError 208669
        
        dValorTotal = dValorTotal + objOcrCasosPC.dValor
    Next
    
    'O valor da condenação tem que bater com a soma das parcelas a pagar
    If Abs(objOcrCasos.dValorCondenacao - dValorTotal) > DELTA_VALORMONETARIO Then gError 208670
    
    'Para gravar tem que abrir a ocorrência, sem essa data só terá os dados importados histórico e anotações
    If objOcrCasos.dtDataAbertura = DATA_NULA Then gError 208671
    
    'Tenta ler o voucher para pegar mais informações
    If Len(Trim(objOcrCasos.sSerie)) > 0 And Len(Trim(objOcrCasos.sTipVou)) > 0 Then
    
        Set objVou = New ClassTRVVouchers
    
        objVou.sTipVou = objOcrCasos.sTipVou
        objVou.sSerie = objOcrCasos.sSerie
        objVou.lNumVou = objOcrCasos.lNumVou
        objVou.sTipoDoc = TRV_TIPODOC_VOU_TEXTO
    
        lErro = CF("TRVVouchers_Le", objVou)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    
        If objVou.iStatus = STATUS_TRV_VOU_CANCELADO Then gError 208871
    
    End If
    
    'Le todas ocorrências desse voucher
    lErro = CF("TRVOcrCasos_Le_NumVou", objOcrCasosBD, colRetorno)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Para cada ocorrência
    For Each vValor In colRetorno
    
        Set objOcrCasosAux = New ClassTRVOcrCasos
    
        objOcrCasosAux.sCodigo = vValor
        
        'Se o código dela for diferente do código atual e ela também estiver aberta avisa
        If UCase(objOcrCasos.sCodigo) <> UCase(objOcrCasosAux.sCodigo) Then
    
            lErro = CF("TRVOcrCasos_Le", objOcrCasosAux)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
            
            If objOcrCasosAux.dtDataAbertura <> DATA_NULA Then
                vbResult = Rotina_Aviso(vbYesNo, "AVISO_TRVOCRCASOS_VOU_COM_OUT_OCR")
                If vbResult = vbNo Then gError ERRO_SEM_MENSAGEM
                Exit For
            End If
    
        End If
    
    Next
    
    If objOcrCasos.dtDataIda <> DATA_NULA And objOcrCasos.dtDataVolta <> DATA_NULA Then
        'Verifica a data de cada evento com relação a data de vigência do voucher se ela estiver preenchida
        For Each objOcrCasosEv In objOcrCasos.colEventos
            If objOcrCasosEv.dtData < objOcrCasos.dtDataIda Or objOcrCasosEv.dtData > objOcrCasos.dtDataVolta Then
                vbResult = Rotina_Aviso(vbYesNo, "AVISO_TRVOCRCASOS_EVENTO_FORA_VIGENCIA")
                If vbResult = vbNo Then gError ERRO_SEM_MENSAGEM
                Exit For
            End If
        Next
    End If
        
    'Se está autorizado
    If objOcrCasos.lCGAutorizadoPor <> 0 Then
        'Tem que ter o valor autorizado
        If objOcrCasos.dValorAutorizadoTotalRS = 0 Then gError 208672
        'Já tem que ter recebido todos os documentos
        If objOcrCasos.dtDataDocsRec = DATA_NULA Then gError 208673
        'Já tem que ter enviado para análise
        If objOcrCasos.dtDataEnvioAnalise = DATA_NULA Then gError 208674
        
        'Já tem que estar com todos os dados para pagamento
        If Len(Trim(objOcrCasos.sFavorecidoCGC)) = 0 Then gError 208675
        If Len(Trim(objOcrCasos.sNomeFavorecido)) = 0 Then gError 208676
        If Len(Trim(objOcrCasos.objEndereco.sLogradouro)) = 0 Then 'gError 208677
            vbResult = Rotina_Aviso(vbYesNo, "AVISO_TRVOCRCASOS_FAV_SEM_ENDERECO")
            If vbResult = vbNo Then gError ERRO_SEM_MENSAGEM
        End If
        If objOcrCasos.iBanco = 0 Then gError 208678
        If Len(Trim(objOcrCasos.sAgencia)) = 0 Then gError 208679
        If Len(Trim(objOcrCasos.sContaCorrente)) = 0 Then gError 208680
    End If
    
    lErro = CGC_Valida(FavorecidoCGC)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If StrParaDate(DataFimProcesso.Text) <> DATA_NULA And StrParaDate(DataIniProcesso.Text) <> DATA_NULA Then
        If StrParaDate(DataFimProcesso.Text) < StrParaDate(DataIniProcesso.Text) Then gError 209172
    End If

    Valida_Dados_Tela = SUCESSO

    Exit Function

Erro_Valida_Dados_Tela:

    Valida_Dados_Tela = gErr

    Select Case gErr
    
        Case 208650
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCRCASOS_TRAZER_TELA", gErr)
            
        Case 208651
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCRCASOS_NAO_CADASTRADO1", gErr, glNumIntDoc)

        Case 208652
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCRCASOS_CODIGO_DIF", gErr)

        Case 208653
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCRCASOS_NUMVOU_DIF", gErr)
            
        Case 208654 To 208657
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCRCASOS_PROC_JA_FAT", gErr)

        Case 208658 To 208660, 208681 To 208683
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCRCASOS_PAX_JA_FAT", gErr)
            
        Case 208661
            Call Rotina_Erro(vbOKOnly, "ERRO_VLRAUTORS_NAO_PREENCHIDO", gErr, iIndice)

        Case 208662, 208663
            Call Rotina_Erro(vbOKOnly, "ERRO_VLRAUTORS_MAIOR_VLRLIMITE", gErr, iIndice)

        Case 208664
            Call Rotina_Erro(vbOKOnly, "ERRO_IIDATAREC_NAO_PREENCHIDO", gErr, iIndice)

        Case 208665
            Call Rotina_Erro(vbOKOnly, "ERRO_IIVALORMOEDA_NAO_PREENCHIDO", gErr, iIndice)

        Case 208666
            Call Rotina_Erro(vbOKOnly, "ERRO_OFDATAREC_NAO_PREENCHIDO", gErr, iIndice)

        Case 208667
            Call Rotina_Erro(vbOKOnly, "ERRO_OFVALORRS_NAO_PREENCHIDO", gErr, iIndice)

        Case 208668
            Call Rotina_Erro(vbOKOnly, "ERRO_PCVENCIMENTO_NAO_PREENCHIDO", gErr, iIndice)

        Case 208669
            Call Rotina_Erro(vbOKOnly, "ERRO_PCVALOR_NAO_PREENCHIDO", gErr, iIndice)
            
        Case 208670
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCRCASOS_VLRCOND_DIF", gErr, Format(objOcrCasos.dValorCondenacao, "STANDARD"), Format(dValorTotal, "STANDARD"))

        Case 208671
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCRCASOS_DATAABERTURA_NAO_PREENCHIDA", gErr)

        Case 208672 To 208680
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCRCASOS_AUTO_NAO_PREENCHIDO", gErr)

        Case 208871
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCR_VOUCHER_CANCELADO", gErr)
            
        Case 209172
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCR_DATA_PROC_FIM_MENOR_INI", gErr)

        Case 209265
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCRCASOS_ALTER_DATALIMITE_PAX_JA_FAT", gErr)

        Case 209266
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCRCASOS_ALTER_DATAPROG_PAX_JA_FAT", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208587)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objOcrCasos As New ClassTRVOcrCasos

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TRVOcrCasos"

    objOcrCasos.sCodigo = Trim(Codigo.Text)

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objOcrCasos.sCodigo, STRING_TRV_MAXIMO, "Codigo"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208566)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objOcrCasos As New ClassTRVOcrCasos

On Error GoTo Erro_Tela_Preenche

    objOcrCasos.sCodigo = colCampoValor.Item("Codigo").vValor

    If Len(Trim(objOcrCasos.sCodigo)) > 0 Then

        lErro = Traz_TRVOcrCasos_Tela(objOcrCasos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208567)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objOcrCasos As New ClassTRVOcrCasos

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Preenche o objOcrCasos
    lErro = Move_Tela_Memoria(objOcrCasos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Valida_Dados_Tela(objOcrCasos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Grava o/a TRVOcrCasos no Banco de Dados
    lErro = CF("TRVOcrCasos_Grava", objOcrCasos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208568)

    End Select

    Exit Function

End Function

Function Limpa_Tela_TRVOcrCasos() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_TRVOcrCasos

    Call Grid_Limpa(objGridAnot)
    Call Grid_Limpa(objGridEv)
    Call Grid_Limpa(objGridHist)
    Call Grid_Limpa(objGridII)
    Call Grid_Limpa(objGridOF)
    Call Grid_Limpa(objGridPC)
    Call Grid_Limpa(objGridSrv)
    Call Grid_Limpa(objGridDN)
    Call Grid_Limpa(objGridGA)
    
    gsCEPAnt = ""
    gbMudouCEP = False
    
    glNumIntDoc = 0
    glNumIntDocTitPagProc = 0
    glNumIntDocTitPagCober = 0
    glNumIntDocTitRecReembolso = 0
    gsProdAnt = ""
    
    gdCambioAnt = 0
    
    Pais.ListIndex = giIndexBrasil
    giPaisAnt = Codigo_Extrai(Pais.Text)

    Analise.Caption = ""
    Status.Caption = ""
    AutorizadoPor.Caption = ""
    TotalAutoRS.Caption = ""
    TotalAutoUS.Caption = ""
    DataEnvioFinanc.Caption = ""
    VlrCond.Caption = ""
    DataPagtoReemb.Caption = ""
    FatJuridico.Caption = ""
    FatCobertura.Caption = ""
    FatReembolso.Caption = ""
    GATotal.Caption = ""
    FornFavorecido.Caption = ""
    
    DataAbertura.Caption = ""
    
    CGAnalise.ListIndex = -1
    CGStatus.ListIndex = -1
    CGAutorizadoPor.ListIndex = -1
    Estado.ListIndex = -1
    OrigemHist.ListIndex = -1
    
    SrvTotalAssistRS.Caption = ""
    SrvTotalSegRS.Caption = ""
    SrvTotalSegSegRS.Caption = ""
    SrvTotalAssistUS.Caption = ""
    SrvTotalSegUS.Caption = ""
    
    Judicial.Value = vbUnchecked
    JudicialE.Value = vbUnchecked
    AntecPagto.Value = vbUnchecked
    Condenado.Value = vbUnchecked
    
    IITotalRS.Caption = ""
    IITotalUS.Caption = ""
    OFTotalUS.Caption = ""
    OFTotalRS.Caption = ""
    NumeroFatProc.Caption = ""
    DataEmiProc.Caption = ""
    
    Call Trata_Juridico
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    Call DateParaMasked(DataHist, gdtDataAtual)

    iAlterado = 0
    
    iValorSolAlterado = 0
    iValorAutoAlterado = 0
    iTipoProc = TIPO_PROC_PERGUNTAR

    Limpa_Tela_TRVOcrCasos = SUCESSO

    Exit Function

Erro_Limpa_Tela_TRVOcrCasos:

    Limpa_Tela_TRVOcrCasos = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208569)

    End Select

    Exit Function

End Function

Function Traz_TRVOcrCasos_Tela(ByVal objOcrCasos As ClassTRVOcrCasos, Optional ByVal bTrazSemLer As Boolean = False) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_TRVOcrCasos_Tela

    gbTrazendoDados = True

    Call Limpa_Tela_TRVOcrCasos
    
    Codigo.PromptInclude = False
    Codigo.Text = objOcrCasos.sCodigo
    Codigo.PromptInclude = True

    If Not bTrazSemLer Then
        'Lê o TRVOcrCasos que está sendo Passado
        lErro = CF("TRVOcrCasos_Le", objOcrCasos)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    End If

    If lErro = SUCESSO Then
    
        glNumIntDoc = objOcrCasos.lNumIntDoc
        glNumIntDocTitPagCober = objOcrCasos.lNumIntDocTitPagCobertura
        glNumIntDocTitPagProc = objOcrCasos.lNumIntDocTitPagProcesso

        TipVou.Text = Trim(objOcrCasos.sTipVou)
        Serie.Text = Trim(objOcrCasos.sSerie)

        NumVou.PromptInclude = False
        If objOcrCasos.lNumVou <> 0 Then
            NumVou.Text = CStr(objOcrCasos.lNumVou)
        Else
            NumVou.Text = objOcrCasos.sNumVouTexto
        End If
        NumVou.PromptInclude = True

        VouPaxNome.Text = objOcrCasos.spaxnome
        VouTitular.Text = objOcrCasos.sTitularNome

        If objOcrCasos.lClienteVou <> 0 Then
            ClienteVou.Text = CStr(objOcrCasos.lClienteVou)
            Call ClienteVou_Validate(bSGECancelDummy)
        End If

        If objOcrCasos.dtDataEmissao <> DATA_NULA Then
            DataEmissao.PromptInclude = False
            DataEmissao.Text = Format(objOcrCasos.dtDataEmissao, "dd/mm/yy")
            DataEmissao.PromptInclude = True
        End If

        If objOcrCasos.dtDataIda <> DATA_NULA Then
            DataIda.PromptInclude = False
            DataIda.Text = Format(objOcrCasos.dtDataIda, "dd/mm/yy")
            DataIda.PromptInclude = True
        End If

        If objOcrCasos.dtDataVolta <> DATA_NULA Then
            DataVolta.PromptInclude = False
            DataVolta.Text = Format(objOcrCasos.dtDataVolta, "dd/mm/yy")
            DataVolta.PromptInclude = True
        End If

        Produto.Text = Trim(objOcrCasos.sProduto)
        gsProdAnt = Produto.Text

        If objOcrCasos.iQtdPax <> 0 Then
            QtdePax.PromptInclude = False
            QtdePax.Text = CStr(objOcrCasos.iQtdPax)
            QtdePax.PromptInclude = True
        End If

        If objOcrCasos.dtDataAbertura <> DATA_NULA Then
            DataAbertura.Caption = Format(objOcrCasos.dtDataAbertura, "dd/mm/yyyy")
        Else
            DataAbertura.Caption = ""
        End If

        If objOcrCasos.dtDataDocsRec <> DATA_NULA Then
            DataDocsRec.PromptInclude = False
            DataDocsRec.Text = Format(objOcrCasos.dtDataDocsRec, "dd/mm/yy")
            DataDocsRec.PromptInclude = True
        End If

        If objOcrCasos.dtDataEnvioAnalise <> DATA_NULA Then
            DataEnvioAnalise.PromptInclude = False
            DataEnvioAnalise.Text = Format(objOcrCasos.dtDataEnvioAnalise, "dd/mm/yy")
            DataEnvioAnalise.PromptInclude = True
        End If
        
        Call Combo_Seleciona_ItemData(CGAnalise, objOcrCasos.lCGAnalise)
        Call Combo_Seleciona_ItemData(CGStatus, objOcrCasos.lCGStatus)
        Call Combo_Seleciona_ItemData(CGAutorizadoPor, objOcrCasos.lCGAutorizadoPor)
        
        Analise.Caption = CGAnalise.Text
        Status.Caption = CGStatus.Text
        AutorizadoPor.Caption = CGAutorizadoPor.Text

        TotalAutoRS.Caption = Format(objOcrCasos.dValorAutorizadoTotalRS, "STANDARD")
        TotalAutoUS.Caption = Format(objOcrCasos.dValorAutorizadoTotalUS, "STANDARD")

        If objOcrCasos.dtDataLimite <> DATA_NULA Then
            DataLimite.PromptInclude = False
            DataLimite.Text = Format(objOcrCasos.dtDataLimite, "dd/mm/yy")
            DataLimite.PromptInclude = True
        End If

        If objOcrCasos.dtDataEnvioFinac <> DATA_NULA Then DataEnvioFinanc.Caption = Format(objOcrCasos.dtDataEnvioFinac, "dd/mm/yyyy")

        If objOcrCasos.dtDataProgFinanc <> DATA_NULA Then
            DataProgrFin.PromptInclude = False
            DataProgrFin.Text = Format(objOcrCasos.dtDataProgFinanc, "dd/mm/yy")
            DataProgrFin.PromptInclude = True
        End If
        
        If objOcrCasos.dtDataPagtoPax <> DATA_NULA Then
            DataPagtoPax.PromptInclude = False
            DataPagtoPax.Text = Format(objOcrCasos.dtDataPagtoPax, "dd/mm/yy")
            DataPagtoPax.PromptInclude = True
        End If

        If objOcrCasos.iJudicial <> 0 Then
            Judicial.Value = vbChecked
            JudicialE.Value = vbChecked
        Else
            Judicial.Value = vbUnchecked
            JudicialE.Value = vbUnchecked
        End If

        NumProcesso.PromptInclude = False
        NumProcesso.Text = objOcrCasos.sNumProcesso
        NumProcesso.PromptInclude = False
        
        If objOcrCasos.iCondenado = MARCADO Then
            Condenado.Value = vbChecked
            If objOcrCasos.iPerdaTipo = TRV_OCRCASOS_PERDA_ACORDO Then
                OptAcordo.Value = True
            Else
                OptCondenacao.Value = True
            End If
        Else
            Condenado.Value = vbUnchecked
        End If
        
        If objOcrCasos.iProcon = MARCADO Then
            Procon.Value = vbChecked
        Else
            Procon.Value = vbUnchecked
        End If

        If objOcrCasos.dValorCondenacao <> 0 Then
            ValorCondenacao.PromptInclude = False
            ValorCondenacao.Text = Format(objOcrCasos.dValorCondenacao, "STANDARD")
            ValorCondenacao.PromptInclude = True
        End If
        
        VlrCond.Caption = ValorCondenacao.Text
        
        If objOcrCasos.dProcessoDanoMaterial <> 0 Then
            DanoMaterial.PromptInclude = False
            DanoMaterial.Text = Format(objOcrCasos.dProcessoDanoMaterial, "STANDARD")
            DanoMaterial.PromptInclude = True
        End If
        If objOcrCasos.dProcessoDanoMoral <> 0 Then
            DanoMoral.PromptInclude = False
            DanoMoral.Text = Format(objOcrCasos.dProcessoDanoMoral, "STANDARD")
            DanoMoral.PromptInclude = True
        End If

        Comarca.Text = objOcrCasos.sComarca

        If objOcrCasos.dtDataFimProcesso <> DATA_NULA Then
            DataFimProcesso.PromptInclude = False
            DataFimProcesso.Text = Format(objOcrCasos.dtDataFimProcesso, "dd/mm/yy")
            DataFimProcesso.PromptInclude = True
        End If
        
        If objOcrCasos.dtDataIniProcesso <> DATA_NULA Then
            DataIniProcesso.PromptInclude = False
            DataIniProcesso.Text = Format(objOcrCasos.dtDataIniProcesso, "dd/mm/yy")
            DataIniProcesso.PromptInclude = True
        End If

        If objOcrCasos.objPreReceber.dtDataPagto <> DATA_NULA Then DataPagtoReemb.Caption = Format(objOcrCasos.objPreReceber.dtDataPagto, "dd/mm/yyyy")

        SrvTotalSegRS.Caption = Format(objOcrCasos.dValorAutorizadoSeguroRS, "STANDARD")
        SrvTotalSegUS.Caption = Format(objOcrCasos.dValorAutorizadoSeguroUS, "STANDARD")
        SrvTotalAssistRS.Caption = Format(objOcrCasos.dValorAutorizadoAssistRS, "STANDARD")
        SrvTotalAssistUS.Caption = Format(objOcrCasos.dValorAutorizadoAssistUS, "STANDARD")

        If objOcrCasos.dValorAutoSegRespTrvRS <> 0 Then SrvTotalSegTrvRS.Text = Format(objOcrCasos.dValorAutoSegRespTrvRS, "STANDARD")
        SrvTotalSegSegRS.Caption = Format(objOcrCasos.dValorAutorizadoSeguroRS - objOcrCasos.dValorAutoSegRespTrvRS, "STANDARD")

        If objOcrCasos.dCambio <> 0 Then
            Cambio.PromptInclude = False
            Cambio.Text = Format(objOcrCasos.dCambio, Cambio.Format)
            Cambio.PromptInclude = True
        End If
        
        If objOcrCasos.iAnteciparPagtoSeguro <> 0 Then
            AntecPagto.Value = vbChecked
        Else
            AntecPagto.Value = vbUnchecked
        End If

        If objOcrCasos.iBanco <> 0 Then
            Banco.PromptInclude = False
            Banco.Text = CStr(objOcrCasos.iBanco)
            Banco.PromptInclude = True
        End If

        Agencia.Text = objOcrCasos.sAgencia
        ContaCorrente.Text = objOcrCasos.sContaCorrente
        NomeFavorecido.Text = objOcrCasos.sNomeFavorecido

        If objOcrCasos.lCodFornFavorecido <> 0 Then FornFavorecido.Caption = CStr(objOcrCasos.lCodFornFavorecido)

        FavorecidoCGC.Text = objOcrCasos.sFavorecidoCGC
        Call FavorecidoCGC_Validate(bSGECancelDummy)

        IITotalRS.Caption = Format(objOcrCasos.dValorInvoicesTotal, "STANDARD")
        IITotalUS.Caption = Format(objOcrCasos.dValorInvoicesTotalUS, "STANDARD")
        OFTotalRS.Caption = Format(objOcrCasos.dValorDespesasTotalRS, "STANDARD")
        OFTotalUS.Caption = Format(objOcrCasos.dValorDespesasTotalUS, "STANDARD")

        If objOcrCasos.lNumFatProcesso <> 0 Then NumeroFatProc.Caption = CStr(objOcrCasos.lNumFatProcesso)
        If objOcrCasos.lNumFatCobertura <> 0 Then FatCobertura.Caption = CStr(objOcrCasos.lNumFatCobertura)
        If objOcrCasos.lNumFatProcesso <> 0 Then FatJuridico.Caption = CStr(objOcrCasos.lNumFatProcesso)
        If objOcrCasos.objPreReceber.lNumFatTitRecReembolso <> 0 Then FatReembolso.Caption = CStr(objOcrCasos.objPreReceber.lNumFatTitRecReembolso)
        
        glNumIntDocTitRecReembolso = objOcrCasos.objPreReceber.lNumIntDocTitRecReembolso
        
        Obs.Text = objOcrCasos.sOBS
        
        lErro = Traz_Endereco_Tela(objOcrCasos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = CF("TRVOcrCasosSrv_Carrega", objOcrCasos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = Traz_Srv_Tela(objOcrCasos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = Traz_II_Tela(objOcrCasos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = Traz_OF_Tela(objOcrCasos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        lErro = Traz_PC_Tela(objOcrCasos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = Traz_EV_Tela(objOcrCasos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = Traz_Hist_Tela(objOcrCasos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = Traz_Anot_Tela(objOcrCasos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = CF("TRVOcrCasosDocs_Carrega", objOcrCasos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = Traz_Docs_Tela(objOcrCasos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = Traz_GAdv_Tela(objOcrCasos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    Call Trata_Juridico
    
    Call Trata_Condenacao

    gbTrazendoDados = False
    
    iAlterado = 0

    Traz_TRVOcrCasos_Tela = SUCESSO

    Exit Function

Erro_Traz_TRVOcrCasos_Tela:

    gbTrazendoDados = False

    Traz_TRVOcrCasos_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208570)

    End Select

    Exit Function

End Function

Function Traz_Endereco_Tela(objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim objEndereco As ClassEndereco

On Error GoTo Erro_Traz_Endereco_Tela

    Set objEndereco = objOcrCasos.objEndereco

    Pais.Text = objEndereco.iCodigoPais
    Call Pais_Validate(bSGECancelDummy)
    
    Bairro.Text = objEndereco.sBairro
    Cidade.Text = objEndereco.sCidade
    
    gsCEPAnt = objEndereco.sCEP
    
    CEP.Text = objEndereco.sCEP
    
    If objEndereco.sSiglaEstado = "" Then objEndereco.sSiglaEstado = "SP"
    Estado.Text = objEndereco.sSiglaEstado
    
    Call Estado_Validate(bSGECancelDummy)
    If objEndereco.iCodigoPais = 0 Then objEndereco.iCodigoPais = PAIS_BRASIL
    
    Email1.Text = objEndereco.sEmail
    Contato.Text = objEndereco.sContato
    
    Logradouro.Text = objEndereco.sLogradouro

    If objEndereco.lNumero <> 0 Then
        Numero.Text = CStr(objEndereco.lNumero)
    Else
        Numero.Text = ""
    End If
    
    Complemento.Text = objEndereco.sComplemento
    
    Telefone1.Text = objEndereco.sTelNumero1
    Telefone2.Text = objEndereco.sTelNumero2
    
    Traz_Endereco_Tela = SUCESSO

    Exit Function

Erro_Traz_Endereco_Tela:

    Traz_Endereco_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208594)

    End Select
    
    Resume Next

    Exit Function

End Function

Function Traz_Srv_Tela(objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objOcrCasosSrv As ClassTRVOcrCasosSrv

On Error GoTo Erro_Traz_Srv_Tela

    Call Grid_Limpa(objGridSrv)

    iLinha = 0
    For Each objOcrCasosSrv In objOcrCasos.colCoberturas
    
        iLinha = iLinha + 1
    
        If objOcrCasosSrv.dValorAutorizadoRS <> 0 Then GridSrv.TextMatrix(iLinha, iGrid_SrvVlrAutoRS_Col) = Format(objOcrCasosSrv.dValorAutorizadoRS, "STANDARD")
        If objOcrCasosSrv.dValorAutorizadoUS <> 0 Then GridSrv.TextMatrix(iLinha, iGrid_SrvVlrAutoUS_Col) = Format(objOcrCasosSrv.dValorAutorizadoUS, "STANDARD")
        If objOcrCasosSrv.dValorLimite <> 0 Then GridSrv.TextMatrix(iLinha, iGrid_SrvVlrLimite_Col) = Format(objOcrCasosSrv.dValorLimite, "STANDARD")
        If objOcrCasosSrv.dValorSolicitadoRS <> 0 Then GridSrv.TextMatrix(iLinha, iGrid_SrvVlrSolRS_Col) = Format(objOcrCasosSrv.dValorSolicitadoRS, "STANDARD")
        If objOcrCasosSrv.dValorSolicitadoUS <> 0 Then GridSrv.TextMatrix(iLinha, iGrid_SrvVlrSolUS_Col) = Format(objOcrCasosSrv.dValorSolicitadoUS, "STANDARD")
        
        Call Combo_Seleciona_ItemData(IIMoeda, objOcrCasosSrv.iMoeda)
        
        GridSrv.TextMatrix(iLinha, iGrid_SrvMoeda_Col) = IIMoeda.Text
        
        If objOcrCasosSrv.iTipo = TRV_SRV_TIPO_SEGURO Then
            GridSrv.TextMatrix(iLinha, iGrid_SrvTipo_Col) = CStr(TRV_SRV_TIPO_SEGURO) & SEPARADOR & TRV_SRV_TIPO_SEGURO_TEXTO
        Else
            GridSrv.TextMatrix(iLinha, iGrid_SrvTipo_Col) = CStr(TRV_SRV_TIPO_ASSIST) & SEPARADOR & TRV_SRV_TIPO_ASSIST_TEXTO
        End If
        
        GridSrv.TextMatrix(iLinha, iGrid_SrvDescricao_Col) = objOcrCasosSrv.lCodigoServ & SEPARADOR & objOcrCasosSrv.sDescricao

        GridSrv.TextMatrix(iLinha, iGrid_SrvAuto_Col) = CStr(objOcrCasosSrv.iAutorizado)
        GridSrv.TextMatrix(iLinha, iGrid_SrvSol_Col) = CStr(objOcrCasosSrv.iSolicitado)
    
    Next
    
    objGridSrv.iLinhasExistentes = objOcrCasos.colCoberturas.Count
    
    Call Grid_Refresh_Checkbox(objGridSrv)
    
    Call Totaliza_Valores_Srv
    
    Traz_Srv_Tela = SUCESSO

    Exit Function

Erro_Traz_Srv_Tela:

    Traz_Srv_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208594)

    End Select

    Exit Function

End Function

Private Sub Totaliza_Valores_Srv()

Dim iLinha As Integer
Dim dVlrSolRS As Double
Dim dVlrSolUS As Double
Dim dVlrAutoAssistRS As Double
Dim dVlrAutoAssistUS As Double
Dim dVlrAutoSeguroRS As Double
Dim dVlrAutoSeguroUS As Double

On Error GoTo Erro_Totaliza_Valores_Srv

    For iLinha = 1 To objGridSrv.iLinhasExistentes
    
        dVlrSolRS = dVlrSolRS + StrParaDbl(GridSrv.TextMatrix(iLinha, iGrid_SrvVlrSolRS_Col))
        dVlrSolUS = dVlrSolUS + StrParaDbl(GridSrv.TextMatrix(iLinha, iGrid_SrvVlrSolUS_Col))
    
        If Codigo_Extrai(GridSrv.TextMatrix(iLinha, iGrid_SrvTipo_Col)) = TRV_SRV_TIPO_ASSIST Then
            dVlrAutoAssistRS = dVlrAutoAssistRS + StrParaDbl(GridSrv.TextMatrix(iLinha, iGrid_SrvVlrAutoRS_Col))
            dVlrAutoAssistUS = dVlrAutoAssistUS + StrParaDbl(GridSrv.TextMatrix(iLinha, iGrid_SrvVlrAutoUS_Col))
        Else
            dVlrAutoSeguroRS = dVlrAutoSeguroRS + StrParaDbl(GridSrv.TextMatrix(iLinha, iGrid_SrvVlrAutoRS_Col))
            dVlrAutoSeguroUS = dVlrAutoSeguroUS + StrParaDbl(GridSrv.TextMatrix(iLinha, iGrid_SrvVlrAutoUS_Col))
        End If

    Next
    
    SrvTotalSegRS.Caption = Format(dVlrAutoSeguroRS, "STANDARD")
    SrvTotalSegUS.Caption = Format(dVlrAutoSeguroUS, "STANDARD")
    SrvTotalAssistRS.Caption = Format(dVlrAutoAssistRS, "STANDARD")
    SrvTotalAssistUS.Caption = Format(dVlrAutoAssistUS, "STANDARD")

    TotalAutoRS.Caption = Format(dVlrAutoSeguroRS + dVlrAutoAssistRS, "STANDARD")
    TotalAutoUS.Caption = Format(dVlrAutoSeguroUS + dVlrAutoAssistUS, "STANDARD")
    
    If dVlrAutoSeguroRS < StrParaDbl(SrvTotalSegTrvRS.Text) Then
        Call Rotina_Aviso(vbOKOnly, "ERRO_TRV_SEGURO_RESPTRV_MAIOR_TOTAL")
        SrvTotalSegTrvRS.Text = ""
    End If
    
    SrvTotalSegSegRS.Caption = Format(dVlrAutoSeguroRS - StrParaDbl(SrvTotalSegTrvRS.Text), "STANDARD")

    Exit Sub

Erro_Totaliza_Valores_Srv:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208608)

    End Select

    Exit Sub
    
End Sub

Function Traz_II_Tela(objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objOcrCasosII As ClassTRvOcrCasosInvoices

On Error GoTo Erro_Traz_II_Tela

    Call Grid_Limpa(objGridII)

    iLinha = 0
    For Each objOcrCasosII In objOcrCasos.colInvoices
    
        iLinha = iLinha + 1
    
        If objOcrCasosII.dtDataFatura <> DATA_NULA Then GridII.TextMatrix(iLinha, iGrid_IIDataEmi_Col) = Format(objOcrCasosII.dtDataFatura, "dd/mm/yyyy")
        If objOcrCasosII.dtDataRecepcao <> DATA_NULA Then GridII.TextMatrix(iLinha, iGrid_IIDataRec_Col) = Format(objOcrCasosII.dtDataRecepcao, "dd/mm/yyyy")
        If objOcrCasosII.dValorMoeda <> 0 Then GridII.TextMatrix(iLinha, iGrid_IIValorMoeda_Col) = Format(objOcrCasosII.dValorMoeda, "STANDARD")
        If objOcrCasosII.dValorRS <> 0 Then GridII.TextMatrix(iLinha, iGrid_IIValorRS_Col) = Format(objOcrCasosII.dValorRS, "STANDARD")
        
'        Call Combo_Seleciona_ItemData(IIMoeda, objOcrCasosII.iMoeda)
'
'        GridII.TextMatrix(iLinha, iGrid_IIMoeda_Col) = IIMoeda.Text
    
        GridII.TextMatrix(iLinha, iGrid_IINumero_Col) = objOcrCasosII.sNumero
        GridII.TextMatrix(iLinha, iGrid_IIObs_Col) = objOcrCasosII.sOBS
    
    Next
    
    objGridII.iLinhasExistentes = objOcrCasos.colInvoices.Count
    
    Call Soma_Coluna_Grid(objGridII, iGrid_IIValorRS_Col, IITotalRS, False)
    Call Soma_Coluna_Grid(objGridII, iGrid_IIValorMoeda_Col, IITotalUS, False)
    
    Traz_II_Tela = SUCESSO

    Exit Function

Erro_Traz_II_Tela:

    Traz_II_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208609)

    End Select

    Exit Function

End Function

Function Traz_OF_Tela(objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objOcrCasosOF As ClassTRVOcrCasosOutrasFat

On Error GoTo Erro_Traz_OF_Tela

    Call Grid_Limpa(objGridOF)

    iLinha = 0
    For Each objOcrCasosOF In objOcrCasos.colOutrasFaturas
    
        iLinha = iLinha + 1
    
        If objOcrCasosOF.dtDataFatura <> DATA_NULA Then GridOF.TextMatrix(iLinha, iGrid_OFDataEmi_Col) = Format(objOcrCasosOF.dtDataFatura, "dd/mm/yyyy")
        If objOcrCasosOF.dtDataRecepcao <> DATA_NULA Then GridOF.TextMatrix(iLinha, iGrid_OFDataRec_Col) = Format(objOcrCasosOF.dtDataRecepcao, "dd/mm/yyyy")
        If objOcrCasosOF.dValorUS <> 0 Then GridOF.TextMatrix(iLinha, iGrid_OFValorUS_Col) = Format(objOcrCasosOF.dValorUS, "STANDARD")
        If objOcrCasosOF.dValorRS <> 0 Then GridOF.TextMatrix(iLinha, iGrid_OFValorRS_Col) = Format(objOcrCasosOF.dValorRS, "STANDARD")
        GridOF.TextMatrix(iLinha, iGrid_OFNumero_Col) = objOcrCasosOF.sNumero
        GridOF.TextMatrix(iLinha, iGrid_OFDescricao_Col) = objOcrCasosOF.sDescricao
        GridOF.TextMatrix(iLinha, iGrid_OFConsiderar_Col) = CStr(objOcrCasosOF.iConsiderar)
    
    Next
    
    objGridOF.iLinhasExistentes = objOcrCasos.colOutrasFaturas.Count
    
    Call Grid_Refresh_Checkbox(objGridOF)
    
    Call Soma_Coluna_Grid(objGridOF, iGrid_OFValorRS_Col, OFTotalRS, False, iGrid_OFConsiderar_Col)
    Call Soma_Coluna_Grid(objGridOF, iGrid_OFValorUS_Col, OFTotalUS, False, iGrid_OFConsiderar_Col)
    
    Traz_OF_Tela = SUCESSO

    Exit Function

Erro_Traz_OF_Tela:

    Traz_OF_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208610)

    End Select

    Exit Function

End Function

Function Traz_PC_Tela(objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objOcrCasosPC As ClassTRVOcrCasosParcCond

On Error GoTo Erro_Traz_PC_Tela

    Call Grid_Limpa(objGridPC)

    iLinha = 0
    For Each objOcrCasosPC In objOcrCasos.colParcProcesso
    
        iLinha = iLinha + 1
    
        If objOcrCasosPC.dtDataPagto <> DATA_NULA Then GridPC.TextMatrix(iLinha, iGrid_PCPagamento_Col) = Format(objOcrCasosPC.dtDataPagto, "dd/mm/yyyy")
        If objOcrCasosPC.dtDataVencimento <> DATA_NULA Then GridPC.TextMatrix(iLinha, iGrid_PCVencimento_Col) = Format(objOcrCasosPC.dtDataVencimento, "dd/mm/yyyy")
        If objOcrCasosPC.dValor <> 0 Then GridPC.TextMatrix(iLinha, iGrid_PCValor_Col) = Format(objOcrCasosPC.dValor, "STANDARD")
    
    Next
    
    objGridPC.iLinhasExistentes = objOcrCasos.colParcProcesso.Count
        
    Traz_PC_Tela = SUCESSO

    Exit Function

Erro_Traz_PC_Tela:

    Traz_PC_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208611)

    End Select

    Exit Function

End Function

Function Traz_EV_Tela(objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objOcrCasosEv As ClassTRVOcrCasoImport

On Error GoTo Erro_Traz_EV_Tela

    Call Grid_Limpa(objGridEv)

    iLinha = 0
    For Each objOcrCasosEv In objOcrCasos.colEventos
    
        iLinha = iLinha + 1
    
        GridEv.TextMatrix(iLinha, iGrid_EvData_Col) = Format(objOcrCasosEv.dtData, "dd/mm/yyyy")
        GridEv.TextMatrix(iLinha, iGrid_EvCidade_Col) = objOcrCasosEv.sCidadeOCR
        GridEv.TextMatrix(iLinha, iGrid_EvEstado_Col) = objOcrCasosEv.sEstadoOCR
        GridEv.TextMatrix(iLinha, iGrid_EvPais_Col) = objOcrCasosEv.sPaisOCR
        GridEv.TextMatrix(iLinha, iGrid_EvTelefone_Col) = objOcrCasosEv.sTelefone
        GridEv.TextMatrix(iLinha, iGrid_EvTipo_Col) = objOcrCasosEv.sCarater
    
    Next
    
    objGridEv.iLinhasExistentes = objOcrCasos.colEventos.Count
        
    Traz_EV_Tela = SUCESSO

    Exit Function

Erro_Traz_EV_Tela:

    Traz_EV_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208612)

    End Select

    Exit Function

End Function

Function Traz_Hist_Tela(objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objOcrCasosHist As ClassTRVOcrCasosHist

On Error GoTo Erro_Traz_Hist_Tela

    Call Grid_Limpa(objGridHist)

    iLinha = 0
    For Each objOcrCasosHist In objOcrCasos.colHistorico
    
        iLinha = iLinha + 1
    
        GridHist.TextMatrix(iLinha, iGrid_HistData_Col) = Format(objOcrCasosHist.dtData, "dd/mm/yyyy")
        GridHist.TextMatrix(iLinha, iGrid_HistHora_Col) = Format(objOcrCasosHist.dHora, "HH:MM:SS")
        GridHist.TextMatrix(iLinha, iGrid_HistUsuario_Col) = objOcrCasosHist.sUsuario
        GridHist.TextMatrix(iLinha, iGrid_HistAssunto_Col) = objOcrCasosHist.sAssunto
        
        If objOcrCasosHist.iOrigem = RELACIONAMENTOCLIENTES_ORIGEM_CLIENTE Then
            GridHist.TextMatrix(iLinha, iGrid_HistOrigem_Col) = RELACIONAMENTOCLIENTES_ORIGEM_CLIENTE_TEXTO
        Else
            GridHist.TextMatrix(iLinha, iGrid_HistOrigem_Col) = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA_TEXTO
        End If
    
    Next
    
    objGridHist.iLinhasExistentes = objOcrCasos.colHistorico.Count
        
    Traz_Hist_Tela = SUCESSO

    Exit Function

Erro_Traz_Hist_Tela:

    Traz_Hist_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208613)

    End Select

    Exit Function

End Function

Function Traz_Anot_Tela(objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objOcrCasosAnot As ClassTRVOcrCasosAnotacoes

On Error GoTo Erro_Traz_Anot_Tela

    Call Grid_Limpa(objGridAnot)

    iLinha = 0
    For Each objOcrCasosAnot In objOcrCasos.colAnotacoes
    
        iLinha = iLinha + 1
    
        GridAnot.TextMatrix(iLinha, iGrid_AnotData_Col) = Format(objOcrCasosAnot.dtData, "dd/mm/yyyy")
        GridAnot.TextMatrix(iLinha, iGrid_AnotHora_Col) = Format(objOcrCasosAnot.dHora, "HH:MM:SS")
        GridAnot.TextMatrix(iLinha, iGrid_AnotUsuario_Col) = objOcrCasosAnot.sUsuario
        GridAnot.TextMatrix(iLinha, iGrid_AnotTexto_Col) = objOcrCasosAnot.sTexto
    
    Next
    
    objGridAnot.iLinhasExistentes = objOcrCasos.colAnotacoes.Count
        
    Traz_Anot_Tela = SUCESSO

    Exit Function

Erro_Traz_Anot_Tela:

    Traz_Anot_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208614)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Limpa Tela
    Call Limpa_Tela_TRVOcrCasos

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208571)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208572)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Limpa_Tela_TRVOcrCasos

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208573)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Codigo_GotFocus()
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
End Sub

Private Sub TipVou_Change()
    'iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TipVou_GotFocus()
    Call MaskEdBox_TrataGotFocus(TipVou, iAlterado)
End Sub

Private Sub Serie_Change()
    'iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Serie_GotFocus()
    Call MaskEdBox_TrataGotFocus(Serie, iAlterado)
End Sub

Private Sub NumVou_GotFocus()
    Call MaskEdBox_TrataGotFocus(NumVou, iAlterado)
End Sub

Private Sub NumVou_Change()
'    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOcrCasos As ClassTRVOcrCasos

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objOcrCasos = obj1

    'Mostra os dados do TRVOcrCasos na tela
    lErro = Traz_TRVOcrCasos_Tela(objOcrCasos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208574)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objOcrCasos As New ClassTRVOcrCasos
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then
        objOcrCasos.sCodigo = Codigo.Text
    End If

    Call Chama_Tela("TRVOcrCasosLista", colSelecao, objOcrCasos, objEventoCodigo, "", "Codigo")

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208575)

    End Select

    Exit Sub

End Sub

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

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208576)

    End Select

    Exit Sub

End Sub

Private Sub BotaoTrazerVou_Click()

Dim lErro As Long
Dim objOcrCaso As New ClassTRVOcrCasos
Dim colRetorno As New Collection
Dim colSelecao As New Collection, sFiltro As String

On Error GoTo Erro_BotaoTrazerVou_Click

    objOcrCaso.sSerie = Trim(Serie.Text)
    objOcrCaso.sTipVou = Trim(TipVou.Text)
    
    If Len(Trim(NumVou.Text)) < 10 Then
        objOcrCaso.lNumVou = StrParaLong(NumVou.Text)
    Else
        objOcrCaso.sNumVouTexto = Trim(NumVou.Text)
    End If
    
    'If objOcrCaso.lNumVou = 0 Then gError 208577 'Preencha o número do voucher

    lErro = CF("TRVOcrCasos_Le_NumVou", objOcrCaso, colRetorno)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If colRetorno.Count = 0 Then gError 208578 'Não existem Ocrs cadastradas para esse vou

    If colRetorno.Count = 1 Then
    
        objOcrCaso.sCodigo = colRetorno.Item(1)
        
        lErro = Traz_TRVOcrCasos_Tela(objOcrCaso)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If

    If colRetorno.Count > 1 Then
    
        If objOcrCaso.lNumVou <> 0 Then
            sFiltro = "NumVou = ?"
            colSelecao.Add objOcrCaso.lNumVou
        Else
            sFiltro = "NumVouTexto = ?"
            colSelecao.Add objOcrCaso.sNumVouTexto
        End If
        
        If Len(Trim(objOcrCaso.sTipVou)) > 0 Then
            colSelecao.Add objOcrCaso.sTipVou
            sFiltro = sFiltro & " AND (TipVou = ? OR TipoVou = '')"
        End If
        
        If Len(Trim(objOcrCaso.sSerie)) > 0 Then
            colSelecao.Add objOcrCaso.sSerie
            sFiltro = sFiltro & " AND (Serie = ? OR Serie = '')"
        End If

        Call Chama_Tela("TRVOcrCasosLista", colSelecao, objOcrCaso, objEventoCodigo, sFiltro, "Codigo")

    End If
    
    Exit Sub

Erro_BotaoTrazerVou_Click:

    Select Case gErr
    
        Case 208577
            Call Rotina_Erro(vbOKOnly, "ERRO_TRV_NUMVOU_NAO_PREENCHIDO", gErr)
        
        Case 208578
            Call Rotina_Erro(vbOKOnly, "ERRO_OCRCASO_NAO_ENCONTRADA_VOU", gErr)
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208579)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelNumVou_Click()

Dim lErro As Long
Dim objOcrCasos As New ClassTRVOcrCasos
Dim colSelecao As New Collection

On Error GoTo Erro_LabelNumVou_Click

    If Len(Trim(NumVou.Text)) < 10 Then objOcrCasos.lNumVou = StrParaLong(NumVou.Text)

    Call Chama_Tela("TRVOcrCasosLista", colSelecao, objOcrCasos, objEventoCodigo, "", "NumVou")

    Exit Sub

Erro_LabelNumVou_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208580)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoTrazerCaso_Click()

Dim lErro As Long
Dim objOcrCaso As New ClassTRVOcrCasos

On Error GoTo Erro_BotaoTrazerVou_Click

    If Len(Trim(Codigo.Text)) = 0 Then gError 208581 'Preencha o número da ocorrência
   
    objOcrCaso.sCodigo = Trim(Codigo.Text)
    
    lErro = Traz_TRVOcrCasos_Tela(objOcrCaso)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Exit Sub

Erro_BotaoTrazerVou_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 208581
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TRVOCORRENCIAS_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208582)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoAbrirVou_Click()

Dim objVoucher As New ClassTRVVouchers

On Error GoTo Erro_BotaoVoucher_Click

    If Len(Trim(NumVou.Text)) = 0 Then gError 208583 'Preencha o número do voucher

    If Len(Trim(NumVou.Text)) < 10 Then objVoucher.lNumVou = StrParaLong(NumVou.Text)
    objVoucher.sSerie = Serie.Text
    objVoucher.sTipVou = TipVou.Text

    Call Chama_Tela("TRVVoucher", objVoucher)

    Exit Sub

Erro_BotaoVoucher_Click:

    Select Case gErr

        Case 208583
            Call Rotina_Erro(vbOKOnly, "ERRO_TRV_NUMVOU_NAO_PREENCHIDO", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208584)

    End Select

    Exit Sub
    
End Sub

Private Sub ClienteVou_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objcliente As New ClassCliente

On Error GoTo Erro_ClienteVou_Validate

    If Len(Trim(ClienteVou.Text)) > 0 Then
    
        ClienteVou.Text = LCodigo_Extrai(ClienteVou.Text)

        lErro = TP_Cliente_Le2(ClienteVou, objcliente)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    Exit Sub

Erro_ClienteVou_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208591)
    
    End Select

End Sub

Public Sub FavorecidoCGC_Validate(Cancel As Boolean)
Dim lErro As Long
    lErro = CGC_Valida(FavorecidoCGC)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Function CGC_Valida(ByVal objControle As Object) As Long

Dim lErro As Long

On Error GoTo Erro_CGC_Valida
    
    'Se CGC/CPF não foi preenchido
    If Len(Trim(objControle.Text)) <> 0 Then
    
        Select Case Len(Trim(objControle.Text))
    
            Case STRING_CPF 'CPF
                
                'Critica Cpf
                lErro = Cpf_Critica(objControle.Text)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                'Formata e coloca na Tela
                objControle.Format = "000\.000\.000-00; ; ; "
                objControle.Text = objControle.Text
    
            Case STRING_CGC 'CGC
                
                'Critica CGC
                lErro = Cgc_Critica(objControle.Text)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                'Formata e Coloca na Tela
                objControle.Format = "00\.000\.000\/0000-00; ; ; "
                objControle.Text = objControle.Text
    
            Case Else
                    
                gError 208592
    
        End Select
        
    End If
    
    CGC_Valida = SUCESSO

    Exit Function

Erro_CGC_Valida:

    CGC_Valida = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 208592
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208593)

    End Select

    Exit Function

End Function

Private Sub objEventoPais_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPais As New ClassPais

On Error GoTo Erro_objEventoPais_evSelecao

    Set objPais = obj1

    Pais.Text = CStr(objPais.iCodigo)
    Call Pais_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoPais_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208595)

    End Select

    Exit Sub

End Sub

Public Sub Pais_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Pais_Click()
    iAlterado = REGISTRO_ALTERADO
    Call Trata_Alteracao_UF
End Sub

Public Sub Trata_Alteracao_UF()

On Error GoTo Erro_Trata_Alteracao_UF

    If giPaisAnt <> Codigo_Extrai(Pais.Text) Then
        giPaisAnt = Codigo_Extrai(Pais.Text)
        If Codigo_Extrai(Pais.Text) = PAIS_BRASIL Then
            Estado.Enabled = True
            If Estado.Text = "EX" Then
                Estado.ListIndex = -1
                Estado.Text = ""
            End If
        Else
            Estado.Enabled = False
            Estado.Text = "EX"
        End If
    End If

    Exit Sub

Erro_Trata_Alteracao_UF:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208596)

    End Select

    Exit Sub

End Sub

Public Sub Pais_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objPais As New ClassPais

On Error GoTo Erro_Pais_Validate

    'Verifica se foi preenchida a Combo Pais
    If Len(Trim(Pais.Text)) <> 0 Then

        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
        lErro = Combo_Seleciona(Pais, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM
    
        'Nao existe o item com o CODIGO na List da ComboBox
        If lErro = 6730 Then
    
            objPais.iCodigo = iCodigo
    
            'Tenta ler Pais com esse codigo no BD
            lErro = CF("Paises_Le", objPais)
            If lErro <> SUCESSO And lErro <> 47876 Then gError ERRO_SEM_MENSAGEM
            If lErro <> SUCESSO Then gError 208597
    
            Pais.Text = CStr(iCodigo) & SEPARADOR & objPais.sNome
    
        End If
    
        'Nao existe o item com a STRING na List da ComboBox
        If lErro = 6731 Then gError 208598
        
    End If
    
    Call Trata_Alteracao_UF

    Exit Sub

Erro_Pais_Validate:

    Cancel = True

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 208597
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PAIS", objPais.iCodigo)
            If vbMsgRes = vbYes Then Call Chama_Tela("Paises", objPais)

        Case 208598
            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO1", gErr, Trim(Pais.Text))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208599)

    End Select

    Exit Sub

End Sub

Public Sub PaisLabel_Click()

Dim objPais As New ClassPais
Dim colSelecao As Collection

    objPais.iCodigo = Codigo_Extrai(Pais.Text)

    'Chama a Tela de PaisesLista
    Call Chama_Tela("PaisesLista", colSelecao, objPais, objEventoPais)

End Sub

Public Sub LabelCidade_Click()

Dim objCidade As New ClassCidades
Dim colSelecao As Collection

    objCidade.sDescricao = Cidade.Text

    'Chama a Tela de browse
    Call Chama_Tela("CidadeLista", colSelecao, objCidade, objEventoCidade)

End Sub

Private Sub objEventoCidade_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCidade As ClassCidades

On Error GoTo Erro_objEventoCidade_evSelecao

    Set objCidade = obj1

    Cidade.Text = objCidade.sDescricao

    Me.Show

    Exit Sub

Erro_objEventoCidade_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208600)

    End Select

    Exit Sub

End Sub

Public Sub Cidade_Validate(Cancel As Boolean)

Dim lErro As Long, objCidade As New ClassCidades
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Cidade_Validate

    If Len(Trim(Cidade.Text)) = 0 Then Exit Sub

    objCidade.sDescricao = Cidade.Text
    
    lErro = CF("Cidade_Le_Nome", objCidade)
    If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError ERRO_SEM_MENSAGEM

    If lErro <> SUCESSO Then gError 208601

    Exit Sub

Erro_Cidade_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 208601
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CIDADE")
            If vbMsgRes = vbYes Then
                Call Chama_Tela("CidadeCadastro", objCidade)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208602)

    End Select

    Exit Sub

End Sub

Public Sub CEP_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objEndereco As New ClassEndereco
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_CEP_Validate

    If Len(Trim(CEP.Text)) = 0 Then Exit Sub

    objEndereco.sCEP = CEP.Text
    
    lErro = CF("Endereco_Le_CEP", objEndereco)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    
    If lErro = SUCESSO Then
        
        vbMsgRes = vbYes
        If Len(Trim(Logradouro.Text)) <> 0 And gsCEPAnt <> CEP.Text Then
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUIR_ENDERECO_ATUAL")
        End If
        
        If vbMsgRes = vbYes And gsCEPAnt <> CEP.Text Then
        
            Bairro.Text = UCase(objEndereco.sBairro)
            Cidade.Text = UCase(objEndereco.sCidade)
            Estado.Text = UCase(objEndereco.sSiglaEstado)
            Call Estado_Validate(bSGECancelDummy)
            Pais.Text = PAIS_BRASIL
            Call Pais_Validate(bSGECancelDummy)
            Logradouro.Text = UCase(objEndereco.sTipoLogradouro & " " & objEndereco.sLogradouro)
            
            gbMudouCEP = True
        
        End If
    
    End If
    
    gsCEPAnt = CEP.Text

    Exit Sub

Erro_CEP_Validate:

    Cancel = True

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208603)

    End Select

    Exit Sub

End Sub

Public Sub CEP_LostFocus()
    If gbMudouCEP Then Call Cidade.SetFocus
End Sub

Public Sub Estado_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CEP_GotFocus()
    Call MaskEdBox_TrataGotFocus(CEP, iAlterado)
End Sub

Public Sub Estado_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Estado_Validate

    'Verifica se foi preenchido o Estado
    If Len(Trim(Estado.Text)) <> 0 Then

        'Verifica se está preenchida com o ítem selecionado na ComboBox Estado
        If Estado.Text = Estado.List(Estado.ListIndex) Then Exit Sub
    
        'Verifica se existe o ítem na Combo Estado, se existir seleciona o item
        lErro = Combo_Item_Igual_CI(Estado)
        If lErro <> SUCESSO And lErro <> 58583 Then gError ERRO_SEM_MENSAGEM
    
        'Não existe o ítem na ComboBox Estado
        If lErro = 58583 And Codigo_Extrai(Pais) = PAIS_BRASIL Then gError 208604
    
    End If
    
    Call Trata_Alteracao_UF
    
    Exit Sub

Erro_Estado_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 208604
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", gErr, Estado.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208605)

    End Select
    
    Estado.ListIndex = -1
    Estado.Text = ""

    Exit Sub

End Sub

Function Carrega_Moeda()

Dim lErro As Long
Dim objMoeda As ClassMoedas
Dim colMoedas As New Collection

On Error GoTo Erro_Carrega_Moeda
    
    lErro = CF("Moedas_Le_Todas", colMoedas)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'se não existem moedas cadastradas
    If colMoedas.Count = 0 Then gError 208606
    
    IIMoeda.Clear
    For Each objMoeda In colMoedas
    
        IIMoeda.AddItem objMoeda.iCodigo & SEPARADOR & objMoeda.sNome
        IIMoeda.ItemData(IIMoeda.NewIndex) = objMoeda.iCodigo
    
    Next

    Carrega_Moeda = SUCESSO
    
    Exit Function
    
Erro_Carrega_Moeda:

    Carrega_Moeda = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 208606
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDAS_NAO_CADASTRADAS", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208607)
    
    End Select

End Function

Private Sub GridEv_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridEv, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridEv, iAlterado)
    End If

End Sub

Private Sub GridEv_GotFocus()
    Call Grid_Recebe_Foco(objGridEv)
End Sub

Private Sub GridEv_EnterCell()
    Call Grid_Entrada_Celula(objGridEv, iAlterado)
End Sub

Private Sub GridEv_LeaveCell()
    Call Saida_Celula(objGridEv)
End Sub

Private Sub GridEv_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridEv, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridEv, iAlterado)
    End If

End Sub

Private Sub GridEv_RowColChange()
    Call Grid_RowColChange(objGridEv)
End Sub

Private Sub GridEv_Scroll()
    Call Grid_Scroll(objGridEv)
End Sub

Private Sub GridEv_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridEv)

End Sub

Private Sub GridEv_LostFocus()
    Call Grid_Libera_Foco(objGridEv)
End Sub

Private Sub GridSrv_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridSrv, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridSrv, iAlterado)
    End If

End Sub

Private Sub GridSrv_GotFocus()
    Call Grid_Recebe_Foco(objGridSrv)
End Sub

Private Sub GridSrv_EnterCell()
    Call Grid_Entrada_Celula(objGridSrv, iAlterado)
End Sub

Private Sub GridSrv_LeaveCell()
    Call Saida_Celula(objGridSrv)
End Sub

Private Sub GridSrv_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridSrv, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridSrv, iAlterado)
    End If

End Sub

Private Sub GridSrv_RowColChange()
    Call Grid_RowColChange(objGridSrv)
    Call Exibe_CampoDet_Grid(objGridSrv, iGrid_SrvDescricao_Col, SrvDescricaoDet)
End Sub

Private Sub GridSrv_Scroll()
    Call Grid_Scroll(objGridSrv)
End Sub

Private Sub GridSrv_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridSrv)

End Sub

Private Sub GridSrv_LostFocus()
    Call Grid_Libera_Foco(objGridSrv)
End Sub

Private Sub GridII_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridII, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridII, iAlterado)
    End If

End Sub

Private Sub GridII_GotFocus()
    Call Grid_Recebe_Foco(objGridII)
End Sub

Private Sub GridII_EnterCell()
    Call Grid_Entrada_Celula(objGridII, iAlterado)
End Sub

Private Sub GridII_LeaveCell()
    Call Saida_Celula(objGridII)
End Sub

Private Sub GridII_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridII, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridII, iAlterado)
    End If

End Sub

Private Sub GridII_RowColChange()
    Call Grid_RowColChange(objGridII)
End Sub

Private Sub GridII_Scroll()
    Call Grid_Scroll(objGridII)
End Sub

Private Sub GridII_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnt As Integer

    'Guarda o número de linhas existentes
    iLinhasExistentesAnt = objGridII.iLinhasExistentes

    Call Grid_Trata_Tecla1(KeyCode, objGridII)
    
    If objGridII.iLinhasExistentes < iLinhasExistentesAnt Then
        Call Soma_Coluna_Grid(objGridII, iGrid_IIValorRS_Col, IITotalRS, False)
        Call Soma_Coluna_Grid(objGridII, iGrid_IIValorMoeda_Col, IITotalUS, False)
    End If

End Sub

Private Sub GridII_LostFocus()
    Call Grid_Libera_Foco(objGridII)
End Sub

Private Sub GridOF_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridOF, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridOF, iAlterado)
    End If

End Sub

Private Sub GridOF_GotFocus()
    Call Grid_Recebe_Foco(objGridOF)
End Sub

Private Sub GridOF_EnterCell()
    Call Grid_Entrada_Celula(objGridOF, iAlterado)
End Sub

Private Sub GridOF_LeaveCell()
    Call Saida_Celula(objGridOF)
End Sub

Private Sub GridOF_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridOF, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridOF, iAlterado)
    End If

End Sub

Private Sub GridOF_RowColChange()
    Call Grid_RowColChange(objGridOF)
End Sub

Private Sub GridOF_Scroll()
    Call Grid_Scroll(objGridOF)
End Sub

Private Sub GridOF_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnt As Integer

    'Guarda o número de linhas existentes
    iLinhasExistentesAnt = objGridOF.iLinhasExistentes

    Call Grid_Trata_Tecla1(KeyCode, objGridOF)
    
    If objGridOF.iLinhasExistentes < iLinhasExistentesAnt Then
        Call Trata_Reembolso
'        Call Soma_Coluna_Grid(objGridOF, iGrid_OFValorRS_Col, OFTotalRS, False, iGrid_OFConsiderar_Col)
'        Call Soma_Coluna_Grid(objGridOF, iGrid_OFValorUS_Col, OFTotalUS, False, iGrid_OFConsiderar_Col)
    End If

End Sub

Private Sub GridOF_LostFocus()
    Call Grid_Libera_Foco(objGridOF)
End Sub

Private Sub GridPC_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridPC, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPC, iAlterado)
    End If

End Sub

Private Sub GridPC_GotFocus()
    Call Grid_Recebe_Foco(objGridPC)
End Sub

Private Sub GridPC_EnterCell()
    Call Grid_Entrada_Celula(objGridPC, iAlterado)
End Sub

Private Sub GridPC_LeaveCell()
    Call Saida_Celula(objGridPC)
End Sub

Private Sub GridPC_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridPC, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPC, iAlterado)
    End If

End Sub

Private Sub GridPC_RowColChange()
    Call Grid_RowColChange(objGridPC)
End Sub

Private Sub GridPC_Scroll()
    Call Grid_Scroll(objGridPC)
End Sub

Private Sub GridPC_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridPC)

End Sub

Private Sub GridPC_LostFocus()
    Call Grid_Libera_Foco(objGridPC)
End Sub

Private Sub GridHist_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridHist, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridHist, iAlterado)
    End If

End Sub

Private Sub GridHist_GotFocus()
    Call Grid_Recebe_Foco(objGridHist)
End Sub

Private Sub GridHist_EnterCell()
    Call Grid_Entrada_Celula(objGridHist, iAlterado)
End Sub

Private Sub GridHist_LeaveCell()
    Call Saida_Celula(objGridHist)
End Sub

Private Sub GridHist_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridHist, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridHist, iAlterado)
    End If

End Sub

Private Sub GridHist_RowColChange()
    Call Grid_RowColChange(objGridHist)
    Call Exibe_CampoDet_Grid(objGridHist, iGrid_HistAssunto_Col, HistAssuntoDet)
End Sub

Private Sub GridHist_Scroll()
    Call Grid_Scroll(objGridHist)
End Sub

Private Sub GridHist_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridHist)

End Sub

Private Sub GridHist_LostFocus()
    Call Grid_Libera_Foco(objGridHist)
End Sub

Private Sub GridAnot_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridAnot, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAnot, iAlterado)
    End If

End Sub

Private Sub GridAnot_GotFocus()
    Call Grid_Recebe_Foco(objGridAnot)
End Sub

Private Sub GridAnot_EnterCell()
    Call Grid_Entrada_Celula(objGridAnot, iAlterado)
End Sub

Private Sub GridAnot_LeaveCell()
    Call Saida_Celula(objGridAnot)
End Sub

Private Sub GridAnot_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridAnot, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAnot, iAlterado)
    End If

End Sub

Private Sub GridAnot_RowColChange()
    Call Grid_RowColChange(objGridAnot)
    Call Exibe_CampoDet_Grid(objGridAnot, iGrid_AnotTexto_Col, AnotTextoDet)
End Sub

Private Sub GridAnot_Scroll()
    Call Grid_Scroll(objGridAnot)
End Sub

Private Sub GridAnot_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridAnot)

End Sub

Private Sub GridAnot_LostFocus()
    Call Grid_Libera_Foco(objGridAnot)
End Sub

Private Sub VouPaxNome_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub VouTitular_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ClienteVou_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Produto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iPreenchido As Integer, sProduto As String
Dim vbResult As VbMsgBoxResult
Dim objOcrCaso As New ClassTRVOcrCasos

On Error GoTo Erro_Produto_Validate

    If Len(Trim(Produto.ClipText)) > 0 Then

        'Critica o Produto
        lErro = CF("Produto_Critica_Filial2", Produto.Text, objProduto, iPreenchido)
        If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError ERRO_SEM_MENSAGEM
                       
        'Se o produto não foi encontrado ==> Pergunta se deseja criar
        If lErro = 51381 Then gError 208615
        
    End If
    
    If gsProdAnt <> Produto.Text Then
        If gsProdAnt <> "" Then
            vbResult = Rotina_Aviso(vbYesNo, "AVISO_TRVOCRCASO_TROCA_PROD") 'Essa troca vai recalcular os dados da cobertura, confirma ?
            If vbResult = vbNo Then gError ERRO_SEM_MENSAGEM
        End If
        lErro = CF("Produto_Formata", Produto.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objOcrCaso.sProduto = sProduto
        
        lErro = CF("TRVOcrCasosSrv_Carrega", objOcrCaso)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = Traz_Srv_Tela(objOcrCaso)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        gsProdAnt = Produto.Text
        
    End If
    
    Exit Sub

Erro_Produto_Validate:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 208615 'Produto não cadastrado
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, Produto.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208616)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissao_Change()
    Call Data_Change
End Sub

Private Sub DataEmissao_GotFocus()
    Call Data_GotFocus
End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)
    Call Data_Validate(Cancel)
End Sub

'Private Sub DataAbertura_Change()
'    Call Data_Change
'End Sub
'
'Private Sub DataAbertura_GotFocus()
'    Call Data_GotFocus
'End Sub
'
'Private Sub DataAbertura_Validate(Cancel As Boolean)
'    Call Data_Validate(Cancel)
'End Sub

Private Sub DataDocsRec_Change()
    Call Data_Change
End Sub

Private Sub DataDocsRec_GotFocus()
    Call Data_GotFocus
End Sub

Private Sub DataDocsRec_Validate(Cancel As Boolean)
    Call Data_Validate(Cancel)
End Sub

Private Sub DataEnvioAnalise_Change()
    Call Data_Change
End Sub

Private Sub DataEnvioAnalise_GotFocus()
    Call Data_GotFocus
End Sub

Private Sub DataEnvioAnalise_Validate(Cancel As Boolean)
    Call Data_Validate(Cancel)
End Sub

Private Sub DataPagtoPax_Change()
    Call Data_Change
End Sub

Private Sub DataPagtoPax_GotFocus()
    Call Data_GotFocus
End Sub

Private Sub DataPagtoPax_Validate(Cancel As Boolean)
    Call Data_Validate(Cancel)
End Sub

Private Sub DataLimite_Change()
    Call Data_Change
End Sub

Private Sub DataLimite_GotFocus()
    Call Data_GotFocus
End Sub

Private Sub DataLimite_Validate(Cancel As Boolean)
    Call Data_Validate(Cancel)
End Sub

Private Sub DataProgrFin_Change()
    Call Data_Change
End Sub

Private Sub DataProgrFin_GotFocus()
    Call Data_GotFocus
End Sub

Private Sub DataProgrFin_Validate(Cancel As Boolean)
    Call Data_Validate(Cancel)
End Sub

Private Sub DataFimProcesso_Change()
    Call Data_Change
End Sub

Private Sub DataFimProcesso_GotFocus()
    Call Data_GotFocus
End Sub

Private Sub DataFimProcesso_Validate(Cancel As Boolean)
    Call Data_Validate(Cancel)
End Sub

Private Sub DataHist_GotFocus()
    Dim iAlteradoAux As Integer
    Call MaskEdBox_TrataGotFocus(DataHist, iAlteradoAux)
End Sub

Private Sub DataHist_Validate(Cancel As Boolean)
    Call Data_Validate(Cancel)
End Sub

Private Sub DataIda_Change()
    Call Data_Change
End Sub

Private Sub DataIda_GotFocus()
    Call Data_GotFocus
End Sub

Private Sub DataIda_Validate(Cancel As Boolean)
    Call Data_Validate(Cancel)
End Sub

Private Sub DataVolta_Change()
    Call Data_Change
End Sub

Private Sub DataVolta_GotFocus()
    Call Data_GotFocus
End Sub

Private Sub DataVolta_Validate(Cancel As Boolean)
    Call Data_Validate(Cancel)
End Sub

Private Sub Data_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Me.ActiveControl, iAlterado)
    
End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Trim(Me.ActiveControl.ClipText)) <> 0 Then

        lErro = Data_Critica(Me.ActiveControl.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208617)

    End Select

    Exit Sub

End Sub

Private Sub UpDown_DownClick(Index As Integer)

Dim lErro As Long
Dim sData As String
Dim objControl As Object

On Error GoTo Erro_UpDown_DownClick

    Call Obtem_ControleData(Index, objControl)

    objControl.SetFocus

    If Len(objControl.ClipText) > 0 Then

        sData = objControl.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        objControl.Text = sData

    End If

    Exit Sub

Erro_UpDown_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208618)

    End Select

    Exit Sub

End Sub

Private Sub UpDown_UpClick(Index As Integer)

Dim lErro As Long
Dim sData As String
Dim objControl As Object

On Error GoTo Erro_UpDown_UpClick

    Call Obtem_ControleData(Index, objControl)

    objControl.SetFocus

    If Len(Trim(objControl.ClipText)) > 0 Then

        sData = objControl.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        objControl.Text = sData

    End If

    Exit Sub

Erro_UpDown_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208619)

    End Select

    Exit Sub

End Sub

Sub Obtem_ControleData(ByVal Index As Integer, objControl As Object)

    Select Case Index
    
        Case 0
            Set objControl = DataAbertura
        Case 1
            Set objControl = DataDocsRec
        Case 2
            Set objControl = DataEnvioAnalise
        Case 3
            Set objControl = DataPagtoPax
        Case 4
            Set objControl = DataLimite
        Case 5
            Set objControl = DataProgrFin
        Case 6
            Set objControl = DataFimProcesso
        Case 7
            Set objControl = DataHist
        Case 8
            Set objControl = DataEmissao
        Case 9
            Set objControl = DataIda
        Case 10
            Set objControl = DataVolta
        Case 11
            Set objControl = DataIniProcesso
    End Select

End Sub

Private Sub BotaoHoje_Click(Index As Integer)

Dim lErro As Long
Dim objControl As Object

On Error GoTo Erro_BotaoHoje_Click

    Call Obtem_ControleData(Index, objControl)

    Call DateParaMasked(objControl, gdtDataAtual)

    Exit Sub

Erro_BotaoHoje_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208620)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoProg_Click(Index As Integer)

Dim lErro As Long
Dim objControl As Object
Dim objProgData As New ProgData
Dim objProgData2 As New ProgData2
Dim dValorPagar As Double

On Error GoTo Erro_BotaoProg_Click

    Call Obtem_ControleData(Index, objControl)
    
    If StrParaDate(DataDocsRec.Text) = DATA_NULA Then gError 209143
    
    Select Case objControl.Name
    
        Case DataLimite.Name
            Load objProgData
            Call objProgData.Trata_Parametros(DataLimite, StrParaDate(DataDocsRec.Text), 30, CONDPAGTO_TIPOINTERVALO_DIAS_UTEIS)
            objProgData.Show vbModal
            
        Case DataProgrFin.Name
        
            If StrParaLong(Codigo.Text) = 0 Then gError 209140
            If StrParaDate(DataLimite.Text) = DATA_NULA Then gError 209141
            
            dValorPagar = StrParaDbl(SrvTotalAssistRS.Caption) + IIf(AntecPagto.Value = vbChecked, StrParaDbl(SrvTotalSegRS.Caption), StrParaDbl(SrvTotalSegTrvRS.Text))
        
            If dValorPagar = 0 Then gError 209142
        
            Load objProgData2
            Call objProgData2.Trata_Parametros(DataProgrFin, Codigo.Text, StrParaDate(DataDocsRec.Text), StrParaDate(DataLimite.Text), StrParaDbl(SrvTotalAssistRS.Caption), StrParaDbl(SrvTotalSegRS.Caption), IIf(AntecPagto.Value = vbChecked, MARCADO, DESMARCADO), StrParaDbl(SrvTotalSegTrvRS.Text))
            objProgData2.Show vbModal
    
        Case Else
            Call DateParaMasked(objControl, gdtDataAtual)
    
    End Select

    Exit Sub

Erro_BotaoProg_Click:

    Select Case gErr

        Case 209140
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCRCASOS_PROG_CODIGO_NAO_PREENCHIDO", gErr)

        Case 209141
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCRCASOS_PROG_DATALIMITE_NAO_PREENCHIDA", gErr)

        Case 209142
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCRCASOS_PROG_VALORPAGAR_NAO_PREENCHIDO", gErr)

        Case 209143
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCRCASOS_PROG_DATADOCSREC_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208621)

    End Select

    Exit Sub
    
End Sub

Private Sub QtdePax_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Contato_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Email1_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Telefone1_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Telefone2_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Judicial_Click()
    If Judicial.Value <> JudicialE.Value Then
        JudicialE.Value = Judicial.Value
    End If
    iAlterado = REGISTRO_ALTERADO
    Call Trata_Juridico
End Sub

Private Sub JudicialE_Click()
    If Judicial.Value <> JudicialE.Value Then
        Judicial.Value = JudicialE.Value
    End If
    iAlterado = REGISTRO_ALTERADO
    Call Trata_Juridico
End Sub

Private Sub Trata_Juridico()
    If Judicial.Value = vbChecked Then
        FrameProc.Enabled = True
    Else
        FrameProc.Enabled = False
        Condenado.Value = vbUnchecked
        Call Trata_Condenacao
        NumProcesso.PromptInclude = False
        NumProcesso.Text = ""
        NumProcesso.PromptInclude = True
        DataFimProcesso.PromptInclude = False
        DataFimProcesso.Text = ""
        DataFimProcesso.PromptInclude = True
        Comarca.Text = ""
        Procon.Value = vbUnchecked
    End If
End Sub

Private Sub CGAnalise_Change()
    iAlterado = REGISTRO_ALTERADO
    Analise = CGAnalise.Text
End Sub

Private Sub CGAnalise_Click()
    iAlterado = REGISTRO_ALTERADO
    Analise = CGAnalise.Text
End Sub

Public Sub CGAnalise_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CGAnalise_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_TRVOCRCASO_ANALISE, CGAnalise, "AVISO_CRIAR_TRVOCRCASO")
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Analise = CGAnalise.Text
    
    Exit Sub

Erro_CGAnalise_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208622)

    End Select

End Sub

Private Sub CGStatus_Change()
    iAlterado = REGISTRO_ALTERADO
    Status = CGStatus.Text
End Sub

Private Sub CGStatus_Click()
    iAlterado = REGISTRO_ALTERADO
    Status = CGStatus.Text
End Sub

Public Sub CGStatus_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CGStatus_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_TRVOCRCASO_STATUS, CGStatus, "AVISO_CRIAR_TRVOCRCASO")
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Status = CGStatus.Text
    
    Exit Sub

Erro_CGStatus_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208623)

    End Select

End Sub

Private Sub CGAutorizadoPor_Change()
    iAlterado = REGISTRO_ALTERADO
    AutorizadoPor = CGAutorizadoPor.Text
End Sub

Private Sub CGAutorizadoPor_Click()
    iAlterado = REGISTRO_ALTERADO
    AutorizadoPor = CGAutorizadoPor.Text
End Sub

Public Sub CGAutorizadoPor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CGAutorizadoPor_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_TRVOCRCASO_AUTOPOR, CGAutorizadoPor, "AVISO_CRIAR_TRVOCRCASO")
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    AutorizadoPor = CGAutorizadoPor.Text
    
    Exit Sub

Erro_CGAutorizadoPor_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208624)

    End Select

End Sub

Private Sub SrvAuto_Click()
    iAlterado = REGISTRO_ALTERADO
    Call Trata_SrvAuto_Click
End Sub

Public Sub SrvAuto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridSrv)
End Sub

Public Sub SrvAuto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridSrv)
End Sub

Public Sub SrvAuto_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridSrv.objControle = SrvAuto
    lErro = Grid_Campo_Libera_Foco(objGridSrv)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub SrvSol_Click()
    iAlterado = REGISTRO_ALTERADO
    Call Trata_SrvSol_Click
End Sub

Public Sub SrvSol_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridSrv)
End Sub

Public Sub SrvSol_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridSrv)
End Sub

Public Sub SrvSol_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridSrv.objControle = SrvSol
    lErro = Grid_Campo_Libera_Foco(objGridSrv)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub SrvVlrSolRS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub SrvVlrSolRS_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridSrv)
End Sub

Public Sub SrvVlrSolRS_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridSrv)
End Sub

Public Sub SrvVlrSolRS_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridSrv.objControle = SrvVlrSolRS
    lErro = Grid_Campo_Libera_Foco(objGridSrv)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub SrvVlrSolUS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub SrvVlrSolUS_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridSrv)
End Sub

Public Sub SrvVlrSolUS_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridSrv)
End Sub

Public Sub SrvVlrSolUS_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridSrv.objControle = SrvVlrSolUS
    lErro = Grid_Campo_Libera_Foco(objGridSrv)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub SrvVlrAutoUS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub SrvVlrAutoUS_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridSrv)
End Sub

Public Sub SrvVlrAutoUS_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridSrv)
End Sub

Public Sub SrvVlrAutoUS_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridSrv.objControle = SrvVlrAutoUS
    lErro = Grid_Campo_Libera_Foco(objGridSrv)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub SrvVlrAutoRS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub SrvVlrAutoRS_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridSrv)
End Sub

Public Sub SrvVlrAutoRS_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridSrv)
End Sub

Public Sub SrvVlrAutoRS_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridSrv.objControle = SrvVlrAutoRS
    lErro = Grid_Campo_Libera_Foco(objGridSrv)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub AntecPagto_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Cambio_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Cambio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Cambio_Validate

    'Veifica se Previsao está preenchida
    If Len(Trim(Cambio.Text)) <> 0 Then

       'Critica a Previsao
       lErro = Valor_Positivo_Critica(Cambio.Text)
       If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
       
       Call Trata_Mudanca_Cambio
        
    End If

    Exit Sub

Erro_Cambio_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208625)

    End Select

    Exit Sub
    
End Sub

Private Sub Trata_Mudanca_Cambio()

Dim lErro As Long, iLinha As Integer
Dim objOcrCasos As ClassTRVOcrCasos
Dim objOcrCasosSrv As ClassTRVOcrCasosSrv
Dim objOcrCasosOF As ClassTRVOcrCasosOutrasFat
Dim objOcrCasosII As ClassTRvOcrCasosInvoices
Dim dCambio As Double

On Error GoTo Erro_Trata_Mudanca_Cambio

    dCambio = StrParaDbl(Cambio.Text)

    If Abs(gdCambioAnt - dCambio) > DELTA_VALORMONETARIO2 Then
    
        gdCambioAnt = dCambio
        
        Set objOcrCasos = New ClassTRVOcrCasos
    
        lErro = Move_Srv_Memoria(objOcrCasos, True)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        lErro = Move_II_Memoria(objOcrCasos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        lErro = Move_OF_Memoria(objOcrCasos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        For Each objOcrCasosSrv In objOcrCasos.colCoberturas
            If objOcrCasosSrv.dValorAutorizadoUS > 0 Then
                objOcrCasosSrv.dValorAutorizadoRS = objOcrCasosSrv.dValorAutorizadoUS * gdCambioAnt
            ElseIf gdCambioAnt > 0 Then
                objOcrCasosSrv.dValorAutorizadoUS = objOcrCasosSrv.dValorAutorizadoRS / gdCambioAnt
            End If
            If objOcrCasosSrv.dValorSolicitadoUS > 0 Then
                objOcrCasosSrv.dValorSolicitadoRS = objOcrCasosSrv.dValorSolicitadoUS * gdCambioAnt
            ElseIf gdCambioAnt > 0 Then
                objOcrCasosSrv.dValorSolicitadoUS = objOcrCasosSrv.dValorSolicitadoRS / gdCambioAnt
            End If
        Next
    
        For Each objOcrCasosII In objOcrCasos.colInvoices
            If objOcrCasosII.iMoeda = MOEDA_DOLAR Then
                If objOcrCasosII.dValorMoeda > 0 Then
                    objOcrCasosII.dValorRS = objOcrCasosII.dValorMoeda * gdCambioAnt
                ElseIf gdCambioAnt > 0 Then
                    objOcrCasosII.dValorMoeda = objOcrCasosII.dValorRS / gdCambioAnt
                End If
            End If
        Next
        
        For Each objOcrCasosOF In objOcrCasos.colOutrasFaturas
            If objOcrCasosOF.dValorUS > 0 Then
                objOcrCasosOF.dValorRS = objOcrCasosOF.dValorUS * gdCambioAnt
            ElseIf gdCambioAnt > 0 Then
                objOcrCasosOF.dValorUS = objOcrCasosOF.dValorRS / gdCambioAnt
            End If
        Next
    
        lErro = Traz_Srv_Tela(objOcrCasos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        lErro = Traz_II_Tela(objOcrCasos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        lErro = Traz_OF_Tela(objOcrCasos)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If

    Exit Sub

Erro_Trata_Mudanca_Cambio:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208626)

    End Select

    Exit Sub
    
End Sub

Public Sub LabelBanco_Click()

Dim objBanco As New ClassBanco
Dim colSelecao As New Collection

    objBanco.iCodBanco = StrParaInt(Banco.Text)

    Call Chama_Tela("BancoLista", colSelecao, objBanco, objEventoBanco)

End Sub

Private Sub objEventoBanco_evSelecao(obj1 As Object)

Dim objBanco As ClassBanco
Dim bCancel As Boolean

    Set objBanco = obj1

    Banco.PromptInclude = False
    Banco.Text = CStr(objBanco.iCodBanco)
    Banco.PromptInclude = True

    Me.Show

    Exit Sub

End Sub

Private Sub Banco_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Banco_GotFocus()
    Call MaskEdBox_TrataGotFocus(Banco, iAlterado)
End Sub

Private Sub Banco_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Banco_Validate

    'Verifica se foi preenchido o campo Banco
    If Len(Trim(Banco.Text)) = 0 Then Exit Sub

    'Critica se é do tipo positivo
    lErro = Valor_Positivo_Critica(Banco.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_Banco_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208627)

    End Select

    Exit Sub

End Sub

Private Sub Agencia_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ContaCorrente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub FavorecidoCGC_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub FavorecidoCGC_GotFocus()
    Call MaskEdBox_TrataGotFocus(FavorecidoCGC, iAlterado)
End Sub

Private Sub NomeFavorecido_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub BotaoAbrirFav_Click()

Dim objForn As New ClassFornecedor

On Error GoTo Erro_BotaoAbrirFav_Click

    objForn.lCodigo = StrParaLong(FornFavorecido.Caption)

    Call Chama_Tela("Fornecedores", objForn)

    Exit Sub

Erro_BotaoAbrirFav_Click:

    Select Case gErr
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208628)

    End Select

    Exit Sub
    
End Sub

Private Sub Trata_SrvAuto_Click()

Dim dCambio As Double, iLinha As Integer, iMoeda As Integer
Dim dValorLimiteRS As Double, dValorLimiteUS As Double, dValorLimite As Double
Dim dValorSolUS As Double, dValorSolRS As Double
Dim dValorAutoUS As Double, dValorAutoRS As Double

On Error GoTo Erro_Trata_SrvAuto_Click

    dCambio = StrParaDbl(Cambio.Text)
    
    iLinha = GridSrv.Row
    
    If iLinha > 0 And iLinha <= objGridSrv.iLinhasExistentes Then
    
        If StrParaInt(GridSrv.TextMatrix(iLinha, iGrid_SrvAuto_Col)) = MARCADO Then
    
            iMoeda = Codigo_Extrai(GridSrv.TextMatrix(iLinha, iGrid_SrvMoeda_Col))
            dValorLimite = StrParaDbl(GridSrv.TextMatrix(iLinha, iGrid_SrvVlrLimite_Col))
            
            If iMoeda = MOEDA_REAL Then
                dValorLimiteRS = dValorLimite
                If dCambio > 0 Then dValorLimiteUS = dValorLimiteRS / dCambio
            ElseIf iMoeda = MOEDA_DOLAR Then
                dValorLimiteUS = dValorLimite
                dValorLimiteRS = dValorLimiteUS * dCambio
            End If
            
            dValorSolUS = StrParaDbl(OFTotalUS.Caption) 'StrParaDbl(GridSrv.TextMatrix(iLinha, iGrid_SrvVlrSolUS_Col))
            dValorSolRS = StrParaDbl(OFTotalRS.Caption) 'StrParaDbl(GridSrv.TextMatrix(iLinha, iGrid_SrvVlrSolRS_Col))
    
            If dValorSolUS < dValorLimiteUS Or dValorLimiteUS = 0 Then
                GridSrv.TextMatrix(iLinha, iGrid_SrvVlrAutoUS_Col) = Format(dValorSolUS, "STANDARD")
                GridSrv.TextMatrix(iLinha, iGrid_SrvVlrAutoRS_Col) = Format(dValorSolRS, "STANDARD")
            Else
                GridSrv.TextMatrix(iLinha, iGrid_SrvVlrAutoUS_Col) = Format(dValorLimiteUS, "STANDARD")
                GridSrv.TextMatrix(iLinha, iGrid_SrvVlrAutoRS_Col) = Format(dValorLimiteRS, "STANDARD")
            End If
        
        Else
        
            GridSrv.TextMatrix(iLinha, iGrid_SrvVlrAutoUS_Col) = ""
            GridSrv.TextMatrix(iLinha, iGrid_SrvVlrAutoRS_Col) = ""
        
        End If
        
    End If
    
    Call Totaliza_Valores_Srv

    Exit Sub

Erro_Trata_SrvAuto_Click:

    Select Case gErr
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208629)

    End Select

    Exit Sub
    
End Sub

Function Saida_Celula_Valor(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'Faz a crítica da célula Data que está deixando de ser a corrente

Dim lErro As Long
Dim dCambio As Double, dValorAnt As Double, dValorAtual As Double
Dim iMoeda As Integer, iLinha As Integer

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = objControle

    If Len(Trim(objControle.Text)) > 0 Then
    
        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        objControle.Text = Format(objControle.Text, "STANDARD")
        
        iLinha = objGridInt.objGrid.Row
        
        dCambio = StrParaDbl(Cambio.Text)
        
        dValorAnt = StrParaDbl(objGridInt.objGrid.TextMatrix(iLinha, objGridInt.objGrid.Col))
        dValorAtual = StrParaDbl(objControle.Text)
        
        If dCambio > 0 And Abs(dValorAtual - dValorAnt) > DELTA_VALORMONETARIO Then
        
            Select Case objControle.Name
            
                'Atualiza os valores em Real com base no valor em dólar e câmbio e vice-versa
                Case SrvVlrAutoUS.Name
                    GridSrv.TextMatrix(iLinha, iGrid_SrvVlrAutoRS_Col) = Format(dValorAtual * dCambio, "STANDARD")
                    iValorAutoAlterado = REGISTRO_ALTERADO
                    
                Case SrvVlrSolUS.Name
                    GridSrv.TextMatrix(iLinha, iGrid_SrvVlrSolRS_Col) = Format(dValorAtual * dCambio, "STANDARD")
                    iValorSolAlterado = REGISTRO_ALTERADO
                
                Case SrvVlrAutoRS.Name
                    GridSrv.TextMatrix(iLinha, iGrid_SrvVlrAutoUS_Col) = Format(dValorAtual / dCambio, "STANDARD")
                    iValorAutoAlterado = REGISTRO_ALTERADO
                    
                Case SrvVlrSolRS.Name
                    GridSrv.TextMatrix(iLinha, iGrid_SrvVlrSolUS_Col) = Format(dValorAtual / dCambio, "STANDARD")
                    iValorSolAlterado = REGISTRO_ALTERADO
    
                Case IIValorRS.Name
                    'iMoeda = Codigo_Extrai(GridII.TextMatrix(iLinha, iGrid_IIMoeda_Col))
'                    If iMoeda = MOEDA_REAL Then
'                        GridSrv.TextMatrix(iLinha, iGrid_IIValorMoeda_Col) = Format(dValorAtual, "STANDARD")
'                    ElseIf iMoeda = MOEDA_DOLAR Then
                        GridSrv.TextMatrix(iLinha, iGrid_IIValorMoeda_Col) = Format(dValorAtual / dCambio, "STANDARD")
'                    End If
                
                Case IIValorMoeda.Name
                    'iMoeda = Codigo_Extrai(GridII.TextMatrix(iLinha, iGrid_IIMoeda_Col))
'                    If iMoeda = MOEDA_REAL Then
'                        GridII.TextMatrix(iLinha, iGrid_IIValorRS_Col) = Format(dValorAtual, "STANDARD")
'                    ElseIf iMoeda = MOEDA_DOLAR Then
                        GridII.TextMatrix(iLinha, iGrid_IIValorRS_Col) = Format(dValorAtual * dCambio, "STANDARD")
'                    End If
                    
                Case OFValorRS.Name
                    GridOF.TextMatrix(iLinha, iGrid_OFValorUS_Col) = Format(dValorAtual / dCambio, "STANDARD")
                
                Case OFValorUS.Name
                    GridOF.TextMatrix(iLinha, iGrid_OFValorRS_Col) = Format(dValorAtual * dCambio, "STANDARD")
                    
                Case Else
                    'Não faz nada
                
            End Select
            
        End If

       
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    
    Select Case objGridInt.objGrid.Name
    
        Case GridSrv.Name
            Call Totaliza_Valores_Srv
    
        Case GridII.Name
            Call Soma_Coluna_Grid(objGridII, iGrid_IIValorRS_Col, IITotalRS, False)
            Call Soma_Coluna_Grid(objGridII, iGrid_IIValorMoeda_Col, IITotalUS, False)
            
        Case GridOF.Name
            Call Trata_Reembolso
'            Call Soma_Coluna_Grid(objGridOF, iGrid_OFValorRS_Col, OFTotalRS, False, iGrid_OFConsiderar_Col)
'            Call Soma_Coluna_Grid(objGridOF, iGrid_OFValorUS_Col, OFTotalUS, False, iGrid_OFConsiderar_Col)
            
        Case GridGA.Name
            Call Soma_Coluna_Grid(objGridGA, iGrid_GAValor_Col, GATotal, False)
    
    End Select

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208630)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Verifica qual é o grid
        If objGridInt.objGrid.Name = GridSrv.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_SrvAuto_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, SrvAuto)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                     
                Case iGrid_SrvVlrAutoRS_Col
                
                    lErro = Saida_Celula_Valor(objGridInt, SrvVlrAutoRS)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                     
                Case iGrid_SrvVlrAutoUS_Col
                
                    lErro = Saida_Celula_Valor(objGridInt, SrvVlrAutoUS)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                     
                Case iGrid_SrvVlrSolRS_Col
                
                    lErro = Saida_Celula_Valor(objGridInt, SrvVlrSolRS)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                     
                Case iGrid_SrvVlrSolUS_Col
                
                    lErro = Saida_Celula_Valor(objGridInt, SrvVlrSolUS)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                     
            End Select
            
        ElseIf objGridInt.objGrid.Name = GridII.Name Then
        
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_IIValorMoeda_Col
                
                    lErro = Saida_Celula_Valor(objGridInt, IIValorMoeda)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                Case iGrid_IIValorRS_Col
                
                    lErro = Saida_Celula_Valor(objGridInt, IIValorRS)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                Case iGrid_IIDataEmi_Col
                
                    lErro = Saida_Celula_Data(objGridInt, IIDataEmi, True)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                Case iGrid_IIDataRec_Col
                
                    lErro = Saida_Celula_Data(objGridInt, IIDataRec, True)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                Case iGrid_IIObs_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, IIObs)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                Case iGrid_IINumero_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, IINumero)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
'                Case iGrid_IIMoeda_Col
'
'                    lErro = Saida_Celula_Padrao(objGridInt, IIMoeda)
'                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            End Select
        
        
        ElseIf objGridInt.objGrid.Name = GridOF.Name Then
        
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_OFValorUS_Col
                
                    lErro = Saida_Celula_Valor(objGridInt, OFValorUS)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                Case iGrid_OFValorRS_Col
                
                    lErro = Saida_Celula_Valor(objGridInt, OFValorRS)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                Case iGrid_OFDataEmi_Col
                
                    lErro = Saida_Celula_Data(objGridInt, OFDataEmi, True)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                Case iGrid_OFDataRec_Col
                
                    lErro = Saida_Celula_Data(objGridInt, OFDataRec, True)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                Case iGrid_OFDescricao_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, OFDescricao)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                Case iGrid_OFNumero_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, OFNumero)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                Case iGrid_OFConsiderar_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, OFConsiderar)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            End Select
        
        ElseIf objGridInt.objGrid.Name = GridPC.Name Then
        
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_PCValor_Col
                
                    lErro = Saida_Celula_Valor(objGridInt, PCValor)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                Case iGrid_PCVencimento_Col
                
                    lErro = Saida_Celula_Data(objGridInt, PCVencimento, True)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            End Select
        
        ElseIf objGridInt.objGrid.Name = GridGA.Name Then
        
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_GAValor_Col
                
                    lErro = Saida_Celula_Valor(objGridInt, GAValor)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                Case iGrid_GAData_Col
                
                    lErro = Saida_Celula_Data(objGridInt, GAData, True)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                Case iGrid_GADesc_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, GADesc)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            End Select
            
        ElseIf objGridInt.objGrid.Name = GridDN.Name Then
        
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_DNNU_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, DNNU)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                Case iGrid_DNRecebido_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, DNRecebido)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                Case iGrid_DNObs_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, DNObs)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            End Select
        
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 208631

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 208631
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208632)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable
              
    Select Case objControl.Name
    
        Case SrvSol.Name
            objControl.Enabled = True
    
        Case SrvAuto.Name, SrvVlrSolRS.Name, SrvVlrSolUS.Name
            If StrParaInt(GridSrv.TextMatrix(iLinha, iGrid_SrvSol_Col)) = MARCADO Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
        Case SrvVlrAutoRS.Name, SrvVlrAutoUS.Name
            If StrParaInt(GridSrv.TextMatrix(iLinha, iGrid_SrvAuto_Col)) = MARCADO Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
        Case PCValor.Name, PCVencimento.Name
            objControl.Enabled = True
    
        Case PCPagamento.Name
                objControl.Enabled = True
    
        Case Else
            If left(objControl.Name, 2) = "II" Or left(objControl.Name, 2) = "OF" Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
    End Select
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208633)

    End Select

    Exit Sub

End Sub

Private Sub Cidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Bairro_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Logradouro_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Numero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Numero_GotFocus()
    Call MaskEdBox_TrataGotFocus(Numero, iAlterado)
End Sub

Private Sub Complemento_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub IIDataRec_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub IIDataRec_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridII)
End Sub

Public Sub IIDataRec_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridII)
End Sub

Public Sub IIDataRec_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridII.objControle = IIDataRec
    lErro = Grid_Campo_Libera_Foco(objGridII)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub IIDataEmi_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub IIDataEmi_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridII)
End Sub

Public Sub IIDataEmi_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridII)
End Sub

Public Sub IIDataEmi_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridII.objControle = IIDataEmi
    lErro = Grid_Campo_Libera_Foco(objGridII)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub IIValorMoeda_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub IIValorMoeda_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridII)
End Sub

Public Sub IIValorMoeda_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridII)
End Sub

Public Sub IIValorMoeda_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridII.objControle = IIValorMoeda
    lErro = Grid_Campo_Libera_Foco(objGridII)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub IIValorRS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub IIValorRS_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridII)
End Sub

Public Sub IIValorRS_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridII)
End Sub

Public Sub IIValorRS_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridII.objControle = IIValorRS
    lErro = Grid_Campo_Libera_Foco(objGridII)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub IIObs_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub IIObs_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridII)
End Sub

Public Sub IIObs_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridII)
End Sub

Public Sub IIObs_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridII.objControle = IIObs
    lErro = Grid_Campo_Libera_Foco(objGridII)
    If lErro <> SUCESSO Then Cancel = True
End Sub

'Private Sub IIMoeda_Change()
'    iAlterado = REGISTRO_ALTERADO
'End Sub
'
'Private Sub IIMoeda_Click()
'    iAlterado = REGISTRO_ALTERADO
'End Sub
'
'Public Sub IIMoeda_GotFocus()
'    Call Grid_Campo_Recebe_Foco(objGridII)
'End Sub
'
'Public Sub IIMoeda_KeyPress(KeyAscii As Integer)
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridII)
'End Sub
'
'Public Sub IIMoeda_Validate(Cancel As Boolean)
'Dim lErro As Long
'    Set objGridII.objControle = IIMoeda
'    lErro = Grid_Campo_Libera_Foco(objGridII)
'    If lErro <> SUCESSO Then Cancel = True
'End Sub

Private Sub IINumero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub IINumero_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridII)
End Sub

Public Sub IINumero_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridII)
End Sub

Public Sub IINumero_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridII.objControle = IINumero
    lErro = Grid_Campo_Libera_Foco(objGridII)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub OFDataRec_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub OFDataRec_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridOF)
End Sub

Public Sub OFDataRec_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOF)
End Sub

Public Sub OFDataRec_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridOF.objControle = OFDataRec
    lErro = Grid_Campo_Libera_Foco(objGridOF)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub OFDataEmi_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub OFDataEmi_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridOF)
End Sub

Public Sub OFDataEmi_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOF)
End Sub

Public Sub OFDataEmi_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridOF.objControle = OFDataEmi
    lErro = Grid_Campo_Libera_Foco(objGridOF)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub OFValorUS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub OFValorUS_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridOF)
End Sub

Public Sub OFValorUS_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOF)
End Sub

Public Sub OFValorUS_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridOF.objControle = OFValorUS
    lErro = Grid_Campo_Libera_Foco(objGridOF)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub OFValorRS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub OFValorRS_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridOF)
End Sub

Public Sub OFValorRS_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOF)
End Sub

Public Sub OFValorRS_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridOF.objControle = OFValorRS
    lErro = Grid_Campo_Libera_Foco(objGridOF)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub OFDescricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub OFDescricao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridOF)
End Sub

Public Sub OFDescricao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOF)
End Sub

Public Sub OFDescricao_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridOF.objControle = OFDescricao
    lErro = Grid_Campo_Libera_Foco(objGridOF)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub OFConsiderar_Click()
    iAlterado = REGISTRO_ALTERADO
    Call Trata_Reembolso
'    Call Soma_Coluna_Grid(objGridOF, iGrid_OFValorRS_Col, OFTotalRS, False, iGrid_OFConsiderar_Col)
'    Call Soma_Coluna_Grid(objGridOF, iGrid_OFValorUS_Col, OFTotalUS, False, iGrid_OFConsiderar_Col)
End Sub

Public Sub OFConsiderar_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridOF)
End Sub

Public Sub OFConsiderar_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOF)
End Sub

Public Sub OFConsiderar_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridOF.objControle = OFConsiderar
    lErro = Grid_Campo_Libera_Foco(objGridOF)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub OFNumero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub OFNumero_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridOF)
End Sub

Public Sub OFNumero_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOF)
End Sub

Public Sub OFNumero_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridOF.objControle = OFNumero
    lErro = Grid_Campo_Libera_Foco(objGridOF)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente

    Set objcliente = obj1

    'Preenche campo Cliente
    ClienteVou.Text = objcliente.sNomeReduzido

    'Executa o Validate
    Call ClienteVou_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Public Sub LabelCliente_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objcliente.sNomeReduzido = ClienteVou.Text

    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente)

End Sub

Private Sub BotaoGravarHist_Click()

Dim lErro As Long
Dim objOrcCasos As New ClassTRVOcrCasos
Dim objOrcCasosHist As New ClassTRVOcrCasosHist

On Error GoTo Erro_BotaoGravarHist_Click

    If glNumIntDoc = 0 Then gError 208648
    If Len(Trim(AssuntoHist.Text)) = 0 Then gError 208638
    If OrigemHist.ListIndex = -1 Then gError 208639
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    If OrigemHist.Text = RELACIONAMENTOCLIENTES_ORIGEM_CLIENTE_TEXTO Then
        objOrcCasosHist.iOrigem = RELACIONAMENTOCLIENTES_ORIGEM_CLIENTE
    ElseIf OrigemHist.Text = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA_TEXTO Then
        objOrcCasosHist.iOrigem = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA
    End If
    
    If Len(Trim(HoraHist.ClipText)) = 0 Then
        objOrcCasosHist.dHora = CDbl(Time)
    Else
        objOrcCasosHist.dHora = CDbl(StrParaDate(HoraHist.Text))
    End If
    
    If StrParaDate(DataHist.ClipText) = DATA_NULA Then
        objOrcCasosHist.dtData = Date
    Else
        objOrcCasosHist.dtData = StrParaDate(DataHist.Text)
    End If
        
    objOrcCasosHist.lNumIntDocOcrCaso = glNumIntDoc
    objOrcCasos.lNumIntDoc = glNumIntDoc
    objOrcCasosHist.sAssunto = AssuntoHist.Text
    
    lErro = CF("TRVOcrCasosHist_Grava", objOrcCasosHist)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
       
    lErro = CF("TRVOcrCasosHist_Le", objOrcCasos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Traz_Hist_Tela(objOrcCasos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    OrigemHist.ListIndex = -1
    HoraHist.PromptInclude = False
    HoraHist.Text = ""
    HoraHist.PromptInclude = True
    AssuntoHist.Text = ""
    Call DateParaMasked(DataHist, gdtDataAtual)
        
    GL_objMDIForm.MousePointer = vbNormal
        
    Exit Sub

Erro_BotaoGravarHist_Click:

    GL_objMDIForm.MousePointer = vbNormal

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 208638
            Call Rotina_Erro(vbOKOnly, "ERRO_ASSUNTO_NAO_PREENCHIDO", gErr)
        
        Case 208639
            Call Rotina_Erro(vbOKOnly, "ERRO_ORIGEM_NAO_PREENCHIDA1", gErr)
       
        Case 208649
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCRCASOS_TRAZER_TELA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208640)

    End Select

    Exit Sub
    
End Sub

Private Sub Exibe_CampoDet_Grid(ByVal objGridInt As AdmGrid, ByVal iColunaExibir As Integer, ByVal objControle As Object)

Dim iLinha As Integer

On Error GoTo Erro_Exibe_CampoDet_Grid

    iLinha = objGridInt.objGrid.Row
    
    If iLinha > 0 And iLinha <= objGridInt.iLinhasExistentes Then
        objControle.Text = objGridInt.objGrid.TextMatrix(iLinha, iColunaExibir)
    Else
        objControle.Text = ""
    End If

    Exit Sub

Erro_Exibe_CampoDet_Grid:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208641)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoGravarAnot_Click()

Dim lErro As Long
Dim objOrcCasos As New ClassTRVOcrCasos
Dim objOrcCasosAnot As New ClassTRVOcrCasosAnotacoes

On Error GoTo Erro_BotaoGravarAnot_Click

    If glNumIntDoc = 0 Then gError 208648
    If Len(Trim(TextoAnot.Text)) = 0 Then gError 208642
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    objOrcCasosAnot.sTexto = TextoAnot.Text
    objOrcCasosAnot.lNumIntDocOcrCaso = glNumIntDoc
    objOrcCasos.lNumIntDoc = glNumIntDoc
    
    objOrcCasosAnot.dtData = Date
    objOrcCasosAnot.dHora = CDbl(Time)
    
    lErro = CF("TRVOcrCasosAnotacoes_Grava", objOrcCasosAnot)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("TRVOcrCasosAnotacoes_Le", objOrcCasos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Traz_Anot_Tela(objOrcCasos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    TextoAnot.Text = ""
        
    GL_objMDIForm.MousePointer = vbNormal
        
    Exit Sub

Erro_BotaoGravarAnot_Click:

    GL_objMDIForm.MousePointer = vbNormal

    Select Case gErr
        
        Case 208642
            Call Rotina_Erro(vbOKOnly, "ERRO_TEXTO_NAO_PREENCHIDO", gErr)
            
        Case 208648
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCRCASOS_TRAZER_TELA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208643)

    End Select

    Exit Sub
    
End Sub

Public Sub HoraHist_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_HoraHist_Validate

    'Verifica se a hora de saida foi digitada
    If Len(Trim(HoraHist.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Hora_Critica(HoraHist.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_HoraHist_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208644)

    End Select

    Exit Sub

End Sub

Private Sub Condenado_Click()
    iAlterado = REGISTRO_ALTERADO
    Call Trata_Condenacao
End Sub

Private Sub Obs_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NumProcesso_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Comarca_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorCondenacao_Change()
    iAlterado = REGISTRO_ALTERADO
    VlrCond.Caption = ValorCondenacao.Text
End Sub

Private Sub ValorCondenacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorCondenacao_Validate

    'Veifica se ValorCondenacao está preenchida
    If Len(Trim(ValorCondenacao.Text)) <> 0 Then

       'Critica a ValorCondenacao
       lErro = Valor_Positivo_Critica(ValorCondenacao.Text)
       If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If

    Exit Sub

Erro_ValorCondenacao_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208645)

    End Select

    Exit Sub
    
End Sub

Private Sub PCVencimento_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PCVencimento_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPC)
End Sub

Public Sub PCVencimento_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPC)
End Sub

Public Sub PCVencimento_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridPC.objControle = PCVencimento
    lErro = Grid_Campo_Libera_Foco(objGridPC)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub PCValor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PCValor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridPC)
End Sub

Public Sub PCValor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPC)
End Sub

Public Sub PCValor_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridPC.objControle = PCValor
    lErro = Grid_Campo_Libera_Foco(objGridPC)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub Trata_Condenacao()

Dim lErro As Long

On Error GoTo Erro_Trata_Condenacao

    If Condenado.Value = vbChecked Then
        ValorCondenacao.Enabled = True
        FramePC.Enabled = True
        OptCondenacao.Visible = True
        OptAcordo.Visible = True
    Else
        ValorCondenacao.Enabled = False
        ValorCondenacao.Text = ""
        FramePC.Enabled = False
        Call Grid_Limpa(objGridPC)
        OptCondenacao.Visible = False
        OptAcordo.Visible = False
        OptCondenacao.Value = True
    End If

    Exit Sub

Erro_Trata_Condenacao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208646)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoAbrirFatProc_Click()

Dim lErro As Long
Dim objFat As New ClassTituloPagar

On Error GoTo Erro_BotaoAbrirFatProc_Click

    If StrParaLong(NumeroFatProc.Caption) <> 0 Then

        objFat.lNumIntDoc = glNumIntDocTitPagProc
        
        Call Chama_Tela(TRV_TIPO_DOC_DESTINO_TITPAG_TELA, objFat)
        
    End If
    
    Exit Sub

Erro_BotaoAbrirFatProc_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208647)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoLimparContato_Click()
    Contato.Text = ""
    Telefone1.Text = ""
    Telefone2.Text = ""
    Email1.Text = ""
    Contato.SetFocus
End Sub

Private Sub BotaoLimparFavorecido_Click()
    NomeFavorecido.Text = ""
    CEP.Text = ""
    Pais.ListIndex = giIndexBrasil
    Estado.ListIndex = -1
    Bairro.Text = ""
    Cidade.Text = ""
    Numero.Text = ""
    Complemento.Text = ""
    Logradouro.Text = ""
    FavorecidoCGC.Text = ""
    NomeFavorecido.SetFocus
End Sub

Private Sub BotaoFatJur_Click()
    Call BotaoAbrirFatProc_Click
End Sub

Private Sub BotaoFatCobr_Click()

Dim lErro As Long
Dim objFat As New ClassTituloPagar

On Error GoTo Erro_BotaoFatCobr_Click

    If StrParaLong(FatCobertura.Caption) <> 0 Then

        objFat.lNumIntDoc = glNumIntDocTitPagCober
        
        Call Chama_Tela(TRV_TIPO_DOC_DESTINO_TITPAG_TELA, objFat)
        
    End If
    
    Exit Sub

Erro_BotaoFatCobr_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208705)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoFatReemb_Click()

Dim lErro As Long
Dim objFat As New ClassTituloReceber

On Error GoTo Erro_BotaoFatReemb_Click

    If StrParaLong(FatReembolso.Caption) <> 0 Then

        objFat.lNumIntDoc = glNumIntDocTitRecReembolso
        
        Call Chama_Tela("TRVOcrCasosReemb", objFat)
        
    End If
    
    Exit Sub

Erro_BotaoFatReemb_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208755)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoLibera_Click()

Dim objOcr As New ClassTRVOcrCasos

    objOcr.sCodigo = Trim(Codigo.Text)

    If glNumIntDocTitPagCober = 0 Then
        Call Chama_Tela("TRVLibCoberOcrCasos", objOcr)
    ElseIf glNumIntDocTitPagProc = 0 And Judicial.Value = vbChecked Then
        Call Chama_Tela("TRVLibJurOcrCasos", objOcr)
    End If
    
End Sub

Private Sub ProdutoLabel_Click()

Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long
Dim sSelecao  As String

On Error GoTo Erro_ProdutoLabel_Click
   
    'Verifica se Produto está preenchido
    If Len(Trim(Produto.ClipText)) > 0 Then

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objProduto.sCodigo = sProdutoFormatado
        Else
            objProduto.sCodigo = ""
        End If

    End If

    'Chama a tela de browse
    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoProduto, sSelecao)

    Exit Sub
    
Erro_ProdutoLabel_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209090)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProdutoEnxuto As String

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1
    
    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Coloca o Codigo na tela
    Produto.PromptInclude = False
    Produto.Text = sProdutoEnxuto
    Produto.PromptInclude = True

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    iAlterado = 0

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209091)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_GA(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_GA

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Descrição")

    'campos de edição do grid
    objGridInt.colCampo.Add (GAData.Name)
    objGridInt.colCampo.Add (GAValor.Name)
    objGridInt.colCampo.Add (GADesc.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_GAData_Col = 1
    iGrid_GAValor_Col = 2
    iGrid_GADesc_Col = 3

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridGA

    'Linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 7
    
    'Largura da primeira coluna
    objGridInt.objGrid.ColWidth(0) = 300
    
    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_GA = SUCESSO

    Exit Function

Erro_Inicializa_Grid_GA:

    Inicializa_Grid_GA = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209092)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_DN(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_DN

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("NS")
    objGridInt.colColuna.Add ("NU")
    objGridInt.colColuna.Add ("Recebido")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Observação")

    'campos de edição do grid
    objGridInt.colCampo.Add (DNNS.Name)
    objGridInt.colCampo.Add (DNNU.Name)
    objGridInt.colCampo.Add (DNRecebido.Name)
    objGridInt.colCampo.Add (DNDesc.Name)
    objGridInt.colCampo.Add (DNObs.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_DNNS_Col = 1
    iGrid_DNNU_Col = 2
    iGrid_DNRecebido_Col = 3
    iGrid_DNDesc_Col = 4
    iGrid_DNObs_Col = 5

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridDN

    'Linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 12
    
    'Largura da primeira coluna
    objGridInt.objGrid.ColWidth(0) = 300
    
    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_DN = SUCESSO

    Exit Function

Erro_Inicializa_Grid_DN:

    Inicializa_Grid_DN = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209093)

    End Select

    Exit Function

End Function

Private Sub Trata_SrvSol_Click()

Dim lErro As Long
Dim iLinha As Integer
Dim objOcrCaso As New ClassTRVOcrCasos
Dim sVlrSolUS As String, sVlrSolRS As String

On Error GoTo Erro_Trata_SrvSol_Click
    
    iLinha = GridSrv.Row
    
    If iLinha > 0 And iLinha <= objGridSrv.iLinhasExistentes Then
    
        If StrParaInt(GridSrv.TextMatrix(iLinha, iGrid_SrvSol_Col)) = DESMARCADO Then
           
            GridSrv.TextMatrix(iLinha, iGrid_SrvVlrSolRS_Col) = ""
            GridSrv.TextMatrix(iLinha, iGrid_SrvVlrSolUS_Col) = ""
            GridSrv.TextMatrix(iLinha, iGrid_SrvVlrAutoUS_Col) = ""
            GridSrv.TextMatrix(iLinha, iGrid_SrvVlrAutoRS_Col) = ""
            GridSrv.TextMatrix(iLinha, iGrid_SrvAuto_Col) = DESMARCADO
            
        Else
        
            Call Soma_Coluna_Grid(objGridOF, iGrid_OFValorRS_Col, OFTotalRS, False)
            Call Soma_Coluna_Grid(objGridOF, iGrid_OFValorUS_Col, OFTotalUS, False)
            
            sVlrSolRS = OFTotalRS.Caption
            sVlrSolUS = OFTotalUS.Caption
            
            Call Soma_Coluna_Grid(objGridOF, iGrid_OFValorRS_Col, OFTotalRS, False, iGrid_OFConsiderar_Col)
            Call Soma_Coluna_Grid(objGridOF, iGrid_OFValorUS_Col, OFTotalUS, False, iGrid_OFConsiderar_Col)

            GridSrv.TextMatrix(iLinha, iGrid_SrvVlrSolRS_Col) = sVlrSolRS
            GridSrv.TextMatrix(iLinha, iGrid_SrvVlrSolUS_Col) = sVlrSolUS
        End If
        
    End If
    
    Call Grid_Refresh_Checkbox(objGridSrv)

    Call Totaliza_Valores_Srv
    
    lErro = Move_Srv_Memoria(objOcrCaso)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Move_Docs_Memoria(objOcrCaso)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("TRVOcrCasosDocs_Carrega", objOcrCaso)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Traz_Docs_Tela(objOcrCaso)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_Trata_SrvSol_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209135)

    End Select

    Exit Sub
    
End Sub

Function Traz_Docs_Tela(objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objOcrCasosDocs As ClassTRVOcrCasosDocs

On Error GoTo Erro_Traz_Docs_Tela

    Call Grid_Limpa(objGridDN)

    iLinha = 0
    For Each objOcrCasosDocs In objOcrCasos.colDocs
    
        iLinha = iLinha + 1

        GridDN.TextMatrix(iLinha, iGrid_DNDesc_Col) = objOcrCasosDocs.lCodigoDoc & SEPARADOR & objOcrCasosDocs.sDescricao
        GridDN.TextMatrix(iLinha, iGrid_DNObs_Col) = objOcrCasosDocs.sObservacao

        GridDN.TextMatrix(iLinha, iGrid_DNNS_Col) = CStr(objOcrCasosDocs.iNecessSist)
        GridDN.TextMatrix(iLinha, iGrid_DNNU_Col) = CStr(objOcrCasosDocs.iNecessUsu)
        GridDN.TextMatrix(iLinha, iGrid_DNRecebido_Col) = CStr(objOcrCasosDocs.iRecebido)
    
    Next
    
    objGridDN.iLinhasExistentes = objOcrCasos.colDocs.Count
    
    Call Grid_Refresh_Checkbox(objGridDN)
    
    Traz_Docs_Tela = SUCESSO

    Exit Function

Erro_Traz_Docs_Tela:

    Traz_Docs_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209136)

    End Select

    Exit Function

End Function

Function Traz_GAdv_Tela(objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objOcrCasosGAdv As ClassTRVOcrCasosGastosAdv

On Error GoTo Erro_Traz_GAdv_Tela

    Call Grid_Limpa(objGridGA)

    iLinha = 0
    For Each objOcrCasosGAdv In objOcrCasos.colGastosAdvs
    
        iLinha = iLinha + 1

        GridGA.TextMatrix(iLinha, iGrid_GADesc_Col) = objOcrCasosGAdv.sDescricao
        GridGA.TextMatrix(iLinha, iGrid_GAValor_Col) = Format(objOcrCasosGAdv.dValor, "STANDARD")
        GridGA.TextMatrix(iLinha, iGrid_GAData_Col) = Format(objOcrCasosGAdv.dtData, "dd/mm/yyyy")
    
    Next
    
    objGridGA.iLinhasExistentes = objOcrCasos.colGastosAdvs.Count
    
    Call Soma_Coluna_Grid(objGridGA, iGrid_GAValor_Col, GATotal, False)
    
    Traz_GAdv_Tela = SUCESSO

    Exit Function

Erro_Traz_GAdv_Tela:

    Traz_GAdv_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209137)

    End Select

    Exit Function

End Function

Private Sub GridGA_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridGA, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridGA, iAlterado)
    End If

End Sub

Private Sub GridGA_GotFocus()
    Call Grid_Recebe_Foco(objGridGA)
End Sub

Private Sub GridGA_EnterCell()
    Call Grid_Entrada_Celula(objGridGA, iAlterado)
End Sub

Private Sub GridGA_LeaveCell()
    Call Saida_Celula(objGridGA)
End Sub

Private Sub GridGA_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridGA, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridGA, iAlterado)
    End If

End Sub

Private Sub GridGA_RowColChange()
    Call Grid_RowColChange(objGridGA)
End Sub

Private Sub GridGA_Scroll()
    Call Grid_Scroll(objGridGA)
End Sub

Private Sub GridGA_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridGA)

End Sub

Private Sub GridGA_LostFocus()
    Call Grid_Libera_Foco(objGridGA)
End Sub

Private Sub GridDN_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridDN, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDN, iAlterado)
    End If

End Sub

Private Sub GridDN_GotFocus()
    Call Grid_Recebe_Foco(objGridDN)
End Sub

Private Sub GridDN_EnterCell()
    Call Grid_Entrada_Celula(objGridDN, iAlterado)
End Sub

Private Sub GridDN_LeaveCell()
    Call Saida_Celula(objGridDN)
End Sub

Private Sub GridDN_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridDN, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDN, iAlterado)
    End If

End Sub

Private Sub GridDN_RowColChange()
    Call Grid_RowColChange(objGridDN)
    Call Exibe_CampoDet_Grid(objGridDN, iGrid_DNDesc_Col, DNDescDet)
    Call Exibe_CampoDet_Grid(objGridDN, iGrid_DNObs_Col, DNObsDet)
End Sub

Private Sub GridDN_Scroll()
    Call Grid_Scroll(objGridDN)
End Sub

Private Sub GridDN_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridDN)

End Sub

Private Sub GridDN_LostFocus()
    Call Grid_Libera_Foco(objGridDN)
End Sub

Private Sub GAValor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub GAValor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridGA)
End Sub

Public Sub GAValor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridGA)
End Sub

Public Sub GAValor_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridGA.objControle = GAValor
    lErro = Grid_Campo_Libera_Foco(objGridGA)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub GAData_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub GAData_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridGA)
End Sub

Public Sub GAData_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridGA)
End Sub

Public Sub GAData_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridGA.objControle = GAData
    lErro = Grid_Campo_Libera_Foco(objGridGA)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub GADesc_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub GADesc_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridGA)
End Sub

Public Sub GADesc_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridGA)
End Sub

Public Sub GADesc_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridGA.objControle = GADesc
    lErro = Grid_Campo_Libera_Foco(objGridGA)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub DNObs_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DNObs_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDN)
End Sub

Public Sub DNObs_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDN)
End Sub

Public Sub DNObs_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridDN.objControle = DNObs
    lErro = Grid_Campo_Libera_Foco(objGridDN)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub DNNS_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DNNS_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDN)
End Sub

Public Sub DNNS_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDN)
End Sub

Public Sub DNNS_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridDN.objControle = DNNS
    lErro = Grid_Campo_Libera_Foco(objGridDN)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub DNNU_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DNNU_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDN)
End Sub

Public Sub DNNU_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDN)
End Sub

Public Sub DNNU_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridDN.objControle = DNNU
    lErro = Grid_Campo_Libera_Foco(objGridDN)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub DNRecebido_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DNRecebido_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridDN)
End Sub

Public Sub DNRecebido_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDN)
End Sub

Public Sub DNRecebido_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridDN.objControle = DNRecebido
    lErro = Grid_Campo_Libera_Foco(objGridDN)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Function Move_Docs_Memoria(objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim iLinha As Integer, bValido As Boolean
Dim objOcrCasosDocs As ClassTRVOcrCasosDocs

On Error GoTo Erro_Move_Docs_Memoria

    For iLinha = 1 To objGridDN.iLinhasExistentes
    
        Set objOcrCasosDocs = New ClassTRVOcrCasosDocs
    
        objOcrCasosDocs.sObservacao = GridDN.TextMatrix(iLinha, iGrid_DNObs_Col)
        objOcrCasosDocs.iNecessSist = StrParaInt(GridDN.TextMatrix(iLinha, iGrid_DNNS_Col))
        objOcrCasosDocs.iNecessUsu = StrParaInt(GridDN.TextMatrix(iLinha, iGrid_DNNU_Col))
        objOcrCasosDocs.iRecebido = StrParaInt(GridDN.TextMatrix(iLinha, iGrid_DNRecebido_Col))
        objOcrCasosDocs.lCodigoDoc = LCodigo_Extrai(GridDN.TextMatrix(iLinha, iGrid_DNDesc_Col))
        
        objOcrCasosDocs.iSeq = iLinha
        
        bValido = False
        If objOcrCasosDocs.iNecessSist = MARCADO Or objOcrCasosDocs.iNecessUsu = MARCADO Or objOcrCasosDocs.iRecebido = MARCADO Or Len(Trim(objOcrCasosDocs.sObservacao)) > 0 Then bValido = True
        
        If bValido Then objOcrCasos.colDocs.Add objOcrCasosDocs

    Next

    Move_Docs_Memoria = SUCESSO

    Exit Function

Erro_Move_Docs_Memoria:

    Move_Docs_Memoria = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209138)

    End Select

    Exit Function
    
End Function

Function Move_GAdv_Memoria(objOcrCasos As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objOcrCasosGAdv As ClassTRVOcrCasosGastosAdv

On Error GoTo Erro_Move_GAdv_Memoria

    For iLinha = 1 To objGridGA.iLinhasExistentes
    
        Set objOcrCasosGAdv = New ClassTRVOcrCasosGastosAdv
    
        objOcrCasosGAdv.sDescricao = GridGA.TextMatrix(iLinha, iGrid_GADesc_Col)
        objOcrCasosGAdv.dtData = StrParaDate(GridGA.TextMatrix(iLinha, iGrid_GAData_Col))
        objOcrCasosGAdv.dValor = StrParaDbl(GridGA.TextMatrix(iLinha, iGrid_GAValor_Col))
        
        objOcrCasosGAdv.iSeq = iLinha
               
        objOcrCasos.colGastosAdvs.Add objOcrCasosGAdv

    Next

    Move_GAdv_Memoria = SUCESSO

    Exit Function

Erro_Move_GAdv_Memoria:

    Move_GAdv_Memoria = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209139)

    End Select

    Exit Function
    
End Function

Private Sub NumProcesso_GotFocus()
    Call MaskEdBox_TrataGotFocus(NumProcesso, iAlterado)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is ClienteVou Then
            Call LabelCliente_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call ProdutoLabel_Click
        ElseIf Me.ActiveControl Is NumVou Then
            Call LabelNumVou_Click
        ElseIf Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        End If
    
    End If
    
End Sub

Private Sub DanoMoral_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DanoMoral_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DanoMoral_Validate

    'Veifica se DanoMoral está preenchida
    If Len(Trim(DanoMoral.Text)) <> 0 Then

       'Critica a DanoMoral
       lErro = Valor_Positivo_Critica(DanoMoral.Text)
       If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If

    Exit Sub

Erro_DanoMoral_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208645)

    End Select

    Exit Sub
    
End Sub

Private Sub DanoMaterial_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DanoMaterial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DanoMaterial_Validate

    'Veifica se DanoMaterial está preenchida
    If Len(Trim(DanoMaterial.Text)) <> 0 Then

       'Critica a DanoMaterial
       lErro = Valor_Positivo_Critica(DanoMaterial.Text)
       If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If

    Exit Sub

Erro_DanoMaterial_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208645)

    End Select

    Exit Sub
    
End Sub

Private Sub DataIniProcesso_Change()
    Call Data_Change
End Sub

Private Sub DataIniProcesso_GotFocus()
    Call Data_GotFocus
End Sub

Private Sub DataIniProcesso_Validate(Cancel As Boolean)
    Call Data_Validate(Cancel)
End Sub

Private Sub Trata_Reembolso()

Dim iLinhaSrv As Integer
Dim sVlrSolRS As String, sVlrSolUS As String
Dim vbResult As VbMsgBoxResult

On Error GoTo Erro_Trata_Reembolso

    If Not gbTrazendoDados Then
        
        Call Soma_Coluna_Grid(objGridOF, iGrid_OFValorRS_Col, OFTotalRS, False)
        Call Soma_Coluna_Grid(objGridOF, iGrid_OFValorUS_Col, OFTotalUS, False)
        
        sVlrSolRS = OFTotalRS.Caption
        sVlrSolUS = OFTotalUS.Caption
        
        Call Soma_Coluna_Grid(objGridOF, iGrid_OFValorRS_Col, OFTotalRS, False, iGrid_OFConsiderar_Col)
        Call Soma_Coluna_Grid(objGridOF, iGrid_OFValorUS_Col, OFTotalUS, False, iGrid_OFConsiderar_Col)
    
        If iAssistCalcValorAuto = MARCADO Then
            
            For iLinhaSrv = 1 To objGridSrv.iLinhasExistentes
                If StrParaInt(GridSrv.TextMatrix(iLinhaSrv, iGrid_SrvSol_Col)) = MARCADO Then
                
                    If iValorSolAlterado = REGISTRO_ALTERADO Then
                        If iTipoProc = TIPO_PROC_PERGUNTAR Then
                            vbResult = Rotina_Aviso(vbYesNo, "AVISO_TRV_OCR_ALTERACAO_VALOR_SOL")
                            If vbResult = vbNo Then
                                iTipoProc = TIPO_PROC_NAO_PERGUNTAR_E_NAO_ALTERAR
                            Else
                                iTipoProc = TIPO_PROC_NAO_PERGUNTAR_E_ALTERAR
                            End If
                        End If
                    End If
                    If iTipoProc = TIPO_PROC_NAO_PERGUNTAR_E_NAO_ALTERAR Then gError ERRO_SEM_MENSAGEM
                    
                    GridSrv.TextMatrix(iLinhaSrv, iGrid_SrvVlrSolRS_Col) = sVlrSolRS
                    GridSrv.TextMatrix(iLinhaSrv, iGrid_SrvVlrSolUS_Col) = sVlrSolUS
                
                    If StrParaInt(GridSrv.TextMatrix(iLinhaSrv, iGrid_SrvAuto_Col)) = MARCADO Then
                    
                        If iValorAutoAlterado = REGISTRO_ALTERADO Then
                            If iTipoProc = TIPO_PROC_PERGUNTAR Then
                                vbResult = Rotina_Aviso(vbYesNo, "AVISO_TRV_OCR_ALTERACAO_VALOR_AUTO")
                                If vbResult = vbNo Then
                                    iTipoProc = TIPO_PROC_NAO_PERGUNTAR_E_NAO_ALTERAR
                                Else
                                    iTipoProc = TIPO_PROC_NAO_PERGUNTAR_E_ALTERAR
                                End If
                            End If
                        End If
                        If iTipoProc = TIPO_PROC_NAO_PERGUNTAR_E_NAO_ALTERAR Then gError ERRO_SEM_MENSAGEM
                        
                        GridSrv.TextMatrix(iLinhaSrv, iGrid_SrvVlrAutoRS_Col) = OFTotalRS.Caption
                        GridSrv.TextMatrix(iLinhaSrv, iGrid_SrvVlrAutoUS_Col) = OFTotalUS.Caption
                    
                    End If
                    Exit For
                End If
            Next
        End If
    End If
    
    Exit Sub

Erro_Trata_Reembolso:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208645)

    End Select

    Exit Sub
End Sub

Private Sub BotaoAtualizarVou_Click()

Dim lErro As Long
Dim objOcrCaso As New ClassTRVOcrCasos
Dim objVou As New ClassTRVVouchers
Dim objVouAux As New ClassTRVVoucherInfo
Dim objCidade As New ClassCidades
Dim iCodigoCidadeNovo As Integer
Dim Cancel As Boolean

On Error GoTo Erro_BotaoAtualizarVou_Click
    
    objOcrCaso.sTipVou = TipVou.Text
    objOcrCaso.sSerie = Serie.Text
    If Len(Trim(NumVou.Text)) < 10 Then objOcrCaso.lNumVou = StrParaLong(NumVou.Text)
   
    If Len(Trim(objOcrCaso.sSerie)) > 0 And Len(Trim(objOcrCaso.sTipVou)) > 0 Then
    
        lErro = Teste_Salva(Me, iAlterado)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        objOcrCaso.sCodigo = Trim(Codigo.Text)
        
        lErro = CF("TRVOcrCasos_Le", objOcrCaso)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    
        Set objVou = New ClassTRVVouchers
    
        objOcrCaso.sTipVou = TipVou.Text
        objOcrCaso.sSerie = Serie.Text
        If Len(Trim(NumVou.Text)) < 10 Then objOcrCaso.lNumVou = StrParaLong(NumVou.Text)
    
        objVou.sTipVou = objOcrCaso.sTipVou
        objVou.sSerie = objOcrCaso.sSerie
        objVou.lNumVou = objOcrCaso.lNumVou
        objVou.sTipoDoc = TRV_TIPODOC_VOU_TEXTO
    
        lErro = CF("TRVVouchers_Le", objVou)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        
        If lErro = SUCESSO Then
        
            objOcrCaso.dtDataEmissao = objVou.dtData
            objOcrCaso.iQtdPax = objVou.iPax
            objOcrCaso.lClienteVou = objVou.lClienteVou
            objOcrCaso.sTitularNome = objVou.sPassageiroNome & " " & objVou.sPassageiroSobreNome
            objOcrCaso.sNomeFavorecido = objVou.sPassageiroNome & " " & objVou.sPassageiroSobreNome
            objOcrCaso.sProduto = objVou.sProduto
        
            objVouAux.lNumVou = objVou.lNumVou
            objVouAux.sSerie = objVou.sSerie
            objVouAux.sTipo = objVou.sTipVou
            
            lErro = CF("TRVVoucherInfoSigav_Le", objVouAux)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        
            If lErro <> SUCESSO Then gError 208429
        
            objOcrCaso.dtDataIda = objVouAux.dtDataInicio
            objOcrCaso.dtDataVolta = objVouAux.dtDataTermino
            
            If objVouAux.sTipoDoc = "CPF" And Len(Trim(objVouAux.sPassageiroCGC)) = 11 Then
                objOcrCaso.sFavorecidoCGC = objVouAux.sPassageiroCGC
            End If
            
            objOcrCaso.objEndereco.sBairro = left(objVouAux.sPassageiroBairro, STRING_BAIRRO)
            objOcrCaso.objEndereco.sCEP = Replace(objVouAux.sPassageiroCEP, "-", "")
            If Len(Trim(objOcrCaso.objEndereco.sCEP)) <> 8 Then objOcrCaso.objEndereco.sCEP = ""
            objOcrCaso.objEndereco.sCidade = left(objVouAux.sPassageiroCidade, STRING_CIDADE)
            objOcrCaso.objEndereco.sContato = objVouAux.sPassageiroContato
            objOcrCaso.objEndereco.sEmail = left(objVouAux.sPassageiroEmail, STRING_EMAIL)
            objOcrCaso.objEndereco.sEndereco = left(objVouAux.sPassageiroEndereco, STRING_ENDERECO)
            objOcrCaso.objEndereco.sTelefone1 = left(objVouAux.sPassageiroTelefone1, STRING_TELEFONE)
            objOcrCaso.objEndereco.sTelefone2 = left(objVouAux.sPassageiroTelefone2, STRING_TELEFONE)
            objOcrCaso.objEndereco.sSiglaEstado = Trim(objVouAux.sPassageiroUF)
            If Len(Trim(objOcrCaso.objEndereco.sSiglaEstado)) <> 2 Then objOcrCaso.objEndereco.sSiglaEstado = ""
            
            If Len(Trim(objOcrCaso.objEndereco.sCidade)) > 0 Then
                
                objCidade.sDescricao = objOcrCaso.objEndereco.sCidade
            
                lErro = CF("Cidade_Le_Nome", objCidade)
                If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError ERRO_SEM_MENSAGEM
            
                If lErro <> SUCESSO Then
                                    
                    lErro = CF("Config_Obter_Inteiro_Automatico", "FATConfig", "NUM_PROX_CIDADECADASTRO", "Cidades", "Codigo", iCodigoCidadeNovo)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    Set objCidade = New ClassCidades
                
                    objCidade.iCodigo = iCodigoCidadeNovo
                    objCidade.sDescricao = objOcrCaso.objEndereco.sCidade
                    
                    lErro = CF("Cidade_Grava", objCidade)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                End If
                
            End If
            
        End If
        
        lErro = Traz_TRVOcrCasos_Tela(objOcrCaso, True)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    Exit Sub

Erro_BotaoAtualizarVou_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 208429
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVVOUINFOSEIGAV_INEXISTENTE", gErr, objVouAux.sTipo & objVouAux.sSerie & CStr(objVouAux.lNumVou))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208579)

    End Select

    Exit Sub
    
End Sub

Private Sub SrvTotalSegTrvRS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub SrvTotalSegTrvRS_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_SrvTotalSegTrvRS_Validate

    'Veifica se SrvTotalSegTrvRS está preenchida
    If Len(Trim(SrvTotalSegTrvRS.Text)) <> 0 Then

       'Critica a SrvTotalSegTrvRS
       lErro = Valor_Positivo_Critica(SrvTotalSegTrvRS.Text)
       If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If

    If StrParaDbl(SrvTotalSegRS.Caption) < StrParaDbl(SrvTotalSegTrvRS.Text) Then gError 209209
    
    SrvTotalSegSegRS.Caption = Format(StrParaDbl(SrvTotalSegRS.Caption) - StrParaDbl(SrvTotalSegTrvRS.Text), "STANDARD")

    Exit Sub

Erro_SrvTotalSegTrvRS_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 209209
            Call Rotina_Erro(vbOKOnly, "ERRO_TRV_SEGURO_RESPTRV_MAIOR_TOTAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208645)

    End Select

    Exit Sub
    
End Sub

