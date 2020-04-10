VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl NFPag_ConsultaOcx 
   ClientHeight    =   5820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9270
   ScaleHeight     =   5820
   ScaleWidth      =   9270
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4065
      Index           =   1
      Left            =   135
      TabIndex        =   4
      Top             =   1605
      Width           =   9045
      Begin VB.ComboBox Etapa 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5925
         Style           =   2  'Dropdown List
         TabIndex        =   90
         Top             =   1215
         Width           =   2535
      End
      Begin VB.CommandButton BotaoProjetos 
         Caption         =   "..."
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
         Left            =   3945
         TabIndex        =   86
         Top             =   1230
         Width           =   495
      End
      Begin VB.Frame Frame2 
         Caption         =   "Valores"
         Height          =   2370
         Left            =   210
         TabIndex        =   24
         Top             =   1590
         Width           =   8730
         Begin VB.CheckBox INSSRetido 
            Caption         =   "Retido"
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
            Height          =   195
            Left            =   2250
            TabIndex        =   96
            Top             =   1890
            Width           =   930
         End
         Begin VB.Frame Frame7 
            Caption         =   "Retenções"
            Height          =   1005
            Left            =   3435
            TabIndex        =   92
            Top             =   1185
            Width           =   5145
            Begin VB.Label ISSRetido 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   4050
               TabIndex        =   106
               Top             =   270
               Width           =   990
            End
            Begin VB.Label Label25 
               Caption         =   "ISS:"
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
               Left            =   3660
               TabIndex        =   105
               Top             =   315
               Width           =   345
            End
            Begin VB.Label CSLLRetido 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   2520
               TabIndex        =   104
               Top             =   630
               Width           =   1005
            End
            Begin VB.Label Label23 
               Caption         =   "CSLL:"
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
               Left            =   1950
               TabIndex        =   103
               Top             =   675
               Width           =   615
            End
            Begin VB.Label PISRetido 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   2505
               TabIndex        =   102
               Top             =   240
               Width           =   990
            End
            Begin VB.Label Label17 
               Caption         =   "PIS:"
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
               Left            =   2115
               TabIndex        =   101
               Top             =   285
               Width           =   345
            End
            Begin VB.Label COFINSRetido 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   855
               TabIndex        =   100
               Top             =   615
               Width           =   930
            End
            Begin VB.Label Label14 
               Caption         =   "COFINS:"
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
               Left            =   60
               TabIndex        =   99
               Top             =   660
               Width           =   750
            End
            Begin VB.Label ValorIRRF 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   870
               TabIndex        =   95
               Top             =   240
               Width           =   930
            End
            Begin VB.Label Label16 
               Caption         =   "IR:"
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
               TabIndex        =   94
               Top             =   285
               Width           =   270
            End
         End
         Begin VB.Frame Frame4 
            Height          =   600
            Left            =   120
            TabIndex        =   28
            Top             =   195
            Width           =   6195
            Begin VB.CheckBox CreditoICMS 
               Caption         =   "Crédito"
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
               Height          =   300
               Left            =   5175
               TabIndex        =   6
               Top             =   210
               Width           =   930
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "ICMS Subst:"
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
               Left            =   2640
               TabIndex        =   30
               Top             =   270
               Width           =   1065
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "ICMS:"
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
               TabIndex        =   31
               Top             =   270
               Width           =   525
            End
            Begin VB.Label ValorICMS 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   750
               TabIndex        =   32
               Top             =   210
               Width           =   1215
            End
            Begin VB.Label ValorICMSSubst 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   3750
               TabIndex        =   33
               Top             =   210
               Width           =   1215
            End
         End
         Begin VB.Frame Frame3 
            Height          =   570
            Index           =   1
            Left            =   165
            TabIndex        =   25
            Top             =   1185
            Width           =   3120
            Begin VB.CheckBox CreditoIPI 
               Caption         =   "Crédito"
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
               Height          =   195
               Left            =   2100
               TabIndex        =   7
               Top             =   255
               Width           =   930
            End
            Begin VB.Label Label3 
               Caption         =   "IPI:"
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
               Left            =   300
               TabIndex        =   34
               Top             =   240
               Width           =   315
            End
            Begin VB.Label ValorIPI 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   705
               TabIndex        =   35
               Top             =   195
               Width           =   1215
            End
         End
         Begin VB.Label ValorINSS 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   855
            TabIndex        =   98
            Top             =   1830
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "INSS:"
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
            Left            =   255
            TabIndex        =   97
            Top             =   1875
            Width           =   525
         End
         Begin VB.Label Label22 
            Caption         =   "Produtos:"
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
            Left            =   6465
            TabIndex        =   36
            Top             =   435
            Width           =   825
         End
         Begin VB.Label Label20 
            Caption         =   "Frete:"
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
            Left            =   300
            TabIndex        =   37
            Top             =   900
            Width           =   525
         End
         Begin VB.Label Label18 
            Caption         =   "Outras Despesas:"
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
            Left            =   5760
            TabIndex        =   38
            Top             =   900
            Width           =   1530
         End
         Begin VB.Label Label19 
            Caption         =   "Seguro:"
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
            Left            =   3150
            TabIndex        =   39
            Top             =   900
            Width           =   675
         End
         Begin VB.Label ValorProdutos 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7350
            TabIndex        =   40
            Top             =   405
            Width           =   1215
         End
         Begin VB.Label ValorFrete 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   870
            TabIndex        =   41
            Top             =   870
            Width           =   1215
         End
         Begin VB.Label ValorSeguro 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3870
            TabIndex        =   42
            Top             =   870
            Width           =   1215
         End
         Begin VB.Label OutrasDespesas 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7350
            TabIndex        =   43
            Top             =   870
            Width           =   1215
         End
      End
      Begin VB.CommandButton BotaoDocOriginal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   6600
         Picture         =   "NFPag_ConsultaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Consulta de notas fiscais"
         Top             =   60
         Width           =   1815
      End
      Begin MSMask.MaskEdBox Projeto 
         Height          =   300
         Left            =   1425
         TabIndex        =   91
         Top             =   1230
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Etapa:"
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
         Height          =   180
         Index           =   62
         Left            =   5340
         TabIndex        =   88
         Top             =   1260
         Width           =   570
      End
      Begin VB.Label LabelProjeto 
         AutoSize        =   -1  'True
         Caption         =   "Projeto:"
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
         Left            =   660
         TabIndex        =   87
         Top             =   1260
         Width           =   675
      End
      Begin VB.Label Label7 
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   570
         TabIndex        =   85
         Top             =   870
         Width           =   825
      End
      Begin VB.Label Historico 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1440
         TabIndex        =   84
         Top             =   825
         Width           =   6990
      End
      Begin VB.Label DataVencimento 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4395
         TabIndex        =   44
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Vencimento:"
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
         Left            =   3240
         TabIndex        =   45
         Top             =   330
         Width           =   1065
      End
      Begin VB.Label ValorTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1440
         TabIndex        =   46
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
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
         Left            =   885
         TabIndex        =   47
         Top             =   360
         Width           =   510
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   4065
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   1590
      Visible         =   0   'False
      Width           =   9075
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4920
         TabIndex        =   89
         Tag             =   "1"
         Top             =   1920
         Width           =   870
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
         Height          =   375
         Left            =   7560
         TabIndex        =   11
         Top             =   255
         Width           =   1245
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   930
         Left            =   105
         TabIndex        =   26
         Top             =   3090
         Width           =   5895
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   48
            Top             =   585
            Visible         =   0   'False
            Width           =   3720
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   49
            Top             =   255
            Width           =   3720
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
            Left            =   1125
            TabIndex        =   50
            Top             =   270
            Width           =   570
         End
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
            Left            =   240
            TabIndex        =   51
            Top             =   600
            Visible         =   0   'False
            Width           =   1440
         End
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   4500
         MaxLength       =   150
         TabIndex        =   16
         Top             =   1260
         Width           =   3495
      End
      Begin VB.CheckBox CTBAglutina 
         Enabled         =   0   'False
         Height          =   210
         Left            =   8040
         TabIndex        =   18
         Top             =   1320
         Width           =   795
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4140
         TabIndex        =   17
         Top             =   1260
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
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
      Begin MSMask.MaskEdBox CTBConta 
         Bindings        =   "NFPag_ConsultaOcx.ctx":2F16
         Height          =   225
         Left            =   60
         TabIndex        =   12
         Top             =   1230
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBDebito 
         Height          =   225
         Left            =   2970
         TabIndex        =   15
         Top             =   1230
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
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
      Begin MSMask.MaskEdBox CTBCredito 
         Height          =   225
         Left            =   1800
         TabIndex        =   14
         Top             =   1200
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
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
      Begin MSMask.MaskEdBox CTBCcl 
         Height          =   225
         Left            =   1050
         TabIndex        =   13
         Top             =   1200
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         Enabled         =   0   'False
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
      Begin MSFlexGridLib.MSFlexGrid CTBGridContabil 
         Height          =   1830
         Left            =   30
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   945
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   3228
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label CTBLote 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5970
         TabIndex        =   52
         Top             =   75
         Width           =   615
      End
      Begin VB.Label CTBDocumento 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3705
         TabIndex        =   53
         Top             =   75
         Width           =   705
      End
      Begin VB.Label CTBDataContabil 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   750
         TabIndex        =   54
         Top             =   442
         Width           =   1095
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
         Left            =   5460
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   55
         Top             =   135
         Width           =   450
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
         Left            =   2640
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   56
         Top             =   135
         Width           =   1035
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
         Left            =   225
         TabIndex        =   57
         Top             =   495
         Width           =   480
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   58
         Top             =   2790
         Width           =   1155
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   59
         Top             =   2790
         Width           =   1155
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
         TabIndex        =   60
         Top             =   2805
         Width           =   615
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
         TabIndex        =   61
         Top             =   765
         Width           =   1140
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
         Left            =   2775
         TabIndex        =   62
         Top             =   465
         Width           =   870
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3705
         TabIndex        =   63
         Top             =   450
         Width           =   1185
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5970
         TabIndex        =   64
         Top             =   450
         Width           =   1185
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
         Left            =   5160
         TabIndex        =   65
         Top             =   465
         Width           =   735
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   66
         Top             =   90
         Width           =   1530
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
         TabIndex        =   67
         Top             =   135
         Width           =   720
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4065
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1590
      Visible         =   0   'False
      Width           =   9045
      Begin VB.CommandButton BotaoConsultaFatura 
         Caption         =   "Consultar Fatura ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6840
         TabIndex        =   9
         Top             =   3450
         Width           =   1965
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total:"
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
         Left            =   3795
         TabIndex        =   68
         Top             =   1770
         Width           =   1005
      End
      Begin VB.Label ValorTotalFatura 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4905
         TabIndex        =   69
         Top             =   1717
         Width           =   1215
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Saldo:"
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
         TabIndex        =   70
         Top             =   1770
         Width           =   555
      End
      Begin VB.Label SaldoFatura 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1215
         TabIndex        =   71
         Top             =   1717
         Width           =   1215
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4185
         TabIndex        =   72
         Top             =   1088
         Width           =   615
      End
      Begin VB.Label StatusFatura 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4905
         TabIndex        =   73
         Top             =   1038
         Width           =   1215
      End
      Begin VB.Label Label13 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   420
         TabIndex        =   74
         Top             =   1088
         Width           =   765
      End
      Begin VB.Label EmissaoFatura 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1215
         TabIndex        =   75
         Top             =   1038
         Width           =   1215
      End
      Begin VB.Label Label8 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4080
         TabIndex        =   76
         Top             =   413
         Width           =   720
      End
      Begin VB.Label NumeroFatura 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4905
         TabIndex        =   77
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
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
         Left            =   735
         TabIndex        =   78
         Top             =   413
         Width           =   450
      End
      Begin VB.Label TipoFatura 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1215
         TabIndex        =   79
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.ComboBox Filial 
      Height          =   315
      Left            =   4560
      TabIndex        =   1
      Top             =   210
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   6540
      ScaleHeight     =   795
      ScaleWidth      =   2565
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   150
      Width           =   2625
      Begin VB.CommandButton BotaoConsulta 
         Height          =   675
         Left            =   90
         Picture         =   "NFPag_ConsultaOcx.ctx":2F21
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   60
         Width           =   1275
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   675
         Left            =   1455
         Picture         =   "NFPag_ConsultaOcx.ctx":4CE3
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   480
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   675
         Left            =   2025
         Picture         =   "NFPag_ConsultaOcx.ctx":5215
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   480
      End
   End
   Begin MSMask.MaskEdBox NumeroNota 
      Height          =   300
      Left            =   1155
      TabIndex        =   2
      Top             =   690
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   9
      Mask            =   "999999999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Fornecedor 
      Height          =   300
      Left            =   1155
      TabIndex        =   0
      Top             =   210
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   "_"
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   4455
      Left            =   90
      TabIndex        =   27
      Top             =   1260
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   7858
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Nota Fiscal"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fatura"
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
   Begin MSComCtl2.UpDown UpDownEmissao 
      Height          =   300
      Left            =   5670
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   690
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataEmissao 
      Height          =   300
      Left            =   4560
      TabIndex        =   3
      Top             =   690
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   " Filial:"
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
      Left            =   4005
      TabIndex        =   80
      Top             =   270
      Width           =   525
   End
   Begin VB.Label NumeroLabel 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   405
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   81
      Top             =   750
      Width           =   720
   End
   Begin VB.Label FornecedorLabel 
      AutoSize        =   -1  'True
      Caption         =   "Fornecedor:"
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
      Left            =   90
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   82
      Top             =   270
      Width           =   1005
   End
   Begin VB.Label Label4 
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
      Left            =   3765
      TabIndex        =   83
      Top             =   750
      Width           =   765
   End
   Begin VB.Line Line1 
      X1              =   30
      X2              =   9450
      Y1              =   1170
      Y2              =   1170
   End
End
Attribute VB_Name = "NFPag_ConsultaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTNFPag_Consulta
Attribute objCT.VB_VarHelpID = -1

Private Sub UserControl_Initialize()
    Set objCT = New CTNFPag_Consulta
    Set objCT.objUserControl = Me
End Sub

Private Sub BotaoConsulta_Click()
     Call objCT.BotaoConsulta_Click
End Sub

Private Sub BotaoConsultaFatura_Click()
     Call objCT.BotaoConsultaFatura_Click
End Sub

Private Sub BotaoDocOriginal_Click()
     Call objCT.BotaoDocOriginal_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub CTBGridContabil_Click()
     Call objCT.CTBGridContabil_Click
End Sub

Private Sub DataEmissao_GotFocus()
     Call objCT.DataEmissao_GotFocus
End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)
     Call objCT.DataEmissao_Validate(Cancel)
End Sub

Private Sub UpDownEmissao_DownClick()
     Call objCT.UpDownEmissao_DownClick
End Sub

Private Sub UpDownEmissao_UpClick()
     Call objCT.UpDownEmissao_UpClick
End Sub

Private Sub Filial_Validate(Cancel As Boolean)
     Call objCT.Filial_Validate(Cancel)
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub FornecedorLabel_Click()
     Call objCT.FornecedorLabel_Click
End Sub

Private Sub NumeroLabel_Click()
     Call objCT.NumeroLabel_Click
End Sub

Function Trata_Parametros(Optional objNFsPag As ClassNFsPag) As Long
     Trata_Parametros = objCT.Trata_Parametros(objNFsPag)
End Function

Private Sub Fornecedor_Validate(Cancel As Boolean)
     Call objCT.Fornecedor_Validate(Cancel)
End Sub

Private Sub NumeroNota_GotFocus()
     Call objCT.NumeroNota_GotFocus
End Sub

Private Sub NumeroNota_Validate(Cancel As Boolean)
     Call objCT.NumeroNota_Validate(Cancel)
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub Opcao_Click()
     Call objCT.Opcao_Click
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Private Sub CTBBotaoImprimir_Click()
     Call objCT.CTBBotaoImprimir_Click
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
Private Sub ValorICMS_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorICMS, Source, X, Y)
End Sub
Private Sub ValorICMS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorICMS, Button, Shift, X, Y)
End Sub
Private Sub ValorICMSSubst_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorICMSSubst, Source, X, Y)
End Sub
Private Sub ValorICMSSubst_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorICMSSubst, Button, Shift, X, Y)
End Sub
Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub
Private Sub ValorIPI_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorIPI, Source, X, Y)
End Sub
Private Sub ValorIPI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorIPI, Button, Shift, X, Y)
End Sub
Private Sub Label22_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label22, Source, X, Y)
End Sub
Private Sub Label22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label22, Button, Shift, X, Y)
End Sub
Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub
Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
End Sub
Private Sub Label18_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label18, Source, X, Y)
End Sub
Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label18, Button, Shift, X, Y)
End Sub
Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub
Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub
Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub
Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
End Sub
Private Sub ValorProdutos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorProdutos, Source, X, Y)
End Sub
Private Sub ValorProdutos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorProdutos, Button, Shift, X, Y)
End Sub
Private Sub ValorFrete_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorFrete, Source, X, Y)
End Sub
Private Sub ValorFrete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorFrete, Button, Shift, X, Y)
End Sub
Private Sub ValorSeguro_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorSeguro, Source, X, Y)
End Sub
Private Sub ValorSeguro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorSeguro, Button, Shift, X, Y)
End Sub
Private Sub OutrasDespesas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(OutrasDespesas, Source, X, Y)
End Sub
Private Sub OutrasDespesas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(OutrasDespesas, Button, Shift, X, Y)
End Sub
Private Sub ValorIRRF_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorIRRF, Source, X, Y)
End Sub
Private Sub ValorIRRF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorIRRF, Button, Shift, X, Y)
End Sub
Private Sub DataVencimento_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataVencimento, Source, X, Y)
End Sub
Private Sub DataVencimento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataVencimento, Button, Shift, X, Y)
End Sub
Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub
Private Sub ValorTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorTotal, Source, X, Y)
End Sub
Private Sub ValorTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorTotal, Button, Shift, X, Y)
End Sub
Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub
Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub
Private Sub CTBCclDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclDescricao, Source, X, Y)
End Sub
Private Sub CTBCclDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclDescricao, Button, Shift, X, Y)
End Sub
Private Sub CTBContaDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBContaDescricao, Source, X, Y)
End Sub
Private Sub CTBContaDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBContaDescricao, Button, Shift, X, Y)
End Sub
Private Sub CTBLabel7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel7, Source, X, Y)
End Sub
Private Sub CTBLabel7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel7, Button, Shift, X, Y)
End Sub
Private Sub CTBCclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclLabel, Source, X, Y)
End Sub
Private Sub CTBCclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclLabel, Button, Shift, X, Y)
End Sub
Private Sub CTBLote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLote, Source, X, Y)
End Sub
Private Sub CTBLote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLote, Button, Shift, X, Y)
End Sub
Private Sub CTBDocumento_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBDocumento, Source, X, Y)
End Sub
Private Sub CTBDocumento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBDocumento, Button, Shift, X, Y)
End Sub
Private Sub CTBDataContabil_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBDataContabil, Source, X, Y)
End Sub
Private Sub CTBDataContabil_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBDataContabil, Button, Shift, X, Y)
End Sub
Private Sub CTBLabelLote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelLote, Source, X, Y)
End Sub
Private Sub CTBLabelLote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelLote, Button, Shift, X, Y)
End Sub
Private Sub CTBLabelDoc_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelDoc, Source, X, Y)
End Sub
Private Sub CTBLabelDoc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelDoc, Button, Shift, X, Y)
End Sub
Private Sub CTBLabel8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel8, Source, X, Y)
End Sub
Private Sub CTBLabel8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel8, Button, Shift, X, Y)
End Sub
Private Sub CTBTotalCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalCredito, Source, X, Y)
End Sub
Private Sub CTBTotalCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalCredito, Button, Shift, X, Y)
End Sub
Private Sub CTBTotalDebito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalDebito, Source, X, Y)
End Sub
Private Sub CTBTotalDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalDebito, Button, Shift, X, Y)
End Sub
Private Sub CTBLabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelTotais, Source, X, Y)
End Sub
Private Sub CTBLabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelTotais, Button, Shift, X, Y)
End Sub
Private Sub CTBLabel5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel5, Source, X, Y)
End Sub
Private Sub CTBLabel5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel5, Button, Shift, X, Y)
End Sub
Private Sub CTBLabel13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel13, Source, X, Y)
End Sub
Private Sub CTBLabel13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel13, Button, Shift, X, Y)
End Sub
Private Sub CTBExercicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBExercicio, Source, X, Y)
End Sub
Private Sub CTBExercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBExercicio, Button, Shift, X, Y)
End Sub
Private Sub CTBPeriodo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBPeriodo, Source, X, Y)
End Sub
Private Sub CTBPeriodo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBPeriodo, Button, Shift, X, Y)
End Sub
Private Sub CTBLabel14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel14, Source, X, Y)
End Sub
Private Sub CTBLabel14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel14, Button, Shift, X, Y)
End Sub
Private Sub CTBOrigem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBOrigem, Source, X, Y)
End Sub
Private Sub CTBOrigem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBOrigem, Button, Shift, X, Y)
End Sub
Private Sub CTBLabel21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel21, Source, X, Y)
End Sub
Private Sub CTBLabel21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel21, Button, Shift, X, Y)
End Sub
Private Sub Label24_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label24, Source, X, Y)
End Sub
Private Sub Label24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label24, Button, Shift, X, Y)
End Sub
Private Sub ValorTotalFatura_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorTotalFatura, Source, X, Y)
End Sub
Private Sub ValorTotalFatura_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorTotalFatura, Button, Shift, X, Y)
End Sub
Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub
Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub
Private Sub SaldoFatura_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SaldoFatura, Source, X, Y)
End Sub
Private Sub SaldoFatura_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SaldoFatura, Button, Shift, X, Y)
End Sub
Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub
Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub
Private Sub StatusFatura_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(StatusFatura, Source, X, Y)
End Sub
Private Sub StatusFatura_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(StatusFatura, Button, Shift, X, Y)
End Sub
Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub
Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub
Private Sub EmissaoFatura_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(EmissaoFatura, Source, X, Y)
End Sub
Private Sub EmissaoFatura_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(EmissaoFatura, Button, Shift, X, Y)
End Sub
Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub
Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub
Private Sub NumeroFatura_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroFatura, Source, X, Y)
End Sub
Private Sub NumeroFatura_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroFatura, Button, Shift, X, Y)
End Sub
Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub
Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub
Private Sub TipoFatura_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoFatura, Source, X, Y)
End Sub
Private Sub TipoFatura_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoFatura, Button, Shift, X, Y)
End Sub
Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub
Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub
Private Sub NumeroLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroLabel, Source, X, Y)
End Sub
Private Sub NumeroLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroLabel, Button, Shift, X, Y)
End Sub
Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub
Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub
Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub
Private Sub Opcao_BeforeClick(Cancel As Integer)
     Call objCT.Opcao_BeforeClick(Cancel)
End Sub

Private Sub Fornecedor_Change()
     Call objCT.Fornecedor_Change
End Sub

Private Sub CTBGerencial_Click()
     Call objCT.CTBGerencial_Click
End Sub

Private Sub CTBGerencial_GotFocus()
     Call objCT.CTBGerencial_GotFocus
End Sub

Private Sub CTBGerencial_KeyPress(KeyAscii As Integer)
     Call objCT.CTBGerencial_KeyPress(KeyAscii)
End Sub

Private Sub CTBGerencial_Validate(Cancel As Boolean)
     Call objCT.CTBGerencial_Validate(Cancel)
End Sub

Public Function Form_Load_Ocx() As Object

    Call objCT.Form_Load_Ocx
    Set Form_Load_Ocx = Me

End Function

Public Sub Form_UnLoad(Cancel As Integer)
    If Not (objCT Is Nothing) Then
        Call objCT.Form_UnLoad(Cancel)
        If Cancel = False Then
             Set objCT.objUserControl = Nothing
             Set objCT = Nothing
        End If
    End If
End Sub

Private Sub objCT_Unload()
   RaiseEvent Unload
End Sub

Public Function Name() As String
    Name = objCT.Name
End Function

Public Sub Show()
    Call objCT.Show
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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Property Get Caption() As String
    Caption = objCT.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    objCT.Caption = New_Caption
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Call objCT.UserControl_KeyDown(KeyCode, Shift)
End Sub

