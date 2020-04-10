VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl TituloPagar_ConsultaOcx 
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9405
   KeyPreview      =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   9405
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4080
      Index           =   3
      Left            =   120
      TabIndex        =   119
      Top             =   1560
      Visible         =   0   'False
      Width           =   9075
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
         Left            =   7710
         TabIndex        =   132
         Top             =   105
         Width           =   1245
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   930
         Left            =   75
         TabIndex        =   127
         Top             =   3120
         Width           =   5895
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   131
            Top             =   555
            Visible         =   0   'False
            Width           =   3720
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   130
            Top             =   240
            Width           =   3720
         End
         Begin VB.Label CTBLabel7 
            Appearance      =   1  'Flat
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
            TabIndex        =   129
            Top             =   240
            Width           =   570
         End
         Begin VB.Label CTBCclLabel 
            Appearance      =   1  'Flat
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
            TabIndex        =   128
            Top             =   585
            Visible         =   0   'False
            Width           =   1440
         End
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   7740
         Style           =   2  'Dropdown List
         TabIndex        =   126
         Top             =   780
         Width           =   1260
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
         Height          =   375
         Left            =   6330
         TabIndex        =   125
         Top             =   105
         Width           =   1245
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
         Height          =   510
         Left            =   6330
         TabIndex        =   124
         Top             =   615
         Width           =   1245
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   123
         Top             =   2190
         Width           =   1770
      End
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4455
         TabIndex        =   122
         Top             =   2565
         Width           =   870
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
         Left            =   3465
         TabIndex        =   121
         Top             =   795
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2400
         IntegralHeight  =   0   'False
         Left            =   6330
         TabIndex        =   120
         Top             =   1500
         Visible         =   0   'False
         Width           =   2625
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4680
         TabIndex        =   133
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
      Begin MSMask.MaskEdBox CTBConta 
         Height          =   225
         Left            =   525
         TabIndex        =   134
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
         TabIndex        =   135
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
         TabIndex        =   136
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
         TabIndex        =   137
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
         TabIndex        =   138
         TabStop         =   0   'False
         Top             =   420
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   327681
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil 
         Height          =   300
         Left            =   570
         TabIndex        =   139
         Top             =   420
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
         Left            =   5580
         TabIndex        =   140
         Top             =   90
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
         TabIndex        =   141
         Top             =   75
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
         Height          =   1770
         Left            =   45
         TabIndex        =   142
         Top             =   1035
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   3122
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin MSComctlLib.TreeView CTBTvwContas 
         Height          =   2220
         Left            =   6330
         TabIndex        =   143
         Top             =   1500
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   3916
         _Version        =   327682
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView CTBTvwCcls 
         Height          =   2220
         Left            =   6330
         TabIndex        =   144
         Top             =   1500
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   3916
         _Version        =   327682
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.Label CTBLabelLote 
         Appearance      =   1  'Flat
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
         Left            =   5115
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   161
         Top             =   120
         Width           =   450
      End
      Begin VB.Label CTBLabelDoc 
         Appearance      =   1  'Flat
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
         TabIndex        =   160
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
         Left            =   45
         TabIndex        =   159
         Top             =   480
         Width           =   480
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2640
         TabIndex        =   158
         Top             =   2820
         Width           =   1155
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3885
         TabIndex        =   157
         Top             =   2820
         Width           =   1155
      End
      Begin VB.Label CTBLabelTotais 
         Appearance      =   1  'Flat
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
         Height          =   195
         Left            =   1980
         TabIndex        =   156
         Top             =   2850
         Width           =   615
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
         Left            =   7755
         TabIndex        =   155
         Top             =   555
         Width           =   690
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
         TabIndex        =   154
         Top             =   1275
         Visible         =   0   'False
         Width           =   2490
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
         TabIndex        =   153
         Top             =   1275
         Width           =   2340
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
         TabIndex        =   152
         Top             =   1275
         Visible         =   0   'False
         Width           =   1005
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
         TabIndex        =   151
         Top             =   810
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
         Left            =   1995
         TabIndex        =   150
         Top             =   450
         Width           =   870
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   149
         Top             =   435
         Width           =   1185
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   148
         Top             =   435
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
         Left            =   4230
         TabIndex        =   147
         Top             =   450
         Width           =   735
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   146
         Top             =   75
         Width           =   1530
      End
      Begin VB.Label CTBLabel21 
         Appearance      =   1  'Flat
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
         Left            =   30
         TabIndex        =   145
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4065
      Index           =   1
      Left            =   150
      TabIndex        =   13
      Top             =   1560
      Width           =   9165
      Begin VB.CommandButton BotaoContabil 
         Caption         =   "Contabilização..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6630
         TabIndex        =   110
         Top             =   255
         Width           =   2055
      End
      Begin VB.Frame Frame2 
         Caption         =   "Valores"
         Height          =   2160
         Left            =   210
         TabIndex        =   14
         Top             =   1410
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
            Height          =   285
            Left            =   7545
            TabIndex        =   36
            Top             =   1605
            Width           =   930
         End
         Begin VB.Frame Frame3 
            Height          =   570
            Index           =   1
            Left            =   165
            TabIndex        =   19
            Top             =   1425
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
               Left            =   2115
               TabIndex        =   20
               Top             =   225
               Width           =   930
            End
            Begin VB.Label ValorIPI 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   675
               TabIndex        =   107
               Top             =   180
               Width           =   1215
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
               Left            =   285
               TabIndex        =   21
               Top             =   240
               Width           =   315
            End
         End
         Begin VB.Frame Frame4 
            Height          =   600
            Left            =   120
            TabIndex        =   15
            Top             =   240
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
               TabIndex        =   16
               Top             =   210
               Width           =   930
            End
            Begin VB.Label ValorICMSSubst 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   3750
               TabIndex        =   102
               Top             =   187
               Width           =   1215
            End
            Begin VB.Label ValorICMS 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   840
               TabIndex        =   101
               Top             =   187
               Width           =   1215
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
               TabIndex        =   18
               Top             =   240
               Width           =   525
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
               TabIndex        =   17
               Top             =   240
               Width           =   1065
            End
         End
         Begin VB.Label ValorINSS 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6195
            TabIndex        =   109
            Top             =   1590
            Width           =   1215
         End
         Begin VB.Label ValorIRRF 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3915
            TabIndex        =   108
            Top             =   1590
            Width           =   1215
         End
         Begin VB.Label OutrasDespesas 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7380
            TabIndex        =   106
            Top             =   1027
            Width           =   1215
         End
         Begin VB.Label ValorSeguro 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3930
            TabIndex        =   105
            Top             =   1027
            Width           =   1215
         End
         Begin VB.Label ValorFrete 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   840
            TabIndex        =   104
            Top             =   1027
            Width           =   1215
         End
         Begin VB.Label ValorProdutos 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7350
            TabIndex        =   103
            Top             =   457
            Width           =   1215
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
            Left            =   3555
            TabIndex        =   27
            Top             =   1635
            Width           =   270
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
            TabIndex        =   26
            Top             =   1050
            Width           =   675
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
            TabIndex        =   25
            Top             =   1050
            Width           =   1530
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
            Left            =   270
            TabIndex        =   24
            Top             =   1050
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
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   6465
            TabIndex        =   23
            Top             =   480
            Width           =   825
         End
         Begin VB.Label Label2 
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
            Left            =   5565
            TabIndex        =   22
            Top             =   1635
            Width           =   495
         End
      End
      Begin VB.Label LblNumPC 
         AutoSize        =   -1  'True
         Caption         =   "Nº PC:"
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
         Left            =   810
         TabIndex        =   118
         Top             =   930
         Width           =   585
      End
      Begin VB.Label LblFilialPC 
         AutoSize        =   -1  'True
         Caption         =   "Filial PC:"
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
         Left            =   3555
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   117
         Top             =   930
         Width           =   765
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
         Left            =   630
         TabIndex        =   116
         Top             =   300
         Width           =   765
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3810
         TabIndex        =   115
         Top             =   300
         Width           =   510
      End
      Begin VB.Label DataEmissao 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1470
         TabIndex        =   114
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label ValorTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4350
         TabIndex        =   113
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label NumPC 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1440
         TabIndex        =   112
         Top             =   870
         Width           =   810
      End
      Begin VB.Label ComboFilialPC 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4350
         TabIndex        =   111
         Top             =   870
         Width           =   810
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   4050
      Index           =   3
      Left            =   150
      TabIndex        =   37
      Top             =   1590
      Width           =   9075
      Begin VB.ComboBox Sequencial 
         Height          =   315
         Left            =   4290
         Style           =   2  'Dropdown List
         TabIndex        =   98
         Top             =   225
         Width           =   825
      End
      Begin VB.Frame FramePagamento 
         Caption         =   "Dados do Pagamento"
         Height          =   1770
         Index           =   0
         Left            =   2610
         TabIndex        =   81
         Top             =   2145
         Width           =   6375
         Begin VB.Frame Frame10 
            Caption         =   "Meio Pagamento"
            Height          =   585
            Left            =   2532
            TabIndex        =   82
            Top             =   195
            Width           =   3648
            Begin VB.OptionButton TipoMeioPagto 
               Caption         =   "Dinheiro"
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
               Index           =   2
               Left            =   2532
               TabIndex        =   85
               Top             =   285
               Width           =   1035
            End
            Begin VB.OptionButton TipoMeioPagto 
               Caption         =   "Cheque"
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
               Index           =   0
               Left            =   120
               TabIndex        =   84
               Top             =   255
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton TipoMeioPagto 
               Caption         =   "Borderô"
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
               Height          =   204
               Index           =   1
               Left            =   1272
               TabIndex        =   83
               Top             =   288
               Width           =   996
            End
         End
         Begin VB.Label Label29 
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
            Height          =   255
            Left            =   375
            TabIndex        =   95
            Top             =   435
            Width           =   555
         End
         Begin VB.Label Label28 
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
            Height          =   255
            Left            =   105
            TabIndex        =   94
            Top             =   1350
            Width           =   810
         End
         Begin VB.Label Label27 
            Caption         =   "Local de Pagto:"
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
            Left            =   3075
            TabIndex        =   93
            Top             =   1350
            Width           =   1365
         End
         Begin VB.Label Label26 
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
            Height          =   255
            Left            =   390
            TabIndex        =   92
            Top             =   885
            Width           =   495
         End
         Begin VB.Label Label32 
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
            Height          =   240
            Left            =   3690
            TabIndex        =   91
            Top             =   885
            Width           =   750
         End
         Begin VB.Label ContaCorrente 
            Height          =   255
            Left            =   1050
            TabIndex        =   90
            Top             =   435
            Width           =   1590
         End
         Begin VB.Label ValorPagoPagto 
            Height          =   255
            Left            =   1050
            TabIndex        =   89
            Top             =   900
            Width           =   1590
         End
         Begin VB.Label Historico 
            Height          =   255
            Left            =   1035
            TabIndex        =   88
            Top             =   1350
            Width           =   1845
         End
         Begin VB.Label NumOuSequencial 
            Height          =   255
            Left            =   4500
            TabIndex        =   87
            Top             =   900
            Width           =   735
         End
         Begin VB.Label Portador 
            Height          =   255
            Left            =   4500
            TabIndex        =   86
            Top             =   1350
            Width           =   1695
         End
      End
      Begin VB.Frame FramePagamento 
         Caption         =   "Pagamento Antecipado"
         Height          =   1770
         Index           =   1
         Left            =   2610
         TabIndex        =   68
         Top             =   2145
         Visible         =   0   'False
         Width           =   6375
         Begin VB.Label NumeroMP 
            Caption         =   "Numero"
            Height          =   225
            Left            =   4650
            TabIndex        =   80
            Top             =   810
            Width           =   720
         End
         Begin VB.Label CCIntNomeReduzido 
            Caption         =   "CCorrente"
            Height          =   225
            Left            =   4665
            TabIndex        =   79
            Top             =   450
            Width           =   1335
         End
         Begin VB.Label FilialEmpresaPA 
            Caption         =   "FilEmpr"
            Height          =   225
            Left            =   4620
            TabIndex        =   78
            Top             =   1125
            Width           =   525
         End
         Begin VB.Label ValorPA 
            Caption         =   "ValorPagtoAnt"
            Height          =   225
            Left            =   1755
            TabIndex        =   77
            Top             =   1125
            Width           =   1080
         End
         Begin VB.Label MeioPagtoDescricao 
            Caption         =   "MeioPagto"
            Height          =   225
            Left            =   1725
            TabIndex        =   76
            Top             =   795
            Width           =   960
         End
         Begin VB.Label DataMovimento 
            Caption         =   "DataMovto"
            Height          =   225
            Left            =   1740
            TabIndex        =   75
            Top             =   465
            Width           =   1095
         End
         Begin VB.Label Label33 
            Caption         =   "Filial Empresa:"
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
            Left            =   3270
            TabIndex        =   74
            Top             =   1095
            Width           =   1245
         End
         Begin VB.Label Label25 
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
            Height          =   255
            Left            =   1125
            TabIndex        =   73
            Top             =   1095
            Width           =   510
         End
         Begin VB.Label Label24 
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
            Height          =   255
            Left            =   3840
            TabIndex        =   72
            Top             =   795
            Width           =   705
         End
         Begin VB.Label Label23 
            Caption         =   "Meio Pagto:"
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
            TabIndex        =   71
            Top             =   795
            Width           =   1035
         End
         Begin VB.Label Label21 
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
            Height          =   255
            Left            =   3195
            TabIndex        =   70
            Top             =   450
            Width           =   1350
         End
         Begin VB.Label Label15 
            Caption         =   "Data Movimto:"
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
            Left            =   390
            TabIndex        =   69
            Top             =   450
            Width           =   1245
         End
      End
      Begin VB.Frame FramePagamento 
         Caption         =   "Crédito"
         Height          =   1770
         Index           =   2
         Left            =   2610
         TabIndex        =   55
         Top             =   2115
         Visible         =   0   'False
         Width           =   6375
         Begin VB.Label Label34 
            Caption         =   "Data Emissão:"
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
            Left            =   570
            TabIndex        =   67
            Top             =   405
            Width           =   1245
         End
         Begin VB.Label Label48 
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
            Height          =   255
            Left            =   4005
            TabIndex        =   66
            Top             =   405
            Width           =   450
         End
         Begin VB.Label Label37 
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
            Height          =   255
            Left            =   1110
            TabIndex        =   65
            Top             =   810
            Width           =   705
         End
         Begin VB.Label Label38 
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
            Height          =   255
            Left            =   3945
            TabIndex        =   64
            Top             =   810
            Width           =   510
         End
         Begin VB.Label Label39 
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
            Height          =   255
            Left            =   1260
            TabIndex        =   63
            Top             =   1215
            Width           =   555
         End
         Begin VB.Label Label40 
            Caption         =   "Filial Empresa:"
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
            Left            =   3210
            TabIndex        =   62
            Top             =   1215
            Width           =   1245
         End
         Begin VB.Label DataEmissaoCred 
            Caption         =   "DataEmissao"
            Height          =   225
            Left            =   1950
            TabIndex        =   61
            Top             =   405
            Width           =   1095
         End
         Begin VB.Label SiglaDocumentoCR 
            AutoSize        =   -1  'True
            Caption         =   "SiglaDoc"
            Height          =   195
            Left            =   4560
            TabIndex        =   60
            Top             =   405
            Width           =   645
         End
         Begin VB.Label ValorCredito 
            Caption         =   "Valor"
            Height          =   225
            Left            =   4560
            TabIndex        =   59
            Top             =   810
            Width           =   1080
         End
         Begin VB.Label FilialEmpresaCR 
            Caption         =   "FilEmpr"
            Height          =   225
            Left            =   4560
            TabIndex        =   58
            Top             =   1230
            Width           =   525
         End
         Begin VB.Label NumTitulo 
            Caption         =   "Numero"
            Height          =   225
            Left            =   1950
            TabIndex        =   57
            Top             =   810
            Width           =   720
         End
         Begin VB.Label SaldoCredito 
            Caption         =   "Saldo"
            Height          =   225
            Left            =   1980
            TabIndex        =   56
            Top             =   1215
            Width           =   1080
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Dados da Baixa"
         Height          =   1170
         Left            =   105
         TabIndex        =   42
         Top             =   765
         Width           =   8940
         Begin VB.Label DataBaixa 
            Height          =   255
            Left            =   1470
            TabIndex        =   54
            Top             =   375
            Width           =   1125
         End
         Begin VB.Label Label17 
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
            Left            =   900
            TabIndex        =   53
            Top             =   375
            Width           =   480
         End
         Begin VB.Label ValorPago 
            Height          =   225
            Left            =   4260
            TabIndex        =   52
            Top             =   375
            Width           =   945
         End
         Begin VB.Label Juros 
            Height          =   225
            Left            =   6930
            TabIndex        =   51
            Top             =   795
            Width           =   945
         End
         Begin VB.Label Multa 
            Height          =   225
            Left            =   4260
            TabIndex        =   50
            Top             =   795
            Width           =   945
         End
         Begin VB.Label Desconto 
            Height          =   225
            Left            =   1470
            TabIndex        =   49
            Top             =   795
            Width           =   945
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Valor Pago:"
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
            Left            =   3195
            TabIndex        =   48
            Top             =   375
            Width           =   1005
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Juros:"
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
            Left            =   6360
            TabIndex        =   47
            Top             =   795
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Multa:"
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
            Left            =   3660
            TabIndex        =   46
            Top             =   795
            Width           =   540
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Desconto:"
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
            Left            =   480
            TabIndex        =   45
            Top             =   795
            Width           =   885
         End
         Begin VB.Label ValorBaixado 
            Height          =   225
            Left            =   6930
            TabIndex        =   44
            Top             =   345
            Width           =   945
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Valor Baixado:"
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
            Left            =   5610
            TabIndex        =   43
            Top             =   345
            Width           =   1245
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tipo de Baixa"
         Height          =   1770
         Left            =   105
         TabIndex        =   38
         Top             =   2145
         Width           =   2385
         Begin VB.OptionButton Pagamento 
            Caption         =   "Crédito/Devolução"
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
            Index           =   2
            Left            =   105
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   1350
            Width           =   1950
         End
         Begin VB.OptionButton Pagamento 
            Caption         =   "Adiantamento à Forn."
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
            Height          =   210
            Index           =   1
            Left            =   120
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   900
            Width           =   2145
         End
         Begin VB.OptionButton Pagamento 
            Caption         =   "Pagamento"
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
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   450
            Value           =   -1  'True
            Width           =   1290
         End
      End
      Begin MSMask.MaskEdBox Parcela 
         Height          =   300
         Left            =   2055
         TabIndex        =   96
         Top             =   225
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   2
         Mask            =   "99"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownParcela 
         Height          =   300
         Index           =   1
         Left            =   2400
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   225
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   327681
         Enabled         =   -1  'True
      End
      Begin VB.Label Label13 
         Caption         =   "Sequencial:"
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
         Height          =   225
         Left            =   3120
         TabIndex        =   100
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Parcela:"
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
         Left            =   1230
         TabIndex        =   97
         Top             =   285
         Width           =   720
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4020
      Index           =   2
      Left            =   90
      TabIndex        =   28
      Top             =   1620
      Visible         =   0   'False
      Width           =   9015
      Begin VB.Frame Frame6 
         Caption         =   "Parcelas"
         Height          =   3795
         Left            =   270
         TabIndex        =   29
         Top             =   0
         Width           =   8400
         Begin MSMask.MaskEdBox DataVencimentoReal 
            Height          =   225
            Left            =   1650
            TabIndex        =   33
            Top             =   900
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorParcela 
            Height          =   225
            Left            =   2940
            TabIndex        =   34
            Top             =   870
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox DataVencimento 
            Height          =   225
            Left            =   480
            TabIndex        =   35
            Top             =   900
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.CheckBox Suspenso 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   225
            Left            =   6270
            TabIndex        =   31
            Top             =   630
            Width           =   900
         End
         Begin VB.ComboBox TipoCobranca 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4275
            TabIndex        =   30
            Top             =   600
            Width           =   1935
         End
         Begin MSFlexGridLib.MSFlexGrid GridParcelas 
            Height          =   1755
            Left            =   180
            TabIndex        =   32
            Top             =   360
            Width           =   7755
            _ExtentX        =   13679
            _ExtentY        =   3096
            _Version        =   393216
            Rows            =   50
            Cols            =   6
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.ComboBox Filial 
      Height          =   315
      Left            =   4710
      TabIndex        =   5
      Top             =   180
      Width           =   1815
   End
   Begin VB.ComboBox Tipo 
      Height          =   315
      Left            =   1245
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   675
      Width           =   885
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   6690
      ScaleHeight     =   795
      ScaleWidth      =   2565
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   2625
      Begin VB.CommandButton BotaoFechar 
         Height          =   675
         Left            =   2025
         Picture         =   "TituloPagar_Consulta.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   480
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   675
         Left            =   1455
         Picture         =   "TituloPagar_Consulta.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   480
      End
      Begin VB.CommandButton BotaoConsulta 
         Height          =   675
         Left            =   90
         Picture         =   "TituloPagar_Consulta.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   60
         Width           =   1275
      End
   End
   Begin MSMask.MaskEdBox NumeroTitulo 
      Height          =   300
      Left            =   4680
      TabIndex        =   6
      Top             =   675
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   6
      Mask            =   "999999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Fornecedor 
      Height          =   300
      Left            =   1245
      TabIndex        =   7
      Top             =   180
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   "_"
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   4455
      Left            =   60
      TabIndex        =   12
      Top             =   1230
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   7858
      MultiRow        =   -1  'True
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Título"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Parcelas"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Baixa"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Contabilização"
            Key             =   ""
            Object.Tag             =   ""
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
      Left            =   150
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   11
      Top             =   240
      Width           =   1035
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
      Left            =   3930
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   10
      Top             =   735
      Width           =   720
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
      Left            =   4125
      TabIndex        =   9
      Top             =   240
      Width           =   525
   End
   Begin VB.Label LabelTipo 
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
      Height          =   210
      Left            =   735
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   8
      Top             =   720
      Width           =   480
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9420
      Y1              =   1140
      Y2              =   1140
   End
End
Attribute VB_Name = "TituloPagar_ConsultaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Public iAlterado As Integer
Private iFrameAtual As Integer
Dim iFramePagamentoAtual As Integer

Dim objGridParcelas As AdmGrid
Dim iGrid_Parcela_Col As Integer
Dim iGrid_Vencimento_Col As Integer
Dim iGrid_VenctoReal_Col As Integer
Dim iGrid_ValorParcela_Col As Integer
Dim iGrid_Cobranca_Col As Integer
Dim iGrid_Suspenso_Col As Integer

Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1

'inicio contabilidade

Dim objGrid1 As AdmGrid
Dim objContabil As New ClassContabil

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 18392

    'Se não encontra valor que era CÓDIGO
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then Error 18393

        sFornecedor = Fornecedor.Text
        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial",sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then Error 18394
        
        'Se não encontrou
        If lErro = 18272 Then
        
            objFornecedor.sNomeReduzido = sFornecedor
            
            'Le o Código do Fornecedor --> Para Passar para a Tela de Filiais
            lErro = CF("Fornecedor_Le_NomeReduzido",objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then Error 58616
            
            'Passa o Código do Fornecedor
            objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
            
            'Sugere cadastrar nova Filial
            Error 18395
        
        End If

        'Coloca na tela a Filial lida
        Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then Error 18396

    Exit Sub

Erro_Filial_Validate:

    Cancel = True
    
    Select Case Err

       Case 18392, 18394, 58616

       Case 18393
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

       Case 18395
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 18396
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", Err, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175112)

    End Select

    Exit Sub
End Sub

Private Sub NumPC_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    iFramePagamentoAtual = 0
    
    Set objEventoNumero = New AdmEvento
    Set objEventoFornecedor = New AdmEvento

    'Carrega na combo os tipos de cobrança
    lErro = Carrega_TipoCobranca()
    If lErro <> SUCESSO Then Error 18363
    
    Set objGridParcelas = New AdmGrid

    'Inicializa Grid Parcelas
    lErro = Inicializa_Grid_Parcelas(objGridParcelas)
    If lErro <> SUCESSO Then Error 18367

    'Inicialização da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO Then Error 18362
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 18362, 18363, 18365, 18367, 48914 'Tratados nas Rotinas Chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175113)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
Dim colParcelas As New colParcelaPagar

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TitulosPag"

    'Lê os dados da Tela Notas Fiscais a Pagar
    lErro = Move_Tela_Memoria(objTituloPagar, colParcelas)
    If lErro <> SUCESSO Then Error 18360

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", CLng(0), 0, "NumIntDoc"
    colCampoValor.Add "Fornecedor", objTituloPagar.lFornecedor, 0, "Fornecedor"
    colCampoValor.Add "Filial", objTituloPagar.iFilial, 0, "Filial"
    colCampoValor.Add "NumTitulo", objTituloPagar.lNumTitulo, 0, "NumTitulo"
    colCampoValor.Add "DataEmissao", objTituloPagar.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "NumParcelas", objTituloPagar.iNumParcelas, 0, "NumParcelas"
    colCampoValor.Add "ValorTotal", objTituloPagar.dValorTotal, 0, "ValorTotal"
    colCampoValor.Add "ValorSeguro", objTituloPagar.dValorSeguro, 0, "ValorSeguro"
    colCampoValor.Add "ValorFrete", objTituloPagar.dValorFrete, 0, "ValorFrete"
    colCampoValor.Add "OutrasDespesas", objTituloPagar.dOutrasDespesas, 0, "OutrasDespesas"
    colCampoValor.Add "ValorProdutos", objTituloPagar.dValorProdutos, 0, "ValorProdutos"
    colCampoValor.Add "ValorICMS", objTituloPagar.dValorICMS, 0, "ValorICMS"
    colCampoValor.Add "ValorICMSSubst", objTituloPagar.dValorICMS, 0, "ValorICMSSubst"
    colCampoValor.Add "CreditoICMS", objTituloPagar.iCreditoICMS, 0, "CreditoICMS"
    colCampoValor.Add "ValorIPI", objTituloPagar.dValorIPI, 0, "ValorIPI"
    colCampoValor.Add "CreditoIPI", objTituloPagar.iCreditoIPI, 0, "CreditoIPI"
    colCampoValor.Add "ValorIRRF", objTituloPagar.dValorIRRF, 0, "ValorIRRF"
    colCampoValor.Add "ValorINSS", objTituloPagar.dValorINSS, 0, "ValorINSS"
    colCampoValor.Add "FilialPedCompra", objTituloPagar.iFilialPedCompra, 0, "FilialPedCompra"
    colCampoValor.Add "NumPedCompra", objTituloPagar.lNumPedCompra, 0, "NumPedCompra"
    colCampoValor.Add "CondicaoPagto", objTituloPagar.iCondicaoPagto, 0, "CondicaoPagto"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "SiglaDocumento", OP_IGUAL, TIPODOC_NF_FATURA_PAGAR

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 18360

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175114)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar

On Error GoTo Erro_Tela_Preenche

    objTituloPagar.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor

    If objTituloPagar.lNumIntDoc <> 0 Then

        'Carrega objTituloPagar com os dados passados em colCampoValor
        objTituloPagar.lFornecedor = colCampoValor.Item("Fornecedor").vValor
        objTituloPagar.iFilial = colCampoValor.Item("Filial").vValor
        objTituloPagar.dtDataEmissao = colCampoValor.Item("DataEmissao").vValor
        objTituloPagar.lNumTitulo = colCampoValor.Item("NumTitulo").vValor
        objTituloPagar.iNumParcelas = colCampoValor.Item("NumParcelas").vValor
        objTituloPagar.dValorTotal = colCampoValor.Item("ValorTotal").vValor
        objTituloPagar.dValorSeguro = colCampoValor.Item("ValorSeguro").vValor
        objTituloPagar.dValorFrete = colCampoValor.Item("ValorFrete").vValor
        objTituloPagar.dOutrasDespesas = colCampoValor.Item("OutrasDespesas").vValor
        objTituloPagar.dValorProdutos = colCampoValor.Item("ValorProdutos").vValor
        objTituloPagar.dValorICMS = colCampoValor.Item("ValorICMS").vValor
        objTituloPagar.iCreditoICMS = colCampoValor.Item("CreditoICMS").vValor
        objTituloPagar.dValorICMSSubst = colCampoValor.Item("ValorICMSSubst").vValor
        objTituloPagar.dValorIPI = colCampoValor.Item("ValorIPI").vValor
        objTituloPagar.iCreditoIPI = colCampoValor.Item("CreditoIPI").vValor
        objTituloPagar.dValorIRRF = colCampoValor.Item("ValorIRRF").vValor
        objTituloPagar.dValorINSS = colCampoValor.Item("ValorINSS").vValor
        objTituloPagar.iFilialPedCompra = colCampoValor.Item("FilialPedCompra").vValor
        objTituloPagar.lNumPedCompra = colCampoValor.Item("NumPedCompra").vValor
        objTituloPagar.iCondicaoPagto = colCampoValor.Item("CondicaoPagto").vValor
        
        'Traz a Nota para Tela
        lErro = Traz_NFFatPag_Tela(objTituloPagar)
        If lErro <> SUCESSO Then Error 18361

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 18361

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175115)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

 Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Sub FornecedorLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'Preenche NomeReduzido com o fornecedor da tela
    If Len(Trim(Fornecedor.Text)) > 0 Then objFornecedor.sNomeReduzido = Fornecedor.Text

    'Chama Tela FornecedorLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor, Cancel As Boolean

    Set objFornecedor = obj1

    'Preenche campo Fornecedor
    Fornecedor.Text = objFornecedor.sNomeReduzido

    Call Fornecedor_Validate(Cancel)

    Me.Show

    Exit Sub

End Sub

Private Sub NumeroLabel_Click()

Dim objTituloPagar As New ClassTituloPagar
Dim colParcelas As New colParcelaPagar
Dim colSelecao As New Collection
Dim lErro As Long

On Error GoTo Erro_NumeroLabel_Click

    'Se Forncedor estiver vazio, erro
    If Len(Trim(Fornecedor.Text)) = 0 Then Error 18389

    'Se Filial estiver vazia, erro
    If Len(Trim(Filial.Text)) = 0 Then Error 18390

    'Move os dados da Tela para objTituloPagar e colParcelas
    lErro = Move_Tela_Memoria(objTituloPagar, colParcelas)
    If lErro <> SUCESSO Then Error 18391

    'Adiciona filtros: lFornecedor e iFilial
    colSelecao.Add objTituloPagar.lFornecedor
    colSelecao.Add objTituloPagar.iFilial

    'Chama Tela NFFatPagLista
    Call Chama_Tela("NFFatPagLista", colSelecao, objTituloPagar, objEventoNumero)

    Exit Sub

Erro_NumeroLabel_Click:

    Select Case Err

        Case 18389
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 18390
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 18391

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175116)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTituloPagar As ClassTituloPagar

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objTituloPagar = obj1

    'Traz os dados de objTituloPagar para Teal
    lErro = Traz_NFFatPag_Tela(objTituloPagar)
    If lErro <> SUCESSO Then Error 18359

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case Err

        Case 18359

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175117)

    End Select

    Exit Sub


End Sub

Function Trata_Parametros(Optional objTituloPagar As ClassTituloPagar) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se foi passado algum Título
    If Not (objTituloPagar Is Nothing) Then

        'Lê o Título
        lErro = CF("TituloPagar_Le",objTituloPagar)
        If lErro <> SUCESSO And lErro <> 18372 Then Error 18368
        
        If lErro <> SUCESSO Then Error 18384

        'Verifica se é uma Nota Fiscal Fatura
        If objTituloPagar.sSiglaDocumento <> TIPODOC_NF_FATURA_PAGAR Then Error 18373

        'Traz os dados para a Tela
        lErro = Traz_NFFatPag_Tela(objTituloPagar)
        If lErro <> SUCESSO Then Error 18374

    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 18368, 18374

        Case 18373
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULO_NAO_NFFATPAG", Err, objTituloPagar.lNumIntDoc)

        Case 18384
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFFATPAG_NAO_CADASTRADA", Err, objTituloPagar.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175118)

    End Select

    Exit Function

End Function

Private Function Traz_NFFatPag_Tela(objTituloPagar As ClassTituloPagar) As Long
'Traz os dados da Nota Fiscal Fatura para a Tela

Dim lErro As Long
Dim colParcelasPag As New colParcelaPagar
Dim objParcelaPagar As ClassParcelaPagar
Dim iLinha As Integer
Dim iIndice As Integer, Cancel As Boolean
Dim bCancel As Boolean

On Error GoTo Erro_Traz_NFFatPag_Tela
    
    Call Limpa_Tela_NFFatPag
    
    'Coloca os dados do Título Pagar na Tela
    NumeroTitulo.Text = objTituloPagar.lNumTitulo

    Fornecedor.Text = objTituloPagar.lFornecedor
    Call Fornecedor_Validate(Cancel)

    Filial.Text = objTituloPagar.iFilial
    Filial_Validate (bCancel)
    
    If objTituloPagar.iFilialPedCompra <> 0 Then
        ComboFilialPC.Caption = objTituloPagar.iFilialPedCompra
    Else
        ComboFilialPC.Caption = ""
    End If
    
    If objTituloPagar.lNumPedCompra <> 0 Then
        NumPC.Caption = objTituloPagar.lNumPedCompra
    Else
        NumPC.Caption = ""
    End If
    
    DataEmissao.Caption = Format(objTituloPagar.dtDataEmissao, "dd/mm/yyyy")

    ValorTotal.Caption = objTituloPagar.dValorTotal
    ValorICMS.Caption = objTituloPagar.dValorICMS
    ValorICMSSubst.Caption = objTituloPagar.dValorICMSSubst
    CreditoICMS.Value = objTituloPagar.iCreditoICMS
    ValorProdutos.Caption = objTituloPagar.dValorProdutos
    ValorFrete.Caption = objTituloPagar.dValorFrete
    ValorSeguro.Caption = objTituloPagar.dValorSeguro
    OutrasDespesas.Caption = objTituloPagar.dOutrasDespesas
    ValorIPI.Caption = objTituloPagar.dValorIPI
    CreditoIPI.Value = objTituloPagar.iCreditoIPI
    ValorIRRF.Caption = objTituloPagar.dValorIRRF
    ValorINSS.Caption = objTituloPagar.dValorINSS
    
    'Lê as Parcelas a Pagar vinculadas ao Título
    lErro = CF("ParcelasPagar_Le",objTituloPagar, colParcelasPag)
    If lErro <> SUCESSO Then Error 18377

    If colParcelasPag.Count > NUM_MAXIMO_PARCELAS Then Error 18674

    Call Grid_Limpa(objGridParcelas)
    
    iLinha = 0

    'Preenche as linhas do Grid Parcelas com os dados de cada Parcela
    For Each objParcelaPagar In colParcelasPag

        iLinha = iLinha + 1

        GridParcelas.TextMatrix(iLinha, iGrid_Parcela_Col) = objParcelaPagar.iNumParcela
        GridParcelas.TextMatrix(iLinha, iGrid_Vencimento_Col) = Format(objParcelaPagar.dtDataVencimento, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iLinha, iGrid_VenctoReal_Col) = Format(objParcelaPagar.dtDataVencimentoReal, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iLinha, iGrid_ValorParcela_Col) = Format(objParcelaPagar.dValor, "Standard")

        For iIndice = 0 To TipoCobranca.ListCount - 1
            If TipoCobranca.ItemData(iIndice) = objParcelaPagar.iTipoCobranca Then
                GridParcelas.TextMatrix(iLinha, iGrid_Cobranca_Col) = TipoCobranca.List(iIndice)
                Exit For
            End If
        Next

        If objParcelaPagar.iStatus = STATUS_SUSPENSO Then
            GridParcelas.TextMatrix(iLinha, iGrid_Suspenso_Col) = "1"
        Else
            GridParcelas.TextMatrix(iLinha, iGrid_Suspenso_Col) = "0"
        End If

    Next

    'Faz o número de linhas existentes do Grid ser igual ao número de Parcelas
    objGridParcelas.iLinhasExistentes = iLinha

    'Faz refresh nas checkboxes
    Call Grid_Refresh_Checkbox(objGridParcelas)

    'traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objTituloPagar.lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 36326 Then Error 36364

    iAlterado = 0
    
    Traz_NFFatPag_Tela = SUCESSO

    Exit Function

Erro_Traz_NFFatPag_Tela:

    Traz_NFFatPag_Tela = Err

    Select Case Err

        Case 18377, 36364 'Tratados nas Rotinas Chamadas

        Case 18674
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_MAXIMO_PARCELAS_ULTRAPASSADO", Err, colParcelasPag.Count, NUM_MAXIMO_PARCELAS)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175119)

    End Select

    Exit Function

End Function

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objTipoFornecedor As New ClassTipoFornecedor
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim bCancel As Boolean

On Error GoTo Erro_Fornecedor_Validate

        If Len(Trim(Fornecedor.Text)) > 0 Then

            'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then Error 18385

            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor",objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then Error 18386

            'Preenche ComboBox de Filiais
            Call Filial_Preenche(Filial, colCodigoNome)

            'Seleciona filial na Combo Filial
            Call Filial_Seleciona(Filial, iCodFilial)

        ElseIf Len(Trim(Fornecedor.Text)) = 0 Then

            'Limpa Combo de Filial
            Filial.Clear

        End If

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True
    
    Select Case Err

        Case 18385, 18386, 18387, 18588, 18589 'Tratados nas Rotinas chamadas

        Case 18675, 18676
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", Err, objCondicaoPagto.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175120)

    End Select

    Exit Sub

End Sub

Private Sub NumeroTitulo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NumeroTitulo_Validate

    'Verifica se o Numero foi preenchido
    If Len(Trim(NumeroTitulo.ClipText)) = 0 Then Exit Sub

    'Critica se é Long positivo
    lErro = Long_Critica(NumeroTitulo.ClipText)
    If lErro <> SUCESSO Then Error 18398

    Exit Sub

Erro_NumeroTitulo_Validate:

    Cancel = True


    Select Case Err

        Case 18398

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175121)

    End Select

    Exit Sub

End Sub

Private Function GridParcelas_Preenche(objCondicaoPagto As ClassCondicaoPagto) As Long
'Calcula valores e datas de vencimento de Parcelas a partir da Condição de Pagamento e preenche GridParcelas

Dim lErro As Long
Dim dValorPagar As Double
Dim colValorParcelas As New Collection
Dim colDataVencimento As New Collection
Dim dtDataEmissao As Date
Dim dtDataVenctoReal As Date
Dim dValorIRRF As Double
Dim iIndice As Integer

On Error GoTo Erro_GridParcelas_Preenche
    
    Call Grid_Limpa(objGridParcelas)
    
    'obtem Valor a Pagar
    If Len(Trim(ValorTotal)) > 0 Then
        If Len(Trim(ValorIRRF)) > 0 Then dValorIRRF = CDbl(ValorIRRF)
        dValorPagar = CDbl(ValorTotal) - dValorIRRF
    End If

    'Se Valor a Pagar for positivo
    If dValorPagar > 0 Then

        'Calcula os valores das Parcelas
        lErro = CF("Parcelas_Calcula",dValorPagar, objCondicaoPagto.iNumeroParcelas, colValorParcelas)
        If lErro <> SUCESSO Then Error 18417

        'Número de Parcelas
        objGridParcelas.iLinhasExistentes = objCondicaoPagto.iNumeroParcelas

        'Coloca os valores das Parcelas no Grid Parcelas
        For iIndice = 1 To objGridParcelas.iLinhasExistentes
            GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col) = Format(colValorParcelas(iIndice), "Standard")
        Next

    End If

    'Se Data Emissão estiver preenchida
    If Len(Trim(DataEmissao.Caption)) > 0 Then

        dtDataEmissao = CDate(DataEmissao.Caption)

        'Calcula Datas de Vencimento das Parcelas
        lErro = CF("Parcelas_DatasVencimento",objCondicaoPagto, dtDataEmissao, colDataVencimento)
        If lErro <> SUCESSO Then Error 18441

        'Número de Parcelas
        objGridParcelas.iLinhasExistentes = objCondicaoPagto.iNumeroParcelas

        'Loop de preenchimento do Grid Parcelas com Datas de Vencimento
        For iIndice = 1 To objCondicaoPagto.iNumeroParcelas

            'Coloca Data de Vencimento no Grid Parcelas
            GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col) = Format(colDataVencimento(iIndice), "dd/mm/yyyy")

            'Calcula Data Vencimento Real
            lErro = CF("DataVencto_Real",colDataVencimento(iIndice), dtDataVenctoReal)
            If lErro <> SUCESSO Then Error 18443

            'Coloca Data de Vencimento Real no Grid Parcelas
            GridParcelas.TextMatrix(iIndice, iGrid_VenctoReal_Col) = Format(dtDataVenctoReal, "dd/mm/yyyy")

        Next

    End If

    GridParcelas_Preenche = SUCESSO

    Exit Function

Erro_GridParcelas_Preenche:

    GridParcelas_Preenche = Err

    Select Case Err

        Case 18417, 18441, 18443

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175122)

    End Select

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iUltimaLinha As Integer
Dim ColRateioOn As New Collection

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 36242

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 18591

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 18422, 18424, 18440, 18512, 36242

        Case 18591
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 18682

    'Limpa a Tela
    Call Limpa_Tela_NFFatPag

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 18682

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175123)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoNumero = Nothing
    Set objEventoFornecedor = Nothing

    Set objGridParcelas = Nothing

    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing

    Set objGrid1 = Nothing
    Set objContabil = Nothing
    
   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
     
End Sub

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        Frame(Opcao.SelectedItem.Index).Visible = True
        Frame(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index

        'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        If Opcao.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then Call objContabil.Contabil_Carga_Modelo_Padrao

    End If

End Sub

Private Function Carrega_TipoCobranca() As Long

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_TipoCobranca

    'Lê o código e a descrição de todos os Tipos de Cobrança
    lErro = CF("Cod_Nomes_Le","TiposDeCobranca", "Codigo", "Descricao", STRING_TIPOSDECOBRANCA_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 18364

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o ítem na List da Combo TipoCobranca
        TipoCobranca.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        TipoCobranca.ItemData(TipoCobranca.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_TipoCobranca = SUCESSO

    Exit Function

Erro_Carrega_TipoCobranca:

    Carrega_TipoCobranca = Err

    Select Case Err

        Case 18364

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175124)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Parcelas(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Vencto Real")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Cobrança")
    objGridInt.colColuna.Add ("Suspenso")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (DataVencimento.Name)
    objGridInt.colCampo.Add (DataVencimentoReal.Name)
    objGridInt.colCampo.Add (ValorParcela.Name)
    objGridInt.colCampo.Add (TipoCobranca.Name)
    objGridInt.colCampo.Add (Suspenso.Name)

    'Colunas do Grid
    iGrid_Parcela_Col = 0
    iGrid_Vencimento_Col = 1
    iGrid_VenctoReal_Col = 2
    iGrid_ValorParcela_Col = 3
    iGrid_Cobranca_Col = 4
    iGrid_Suspenso_Col = 5

    'Grid do GridInterno
    objGridInt.objGrid = GridParcelas

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_PARCELAS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 8

    'Largura da primeira coluna
    GridParcelas.ColWidth(0) = 900

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Parcelas = SUCESSO

    Exit Function

End Function

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Private Sub Limpa_Tela_NFFatPag()

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Chama função que limpa TextBoxes e MaskedEdits da Tela
    Call Limpa_Tela(Me)

    'Limpa os campos não são limpos pela função acima
    Filial.Clear
    CreditoICMS.Value = 0
    CreditoIPI.Value = 0

    Call Grid_Limpa(objGridParcelas)

    'Limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    iAlterado = 0

End Sub

Private Function Move_Tela_Memoria(objTituloPagar As ClassTituloPagar, colParcelas As colParcelaPagar) As Long
'Move os dados da Tela para objTituloPagar e colParcelas

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Move_Tela_Memoria

    If Len(Trim(Fornecedor.Text)) > 0 Then
        'Codigo do Fornecedor
        objFornecedor.sNomeReduzido = Fornecedor.Text

        lErro = CF("Fornecedor_Le_NomeReduzido",objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then Error 18418
        If lErro <> SUCESSO Then Error 18791

        objTituloPagar.lFornecedor = objFornecedor.lCodigo
    End If

    'A filial do Fornecedor
    If Len(Trim(Filial.Text)) > 0 Then
        objTituloPagar.iFilial = Codigo_Extrai(Filial.Text)
    End If
        
    'FilialPC
    If Len(Trim(ComboFilialPC.Caption)) > 0 Then
        objTituloPagar.iFilialPedCompra = Codigo_Extrai(ComboFilialPC.Caption)
    End If
    
    'NumPC
    If Len(Trim(NumPC.Caption)) > 0 Then objTituloPagar.lNumPedCompra = CLng(NumPC.Caption)
    
    'Numero do Titulo
    If Len(Trim(NumeroTitulo.ClipText)) > 0 Then objTituloPagar.lNumTitulo = CLng(NumeroTitulo.ClipText)

    'Data de Emissao
    If Len(Trim(DataEmissao.Caption)) = 0 Then
        objTituloPagar.dtDataEmissao = DATA_NULA
    Else
        objTituloPagar.dtDataEmissao = CDate(DataEmissao.Caption)
    End If

    'Os Valores, sSiglaDocumento, FilialEmpresa, CondPagamento
    If Len(Trim(ValorTotal.Caption)) > 0 Then objTituloPagar.dValorTotal = CDbl(ValorTotal.Caption)
    objTituloPagar.iNumParcelas = objGridParcelas.iLinhasExistentes
    If Len(Trim(ValorICMS.Caption)) > 0 Then objTituloPagar.dValorICMS = CDbl(ValorICMS.Caption)
    If Len(Trim(ValorICMSSubst.Caption)) > 0 Then objTituloPagar.dValorICMSSubst = CDbl(ValorICMSSubst.Caption)
    objTituloPagar.iCreditoICMS = CreditoICMS.Value
    If Len(Trim(ValorProdutos.Caption)) > 0 Then objTituloPagar.dValorProdutos = CDbl(ValorProdutos.Caption)
    If Len(Trim(OutrasDespesas.Caption)) > 0 Then objTituloPagar.dOutrasDespesas = CDbl(OutrasDespesas.Caption)
    If Len(Trim(ValorSeguro.Caption)) > 0 Then objTituloPagar.dValorSeguro = CDbl(ValorSeguro.Caption)
    If Len(Trim(ValorFrete.Caption)) > 0 Then objTituloPagar.dValorFrete = CDbl(ValorFrete.Caption)
    If Len(Trim(ValorIRRF.Caption)) > 0 Then objTituloPagar.dValorIRRF = CDbl(ValorIRRF.Caption)
    If Len(Trim(ValorIPI.Caption)) > 0 Then objTituloPagar.dValorIPI = CDbl(ValorIPI.Caption)
    If Len(Trim(ValorINSS.Caption)) > 0 Then objTituloPagar.dValorINSS = CDbl(ValorINSS.Caption)
    objTituloPagar.iCreditoIPI = CreditoIPI.Value
    objTituloPagar.sSiglaDocumento = TIPODOC_NF_FATURA_PAGAR
    objTituloPagar.iFilialEmpresa = giFilialEmpresa
    
    'Move para colParcelas os dados do Grid Parcelas
    lErro = Move_GridParcelas_Memoria(colParcelas)
    If lErro <> SUCESSO Then Error 18419

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 18418, 18419

        Case 18791
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175125)

    End Select

    Exit Function

End Function

Private Function Move_GridParcelas_Memoria(colParcelas As colParcelaPagar) As Long
'Move para a memória os dados existentes no Grid

Dim lErro As Long
Dim iIndice As Integer
Dim objParcelaPag As ClassParcelaPagar

On Error GoTo Erro_Move_GridParcelas_Memoria

    For iIndice = 1 To objGridParcelas.iLinhasExistentes

        Set objParcelaPag = New ClassParcelaPagar

        'Preenche objParcelaPag com a linha do GridParcelas
        objParcelaPag.iNumParcela = iIndice
        objParcelaPag.dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))
        objParcelaPag.dtDataVencimentoReal = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_VenctoReal_Col))
        objParcelaPag.dValor = CDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))

        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Cobranca_Col))) = 0 Then
            'Se estiver vazio usamos o Tipo Cobrança DEFAULT
            objParcelaPag.iTipoCobranca = TIPO_COBRANCA_CARTEIRA
        Else
            objParcelaPag.iTipoCobranca = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Cobranca_Col))
        End If

        If GridParcelas.TextMatrix(iIndice, iGrid_Suspenso_Col) = "1" Then
            objParcelaPag.iStatus = STATUS_SUSPENSO
        Else
            objParcelaPag.iStatus = STATUS_ABERTO
        End If

        'Adiciona objParcelaPag à coleção colParcelas
        With objParcelaPag
            colParcelas.Add .lNumIntDoc, .lNumIntTitulo, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, .iPortador, .iProxSeqBaixa, .iTipoCobranca, .iBancoCobrador, .sNossoNumero
        End With
    Next

    Move_GridParcelas_Memoria = SUCESSO

    Exit Function

Erro_Move_GridParcelas_Memoria:

    Move_GridParcelas_Memoria = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175126)

    End Select

    Exit Function

End Function

Private Sub Pagamento_Click(Index As Integer)

    FramePagamento(iFramePagamentoAtual).Visible = False
    FramePagamento(Index).Visible = True
    iFramePagamentoAtual = Index

End Sub

'Parte da Baixa Vinda da Tela de Cancela Baixa
Private Sub Carrega_Dados_Parcela()
'Carrega os dados da Parcela

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim iNumParcela As Integer
Dim objParcPagBaixa As New ClassBaixaParcPagar
Dim objBaixaPagar As New ClassBaixaPagar
Dim iSequencialBaixa As Integer

On Error GoTo Erro_Carrega_Dados_Parcela

    If Len(Trim(Fornecedor.ClipText)) = 0 Or Len(Trim(Filial.Text)) = 0 Or Len(Trim(Tipo.Text)) = 0 Or _
       Len(Trim(NumeroTitulo.ClipText)) = 0 Or Len(Trim(Parcela.ClipText)) = 0 Or Len(Trim(Sequencial.List(Sequencial.ListIndex))) = 0 Then Error 57317

    Call Limpa_Campos_Parcela

    'Lê o Fornecedor
    lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
    If lErro <> SUCESSO Then Error 42801

    objTituloPagar.lFornecedor = objFornecedor.lCodigo
    objTituloPagar.iFilial = Codigo_Extrai(Filial.Text)
    objTituloPagar.sSiglaDocumento = Tipo.Text
    objTituloPagar.lNumTitulo = CLng(NumeroTitulo.Text)
    objTituloPagar.dtDataEmissao = MaskedParaDate(DataEmissao)

    'Lê o Titulo a Pagar Baixado com os dados da tela
    lErro = CF("TituloPagarBaixado_Le_Numero",objTituloPagar)
    If lErro <> SUCESSO And lErro <> 18556 Then Error 42802

    If lErro <> SUCESSO Then

        'Lê o Titulo a Pagar Nao Baixado com os dados da tela
        lErro = CF("TituloPagar_Le_Numero",objTituloPagar)
        If lErro <> SUCESSO And lErro <> 18551 Then Error 42802

        If lErro <> SUCESSO Then Error 42803 'Não encontrou

    End If

        iNumParcela = CInt(Parcela.Text)

        iSequencialBaixa = CInt(Sequencial)

        'Lê a Parcela no BD
        lErro = CF("BaixaPagCancelar_Le_Parcela",objTituloPagar.lNumIntDoc, iNumParcela, iSequencialBaixa, objParcPagBaixa)
        If lErro <> SUCESSO And lErro <> 42807 Then Error 42808

        If lErro = 42807 Then Error 42816 'Não encontrou

        objBaixaPagar.lNumIntBaixa = objParcPagBaixa.lNumIntBaixa

        'Lê  a Baixa
        lErro = CF("BaixaPagar_Le",objBaixaPagar)
        If lErro <> SUCESSO And lErro <> 42812 Then Error 42813

        If lErro <> SUCESSO Then Error 42814 'Não encontrou

        'Coloca na tela os dados da Baixa Lida
        lErro = Traz_Dados_Baixa(objParcPagBaixa, objBaixaPagar)
        If lErro <> SUCESSO Then Error 42815

    Exit Sub

Erro_Carrega_Dados_Parcela:

    Select Case Err

        Case 42801, 42802, 42808, 42813, 42815

        Case 42803
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULO_PAGAR_INEXISTENTE", Err)

        Case 42814
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BAIXAPAG_INEXISTENTE", Err)

        Case 42816
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BAIXAPARCPAG_INEXISTENTE", Err)

        Case 57317
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAMPOS_CANCELARBAIXA_NAO_PREENCHIDOS", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175127)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Campos_Parcela()

    DataBaixa.Caption = ""
    ValorPago.Caption = ""
    ValorBaixado.Caption = ""
    Desconto.Caption = ""
    Multa.Caption = ""
    Juros.Caption = ""

    ContaCorrente.Caption = ""
    ValorPagoPagto.Caption = ""
    Historico.Caption = ""
    Portador.Caption = ""
    NumOuSequencial.Caption = ""
    DataMovimento.Caption = ""
    MeioPagtoDescricao.Caption = ""
    ValorPA.Caption = ""
    CCIntNomeReduzido.Caption = ""
    NumeroMP.Caption = ""
    FilialEmpresaPA.Caption = ""
    DataEmissaoCred.Caption = ""
    NumTitulo.Caption = ""
    ValorCredito.Caption = ""
    SaldoCredito.Caption = ""
    SiglaDocumentoCR.Caption = ""
    FilialEmpresaCR.Caption = ""

End Sub
Private Function Traz_Dados_Baixa(objParcPagBaixa As ClassBaixaParcPagar, objBaixaPagar As ClassBaixaPagar) As Long
'Mostra na tela os dados da baixa

Dim lErro As Long

On Error GoTo Erro_Traz_Dados_Baixa

    'Coloca os dados da Baixa na tela
    Desconto.Caption = Format(objParcPagBaixa.dValorDesconto, "Standard")
    ValorPago.Caption = Format(objParcPagBaixa.dValorBaixado - objParcPagBaixa.dValorDesconto + objParcPagBaixa.dValorMulta + objParcPagBaixa.dValorJuros, "Standard")
    Multa.Caption = Format(objParcPagBaixa.dValorMulta, "Standard")
    ValorBaixado.Caption = Format(objParcPagBaixa.dValorBaixado, "Standard")
    Juros.Caption = Format(objParcPagBaixa.dValorJuros, "Standard")
    DataBaixa.Caption = Format(objBaixaPagar.dtData, "dd/mm/yyyy")

    If objBaixaPagar.iMotivo = MOTIVO_PAGAMENTO Then
        Pagamento(0).Value = True

        'Traz os dados do pagamento
        lErro = Traz_Dados_Pagamento(objBaixaPagar)
        If lErro <> SUCESSO Then Error 42817

    ElseIf objBaixaPagar.iMotivo = MOTIVO_PAGTO_ANTECIPADO Then
        Pagamento(1).Value = True

        'Traz os dados do pagamento antecipado
        lErro = Traz_Dados_Pagamento_Antecipado(objBaixaPagar)
        If lErro <> SUCESSO Then Error 42823

    ElseIf objBaixaPagar.iMotivo = MOTIVO_CREDITO_FORNECEDOR Then
        Pagamento(2).Value = True

        'Traz os dados do crédito
        lErro = Traz_Dados_Credito_Fornecedor(objBaixaPagar)
        If lErro <> SUCESSO Then Error 42820

    End If
    
    Traz_Dados_Baixa = SUCESSO

    Exit Function

Erro_Traz_Dados_Baixa:

    Traz_Dados_Baixa = Err

    Select Case Err

        Case 42817, 42820, 42823, 57363

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175128)

    End Select

    Exit Function

End Function

Private Function Traz_Dados_Pagamento(objBaixaPagar As ClassBaixaPagar) As Long
'Mostra os dados de pagamento na tela

Dim lErro As Long
Dim objMovCCI As New ClassMovContaCorrente
Dim objContaCorrente As New ClassContasCorrentesInternas
Dim objBorderoPagto As New ClassBorderoPagto
Dim objPortador As New ClassPortador

On Error GoTo Erro_Traz_Dados_Pagamento

    objMovCCI.lNumMovto = objBaixaPagar.lNumMovConta

    'Lê o Movimento
    lErro = CF("MovContaCorrente_Le",objMovCCI)
    If lErro <> SUCESSO And lErro <> 11893 Then Error 42818

    'Se não encontrou Movimento --> erro
    If lErro = 11893 Then Error 42825

    'Lê a Conta Corrente
    lErro = CF("ContaCorrenteInt_Le",objMovCCI.iCodConta, objContaCorrente)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 42819

    'Se não encontrou a conta Corrente --> erro
    If lErro <> SUCESSO Then Error 42820

    'Coloca os dados na tela
    ContaCorrente.Caption = objContaCorrente.sNomeReduzido
    ValorPagoPagto.Caption = Format(objMovCCI.dValor, "Standard")
    Historico.Caption = objMovCCI.sHistorico
    NumOuSequencial.Caption = IIf(objMovCCI.lNumero <> 0, CStr(objMovCCI.lNumero), "")

    objPortador.iCodigo = objMovCCI.iPortador

    If objPortador.iCodigo <> 0 Then

        'Lê o Portador
        lErro = CF("Portador_Le",objPortador)
        If lErro <> SUCESSO And lErro <> 15971 Then Error 43143

        'Se não achou o Portador --> erro
        If lErro <> SUCESSO Then Error 43144

        Portador.Caption = objMovCCI.iPortador & SEPARADOR & objPortador.sNomeReduzido

    End If

    objBorderoPagto.dtDataEmissao = objMovCCI.dtDataMovimento
    objBorderoPagto.iCodConta = objMovCCI.iCodConta
    objBorderoPagto.lNumero = objMovCCI.lNumero

    If objMovCCI.iTipoMeioPagto = Cheque Then
        TipoMeioPagto(0).Value = True

    ElseIf objMovCCI.iTipoMeioPagto = BORDERO Then
        TipoMeioPagto(1).Value = True

    ElseIf objMovCCI.iTipoMeioPagto = DINHEIRO Then
        TipoMeioPagto(2).Value = True
    End If

    Traz_Dados_Pagamento = SUCESSO

    Exit Function

Erro_Traz_Dados_Pagamento:

    Traz_Dados_Pagamento = Err

    Select Case Err

        Case 42818, 42819, 43143

        Case 42825
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_CADASTRADO3", Err, objMovCCI.lNumMovto)

        Case 42820
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, objContaCorrente.iCodigo)

        Case 43144
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PORTADOR_NAO_CADASTRADO1", Err, objPortador.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175129)

    End Select

    Exit Function

End Function

Private Function Traz_Dados_Credito_Fornecedor(objBaixaPagar As ClassBaixaPagar) As Long

Dim lErro As Long
Dim objCreditoPag As New ClassCreditoPagar

On Error GoTo Erro_Traz_Dados_Credito_Fornecedor

    objCreditoPag.lNumIntDoc = objBaixaPagar.lNumIntDoc

    'Lê o Crédito Pagar
    lErro = CF("CreditoPagar_Le",objCreditoPag)
    If lErro <> AD_SQL_SUCESSO And 17071 Then Error 42821
    If lErro <> SUCESSO Then Error 42822

    'Coloca os dados na tela
    DataEmissaoCred.Caption = Format(objCreditoPag.dtDataEmissao, "dd/mm/yyyy")
    NumTitulo.Caption = objCreditoPag.lNumTitulo
    SaldoCredito.Caption = Format(objCreditoPag.dSaldo, "Standard")
    SiglaDocumentoCR.Caption = objCreditoPag.sSiglaDocumento
    ValorCredito.Caption = Format(objCreditoPag.dValorTotal, "Standard")
    FilialEmpresaCR.Caption = objCreditoPag.iFilialEmpresa

    Traz_Dados_Credito_Fornecedor = SUCESSO

    Exit Function

Erro_Traz_Dados_Credito_Fornecedor:

    Traz_Dados_Credito_Fornecedor = Err

    Select Case Err

        Case 42821

        Case 42822
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CREDITO_PAG_FORN_INEXISTENTE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175130)

    End Select

    Exit Function

End Function

Private Function Traz_Dados_Pagamento_Antecipado(objBaixaPagar As ClassBaixaPagar) As Long

Dim lErro As Long
Dim objMovCCI As New ClassMovContaCorrente
Dim objCCI As New ClassContasCorrentesInternas
Dim objAntecipPag As New ClassAntecipPag

On Error GoTo Erro_Traz_Dados_Pagamento_Antecipado

    objMovCCI.lNumMovto = objBaixaPagar.lNumMovConta

    'Lê o movimento  da Baixa
    lErro = CF("MovContaCorrente_Le",objMovCCI)
    If lErro <> SUCESSO And lErro <> 11893 Then Error 42824
    If lErro = 11893 Then Error 42826 'Não encontrou

    'Lê a Conta Corrente
    lErro = CF("ContaCorrenteInt_Le",objMovCCI.iCodConta, objCCI)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 42827
    If lErro <> SUCESSO Then Error 42828 'Não encontrou

    objAntecipPag.lNumMovto = objMovCCI.lNumMovto

    lErro = CF("AntecipPag_Le_NumMovto",objAntecipPag)
    If lErro <> AD_SQL_SUCESSO And lErro <> 42845 Then Error 42841
    If lErro = 42845 Then Error 42846

    'Coloca os dados na tela
    DataMovimento.Caption = Format(objMovCCI.dtDataMovimento, "dd/mm/yyyy")
    ValorPA.Caption = Format(objMovCCI.dValor, "Standard")
    FilialEmpresaPA.Caption = objMovCCI.iFilialEmpresa
    CCIntNomeReduzido.Caption = objCCI.sNomeReduzido
    NumeroMP.Caption = objMovCCI.lNumero
    If objMovCCI.iTipoMeioPagto = DINHEIRO Then
        MeioPagtoDescricao.Caption = "Dinheiro"
    ElseIf objMovCCI.iTipoMeioPagto = Cheque Then
        MeioPagtoDescricao.Caption = "Cheque"
    ElseIf objMovCCI.iTipoMeioPagto = BORDERO Then
        MeioPagtoDescricao.Caption = "Borderô"
    End If

    Traz_Dados_Pagamento_Antecipado = SUCESSO

    Exit Function

Erro_Traz_Dados_Pagamento_Antecipado:

    Traz_Dados_Pagamento_Antecipado = Err

    Select Case Err

        Case 42824, 42827, 42841

        Case 42826
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_CADASTRADO3", Err, objMovCCI.lNumMovto)

        Case 42828
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, objCCI.iCodigo)

        Case 42846
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PAGTO_ANTECIPADO_INEXISTENTE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175131)

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

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Consulta de Títulos a Pagar"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TituloPagar_Consulta"
    
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
        
        If Me.ActiveControl Is Fornecedor Then
            Call FornecedorLabel_Click
        ElseIf Me.ActiveControl Is NumeroTitulo Then
            Call NumeroLabel_Click
        End If
    
    End If
    
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

Private Sub CTBLabel1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel1, Source, X, Y)
End Sub

Private Sub CTBLabel1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel1, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelCcl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelCcl, Source, X, Y)
End Sub

Private Sub CTBLabelCcl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelCcl, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelContas, Source, X, Y)
End Sub

Private Sub CTBLabelContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelContas, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelHistoricos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelHistoricos, Source, X, Y)
End Sub

Private Sub CTBLabelHistoricos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelHistoricos, Button, Shift, X, Y)
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

Private Sub ValorIPI_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorIPI, Source, X, Y)
End Sub

Private Sub ValorIPI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorIPI, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub ValorICMSSubst_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorICMSSubst, Source, X, Y)
End Sub

Private Sub ValorICMSSubst_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorICMSSubst, Button, Shift, X, Y)
End Sub

Private Sub ValorICMS_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorICMS, Source, X, Y)
End Sub

Private Sub ValorICMS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorICMS, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub ValorINSS_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorINSS, Source, X, Y)
End Sub

Private Sub ValorINSS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorINSS, Button, Shift, X, Y)
End Sub

Private Sub ValorIRRF_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorIRRF, Source, X, Y)
End Sub

Private Sub ValorIRRF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorIRRF, Button, Shift, X, Y)
End Sub

Private Sub OutrasDespesas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(OutrasDespesas, Source, X, Y)
End Sub

Private Sub OutrasDespesas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(OutrasDespesas, Button, Shift, X, Y)
End Sub

Private Sub ValorSeguro_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorSeguro, Source, X, Y)
End Sub

Private Sub ValorSeguro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorSeguro, Button, Shift, X, Y)
End Sub

Private Sub ValorFrete_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorFrete, Source, X, Y)
End Sub

Private Sub ValorFrete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorFrete, Button, Shift, X, Y)
End Sub

Private Sub ValorProdutos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorProdutos, Source, X, Y)
End Sub

Private Sub ValorProdutos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorProdutos, Button, Shift, X, Y)
End Sub

Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub

Private Sub Label18_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label18, Source, X, Y)
End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label18, Button, Shift, X, Y)
End Sub

Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub

Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
End Sub

Private Sub Label22_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label22, Source, X, Y)
End Sub

Private Sub Label22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label22, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LblNumPC_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblNumPC, Source, X, Y)
End Sub

Private Sub LblNumPC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblNumPC, Button, Shift, X, Y)
End Sub

Private Sub LblFilialPC_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblFilialPC, Source, X, Y)
End Sub

Private Sub LblFilialPC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblFilialPC, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub DataEmissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataEmissao, Source, X, Y)
End Sub

Private Sub DataEmissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataEmissao, Button, Shift, X, Y)
End Sub

Private Sub ValorTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorTotal, Source, X, Y)
End Sub

Private Sub ValorTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorTotal, Button, Shift, X, Y)
End Sub

Private Sub NumPC_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumPC, Source, X, Y)
End Sub

Private Sub NumPC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumPC, Button, Shift, X, Y)
End Sub

Private Sub ComboFilialPC_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ComboFilialPC, Source, X, Y)
End Sub

Private Sub ComboFilialPC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ComboFilialPC, Button, Shift, X, Y)
End Sub

Private Sub Label29_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label29, Source, X, Y)
End Sub

Private Sub Label29_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label29, Button, Shift, X, Y)
End Sub

Private Sub Label28_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label28, Source, X, Y)
End Sub

Private Sub Label28_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label28, Button, Shift, X, Y)
End Sub

Private Sub Label27_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label27, Source, X, Y)
End Sub

Private Sub Label27_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label27, Button, Shift, X, Y)
End Sub

Private Sub Label26_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label26, Source, X, Y)
End Sub

Private Sub Label26_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label26, Button, Shift, X, Y)
End Sub

Private Sub Label32_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label32, Source, X, Y)
End Sub

Private Sub Label32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label32, Button, Shift, X, Y)
End Sub

Private Sub ContaCorrente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaCorrente, Source, X, Y)
End Sub

Private Sub ContaCorrente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaCorrente, Button, Shift, X, Y)
End Sub

Private Sub ValorPagoPagto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorPagoPagto, Source, X, Y)
End Sub

Private Sub ValorPagoPagto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorPagoPagto, Button, Shift, X, Y)
End Sub

Private Sub Historico_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Historico, Source, X, Y)
End Sub

Private Sub Historico_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Historico, Button, Shift, X, Y)
End Sub

Private Sub NumOuSequencial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumOuSequencial, Source, X, Y)
End Sub

Private Sub NumOuSequencial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumOuSequencial, Button, Shift, X, Y)
End Sub

Private Sub Portador_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Portador, Source, X, Y)
End Sub

Private Sub Portador_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Portador, Button, Shift, X, Y)
End Sub

Private Sub NumeroMP_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroMP, Source, X, Y)
End Sub

Private Sub NumeroMP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroMP, Button, Shift, X, Y)
End Sub

Private Sub CCIntNomeReduzido_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CCIntNomeReduzido, Source, X, Y)
End Sub

Private Sub CCIntNomeReduzido_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CCIntNomeReduzido, Button, Shift, X, Y)
End Sub

Private Sub FilialEmpresaPA_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialEmpresaPA, Source, X, Y)
End Sub

Private Sub FilialEmpresaPA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialEmpresaPA, Button, Shift, X, Y)
End Sub

Private Sub ValorPA_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorPA, Source, X, Y)
End Sub

Private Sub ValorPA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorPA, Button, Shift, X, Y)
End Sub

Private Sub MeioPagtoDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MeioPagtoDescricao, Source, X, Y)
End Sub

Private Sub MeioPagtoDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MeioPagtoDescricao, Button, Shift, X, Y)
End Sub

Private Sub DataMovimento_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataMovimento, Source, X, Y)
End Sub

Private Sub DataMovimento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataMovimento, Button, Shift, X, Y)
End Sub

Private Sub Label33_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label33, Source, X, Y)
End Sub

Private Sub Label33_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label33, Button, Shift, X, Y)
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

Private Sub Label23_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label23, Source, X, Y)
End Sub

Private Sub Label23_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label23, Button, Shift, X, Y)
End Sub

Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

Private Sub Label34_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label34, Source, X, Y)
End Sub

Private Sub Label34_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label34, Button, Shift, X, Y)
End Sub

Private Sub Label48_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label48, Source, X, Y)
End Sub

Private Sub Label48_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label48, Button, Shift, X, Y)
End Sub

Private Sub Label37_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label37, Source, X, Y)
End Sub

Private Sub Label37_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label37, Button, Shift, X, Y)
End Sub

Private Sub Label38_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label38, Source, X, Y)
End Sub

Private Sub Label38_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label38, Button, Shift, X, Y)
End Sub

Private Sub Label39_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label39, Source, X, Y)
End Sub

Private Sub Label39_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label39, Button, Shift, X, Y)
End Sub

Private Sub Label40_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label40, Source, X, Y)
End Sub

Private Sub Label40_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label40, Button, Shift, X, Y)
End Sub

Private Sub DataEmissaoCred_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataEmissaoCred, Source, X, Y)
End Sub

Private Sub DataEmissaoCred_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataEmissaoCred, Button, Shift, X, Y)
End Sub

Private Sub SiglaDocumentoCR_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SiglaDocumentoCR, Source, X, Y)
End Sub

Private Sub SiglaDocumentoCR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SiglaDocumentoCR, Button, Shift, X, Y)
End Sub

Private Sub ValorCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorCredito, Source, X, Y)
End Sub

Private Sub ValorCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorCredito, Button, Shift, X, Y)
End Sub

Private Sub FilialEmpresaCR_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialEmpresaCR, Source, X, Y)
End Sub

Private Sub FilialEmpresaCR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialEmpresaCR, Button, Shift, X, Y)
End Sub

Private Sub NumTitulo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumTitulo, Source, X, Y)
End Sub

Private Sub NumTitulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumTitulo, Button, Shift, X, Y)
End Sub

Private Sub SaldoCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SaldoCredito, Source, X, Y)
End Sub

Private Sub SaldoCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SaldoCredito, Button, Shift, X, Y)
End Sub

Private Sub DataBaixa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataBaixa, Source, X, Y)
End Sub

Private Sub DataBaixa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataBaixa, Button, Shift, X, Y)
End Sub

Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub

Private Sub ValorPago_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorPago, Source, X, Y)
End Sub

Private Sub ValorPago_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorPago, Button, Shift, X, Y)
End Sub

Private Sub Juros_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Juros, Source, X, Y)
End Sub

Private Sub Juros_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Juros, Button, Shift, X, Y)
End Sub

Private Sub Multa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Multa, Source, X, Y)
End Sub

Private Sub Multa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Multa, Button, Shift, X, Y)
End Sub

Private Sub Desconto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Desconto, Source, X, Y)
End Sub

Private Sub Desconto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Desconto, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub ValorBaixado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorBaixado, Source, X, Y)
End Sub

Private Sub ValorBaixado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorBaixado, Button, Shift, X, Y)
End Sub

Private Sub Label31_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label31, Source, X, Y)
End Sub

Private Sub Label31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label31, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub

Private Sub NumeroLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroLabel, Source, X, Y)
End Sub

Private Sub NumeroLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroLabel, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub LabelTipo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTipo, Source, X, Y)
End Sub

Private Sub LabelTipo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTipo, Button, Shift, X, Y)
End Sub


Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
   Height = UserControl.Height
End Property

Public Property Get Width() As Long
   Width = UserControl.Width
End Property

Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel,Opcao)
End Sub

