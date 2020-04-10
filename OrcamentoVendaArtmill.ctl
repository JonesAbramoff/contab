VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl OrcamentoVenda 
   ClientHeight    =   5880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4590
      Index           =   2
      Left            =   105
      TabIndex        =   15
      Top             =   1125
      Visible         =   0   'False
      Width           =   9225
      Begin VB.CommandButton BotaoGrade 
         Caption         =   "Grade ..."
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
         Left            =   345
         TabIndex        =   181
         Top             =   4140
         Width           =   1365
      End
      Begin VB.CommandButton BotaoProdutos 
         Caption         =   "Produtos"
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
         Left            =   7680
         TabIndex        =   29
         Top             =   4140
         Width           =   1365
      End
      Begin VB.Frame Frame2 
         Caption         =   "Valores"
         Height          =   1290
         Index           =   4
         Left            =   225
         TabIndex        =   55
         Top             =   2745
         Width           =   8865
         Begin MSMask.MaskEdBox ValorFrete 
            Height          =   285
            Left            =   1695
            TabIndex        =   26
            Top             =   900
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDesconto 
            Height          =   285
            Left            =   450
            TabIndex        =   25
            Top             =   450
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDespesas 
            Height          =   285
            Left            =   4320
            TabIndex        =   28
            Top             =   900
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorSeguro 
            Height          =   285
            Left            =   2985
            TabIndex        =   27
            Top             =   900
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Desconto        Base ICMS          ICMS         Base ICMS Subst    ICMS Subst       Produtos"
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
            Left            =   495
            TabIndex        =   57
            Top             =   225
            Width           =   7695
         End
         Begin VB.Label ICMSSubstValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6000
            TabIndex        =   58
            Top             =   450
            Width           =   1125
         End
         Begin VB.Label ICMSSubstBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4320
            TabIndex        =   59
            Top             =   450
            Width           =   1500
         End
         Begin VB.Label ICMSValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3015
            TabIndex        =   60
            Top             =   450
            Width           =   1125
         End
         Begin VB.Label ICMSBase1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1710
            TabIndex        =   61
            Top             =   450
            Width           =   1125
         End
         Begin VB.Label ValorProdutos 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7305
            TabIndex        =   62
            Top             =   450
            Width           =   1125
         End
         Begin VB.Label IPIValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5985
            TabIndex        =   63
            Top             =   900
            Width           =   1125
         End
         Begin VB.Label ValorTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7320
            TabIndex        =   64
            Top             =   900
            Width           =   1125
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Frete             Seguro              Despesas               IPI                Total"
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
            Left            =   2025
            TabIndex        =   65
            Top             =   720
            Width           =   6030
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Itens"
         Height          =   2685
         Index           =   3
         Left            =   225
         TabIndex        =   54
         Top             =   0
         Width           =   8865
         Begin VB.TextBox DescricaoProduto2 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   4245
            MaxLength       =   50
            TabIndex        =   192
            Top             =   1185
            Width           =   1305
         End
         Begin VB.TextBox DescricaoProduto3 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   5685
            MaxLength       =   50
            TabIndex        =   191
            Top             =   1170
            Width           =   1305
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   225
            Left            =   330
            TabIndex        =   16
            Top             =   360
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.TextBox DescricaoProduto 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   4005
            MaxLength       =   50
            TabIndex        =   23
            Top             =   765
            Width           =   2655
         End
         Begin VB.ComboBox UnidadeMed 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "OrcamentoVendaArtmill.ctx":0000
            Left            =   1575
            List            =   "OrcamentoVendaArtmill.ctx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   240
            Width           =   720
         End
         Begin MSMask.MaskEdBox DataEntrega 
            Height          =   225
            Left            =   2640
            TabIndex        =   21
            Top             =   660
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto 
            Height          =   225
            Left            =   1440
            TabIndex        =   19
            Top             =   585
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
         Begin MSMask.MaskEdBox PercentDesc 
            Height          =   225
            Left            =   270
            TabIndex        =   17
            Top             =   675
            Width           =   990
            _ExtentX        =   1746
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
         Begin MSMask.MaskEdBox PrecoUnitario 
            Height          =   225
            Left            =   4185
            TabIndex        =   22
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
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
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   225
            Left            =   2580
            TabIndex        =   20
            Top             =   315
            Width           =   1500
            _ExtentX        =   2646
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrecoTotal 
            Height          =   225
            Left            =   5670
            TabIndex        =   24
            Top             =   360
            Width           =   1185
            _ExtentX        =   2090
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
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   1455
            Left            =   180
            TabIndex        =   47
            Top             =   240
            Width           =   8565
            _ExtentX        =   15108
            _ExtentY        =   2566
            _Version        =   393216
            Rows            =   21
            Cols            =   4
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
      Caption         =   "Tributacao"
      Height          =   4680
      Index           =   4
      Left            =   135
      TabIndex        =   73
      Top             =   1100
      Visible         =   0   'False
      Width           =   9195
      Begin VB.Frame FrameTributacao 
         BorderStyle     =   0  'None
         Caption         =   "Resumo"
         Height          =   4020
         Index           =   1
         Left            =   270
         TabIndex        =   135
         Top             =   480
         Width           =   8700
         Begin VB.Frame Frame10 
            Caption         =   "CSLL"
            Height          =   570
            Index           =   8
            Left            =   4335
            TabIndex        =   188
            Top             =   3300
            Width           =   1860
            Begin MSMask.MaskEdBox CSLLRetido 
               Height          =   285
               Left            =   750
               TabIndex        =   189
               Top             =   195
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Retido:"
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
               Left            =   75
               TabIndex        =   190
               Top             =   270
               Width           =   630
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "COFINS"
            Height          =   570
            Index           =   7
            Left            =   2325
            TabIndex        =   185
            Top             =   3300
            Width           =   1860
            Begin MSMask.MaskEdBox COFINSRetido 
               Height          =   285
               Left            =   810
               TabIndex        =   186
               Top             =   195
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Retido:"
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
               Index           =   21
               Left            =   135
               TabIndex        =   187
               Top             =   270
               Width           =   630
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "PIS"
            Height          =   570
            Index           =   19
            Left            =   315
            TabIndex        =   182
            Top             =   3300
            Width           =   1860
            Begin MSMask.MaskEdBox PISRetido 
               Height          =   285
               Left            =   765
               TabIndex        =   183
               Top             =   195
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Retido:"
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
               Index           =   25
               Left            =   90
               TabIndex        =   184
               Top             =   270
               Width           =   630
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "ICMS"
            Height          =   984
            Left            =   330
            TabIndex        =   151
            Top             =   2205
            Width           =   7185
            Begin VB.Frame Frame10 
               Caption         =   "Substituicao"
               Height          =   780
               Index           =   0
               Left            =   3300
               TabIndex        =   152
               Top             =   132
               Width           =   3600
               Begin VB.Label ICMSSubstValor 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   312
                  Left            =   1920
                  TabIndex        =   156
                  Top             =   396
                  Width           =   1080
               End
               Begin VB.Label ICMSSubstBase 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   405
                  TabIndex        =   155
                  Top             =   390
                  Width           =   1080
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  Caption         =   "Base"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   192
                  Index           =   0
                  Left            =   384
                  TabIndex        =   154
                  Top             =   192
                  Width           =   444
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Valor"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   192
                  Index           =   4
                  Left            =   1956
                  TabIndex        =   153
                  Top             =   192
                  Width           =   456
               End
            End
            Begin VB.Label ICMSValor 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1830
               TabIndex        =   160
               Top             =   525
               Width           =   1080
            End
            Begin VB.Label ICMSBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   270
               TabIndex        =   159
               Top             =   525
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Base"
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
               Left            =   300
               TabIndex        =   158
               Top             =   300
               Width           =   450
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor"
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
               Index           =   14
               Left            =   1845
               TabIndex        =   157
               Top             =   300
               Width           =   450
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   "ISS"
            Height          =   1455
            Left            =   2460
            TabIndex        =   146
            Top             =   750
            Width           =   2955
            Begin VB.CheckBox ISSIncluso 
               Caption         =   "Incluso"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   1845
               TabIndex        =   165
               Top             =   288
               Width           =   1020
            End
            Begin MSMask.MaskEdBox ISSAliquota 
               Height          =   312
               Left            =   660
               TabIndex        =   166
               Top             =   648
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ISSValor 
               Height          =   315
               Left            =   660
               TabIndex        =   167
               Top             =   1050
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label ISSBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   660
               TabIndex        =   150
               Top             =   255
               Width           =   1080
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               Caption         =   "Base:"
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
               Left            =   120
               TabIndex        =   149
               Top             =   285
               Width           =   495
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               Caption         =   "%:"
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
               Left            =   405
               TabIndex        =   148
               Top             =   705
               Width           =   210
            End
            Begin VB.Label Label44 
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
               Height          =   195
               Index           =   2
               Left            =   105
               TabIndex        =   147
               Top             =   1095
               Width           =   510
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "IPI"
            Height          =   1455
            Index           =   0
            Left            =   330
            TabIndex        =   141
            Top             =   750
            Width           =   2028
            Begin VB.Label IPIValor 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   690
               TabIndex        =   145
               Top             =   870
               Width           =   1080
            End
            Begin VB.Label IPIBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   690
               TabIndex        =   144
               Top             =   345
               Width           =   1080
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               Caption         =   "Base:"
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
               Left            =   135
               TabIndex        =   143
               Top             =   375
               Width           =   495
            End
            Begin VB.Label Label44 
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
               Height          =   195
               Index           =   7
               Left            =   135
               TabIndex        =   142
               Top             =   930
               Width           =   510
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "IR"
            Height          =   1455
            Left            =   5550
            TabIndex        =   136
            Top             =   750
            Width           =   1965
            Begin MSMask.MaskEdBox IRAliquota 
               Height          =   315
               Left            =   690
               TabIndex        =   168
               Top             =   675
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#0.#0\%"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ValorIRRF 
               Height          =   315
               Left            =   690
               TabIndex        =   170
               Top             =   1065
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label IRBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   675
               TabIndex        =   140
               Top             =   285
               Width           =   1080
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               Caption         =   "Base:"
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
               Left            =   135
               TabIndex        =   139
               Top             =   315
               Width           =   495
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               Caption         =   "%:"
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
               Left            =   390
               TabIndex        =   138
               Top             =   735
               Width           =   210
            End
            Begin VB.Label Label44 
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
               Height          =   195
               Index           =   5
               Left            =   105
               TabIndex        =   137
               Top             =   1110
               Width           =   510
            End
         End
         Begin VB.CommandButton TributacaoRecalcular 
            Caption         =   "Recalcular Tributação"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   6270
            Picture         =   "OrcamentoVendaArtmill.ctx":0004
            Style           =   1  'Graphical
            TabIndex        =   172
            Top             =   3255
            Width           =   1260
         End
         Begin MSMask.MaskEdBox TipoTributacao 
            Height          =   300
            Left            =   2196
            TabIndex        =   164
            Top             =   408
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   529
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
         Begin VB.Label LblNatOpEspelho 
            AutoSize        =   -1  'True
            Caption         =   "Natureza de Oper.:"
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
            TabIndex        =   171
            Top             =   30
            Width           =   1575
         End
         Begin VB.Label DescNatOp 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2850
            TabIndex        =   169
            Top             =   15
            Width           =   5235
         End
         Begin VB.Label NatOpEspelho 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2190
            TabIndex        =   163
            Top             =   15
            Width           =   525
         End
         Begin VB.Label LblTipoTrib 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Tributação:"
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
            Left            =   435
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   162
            Top             =   450
            Width           =   1695
         End
         Begin VB.Label DescTipoTrib 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2805
            TabIndex        =   161
            Top             =   405
            Width           =   4710
         End
      End
      Begin VB.Frame FrameTributacao 
         BorderStyle     =   0  'None
         Caption         =   "Detalhamento"
         Height          =   4140
         Index           =   2
         Left            =   270
         TabIndex        =   78
         Top             =   420
         Visible         =   0   'False
         Width           =   8700
         Begin VB.Frame Frame1 
            Caption         =   "Sobre"
            Height          =   1185
            Index           =   0
            Left            =   132
            TabIndex        =   113
            Top             =   -15
            Width           =   8490
            Begin VB.Frame FrameOutrosTrib 
               Height          =   645
               Left            =   120
               TabIndex        =   126
               Top             =   465
               Visible         =   0   'False
               Width           =   8235
               Begin VB.Label Label1 
                  Caption         =   "Outras Desp.:"
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
                  Index           =   8
                  Left            =   3750
                  TabIndex        =   134
                  Top             =   285
                  Width           =   1185
               End
               Begin VB.Label LabelValorOutrasDespesas 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Left            =   4950
                  TabIndex        =   133
                  Top             =   270
                  Width           =   1140
               End
               Begin VB.Label Label1 
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
                  Height          =   225
                  Index           =   10
                  Left            =   1860
                  TabIndex        =   132
                  Top             =   285
                  Width           =   705
               End
               Begin VB.Label LabelValorSeguro 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Left            =   2565
                  TabIndex        =   131
                  Top             =   270
                  Width           =   1140
               End
               Begin VB.Label Label1 
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
                  Height          =   225
                  Index           =   11
                  Left            =   6135
                  TabIndex        =   130
                  Top             =   285
                  Width           =   870
               End
               Begin VB.Label LabelValorDesconto 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Left            =   7035
                  TabIndex        =   129
                  Top             =   255
                  Width           =   1140
               End
               Begin VB.Label Label1 
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
                  Height          =   225
                  Index           =   15
                  Left            =   75
                  TabIndex        =   128
                  Top             =   285
                  Width           =   510
               End
               Begin VB.Label LabelValorFrete 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   270
                  Left            =   600
                  TabIndex        =   127
                  Top             =   270
                  Width           =   1140
               End
            End
            Begin VB.Frame FrameItensTrib 
               Caption         =   "Item"
               Height          =   645
               Left            =   120
               TabIndex        =   119
               Top             =   465
               Width           =   8235
               Begin VB.ComboBox ComboItensTrib 
                  Height          =   315
                  Left            =   144
                  Style           =   2  'Dropdown List
                  TabIndex        =   120
                  Top             =   228
                  Width           =   3015
               End
               Begin VB.Label LabelUMItem 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   7380
                  TabIndex        =   125
                  Top             =   228
                  Width           =   765
               End
               Begin VB.Label LabelQtdeItem 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   6495
                  TabIndex        =   124
                  Top             =   228
                  Width           =   840
               End
               Begin VB.Label LabelValorItem 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   4095
                  TabIndex        =   123
                  Top             =   210
                  Width           =   1140
               End
               Begin VB.Label Label1 
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
                  Height          =   225
                  Index           =   6
                  Left            =   3495
                  TabIndex        =   122
                  Top             =   285
                  Width           =   570
               End
               Begin VB.Label Label1 
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
                  Index           =   3
                  Left            =   5370
                  TabIndex        =   121
                  Top             =   270
                  Width           =   1065
               End
            End
            Begin VB.OptionButton TribSobreOutrasDesp 
               Caption         =   "Outras Despesas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   4560
               TabIndex        =   118
               Top             =   210
               Width           =   1965
            End
            Begin VB.OptionButton TribSobreSeguro 
               Caption         =   "Seguro"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   2958
               TabIndex        =   117
               Top             =   210
               Width           =   960
            End
            Begin VB.OptionButton TribSobreDesconto 
               Caption         =   "Desconto"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   7170
               TabIndex        =   116
               Top             =   225
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.OptionButton TribSobreFrete 
               Caption         =   "Frete"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   1500
               TabIndex        =   115
               Top             =   210
               Width           =   816
            End
            Begin VB.OptionButton TribSobreItem 
               Caption         =   "Item"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   108
               TabIndex        =   114
               Top             =   210
               Width           =   750
            End
         End
         Begin VB.Frame Frame15 
            Height          =   2700
            Left            =   135
            TabIndex        =   79
            Top             =   1185
            Width           =   8508
            Begin VB.Frame Frame10 
               Caption         =   "ICMS"
               Height          =   1620
               Index           =   1
               Left            =   135
               TabIndex        =   90
               Top             =   975
               Width           =   5865
               Begin VB.Frame Frame2 
                  Caption         =   "Substituição"
                  Height          =   1368
                  Index           =   1
                  Left            =   3630
                  TabIndex        =   92
                  Top             =   144
                  Width           =   2004
                  Begin MSMask.MaskEdBox ICMSSubstValorItem 
                     Height          =   285
                     Left            =   675
                     TabIndex        =   93
                     Top             =   1005
                     Width           =   1110
                     _ExtentX        =   1958
                     _ExtentY        =   503
                     _Version        =   393216
                     PromptInclude   =   0   'False
                     MaxLength       =   15
                     Format          =   "#,##0.0000"
                     PromptChar      =   " "
                  End
                  Begin MSMask.MaskEdBox ICMSSubstAliquotaItem 
                     Height          =   288
                     Left            =   672
                     TabIndex        =   94
                     Top             =   618
                     Width           =   1116
                     _ExtentX        =   1958
                     _ExtentY        =   503
                     _Version        =   393216
                     PromptInclude   =   0   'False
                     MaxLength       =   15
                     Format          =   "#0.#0\%"
                     PromptChar      =   " "
                  End
                  Begin MSMask.MaskEdBox ICMSSubstBaseItem 
                     Height          =   288
                     Left            =   672
                     TabIndex        =   95
                     Top             =   252
                     Width           =   1116
                     _ExtentX        =   1958
                     _ExtentY        =   503
                     _Version        =   393216
                     PromptInclude   =   0   'False
                     MaxLength       =   15
                     Format          =   "#,##0.00"
                     PromptChar      =   " "
                  End
                  Begin VB.Label Label19 
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
                     Height          =   195
                     Index           =   1
                     Left            =   90
                     TabIndex        =   98
                     Top             =   1020
                     Width           =   510
                  End
                  Begin VB.Label Label18 
                     AutoSize        =   -1  'True
                     Caption         =   "Aliq.:"
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
                     Left            =   165
                     TabIndex        =   97
                     Top             =   660
                     Width           =   450
                  End
                  Begin VB.Label Label1 
                     AutoSize        =   -1  'True
                     Caption         =   "Base:"
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
                     Index           =   7
                     Left            =   120
                     TabIndex        =   96
                     Top             =   315
                     Width           =   495
                  End
               End
               Begin VB.ComboBox ComboICMSTipo 
                  Height          =   315
                  Left            =   135
                  Style           =   2  'Dropdown List
                  TabIndex        =   91
                  Top             =   228
                  Width           =   3405
               End
               Begin MSMask.MaskEdBox ICMSValorItem 
                  Height          =   285
                  Left            =   2400
                  TabIndex        =   99
                  Top             =   1035
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.0000"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox ICMSAliquotaItem 
                  Height          =   285
                  Left            =   2385
                  TabIndex        =   100
                  Top             =   630
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox ICMSPercRedBaseItem 
                  Height          =   285
                  Left            =   1065
                  TabIndex        =   101
                  Top             =   1005
                  Width           =   660
                  _ExtentX        =   1164
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox ICMSBaseItem 
                  Height          =   285
                  Left            =   615
                  TabIndex        =   102
                  Top             =   630
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00"
                  PromptChar      =   " "
               End
               Begin VB.Label Label15 
                  AutoSize        =   -1  'True
                  Caption         =   "Red. Base:"
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
                  Left            =   75
                  TabIndex        =   106
                  Top             =   1035
                  Width           =   960
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  Caption         =   "Aliq.:"
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
                  Left            =   1890
                  TabIndex        =   105
                  Top             =   645
                  Width           =   450
               End
               Begin VB.Label Label16 
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
                  Height          =   195
                  Index           =   1
                  Left            =   1830
                  TabIndex        =   104
                  Top             =   1035
                  Width           =   510
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  Caption         =   "Base:"
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
                  Left            =   90
                  TabIndex        =   103
                  Top             =   645
                  Width           =   495
               End
            End
            Begin VB.Frame IPIItemFrame 
               Caption         =   "IPI"
               Height          =   2244
               Left            =   6060
               TabIndex        =   80
               Top             =   345
               Width           =   2376
               Begin VB.ComboBox ComboIPITipo 
                  Height          =   315
                  Left            =   252
                  Style           =   2  'Dropdown List
                  TabIndex        =   81
                  Top             =   240
                  Width           =   1716
               End
               Begin MSMask.MaskEdBox IPIPercRedBaseItem 
                  Height          =   285
                  Left            =   1260
                  TabIndex        =   82
                  Top             =   1035
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox IPIValorItem 
                  Height          =   285
                  Left            =   795
                  TabIndex        =   83
                  Top             =   1836
                  Width           =   1170
                  _ExtentX        =   2064
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.0000"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox IPIAliquotaItem 
                  Height          =   285
                  Left            =   795
                  TabIndex        =   84
                  Top             =   1455
                  Width           =   1170
                  _ExtentX        =   2064
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#0.#0\%"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox IPIBaseItem 
                  Height          =   285
                  Left            =   795
                  TabIndex        =   85
                  Top             =   630
                  Width           =   1170
                  _ExtentX        =   2064
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   15
                  Format          =   "#,##0.00"
                  PromptChar      =   " "
               End
               Begin VB.Label Label17 
                  AutoSize        =   -1  'True
                  Caption         =   "Red. Base:"
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
                  Left            =   255
                  TabIndex        =   89
                  Top             =   1080
                  Width           =   960
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  Caption         =   "Aliq.:"
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
                  Left            =   240
                  TabIndex        =   88
                  Top             =   1500
                  Width           =   450
               End
               Begin VB.Label Label13 
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
                  Height          =   195
                  Index           =   1
                  Left            =   210
                  TabIndex        =   87
                  Top             =   1890
                  Width           =   510
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  Caption         =   "Base:"
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
                  Left            =   255
                  TabIndex        =   86
                  Top             =   675
                  Width           =   495
               End
            End
            Begin MSMask.MaskEdBox NaturezaOpItem 
               Height          =   300
               Left            =   1515
               TabIndex        =   107
               Top             =   195
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox TipoTributacaoItem 
               Height          =   300
               Left            =   1860
               TabIndex        =   108
               Top             =   630
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   529
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
            Begin VB.Label NaturezaItemLabel 
               AutoSize        =   -1  'True
               Caption         =   "Natureza Oper.:"
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
               TabIndex        =   112
               Top             =   255
               Width           =   1365
            End
            Begin VB.Label DescTipoTribItem 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   2385
               TabIndex        =   111
               Top             =   645
               Width           =   3615
            End
            Begin VB.Label LabelDescrNatOpItem 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   2055
               TabIndex        =   110
               Top             =   210
               Width           =   3945
            End
            Begin VB.Label LblTipoTribItem 
               Caption         =   "Tipo de Tributação:"
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
               Left            =   105
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   109
               Top             =   660
               Width           =   1785
            End
         End
      End
      Begin MSComctlLib.TabStrip OpcaoTributacao 
         Height          =   4545
         Left            =   210
         TabIndex        =   77
         Top             =   60
         Width           =   8850
         _ExtentX        =   15610
         _ExtentY        =   8017
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Resumo"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Detalhamento"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4590
      Index           =   1
      Left            =   195
      TabIndex        =   1
      Top             =   1100
      Width           =   9165
      Begin VB.Frame Frame3 
         Caption         =   "Outros"
         Height          =   915
         Left            =   90
         TabIndex        =   74
         Top             =   3480
         Width           =   8865
         Begin MSMask.MaskEdBox PrazoValidade 
            Height          =   300
            Left            =   6615
            TabIndex        =   14
            Top             =   360
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   5
            Mask            =   "#####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Vendedor 
            Height          =   315
            Left            =   1980
            TabIndex        =   13
            Top             =   360
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin VB.Label VendedorLabel 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
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
            Left            =   1035
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   76
            Top             =   405
            Width           =   885
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Prazo de Validade:"
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
            Left            =   4950
            TabIndex        =   75
            Top             =   405
            Width           =   1620
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados do Cliente"
         Height          =   900
         Index           =   6
         Left            =   90
         TabIndex        =   52
         Top             =   1425
         Width           =   8865
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5475
            TabIndex        =   9
            Top             =   360
            Width           =   2145
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   300
            Left            =   1980
            TabIndex        =   8
            Top             =   360
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
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
            Left            =   4950
            TabIndex        =   66
            Top             =   405
            Width           =   465
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
            Left            =   1275
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   67
            Top             =   405
            Width           =   660
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Preços"
         Height          =   900
         Index           =   2
         Left            =   90
         TabIndex        =   53
         Top             =   2445
         Width           =   8865
         Begin VB.ComboBox CondicaoPagamento 
            Height          =   315
            Left            =   4530
            Sorted          =   -1  'True
            TabIndex        =   11
            Top             =   345
            Width           =   1815
         End
         Begin VB.ComboBox TabelaPreco 
            Height          =   315
            Left            =   1305
            TabIndex        =   10
            Top             =   345
            Width           =   1875
         End
         Begin MSMask.MaskEdBox PercAcrescFin 
            Height          =   315
            Left            =   7995
            TabIndex        =   12
            Top             =   345
            Width           =   765
            _ExtentX        =   1349
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
         Begin VB.Label CondPagtoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Cond Pagto:"
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
            Left            =   3390
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   68
            Top             =   405
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "% Acrésc Financ:"
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
            Left            =   6480
            TabIndex        =   69
            Top             =   405
            Width           =   1485
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tabela Preço:"
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
            Left            =   90
            TabIndex        =   70
            Top             =   405
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Identificação"
         Height          =   1200
         Index           =   0
         Left            =   90
         TabIndex        =   51
         Top             =   105
         Width           =   8865
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   3105
            Picture         =   "OrcamentoVendaArtmill.ctx":0176
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Numeração Automática"
            Top             =   300
            Width           =   300
         End
         Begin MSComCtl2.UpDown UpDownEmissao 
            Height          =   300
            Left            =   7305
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   285
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmissao 
            Height          =   300
            Left            =   6255
            TabIndex        =   5
            Top             =   285
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   300
            Left            =   2295
            TabIndex        =   3
            Top             =   285
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NaturezaOp 
            Height          =   300
            Left            =   2295
            TabIndex        =   7
            Top             =   750
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label NaturezaLabel 
            AutoSize        =   -1  'True
            Caption         =   "Natureza Operação:"
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
            Left            =   510
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   173
            Top             =   780
            Width           =   1725
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
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   0
            Left            =   5445
            TabIndex        =   71
            Top             =   330
            Width           =   765
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
            Left            =   1515
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   2
            Top             =   330
            Width           =   720
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   6675
      ScaleHeight     =   450
      ScaleWidth      =   2685
      TabIndex        =   175
      TabStop         =   0   'False
      Top             =   90
      Width           =   2745
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2205
         Picture         =   "OrcamentoVendaArtmill.ctx":0260
         Style           =   1  'Graphical
         TabIndex        =   180
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   1680
         Picture         =   "OrcamentoVendaArtmill.ctx":03DE
         Style           =   1  'Graphical
         TabIndex        =   179
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   345
         Left            =   1155
         Picture         =   "OrcamentoVendaArtmill.ctx":0910
         Style           =   1  'Graphical
         TabIndex        =   178
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   345
         Left            =   660
         Picture         =   "OrcamentoVendaArtmill.ctx":0A9A
         Style           =   1  'Graphical
         TabIndex        =   177
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   345
         Left            =   120
         Picture         =   "OrcamentoVendaArtmill.ctx":0BF4
         Style           =   1  'Graphical
         TabIndex        =   176
         ToolTipText     =   "Imprimir"
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.CheckBox ImprimeOrcamentoGravacao 
      Caption         =   "Imprimir o orçamento ao gravar"
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
      Left            =   240
      TabIndex        =   174
      Top             =   240
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4410
      Index           =   3
      Left            =   90
      TabIndex        =   48
      Top             =   1100
      Visible         =   0   'False
      Width           =   9240
      Begin VB.CheckBox CobrancaAutomatica 
         Caption         =   "Calcula cobrança automaticamente"
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
         Left            =   300
         TabIndex        =   30
         Top             =   150
         Value           =   1  'Checked
         Width           =   3360
      End
      Begin VB.Frame SSFrame3 
         Caption         =   "Cobrança"
         Height          =   3855
         Left            =   150
         TabIndex        =   56
         Top             =   435
         Width           =   8970
         Begin VB.CommandButton BotaoDataReferenciaDown 
            Height          =   150
            Left            =   3960
            Picture         =   "OrcamentoVendaArtmill.ctx":0CF6
            Style           =   1  'Graphical
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   390
            Width           =   240
         End
         Begin VB.CommandButton BotaoDataReferenciaUp 
            Height          =   150
            Left            =   3960
            Picture         =   "OrcamentoVendaArtmill.ctx":0D50
            Style           =   1  'Graphical
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   240
            Width           =   240
         End
         Begin VB.ComboBox TipoDesconto1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3075
            TabIndex        =   36
            Top             =   1215
            Width           =   1965
         End
         Begin VB.ComboBox TipoDesconto2 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3060
            TabIndex        =   39
            Top             =   1530
            Width           =   1965
         End
         Begin VB.ComboBox TipoDesconto3 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3105
            TabIndex        =   43
            Top             =   1845
            Width           =   1965
         End
         Begin MSMask.MaskEdBox Desconto1Percentual 
            Height          =   225
            Left            =   7470
            TabIndex        =   38
            Top             =   1260
            Width           =   900
            _ExtentX        =   1588
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
         Begin MSMask.MaskEdBox Desconto3Valor 
            Height          =   225
            Left            =   6105
            TabIndex        =   45
            Top             =   1905
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto3Ate 
            Height          =   225
            Left            =   4995
            TabIndex        =   44
            Top             =   1905
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto2Valor 
            Height          =   225
            Left            =   6135
            TabIndex        =   41
            Top             =   1590
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto2Ate 
            Height          =   225
            Left            =   4995
            TabIndex        =   40
            Top             =   1590
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto1Valor 
            Height          =   225
            Left            =   6120
            TabIndex        =   49
            Top             =   1260
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto1Ate 
            Height          =   225
            Left            =   4995
            TabIndex        =   37
            Top             =   1260
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataVencimento 
            Height          =   225
            Left            =   570
            TabIndex        =   34
            Top             =   1230
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorParcela 
            Height          =   240
            Left            =   1695
            TabIndex        =   35
            Top             =   1245
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto2Percentual 
            Height          =   225
            Left            =   7500
            TabIndex        =   42
            Top             =   1605
            Width           =   900
            _ExtentX        =   1588
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
         Begin MSMask.MaskEdBox Desconto3Percentual 
            Height          =   225
            Left            =   7455
            TabIndex        =   46
            Top             =   1935
            Width           =   900
            _ExtentX        =   1588
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
         Begin MSMask.MaskEdBox DataReferencia 
            Height          =   300
            Left            =   2850
            TabIndex        =   31
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridParcelas 
            Height          =   2745
            Left            =   180
            TabIndex        =   50
            Top             =   675
            Width           =   8595
            _ExtentX        =   15161
            _ExtentY        =   4842
            _Version        =   393216
            Rows            =   50
            Cols            =   6
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin VB.Label Label1 
            Caption         =   "Data de Referência:"
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
            Index           =   9
            Left            =   1020
            TabIndex        =   72
            Top             =   285
            Width           =   1740
         End
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5070
      Left            =   75
      TabIndex        =   0
      Top             =   735
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   8943
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cobrança"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tributação"
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
Attribute VB_Name = "OrcamentoVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Grid Itens
Public iGrid_Item_Col As Integer
Public iGrid_UMEstoque_Col As Integer
Public iGrid_Responsavel_Col As Integer
Public iGrid_ItemProduto_Col As Integer
Public iGrid_Produto_Col As Integer
Public iGrid_DescProduto_Col As Integer
Public iGrid_DescProduto2_Col As Integer
Public iGrid_DescProduto3_Col As Integer
Public iGrid_UnidadeMed_Col As Integer
Public iGrid_Quantidade_Col As Integer
Public iGrid_PrecoUnitario_Col As Integer
'precodesc
Public iGrid_PrecoDesc_Col As Integer
Public iGrid_PercDesc_Col As Integer
Public iGrid_Desconto_Col As Integer
Public iGrid_PrecoTotal_Col As Integer
Public iGrid_DataEntrega_Col As Integer

'Grid Parcelas
Public iGrid_Vencimento_col  As Integer
Public iGrid_ValorParcela_Col As Integer
Public iGrid_Desc1Codigo_Col As Integer
Public iGrid_Desc1Ate_Col As Integer
Public iGrid_Desc1Valor_Col As Integer
Public iGrid_Desc1Perc_Col As Integer
Public iGrid_Desc2Codigo_Col As Integer
Public iGrid_Desc2Ate_Col As Integer
Public iGrid_Desc2Valor_Col As Integer
Public iGrid_Desc2Perc_Col As Integer
Public iGrid_Desc3Codigo_Col As Integer
Public iGrid_Desc3Ate_Col As Integer
Public iGrid_Desc3Valor_Col As Integer
Public iGrid_Desc3Perc_Col As Integer

Dim giTrazendoTribTela As Integer
Dim giFrameAtual As Integer
Dim giFrameAtualTributacao As Integer
Dim gsCodigoAnterior  As String

Public iAlterado As Integer
Public iVendedorAlterado As Integer

Dim giClienteAlterado As Integer
Dim giFilialAlterada As Integer
Dim gdDesconto As Double
Dim giValorFreteAlterado As Integer
Dim giValorSeguroAlterado As Integer
Dim giValorDescontoAlterado As Integer
Dim giValorDespesasAlterado  As Integer
Dim giDataReferenciaAlterada As Integer
Dim giNaturezaOpAlterada As Integer

Dim giValorDescontoManual As Integer

Dim giPercAcresFinAlterado As Integer


Dim gobjOrcamentoVenda As New ClassOrcamentoVenda 'estrutura mantida para auxiliar a manutencao de informacoes p/tributacao
    'todos os dados relevantes p/tributacao dentro de gobjOrcamentoVenda estarao sincronizados com a tela antes da atualizacao da tributacao

Dim objGridItens As AdmGrid
Dim objGridParcelas As AdmGrid

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoCondPagto As AdmEvento
Attribute objEventoCondPagto.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Public WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoNaturezaOp As AdmEvento
Attribute objEventoNaturezaOp.VB_VarHelpID = -1
Private WithEvents objEventoTiposDeTributacao As AdmEvento
Attribute objEventoTiposDeTributacao.VB_VarHelpID = -1

Dim giLinhaAnterior As Integer
Dim giRecalculandoTributacao As Integer
Dim gcolTiposTribICMS As New Collection
Dim gcolTiposTribIPI As New Collection

'variaveis auxiliares para criacao da contabilizacao
Private gobjContabAutomatica As ClassContabAutomatica
Private gobjNFiscalCTB As ClassNFiscal
Private gobjOrcamentoVendaCTB As ClassOrcamentoVenda
Private giExercicio As Integer, giPeriodo As Integer
Private gcolAlmoxFilial As New Collection
Private gobjGeracaoNFiscal As ClassGeracaoNFiscal

'Constantes públicas dos tabs
Private Const TAB_Principal = 1
Private Const TAB_Itens = 2
Private Const TAB_Cobranca = 3
Private Const TAB_Tributacao = 4

'Property Variables:
Dim m_Caption As String
Dim gbCarregandoTela As Boolean
Dim iFrameAtual As Integer
Dim bTrazendoDoc As Boolean
Dim giPosCargaOk As Integer

Dim iValorIRRFAlterado As Integer

'Incluidos por Leo em 30/04/02 para tratamento da tributação
Dim giISSAliquotaAlterada As Integer
Dim giISSValorAlterado As Integer
Dim giValorIRRFAlterado As Integer
Dim giTipoTributacaoAlterado As Integer
Dim giAliqIRAlterada As Integer
Dim iPISRetidoAlterado As Integer
Dim iCOFINSRetidoAlterado As Integer
Dim iCSLLRetidoAlterado As Integer

Dim giTrazendoTribItemTela As Integer 'por Leo em 02/05/02
Dim giNatOpItemAlterado As Integer
Dim giTipoTributacaoItemAlterado As Integer
Dim giICMSBaseItemAlterado As Integer
Dim giICMSPercRedBaseItemAlterado As Integer
Dim giICMSAliquotaItemAlterado As Integer
Dim giICMSValorItemAlterado As Integer
Dim giICMSSubstBaseItemAlterado As Integer
Dim giICMSSubstAliquotaItemAlterado As Integer
Dim giICMSSubstValorItemAlterado As Integer
Dim giIPIBaseItemAlterado As Integer
Dim giIPIPercRedBaseItemAlterado As Integer
Dim giIPIAliquotaItemAlterado As Integer
Dim giIPIValorItemAlterado As Integer

Event Unload()
    
'Alterada por Luiz Nogueira em 13/01/04
Function Trata_Parametros(Optional objOrcamentoVenda As ClassOrcamentoVenda) As Long

Dim lErro As Long
Dim objOrcamentoVendaAux As New ClassOrcamentoVenda

On Error GoTo Erro_Trata_Parametros

    lErro = CargaPosFormLoad(True)
    If lErro <> SUCESSO Then gError 59288
    
    If Not (objOrcamentoVenda Is Nothing) Then
        
        'Se foi passado o código do orçamento
        If objOrcamentoVenda.lCodigo > 0 Then 'Incluído por Luiz Nogueira em 13/01/04
            
            objOrcamentoVendaAux.lCodigo = objOrcamentoVenda.lCodigo
            objOrcamentoVendaAux.iFilialEmpresa = objOrcamentoVenda.iFilialEmpresa
            
            'Coloca o Pedido de Venda na tela
            
            lErro = Traz_OrcamentoVenda_Tela(objOrcamentoVendaAux)
            If lErro <> SUCESSO And lErro <> 84363 Then gError 84479
    
            If lErro <> SUCESSO Then  'Não encontrou no BD o código de Pedido
    
                'Limpa a tela e coloca o código na Tela
                Call Limpa_OrcamentoVenda
                Codigo.Text = CStr(objOrcamentoVenda.lCodigo)
    
            End If

        '*** Incluído por Luiz Nogueira em 13/01/04 - INÍCIO ***
        'Se foi passado o código do cliente
        ElseIf objOrcamentoVenda.lCliente > 0 Then
        
            'Joga o código do cliente na tela
            Cliente.Text = objOrcamentoVenda.lCliente
            Call Cliente_Validate(bSGECancelDummy)
            
            'Se foi passada uma filial de cliente
            If objOrcamentoVenda.iFilial > 0 Then
            
                'Joga a filial do cliente na tela
                Filial.Text = objOrcamentoVenda.iFilial
                Call Filial_Validate(bSGECancelDummy)
            End If
            
            'Cria um número automático para o orçamento
            Call BotaoProxNum_Click
        '*** Incluído por Luiz Nogueira em 13/01/04 - FIM ***
        End If

    End If

    iAlterado = 0
    iVendedorAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 84479, 59288

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177333)

    End Select

    iAlterado = 0

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim objContainer As Object
Dim objUserControl As Object

On Error GoTo Erro_Form_Load

    'precodesc
    Set objContainer = Frame2(3)
    Set objUserControl = Me
    
    'precodesc
    lErro = CF("Orcamento_Form_Load", objUserControl, objContainer)
    If lErro <> SUCESSO Then gError 126500

    giPosCargaOk = 0
    
    giFrameAtual = 1
    giFrameAtualTributacao = 1
    
    'Preenche Data Referencia e Data de Emissão coma Data Atual
    DataReferencia.PromptInclude = False
    DataReferencia.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataReferencia.PromptInclude = True
    giDataReferenciaAlterada = 0
    
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    
    lErro = Carrega_TipoDesconto
    If lErro <> SUCESSO Then gError 103050

    iAlterado = 0
    
    iVendedorAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 84082, 84083, 84084, 84085, 84086, 84087, 84094, 84307, 101105, 103050, 126500

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177334)

    End Select

    iAlterado = 0

    Exit Sub

End Sub


Private Function Carrega_TipoDesconto() As Long

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_TipoDesconto

    Set colCodigoDescricao = gobjCRFAT.colTiposDesconto

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o item nas List's das Combos de Tipos Desconto
        TipoDesconto1.AddItem objCodDescricao.iCodigo & SEPARADOR & objCodDescricao.sNome
        TipoDesconto1.ItemData(TipoDesconto1.NewIndex) = objCodDescricao.iCodigo
        TipoDesconto2.AddItem objCodDescricao.iCodigo & SEPARADOR & objCodDescricao.sNome
        TipoDesconto2.ItemData(TipoDesconto2.NewIndex) = objCodDescricao.iCodigo
        TipoDesconto3.AddItem objCodDescricao.iCodigo & SEPARADOR & objCodDescricao.sNome
        TipoDesconto3.ItemData(TipoDesconto3.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_TipoDesconto = SUCESSO

    Exit Function

Erro_Carrega_TipoDesconto:

    Carrega_TipoDesconto = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177335)

    End Select

    Exit Function

End Function

Private Function Carrega_CondicaoPagamento() As Long

Dim lErro As Long
Dim colCod_DescReduzida As New AdmColCodigoNome
Dim objCod_DescReduzida As AdmCodigoNome

On Error GoTo Erro_Carrega_CondicaoPagamento

    'Lê o código e a descrição reduzida de todas as Condições de Pagamento
    lErro = CF("CondicoesPagto_Le_Recebimento", colCod_DescReduzida)
    If lErro <> SUCESSO Then gError 84088 '26489

   For Each objCod_DescReduzida In colCod_DescReduzida

        'Adiciona novo item na List da Combo CondicaoPagamento
        CondicaoPagamento.AddItem CInt(objCod_DescReduzida.iCodigo) & SEPARADOR & objCod_DescReduzida.sNome
        CondicaoPagamento.ItemData(CondicaoPagamento.NewIndex) = objCod_DescReduzida.iCodigo

    Next

    Carrega_CondicaoPagamento = SUCESSO

    Exit Function

Erro_Carrega_CondicaoPagamento:

    Carrega_CondicaoPagamento = gErr

    Select Case gErr

        Case 84088

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177336)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjOrcamentoVenda = Nothing

    'trecho incluido por Leo em 20/03/02 - Filtrar
    Set objEventoCliente = Nothing
    Set objEventoNumero = Nothing
    Set objEventoCondPagto = Nothing
    Set objEventoProduto = Nothing
    Set objEventoVendedor = Nothing
    Set objEventoNaturezaOp = Nothing
    Set objEventoTiposDeTributacao = Nothing

    Set objEventoTiposDeTributacao = Nothing
    Set gcolTiposTribICMS = Nothing
    Set gcolTiposTribIPI = Nothing
    
    'Encerra tributacao
    Call TributacaoOV_Terminar
    
    Call ComandoSeta_Liberar(Me.Name)
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Orçamento"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "OrcamentoVenda"

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

Private Function Carrega_TabelaPreco() As Long

Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome
Dim lErro As Long

On Error GoTo Erro_Carrega_TabelaPreco

    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("Cod_Nomes_Le", "TabelasDePreco", "Codigo", "Descricao", STRING_TABELA_PRECO_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 84012

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o item na Lista de Tabela de Preços
        TabelaPreco.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        TabelaPreco.ItemData(TabelaPreco.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_TabelaPreco = SUCESSO

    Exit Function

Erro_Carrega_TabelaPreco:

    Carrega_TabelaPreco = gErr

    Select Case gErr

        Case 84012

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177337)

    End Select

    Exit Function

End Function

Function Carrega_FilialEmpresa() As Long

Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome
Dim lErro As Long

On Error GoTo Erro_Carrega_FilialEmpresa

    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("Cod_Nomes_Le", "FiliaisEmpresa", "FilialEmpresa", "Nome", STRING_FILIAISEMPRESA_NOME, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 84093

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o item na Lista de Tabela de Preços
        Filial.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        Filial.ItemData(Filial.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_FilialEmpresa = SUCESSO

    Exit Function

Erro_Carrega_FilialEmpresa:

    Carrega_FilialEmpresa = gErr

    Select Case gErr

        Case 84093
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_FILIALEMPRESA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177338)

    End Select

    Exit Function

End Function

Private Sub BotaoDataReferenciaDown_Click()

Dim lErro As Long
Dim bCancel As Boolean
Dim sData As String

On Error GoTo Erro_BotaoDataReferenciaDown_Click

    sData = DataReferencia.Text

    'diminui a data em um dia
    lErro = Data_Up_Down_Click(DataReferencia, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 84070 '26715

    Call DataReferencia_Validate(bCancel)

    If bCancel = True Then DataReferencia.Text = sData

    Exit Sub

Erro_BotaoDataReferenciaDown_Click:

    Select Case gErr

        Case 84070

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177339)

    End Select

    Exit Sub

End Sub

Private Sub BotaoDataReferenciaUp_Click()

Dim lErro As Long
Dim sData As String
Dim bCancel As Boolean

On Error GoTo Erro_BotaoDataReferenciaUp_Click

    sData = DataReferencia.Text

    'aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataReferencia, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 84072

    Call DataReferencia_Validate(bCancel)

    If bCancel = True Then DataReferencia.Text = sData

    Exit Sub

Erro_BotaoDataReferenciaUp_Click:

    Select Case gErr

        Case 84072

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177340)

    End Select

    Exit Sub


End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o código foi informado
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 84307 '46130

    objOrcamentoVenda.lCodigo = StrParaLong(Codigo.Text)
    objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa

    'Lê o Orcamento
    lErro = CF("OrcamentoVenda_Le", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 101232 Then gError 84308
    If lErro = 101232 Then gError 101274
    
    'por Leo em 29/04/02 **********
    'Se o orçamento estiver vinculado a um Pedido de Venda, não poderá ser excluido -> Erro.
    If objOrcamentoVenda.lNumIntPedVenda <> 0 Then
        
        gError 94486
    
    'Se o orçamento estiver vinculado a uma Nota Fiscal, não poderá ser excluido -> Erro.
    ElseIf objOrcamentoVenda.lNumIntNFiscal <> 0 Then
        
        gError 94487
    
    End If
    'leo *********
    
    'Pede a confirmação da exclusão do Orcamento de Venda
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_ORCAMENTO_VENDA", objOrcamentoVenda.lCodigo)
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If

    'Faz a exclusão do Orcamento de Venda
    lErro = CF("OrcamentoVenda_Exclui", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 84310 '46139

    'Limpa a Tela de Orcamento de Venda
    Call Limpa_OrcamentoVenda
    
    'fecha o comando de setas
    Call ComandoSeta_Fechar(Me.Name)

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 84307
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_PREENCHIDO", gErr)

        Case 84308, 84310

        Case 101274
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO", gErr, objOrcamentoVenda.lCodigo)

        Case 94486
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ORCAMENTO_VINCULADO_PEDIDO", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)
            
        Case 94487
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ORCAMENTO_VINCULADO_NFISCAL", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)
              
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177341)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub



Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 84243 '26806

    'Limpa a Tela
    Call Limpa_OrcamentoVenda

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 84243

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 177342)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 84065 '26805

    'Limpa a Tela
    Call Limpa_OrcamentoVenda
    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 84065

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177343)

    End Select

    Exit Sub


End Sub

'mario
Private Sub Limpa_OrcamentoVenda()

Dim lErro As Long

On Error GoTo Erro_OrcamentoVenda
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Call Limpa_OrcamentoVenda2

    Call BotaoGravarTrib
    
    iAlterado = 0
          
    iVendedorAlterado = 0
     
    Exit Sub

Erro_OrcamentoVenda:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177344)

    End Select

    Exit Sub

End Sub

'mario
Private Sub Limpa_OrcamentoVenda2()
'Limpa os campos da tela sem fechar o sistema de setas

Dim iIndice As Integer

    Call Limpa_Tela(Me)

    Codigo.Enabled = True
    Codigo.Text = ""
    Filial.Clear

    CobrancaAutomatica.Value = vbChecked
    ValorTotal.Caption = ""
    ValorProdutos.Caption = ""
    CondicaoPagamento.Text = ""
    DataReferencia.PromptInclude = False
    DataReferencia.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataReferencia.PromptInclude = True
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    Call DataEmissao_Validate(bSGECancelDummy)
    'Trecho Incluido por Leo em 22/03/02
    Vendedor.PromptInclude = False
    Vendedor.Text = ""
    Vendedor.PromptInclude = True
    
    If giTipoVersao = VERSAO_FULL Then
        Seleciona_FilialEmpresa
        TabelaPreco.Text = ""
    End If

    'tab de tributacao resumo
    'ISSIncluso.Value = 0
    IPIBase.Caption = ""
    IPIValor.Caption = ""
    ISSBase.Caption = ""
    DescTipoTrib.Caption = ""
    IRBase.Caption = ""
    ICMSBase.Caption = ""
    ICMSValor.Caption = ""
    ICMSSubstBase.Caption = ""
    ICMSSubstValor.Caption = ""

    'tab de tributacao itens
    LabelValorFrete.Caption = ""
    LabelValorDesconto.Caption = ""
    LabelValorSeguro.Caption = ""
    LabelValorOutrasDespesas.Caption = ""
    ComboItensTrib.Clear
    LabelValorItem.Caption = ""
    LabelQtdeItem.Caption = ""
    LabelUMItem.Caption = ""
    LabelDescrNatOpItem.Caption = ""
    DescTipoTribItem.Caption = ""
    
    '************** TRATAMENTO DE GRADE **************
    For iIndice = 1 To objGridItens.iLinhasExistentes
        GridItens.TextMatrix(iIndice, 0) = iIndice
    Next
    '*************************************************
    
    Call Grid_Limpa(objGridItens)
    Call Grid_Limpa(objGridParcelas)

    'Resseta tributação
    Call TributacaoOV_Reset
    
    iAlterado = 0
    giValorDescontoAlterado = 0
    giClienteAlterado = 0
    giFilialAlterada = 0
    giDataReferenciaAlterada = 0

    giValorDescontoManual = 0


    Exit Sub

End Sub

Sub Seleciona_FilialEmpresa()

Dim lErro As Long
Dim iIndice As Integer
Dim iFilialFaturamento As Integer

On Error GoTo Erro_Seleciona_FilialEmpresa

    iFilialFaturamento = gobjFAT.iFilialFaturamento

    Exit Sub

Erro_Seleciona_FilialEmpresa:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177345)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProdutos_Click()

Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim colSelecao As Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoProdutos_Click

    If Me.ActiveControl Is Produto Then

        sProduto1 = Produto.Text

    Else

        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 84074 '58771

        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)

    End If

    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 84075 '58772

    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    'preenche o codigo do produto
    objProduto.sCodigo = sProduto

    'Chama a tela de browse ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case 84074 '58771
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 84075 '58772 Tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177346)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lOrcamentoVenda As Long

On Error GoTo Erro_PedidoDeVenda_Automatico

    lErro = CF("Config_ObterAutomatico", "FatConfig", "NUM_PROX_CODIGO_ORCAMENTOVENDA", "OrcamentoVenda", "Codigo", lOrcamentoVenda)
    If lErro <> SUCESSO Then gError 94422
    
    Codigo.Text = lOrcamentoVenda

    Exit Sub

Erro_PedidoDeVenda_Automatico:

    Select Case gErr

        Case 94422

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177347)

    End Select

    Exit Sub
    
End Sub

Private Sub Cliente_Change()

    iAlterado = REGISTRO_ALTERADO
    giClienteAlterado = 1

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long, sNatOp As String, iTipoTrib As Integer
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objTipoCliente As New ClassTipoCliente
Dim objMensagem As New ClassMensagem

On Error GoTo Erro_Cliente_Validate

    If giClienteAlterado = 1 Then

        'Verifica se o Cliente está preenchido
        If Len(Trim(Cliente.Text)) > 0 Then

            'Busca o Cliente no BD
            lErro = TP_Cliente_Le_Orcamento(Cliente, objCliente, iCodFilial)
            If lErro <> SUCESSO And lErro <> 94449 And lErro <> 94450 And lErro <> 94451 And lErro <> 94452 Then gError 84037
                                        
            If lErro = SUCESSO Then
            
                gobjOrcamentoVenda.lCliente = objCliente.lCodigo
            
                lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
                If lErro <> SUCESSO Then gError 84038 '26136
    
                'Preenche ComboBox de Filiais
                Call CF("Filial_Preenche", Filial, colCodigoNome)
    
                If Not gbCarregandoTela Then
    
                    If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then
    
                        If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ
    
                        'Seleciona filial na Combo Filial
                        Call CF("Filial_Seleciona", Filial, iCodFilial)
    
                    End If
                    
                End If
                
                giValorDescontoManual = 0
                
                'Guarda o valor do desconto do cliente
                If objCliente.dDesconto > 0 Then
                    
                    gdDesconto = objCliente.dDesconto
                
                ElseIf objTipoCliente.dDesconto > 0 Then
                    
                    gdDesconto = objTipoCliente.dDesconto
                
                Else
                    
                    gdDesconto = 0
                
                End If
    
                If Not gbCarregandoTela Then
    
                    Call DescontoGlobal_Recalcula
    
                    'ATualiza o total com o novo desconto
                    lErro = ValorTotal_Calcula()
                    If lErro <> SUCESSO Then gError 84039
    
                    'Incluído por Luiz Nogueira em 26/01/04
                    'Coloca na tela a tabela do cliente
                    If objCliente.iVendedor > 0 Then
                        
                        Vendedor.Text = objCliente.iVendedor
                        Call Vendedor_Validate(bSGECancelDummy)
                    
                    ElseIf objTipoCliente.iVendedor > 0 Then
                        
                        Vendedor.Text = objTipoCliente.iVendedor
                        Call Vendedor_Validate(bSGECancelDummy)
                    
                    End If
                    'Fim Luiz Nogueira - 26/01/04
    
                    'Coloca na tela a tabela do cliente
                    If objCliente.iTabelaPreco > 0 Then
                        
                        TabelaPreco.Text = objCliente.iTabelaPreco
                        Call TabelaPreco_Validate(bSGECancelDummy)
                    
                    ElseIf objTipoCliente.iTabelaPreco > 0 Then
                        
                        TabelaPreco.Text = objTipoCliente.iTabelaPreco
                        Call TabelaPreco_Validate(bSGECancelDummy)
                    
                    End If
                    
                    'Se cobrança automática estiver selecionada preenche a CondPagto e dispara o Validate
                    If CobrancaAutomatica.Value = MARCADO Then
                        
                        If objCliente.iCondicaoPagto > 0 Then
                            
                            CondicaoPagamento.Text = objCliente.iCondicaoPagto
                            Call CondicaoPagamento_Validate(bSGECancelDummy)
                        
                        ElseIf objTipoCliente.iCondicaoPagto > 0 Then
                            
                            CondicaoPagamento.Text = objTipoCliente.iCondicaoPagto
                            Call CondicaoPagamento_Validate(bSGECancelDummy)
                        
                        End If
    
                    End If
    
                End If
                
                giClienteAlterado = 0
    
            'trecho por leo em 17/04/02
            Else
            
                gobjOrcamentoVenda.lCliente = 0
                giValorDescontoManual = 0
                gdDesconto = 0
                                
                If Not gbCarregandoTela Then
                
                    Call DescontoGlobal_Recalcula
    
                    'ATualiza o total com o novo desconto
                    lErro = ValorTotal_Calcula()
                    If lErro <> SUCESSO Then gError 84039 '51034
                    '????? Alterar a numeracao de erro pois está repetido
                          
                          
                    TabelaPreco.ListIndex = -1
                    
                    objCliente.lCodigo = 0
                    
                    Filial.Clear
                            
                End If
                
            End If
            
            'Leo em 17/04/02 até aqui
            
        'Se não estiver preenchido
        ElseIf Len(Trim(Cliente.Text)) = 0 Then

            'Limpa a Combo de Filiais
            Filial.Clear

        End If

        giClienteAlterado = 0
    
    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 84037, 84038, 84039, 84040, 84041

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177348)

    End Select

    Exit Sub

End Sub

Private Sub DescontoGlobal_Recalcula()

Dim dValorDesconto As Double
Dim dValorProdutos As Double

    If Len(Trim(ValorProdutos.Caption)) <> 0 And IsNumeric(ValorProdutos.Caption) Then

        'Se o cliente possui desconto e o campo desconto não foi alterado pelo usuário
        If gdDesconto > 0 And giValorDescontoManual = 0 Then

            Call Calcula_ValorProdutos(dValorProdutos)

            'Calcula o valor do desconto para o cliente e coloca na tela
            dValorDesconto = gdDesconto * dValorProdutos
            ValorDesconto.Text = Format(dValorDesconto, "Standard")
            giValorDescontoAlterado = 0

            'Para tributação
            gobjOrcamentoVenda.dValorDesconto = dValorDesconto

        End If

    End If

End Sub

Public Sub Calcula_ValorProdutos(dValorProdutos As Double)

Dim dValorTotal As Double
Dim dValor As Double
Dim iIndice As Integer

    dValor = 0

    For iIndice = 1 To objGridItens.iLinhasExistentes

        dValorTotal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))

        dValor = dValor + dValorTotal

    Next

    dValorProdutos = dValor

End Sub

Private Sub CobrancaAutomatica_Click()

  iAlterado = REGISTRO_ALTERADO
  
  Call Cobranca_Automatica

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_GotFocus()

    gsCodigoAnterior = Codigo.Text
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

'jones. Revisar
Public Sub ComboICMSTipo_Click()

    If ComboICMSTipo.ListIndex = -1 Then Exit Sub

    If giTrazendoTribItemTela = 0 Then
        Call BotaoGravarTribItem_Click
    End If

    iAlterado = REGISTRO_ALTERADO

End Sub

'jones. Revisar
Public Sub ComboIPITipo_Click()

    If ComboIPITipo.ListIndex = -1 Then Exit Sub

    If giTrazendoTribItemTela = 0 Then
        Call BotaoGravarTribItem_Click
    End If

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CondicaoPagamento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'????? Retirar comentarios com a numeracao de erro antiga. mario
Private Sub CondicaoPagamento_Click()

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim vbMsgRes As VbMsgBoxResult
Dim dPercAcresFin As Double

On Error GoTo Erro_CondicaoPagamento_Click

    'Verifica se alguma Condição foi selecionada
    If CondicaoPagamento.ListIndex = -1 Then Exit Sub

    'Passa o código da Condição para objCondicaoPagto
    objCondicaoPagto.iCodigo = CondicaoPagamento.ItemData(CondicaoPagamento.ListIndex)

    'Lê Condição a partir do código
    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
    If lErro <> SUCESSO And lErro <> 19205 Then gError 84043 '26718
    If lErro = 19205 Then gError 84044 '26720

    'Altera PercAcrescFin
    If Len(Trim(PercAcrescFin.ClipText)) > 0 Then

        dPercAcresFin = StrParaDbl(PercAcrescFin.Text) / 100
        If dPercAcresFin <> objCondicaoPagto.dAcrescimoFinanceiro Then
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUICAO_PERC_ACRESCIMO_FINANCEIRO")
            If vbMsgRes = vbYes Then
                PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
                Call PercAcrescFin_Validate(bSGECancelDummy)
            End If
        End If
    Else
        PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
        Call PercAcrescFin_Validate(bSGECancelDummy)
    End If

    'Testa se ValorTotal está preenchido
    If Len(Trim(ValorTotal)) > 0 Then
        'Testa se DataReferencia está preenchida e ValorTotal é positivo
        If Len(Trim(DataReferencia.ClipText)) > 0 And (CDbl(ValorTotal.Caption) > 0) Then

            'Preenche o GridParcelas
            lErro = Cobranca_Automatica()
            If lErro <> SUCESSO Then gError 84045 '26719

        End If
    End If

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_CondicaoPagamento_Click:

    Select Case gErr

        Case 84045, 84043

        Case 84044
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177349)

      End Select

    Exit Sub

End Sub

'mario
Private Function Cobranca_Automatica() As Long
'recalcula o tab de cobranca

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_Cobranca_Automatica

    If CobrancaAutomatica.Value = vbChecked And Len(Trim(CondicaoPagamento.Text)) <> 0 Then

        objCondicaoPagto.iCodigo = Codigo_Extrai(CondicaoPagamento.Text)

        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 84046
        If lErro <> SUCESSO Then gError 84048

        lErro = GridParcelas_Preenche(objCondicaoPagto)
        If lErro <> SUCESSO Then gError 84047

    End If

    Cobranca_Automatica = SUCESSO

    Exit Function

Erro_Cobranca_Automatica:

    Cobranca_Automatica = gErr

    Select Case gErr

        Case 84046, 84047

        Case 84048
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 177350)

    End Select

    Exit Function

End Function

'mario
Private Function GridParcelas_Preenche(objCondicaoPagto As ClassCondicaoPagto) As Long
'Calcula valores e datas de vencimento de Parcelas a partir da Condição de Pagamento e preenche GridParcelas

Dim lErro As Long
Dim dValorPagar As Double
Dim dValorIRRF As Double, dPISRetido As Double, dCOFINSRetido As Double, dCSLLRetido As Double
Dim iIndice As Integer
Dim dPercAcrescFin As Double
Dim iTamanho As Integer
Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim iColuna As Integer

Dim colDescontoPadrao As New Collection

On Error GoTo Erro_GridParcelas_Preenche

    'Limpa o GridParcelas
    Call Grid_Limpa(objGridParcelas)

    'Número de Parcelas
    objGridParcelas.iLinhasExistentes = objCondicaoPagto.iNumeroParcelas

    If Len(Trim(ValorIRRF.Text)) > 0 Then dValorIRRF = CDbl(ValorIRRF)
    If Len(Trim(PISRetido.Text)) <> 0 And IsNumeric(PISRetido.Text) Then dPISRetido = CDbl(PISRetido.Text)
    If Len(Trim(COFINSRetido.Text)) <> 0 And IsNumeric(COFINSRetido.Text) Then dCOFINSRetido = CDbl(COFINSRetido.Text)
    If Len(Trim(CSLLRetido.Text)) <> 0 And IsNumeric(CSLLRetido.Text) Then dCSLLRetido = CDbl(CSLLRetido.Text)
    
    'Valor a Pagar
    dValorPagar = StrParaDbl(ValorTotal) - (dValorIRRF + dPISRetido + dCOFINSRetido + dCSLLRetido)

    'Se Valor a Pagar for positivo
    If dValorPagar > 0 Then

        objCondicaoPagto.dValorTotal = dValorPagar
        
        'Calcula os valores das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, True, False)
        If lErro <> SUCESSO Then gError 84076 '26721

        'Coloca os valores das Parcelas no Grid Parcelas
        For iIndice = 1 To objGridParcelas.iLinhasExistentes
            GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dValor, "Standard")
        Next

    End If

    'Se Data Referencia estiver preenchida
    If Len(Trim(DataReferencia.ClipText)) > 0 Then

        objCondicaoPagto.dtDataRef = CDate(DataReferencia.Text)

        'Calcula Datas de Vencimento das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, False, True)
        If lErro <> SUCESSO Then gError 84077

        'Loop de preenchimento do Grid Parcelas com Datas de Vencimento
        For iIndice = 1 To objCondicaoPagto.iNumeroParcelas

            'Coloca Data de Vencimento no Grid Parcelas
            GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col) = Format(objCondicaoPagto.colParcelas(iIndice).dtVencimento, "dd/mm/yyyy")

        Next

    End If

    ' Se dValorPagar>0 coloca desconto padrao (quantos houver, se houver) em todas as parcelas.
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        lErro = Preenche_DescontoPadrao(iIndice)
        If lErro <> SUCESSO Then gError 84078
    Next

    GridParcelas_Preenche = SUCESSO

    Exit Function

Erro_GridParcelas_Preenche:

    GridParcelas_Preenche = gErr

    Select Case gErr

        Case 84076 To 84078

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177351)

    End Select

    Exit Function

End Function

Function Preenche_DescontoPadrao(iLinha As Integer) As Long

Dim lErro As Long
Dim colDescontoPadrao As New ColDesconto
Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim iColuna  As Integer
Dim dtDataVencimento As Date
Dim dPercentual As Double
Dim dValorParcela As Double
Dim sValorDesconto As String

On Error GoTo Erro_Preenche_DescontoPadrao

    'Se a data de referencia estiver preenchida
    If Len(Trim(DataReferencia.ClipText)) > 0 Then

        dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(iLinha, iGrid_Vencimento_col))
        lErro = CF("Parcela_GeraDescontoPadrao", colDescontoPadrao, dtDataVencimento)
        If lErro <> SUCESSO Then gError 84098

        If colDescontoPadrao.Count > 0 Then

            'Para cada um dos desontos padrão
            For iIndice1 = 1 To colDescontoPadrao.Count

                'Seleciona a coluna correspondente ao Desconto
                If iIndice1 = 1 Then iColuna = iGrid_Desc1Codigo_Col
                If iIndice1 = 2 Then iColuna = iGrid_Desc2Codigo_Col
                If iIndice1 = 3 Then iColuna = iGrid_Desc3Codigo_Col

                'Seleciona o tipo de desconto
                For iIndice2 = 0 To TipoDesconto1.ListCount - 1
                    If colDescontoPadrao.Item(iIndice1).iCodigo = TipoDesconto1.ItemData(iIndice2) Then
                        GridParcelas.TextMatrix(iLinha, iColuna) = TipoDesconto1.List(iIndice2)
                        GridParcelas.TextMatrix(iLinha, iColuna + 1) = Format(colDescontoPadrao.Item(iIndice1).dtData, "dd/mm/yyyy")
                        GridParcelas.TextMatrix(iLinha, iColuna + 3) = Format(colDescontoPadrao.Item(iIndice1).dValor, "Percent")

                        '*** Inicio colocacao Valor Desconto na tela
                        dPercentual = colDescontoPadrao.Item(iIndice1).dValor
                        dValorParcela = StrParaDbl(GridParcelas.TextMatrix(iLinha, iGrid_ValorParcela_Col))

                        'Coloca Valor do Desconto na tela
                        If dValorParcela > 0 Then
                            sValorDesconto = Format(dPercentual * dValorParcela, "Standard")
                            GridParcelas.TextMatrix(iLinha, iColuna + 2) = sValorDesconto
                        End If
                        '*** Fim colocacao Valor Desconto na tela

                    End If
                Next
            Next

        End If

    End If

    Preenche_DescontoPadrao = SUCESSO

    Exit Function

Erro_Preenche_DescontoPadrao:

    Preenche_DescontoPadrao = gErr

    Select Case gErr

        Case 84098

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177352)

    End Select

    Exit Function

End Function

Private Sub CondicaoPagamento_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim vbMsgRes As VbMsgBoxResult
Dim dPercAcresFin As Double

On Error GoTo Erro_Condicaopagamento_Validate

    'Verifica se a Condicaopagamento foi preenchida
    If Len(Trim(CondicaoPagamento.Text)) = 0 Then Exit Sub

    'Verifica se é uma Condicaopagamento selecionada
    If CondicaoPagamento.Text = CondicaoPagamento.List(CondicaoPagamento.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(CondicaoPagamento, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 84049 '26542

    'Se não encontra valor que contém CÓDIGO, mas extrai o código
    If lErro = 6730 Then

        objCondicaoPagto.iCodigo = iCodigo

        'Lê Condicao Pagamento no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 84050 '26543
        If lErro = 19205 Then gError 84054 '26545

        'Testa se pode ser usada em Contas a Receber
        If objCondicaoPagto.iEmRecebimento = 0 Then gError 84051 '26547

        'Coloca na Tela
        CondicaoPagamento.Text = iCodigo & SEPARADOR & objCondicaoPagto.sDescReduzida

        'Altera PercAcrescFin
        If Len(Trim(PercAcrescFin.ClipText)) > 0 Then
            dPercAcresFin = StrParaDbl(PercAcrescFin.Text) / 100
            If dPercAcresFin <> objCondicaoPagto.dAcrescimoFinanceiro Then
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUICAO_PERC_ACRESCIMO_FINANCEIRO")
                If vbMsgRes = vbYes Then
                    PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
                   Call PercAcrescFin_Validate(bSGECancelDummy)
                End If
            End If
        Else
            PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
            Call PercAcrescFin_Validate(bSGECancelDummy)
        End If

        'Se ValorTotal e DataReferencia estiverem preenchidos, preenche GridParcelas
        If Len(Trim(ValorTotal)) > 0 Then
            If Len(Trim(DataReferencia.ClipText)) > 0 And CLng(ValorTotal.Caption) > 0 Then

                'Preenche o GridParcelas
                lErro = Cobranca_Automatica()
                If lErro <> SUCESSO Then gError 84052 '26544

            End If
        End If

    End If

    'Não encontrou o valor que era STRING
    If lErro = 6731 Then gError 84053 '26546

    Exit Sub

Erro_Condicaopagamento_Validate:

    Cancel = True

    Select Case gErr

       Case 84049, 84050, 84052

        Case 84051
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_PAGAMENTO", gErr, objCondicaoPagto.iCodigo)

        Case 84053
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", gErr, CondicaoPagamento.Text)

       Case 84054
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAOPAGTO", iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177353)

    End Select

    Exit Sub

End Sub

Private Sub CondPagtoLabel_Click()

Dim objCondicaoPagto As New ClassCondicaoPagto
Dim colSelecao As New Collection

    'Se Condição de Pagto estiver preenchida, extrai o código
    If Len(Trim(CondicaoPagamento.Text)) > 0 Then
        objCondicaoPagto.iCodigo = Codigo_Extrai(CondicaoPagamento.Text)
    End If

    'Chama a Tela CondicoesPagamentoCRLista
    Call Chama_Tela("CondicaoPagtoCRLista", colSelecao, objCondicaoPagto, objEventoCondPagto)

End Sub

Private Sub DataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissao_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)

End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then gError 84037

    If gobjOrcamentoVenda.dtDataEmissao <> StrParaDate(DataEmissao.Text) Then
        
        gobjOrcamentoVenda.dtDataEmissao = StrParaDate(DataEmissao.Text)
        
        Call ValorTotal_Calcula
        
    End If
    
    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case 84037

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177354)

    End Select

    Exit Sub

End Sub

Private Sub DataEntrega_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEntrega_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub DataEntrega_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub DataEntrega_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataEntrega
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True


End Sub

Private Sub DataReferencia_Change()

     iAlterado = REGISTRO_ALTERADO
     giDataReferenciaAlterada = REGISTRO_ALTERADO

End Sub

Private Sub DataReferencia_GotFocus()

    Dim iDataAux As Integer

    iDataAux = giDataReferenciaAlterada
    Call MaskEdBox_TrataGotFocus(DataReferencia, iAlterado)
    giDataReferenciaAlterada = iDataAux

End Sub

Private Sub DataReferencia_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtDataEmissao As Date
Dim dtDataReferencia As Date
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_DataReferencia_Validate

    If giDataReferenciaAlterada <> REGISTRO_ALTERADO Then Exit Sub

    If Len(Trim(DataReferencia.ClipText)) > 0 Then

        'Critica a data digitada
        lErro = Data_Critica(DataReferencia.Text)
        If lErro <> SUCESSO Then gError 84062 '26713

        'Compara com data de emissão
        If Len(Trim(DataEmissao.ClipText)) > 0 Then

            dtDataEmissao = CDate(DataEmissao.Text)
            dtDataReferencia = CDate(DataReferencia.Text)

            If dtDataEmissao > dtDataReferencia Then gError 84063

        End If


    End If

    giDataReferenciaAlterada = 0

    'Preenche o GridParcelas
    lErro = Cobranca_Automatica()
    If lErro <> SUCESSO Then gError 84064

    Exit Sub

Erro_DataReferencia_Validate:

    Cancel = True

    Select Case gErr

        Case 84062, 84064 'Tratado na rotina chamada

        Case 84063
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_MAIOR_DATAREFERENCIA", gErr, dtDataReferencia, dtDataEmissao)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177355)

    End Select

    Exit Sub

End Sub

Private Sub DataVencimento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataVencimento_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub DataVencimento_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub DataVencimento_Validate(Cancel As Boolean)

    Dim lErro As Long

    Set objGridParcelas.objControle = DataVencimento
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Desconto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Desconto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Desconto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Desconto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Desconto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Desconto1Ate_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Desconto1Ate_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub Desconto1Ate_KeyPress(KeyAscii As Integer)

   Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub Desconto1Ate_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Desconto1Ate
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Desconto1Percentual_Change()

   iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Desconto1Percentual_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub Desconto1Percentual_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub Desconto1Percentual_Validate(Cancel As Boolean)

    Dim lErro As Long

    Set objGridParcelas.objControle = Desconto1Percentual
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Desconto1Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Desconto1Valor_GotFocus()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Desconto1Valor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub Desconto1Valor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Desconto1Valor
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Desconto2Ate_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Desconto2Ate_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub Desconto2Ate_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub Desconto2Ate_Validate(Cancel As Boolean)

    Dim lErro As Long

    Set objGridParcelas.objControle = Desconto2Ate
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Desconto2Percentual_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Desconto2Percentual_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub Desconto2Percentual_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub Desconto2Percentual_Validate(Cancel As Boolean)

    Dim lErro As Long

    Set objGridParcelas.objControle = Desconto2Percentual
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Desconto2Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Desconto2Valor_GotFocus()

Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub Desconto2Valor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub Desconto2Valor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Desconto2Valor
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Desconto3Ate_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Desconto3Ate_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub Desconto3Ate_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub Desconto3Ate_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Desconto3Ate
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Desconto3Percentual_Change()

    iAlterado = REGISTRO_ALTERADO

 End Sub

Private Sub Desconto3Percentual_GotFocus()

     Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub Desconto3Percentual_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub Desconto3Percentual_Validate(Cancel As Boolean)

    Dim lErro As Long

    Set objGridParcelas.objControle = Desconto3Percentual
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Desconto3Valor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Desconto3Valor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub Desconto3Valor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub Desconto3Valor_Validate(Cancel As Boolean)

    Dim lErro As Long

    Set objGridParcelas.objControle = Desconto3Valor
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescricaoProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescricaoProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub DescricaoProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub DescricaoProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DescricaoProduto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO
    giFilialAlterada = 1

End Sub

Private Sub Filial_Click()

Dim lErro As Long

On Error GoTo Erro_Filial_Click

    iAlterado = REGISTRO_ALTERADO

    'Se nenhuma filial foi selecionada, sai.
    If Filial.ListIndex = -1 Then Exit Sub

    'Faz o tratamento para a filial do cliente selecionada
    lErro = Trata_FilialCliente()
    If lErro <> SUCESSO Then gError 84042 '23581

    Exit Sub

Erro_Filial_Click:

    Select Case gErr

        Case 84042

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177356)

    End Select

    Exit Sub

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
    If Len(Trim(Filial.Text)) = 0 Or giFilialAlterada = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 84000 '26622

    'Se não encontrou o CÓDIGO
    If lErro = 6730 Then

        'Verifica se o cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then gError 84001 '26623

        sCliente = Cliente.Text
        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 84002 '26624

        If lErro = 17660 Then

            'Lê o Cliente
            objCliente.sNomeReduzido = sCliente
            lErro = CF("Cliente_Le_NomeReduzido", objCliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 84003 '25664

            'Se encontrou o Cliente
            If lErro = SUCESSO Then
                
                objFilialCliente.lCodCliente = objCliente.lCodigo

                gError 84005
            
            End If
            
        End If
        
        If iCodigo <> 0 Then
        
            'Coloca na tela a Filial lida
            Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome
        
            lErro = Trata_FilialCliente
            If lErro <> SUCESSO Then gError 84006 '25435
        
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
        If lErro <> SUCESSO And lErro <> 12348 Then gError 94448 '25664
        
        If lErro = SUCESSO Then gError 84007
        
    End If
    
    giFilialAlterada = 0

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 84000, 84002

        Case 84001
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        
        Case 84003, 84006, 94448 'tratado na rotina chamada

        Case 84005
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 84007
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177357)

    End Select

    Exit Sub

End Sub

Function Trata_FilialCliente() As Long

Dim objFilialCliente As New ClassFilialCliente
Dim objCliente As New ClassCliente
Dim objVendedor As New ClassVendedor
Dim objTipoCliente As New ClassTipoCliente
Dim dValorTotal As Double
Dim dValorBase As Double
Dim objTransportadora As New ClassTransportadora
Dim dValorComissao As Double
Dim dValorEmissao As Double
Dim lErro As Long

On Error GoTo Erro_Trata_FilialCliente

    objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)
    objCliente.sNomeReduzido = Trim(Cliente.Text)

    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Trim(Cliente.Text), objFilialCliente)
    If lErro <> SUCESSO And lErro <> 17660 Then gError 84019
    If lErro = 17660 Then gError 84020

    gobjOrcamentoVenda.iFilial = objFilialCliente.iCodFilial
    
    'Calula o valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 124223
            
    Trata_FilialCliente = SUCESSO

    Exit Function

Erro_Trata_FilialCliente:

    Trata_FilialCliente = gErr

    Select Case gErr

        Case 84019, 124223

        Case 84020
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA1", gErr, Cliente.Text, objFilialCliente.iCodFilial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177358)

    End Select

    Exit Function

End Function

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGridItens, iAlterado)

End Sub

Private Sub GridItens_GotFocus()

    Call Grid_Recebe_Foco(objGridItens)

End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iItemAtual As Integer
Dim iIndice As Integer
Dim dValorTotal As Double
Dim lErro As Long

On Error GoTo Erro_GridItens_KeyDown

    'Guarda o número de linhas existentes e a linha atual
    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes
    iItemAtual = GridItens.Row

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

    If objGridItens.iLinhasExistentes < iLinhasExistentesAnterior Then

        '************ grade ************
        'Retira a "#" caso o item excluído tenha sido um de grade
        GridItens.TextMatrix(GridItens.Row, 0) = GridItens.Row
        '*******************************

        Call Tributacao_Remover_Item_Grid(iItemAtual)
        
        'Calcula a soma dos valores de produtos
        For iIndice = 1 To objGridItens.iLinhasExistentes
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
                If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
            End If
        Next

        'Coloca valor total dos produtos na tela
        ValorProdutos.Caption = Format(dValorTotal, "Standard")

        'Calcula o valor total da nota
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 84145

    End If

    Exit Sub

Erro_GridItens_KeyDown:

    Select Case gErr

        Case 84145

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177359)

    End Select

    Exit Sub

End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If


End Sub

Private Sub GridItens_LeaveCell()

    Call Saida_Celula(objGridItens)

End Sub

Private Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridItens)

End Sub

Private Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub

Private Sub GridItens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItens)

End Sub

Private Sub GridParcelas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Private Sub GridParcelas_EnterCell()

     Call Grid_Entrada_Celula(objGridParcelas, iAlterado)

End Sub

Private Sub GridParcelas_GotFocus()

    Call Grid_Recebe_Foco(objGridParcelas)

End Sub

Private Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridParcelas)

End Sub

Private Sub GridParcelas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Private Sub GridParcelas_LeaveCell()

    Call Saida_Celula(objGridParcelas)

End Sub

Private Sub GridParcelas_RowColChange()

    Call Grid_RowColChange(objGridParcelas)

End Sub

Private Sub GridParcelas_Scroll()

    Call Grid_Scroll(objGridParcelas)

End Sub

Private Sub GridParcelas_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridParcelas)

End Sub

Public Sub ICMSAliquotaItem_Change()

    giICMSAliquotaItemAlterado = 1
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ICMSAliquotaItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSAliquotaItem_Validate

    If giICMSAliquotaItemAlterado Then

        If Len(Trim(ICMSAliquotaItem.ClipText)) > 0 Then
            
            lErro = Porcentagem_Critica(ICMSAliquotaItem.Text)
            If lErro <> SUCESSO Then gError 1030167
        
        End If
        
        Call BotaoGravarTribItem_Click

        giICMSAliquotaItemAlterado = 0

    End If
    
    Exit Sub
    
Erro_ICMSAliquotaItem_Validate:

    Cancel = True

    Select Case gErr
    
        Case 1030167
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177360)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub ICMSBase_Change()

    ICMSBase1.Caption = ICMSBase.Caption

End Sub

Public Sub ICMSBaseItem_Change()

    iAlterado = REGISTRO_ALTERADO
    giICMSBaseItemAlterado = 1

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ICMSBaseItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSBaseItem_Validate

    If giICMSBaseItemAlterado Then

        If Len(Trim(ICMSBaseItem.ClipText)) > 0 Then
        
            lErro = Valor_NaoNegativo_Critica(ICMSBaseItem.Text)
            If lErro <> SUCESSO Then gError 103015
        
        End If
        
        Call BotaoGravarTribItem_Click

        giICMSBaseItemAlterado = 0

    End If

    Exit Sub
    
Erro_ICMSBaseItem_Validate:

    Cancel = True

    Select Case gErr
    
        Case 103015
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177361)
    
    End Select
            
    Exit Sub

End Sub

Public Sub ICMSPercRedBaseItem_Change()

    iAlterado = REGISTRO_ALTERADO
    giICMSPercRedBaseItemAlterado = 1

End Sub

Public Sub ICMSPercRedBaseItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSPercRedBaseItem_Validate

    If giICMSPercRedBaseItemAlterado Then

        If Len(Trim(ICMSPercRedBaseItem.Text)) > 0 Then
        
            lErro = Porcentagem_Critica(ICMSPercRedBaseItem.Text)
            If lErro <> SUCESSO Then gError 103016
        
        End If
        
        Call BotaoGravarTribItem_Click

        giICMSPercRedBaseItemAlterado = 0

    End If
    
    Exit Sub
    
Erro_ICMSPercRedBaseItem_Validate:

    Cancel = True

    Select Case gErr
    
        Case 103016
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177362)
            
    End Select

    Exit Sub
    
End Sub

Public Sub ICMSSubstAliquotaItem_Change()

    iAlterado = REGISTRO_ALTERADO
    giICMSSubstAliquotaItemAlterado = 1

End Sub

Public Sub ICMSSubstAliquotaItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSSubstAliquotaItem_Validate

    If giICMSSubstAliquotaItemAlterado Then

        If Len(Trim(ICMSSubstAliquotaItem.ClipText)) > 0 Then
            
            lErro = Porcentagem_Critica(ICMSSubstAliquotaItem.Text)
            If lErro <> SUCESSO Then gError 103017
            
        End If

        Call BotaoGravarTribItem_Click

        giICMSSubstAliquotaItemAlterado = 0

    End If

    Exit Sub
    
Erro_ICMSSubstAliquotaItem_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 103017
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177363)
            
    End Select
        
    Exit Sub
    
End Sub

Private Sub ICMSSubstBase_Change()

    ICMSSubstBase1.Caption = ICMSSubstBase.Caption

End Sub

Public Sub ICMSSubstBaseItem_Change()

    iAlterado = REGISTRO_ALTERADO
    giICMSSubstBaseItemAlterado = 1

End Sub

Public Sub ICMSSubstBaseItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSSubstBaseItem_Validate

    If giICMSSubstBaseItemAlterado Then

        If Len(Trim(ICMSSubstBaseItem.ClipText)) > 0 Then
        
            lErro = Valor_NaoNegativo_Critica(ICMSSubstBaseItem.Text)
            If lErro <> SUCESSO Then gError 103016
'?????? Numeracao de Erro Repetida. mario
        End If

        Call BotaoGravarTribItem_Click

        giICMSSubstBaseItemAlterado = 0

    End If

    Exit Sub
    
Erro_ICMSSubstBaseItem_Validate:

    Cancel = True

    Select Case gErr
    
        Case 103016
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177364)
            
    End Select

    Exit Sub
    
End Sub

Private Sub ICMSSubstValor_Change()

    ICMSSubstValor1.Caption = ICMSSubstValor.Caption

End Sub

Public Sub ICMSSubstValorItem_Change()

    iAlterado = REGISTRO_ALTERADO
    giICMSSubstValorItemAlterado = 1

End Sub

Public Sub ICMSSubstValorItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSSubstValorItem_Validate

    If giICMSSubstValorItemAlterado Then

        If Len(Trim(ICMSSubstValorItem.ClipText)) > 0 Then
        
            lErro = Valor_NaoNegativo_Critica(ICMSSubstValorItem.Text)
            If lErro <> SUCESSO Then gError 103018
        
        End If

        Call BotaoGravarTribItem_Click

        giICMSSubstValorItemAlterado = 0

    End If
    
    Exit Sub
    
Erro_ICMSSubstValorItem_Validate:

    Cancel = True

    Select Case gErr
    
        Case 103018
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177365)
    
    End Select

    Exit Sub

End Sub

Private Sub ICMSValor_Change()

    ICMSValor1.Caption = ICMSValor.Caption

End Sub

Public Sub ICMSValorItem_Change()

    iAlterado = REGISTRO_ALTERADO
    giICMSValorItemAlterado = 1

End Sub

Public Sub ICMSValorItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSValorItem_Validate

    If giICMSValorItemAlterado Then

        If Len(Trim(ICMSValorItem.ClipText)) > 0 Then
        
            lErro = Valor_NaoNegativo_Critica(ICMSValorItem.Text)
            If lErro <> SUCESSO Then gError 1030168
'????? Numeracao de Erro grande demais. mario
        End If

        Call BotaoGravarTribItem_Click

        giICMSValorItemAlterado = 0

    End If
    
    Exit Sub
    
Erro_ICMSValorItem_Validate:

    Cancel = True

    Select Case gErr
    
        Case 1030168
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177366)
            
    End Select
        
    Exit Sub

End Sub

Public Sub IPIAliquotaItem_Change()

    iAlterado = REGISTRO_ALTERADO
    giIPIAliquotaItemAlterado = 1

End Sub

Public Sub IPIAliquotaItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_IPIAliquotaItem_Validate
    
    If giIPIAliquotaItemAlterado Then

        If Len(Trim(IPIAliquotaItem.ClipText)) > 0 Then
            
            lErro = Porcentagem_Critica(IPIAliquotaItem.Text)
            If lErro <> SUCESSO Then gError 103021
            
        End If

        Call BotaoGravarTribItem_Click

        giIPIBaseItemAlterado = 0

    End If
    
    Exit Sub
    
Erro_IPIAliquotaItem_Validate:

    Cancel = True

    Select Case gErr
    
        Case 103021
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177367)
        
    End Select

    Exit Sub

End Sub

Public Sub IPIBaseItem_Change()

    iAlterado = REGISTRO_ALTERADO
    giIPIBaseItemAlterado = 1

End Sub

Public Sub IPIBaseItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_IPIBaseItem_Validate

    If giIPIBaseItemAlterado Then

        If Len(Trim(IPIBaseItem.ClipText)) > 0 Then
            
            lErro = Valor_NaoNegativo_Critica(IPIBaseItem.Text)
            If lErro <> SUCESSO Then gError 103019
        
        End If

        Call BotaoGravarTribItem_Click

        giIPIBaseItemAlterado = 0

    End If
    
    Exit Sub
    
Erro_IPIBaseItem_Validate:

    Cancel = True

    Select Case gErr
    
        Case 103019
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177368)
    
    End Select

    Exit Sub

End Sub

Public Sub IPIPercRedBaseItem_Change()

    iAlterado = REGISTRO_ALTERADO
    giIPIPercRedBaseItemAlterado = 1

End Sub

Public Sub IPIPercRedBaseItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_IPIPercRedBaseItem_Validate

    If giIPIPercRedBaseItemAlterado Then

        If Len(Trim(IPIPercRedBaseItem.Text)) > 0 Then
            
            lErro = Porcentagem_Critica(IPIPercRedBaseItem.Text)
            If lErro <> SUCESSO Then gError 103020
        
        End If
        
        Call BotaoGravarTribItem_Click

        giIPIPercRedBaseItemAlterado = 0

    End If
    
    Exit Sub
    
Erro_IPIPercRedBaseItem_Validate:

    Cancel = True

    Select Case gErr
    
        Case 103020
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177369)
            
    End Select
    
    Exit Sub

End Sub

Private Sub IPIValor_Change()

    IPIValor1.Caption = IPIValor.Caption

End Sub

Public Sub IPIValorItem_Change()

    iAlterado = REGISTRO_ALTERADO
    giIPIValorItemAlterado = 1

End Sub

Public Sub IPIValorItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_IPIValorItem_Validate

    If giIPIValorItemAlterado Then

        If Len(Trim(IPIValorItem.ClipText)) > 0 Then
            
            lErro = Valor_NaoNegativo_Critica(IPIValorItem.Text)
            If lErro <> SUCESSO Then gError 103022
            
        End If

        Call BotaoGravarTribItem_Click

        giIPIValorItemAlterado = 0

    End If

    Exit Sub
    
Erro_IPIValorItem_Validate:

    Cancel = True

    Select Case gErr
    
        Case 103022
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177370)
            
    End Select

    Exit Sub

End Sub

Public Sub IRAliquota_Change()

    iAlterado = REGISTRO_ALTERADO
    giAliqIRAlterada = 1

End Sub

Public Sub IRAliquota_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dIRAliquota As Double, dIRValor As Double

On Error GoTo Erro_IRAliquota_Validate

    If giAliqIRAlterada = 0 Then Exit Sub

    If Len(Trim(IRAliquota.ClipText)) > 0 Then
        
        lErro = Porcentagem_Critica(IRAliquota.Text)
        If lErro <> SUCESSO Then gError 94499
    
    End If
    
    Call BotaoGravarTrib

    giAliqIRAlterada = 0

    Exit Sub

Erro_IRAliquota_Validate:

    Cancel = True

    Select Case gErr
    
        Case 94499

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177371)

    End Select

    Exit Sub

End Sub

Public Sub ISSAliquota_Change()

    iAlterado = REGISTRO_ALTERADO
    giISSAliquotaAlterada = 1

End Sub

Public Sub ISSAliquota_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ISSAliquota_Validate

    If giISSAliquotaAlterada = 0 Then Exit Sub
    
    If Len(Trim(ISSAliquota.ClipText)) > 0 Then
        
        lErro = Porcentagem_Critica(ISSAliquota.Text)
        If lErro <> SUCESSO Then gError 94497
    
    End If

    Call BotaoGravarTrib

    giISSAliquotaAlterada = 0
    
    Exit Sub
    
Erro_ISSAliquota_Validate:

    Cancel = True

    Select Case gErr
    
        Case 94497
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177372)
            
    End Select
        
    Exit Sub
    
End Sub

Public Sub ISSValor_Change()

    iAlterado = REGISTRO_ALTERADO
    giISSValorAlterado = 1

End Sub

Public Sub ISSValor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ISSValor_Validate

    If giISSValorAlterado = 0 Then Exit Sub

    If Len(Trim(ISSValor.ClipText)) > 0 Then
        
        lErro = Valor_NaoNegativo_Critica(ISSValor.Text)
        If lErro <> SUCESSO Then gError 94498
    
    End If

    Call BotaoGravarTrib

    giISSValorAlterado = 0

    Exit Sub
    
Erro_ISSValor_Validate:

    Cancel = True

    Select Case gErr
    
        Case 94498
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177373)
            
    End Select
        
    Exit Sub

End Sub

Private Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objCliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)


End Sub

Public Sub LblTipoTrib_Click()

Dim colSelecao As New Collection
Dim objTipoTrib As New ClassTipoDeTributacaoMovto

    'apenas tipos de saida
    colSelecao.Add "0"
    colSelecao.Add "0"
    
    Call Chama_Tela("TiposDeTribMovtoLista", colSelecao, objTipoTrib, objEventoTiposDeTributacao)

End Sub

Private Sub LblTipoTribItem_Click()

    Call LblTipoTrib_Click

End Sub

Public Sub NaturezaItemLabel_Click()

Dim objNaturezaOp As New ClassNaturezaOp
Dim colSelecao As New Collection
Dim dtDataRef As Date, sSelecao As String

    If Len(Trim(NaturezaOpItem.Text)) > 0 Then objNaturezaOp.sCodigo = NaturezaOpItem.Text

    If Len(Trim(DataEmissao.ClipText)) > 0 Then
        dtDataRef = MaskedParaDate(DataEmissao)
    Else
        dtDataRef = DATA_NULA
    End If
        
    sSelecao = "Codigo >= " & NATUREZA_SAIDA_COD_INICIAL & " AND Codigo <= " & NATUREZA_SAIDA_COD_FINAL & " AND {fn LENGTH(Codigo) } = " & IIf(dtDataRef < DATA_INICIO_CFOP4, "3", "4")
        
    Call Chama_Tela("NaturezaOperacaoLista", colSelecao, objNaturezaOp, objEventoNaturezaOp, sSelecao)

End Sub

Private Sub NaturezaLabel_Click()

Dim objNaturezaOp As New ClassNaturezaOp
Dim colSelecao As New Collection
Dim dtDataRef As Date

    'Se NaturezaOP estiver preenchida coloca no Obj
    objNaturezaOp.sCodigo = NaturezaOp.Text

    If Len(Trim(DataEmissao.ClipText)) > 0 Then
        dtDataRef = MaskedParaDate(DataEmissao)
    Else
        dtDataRef = DATA_NULA
    End If
        
    'selecao p/obter apenas as nat de saida
    colSelecao.Add NATUREZA_SAIDA_COD_INICIAL
    colSelecao.Add NATUREZA_SAIDA_COD_FINAL

    'Chama a Tela de browse de NaturezaOp
    Call Chama_Tela("NaturezaOpLista", colSelecao, objNaturezaOp, objEventoNaturezaOp, "{fn LENGTH(Codigo) } = " & IIf(dtDataRef < DATA_INICIO_CFOP4, "3", "4"))

End Sub

Private Sub NaturezaOp_Change()

    iAlterado = REGISTRO_ALTERADO
    giNaturezaOpAlterada = 1

End Sub

Private Sub NaturezaOp_GotFocus()

Dim iNaturezaAux As Integer
    
    iNaturezaAux = giNaturezaOpAlterada
    Call MaskEdBox_TrataGotFocus(NaturezaOp, iAlterado)
    giNaturezaOpAlterada = iNaturezaAux

End Sub

Private Sub NaturezaOp_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objNaturezaOp As New ClassNaturezaOp
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_NaturezaOp_Validate

    'Se Natureza não está preenchida espelha no frame Tributação
    If Len(Trim(NaturezaOp.ClipText)) = 0 Then
    
        NatOpEspelho.Caption = ""
        DescNatOp.Caption = ""
        
    End If
    
    'Verifica se a NaturezaOP foi informada
    If Len(Trim(NaturezaOp.ClipText)) = 0 Or giNaturezaOpAlterada = 0 Then Exit Sub

    objNaturezaOp.sCodigo = Trim(NaturezaOp.Text)

    If objNaturezaOp.sCodigo < NATUREZA_SAIDA_COD_INICIAL Or objNaturezaOp.sCodigo > NATUREZA_SAIDA_COD_FINAL Then gError 94495
    
    'Lê a NaturezaOp
    lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
    If lErro <> SUCESSO And lErro <> 17958 Then gError 94493

    'Se não existir --> Erro
    If lErro = 17958 Then gError 94494
    
    'Espelha Natureza no frame de Tributação
    NatOpEspelho.Caption = objNaturezaOp.sCodigo
    DescNatOp.Caption = objNaturezaOp.sDescricao
    
    If giTrazendoTribTela = 0 And gbCarregandoTela = False Then Call BotaoGravarTrib
    
    giNaturezaOpAlterada = 0
    
    Exit Sub

Erro_NaturezaOp_Validate:

    Cancel = True

    Select Case gErr

        Case 94493

        Case 94494

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_NATUREZA_OPERACAO", NaturezaOp.Text)
            If vbMsgRes = vbYes Then
                Call Chama_Tela("NaturezaOperacao", objNaturezaOp)
            Else
            End If

        Case 94495
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_SAIDA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177374)

    End Select

    Exit Sub

End Sub

Public Sub NaturezaOpItem_Change()
    
    giNatOpItemAlterado = 1
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub NumeroLabel_Click()

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim colSelecao As Collection

On Error GoTo Erro_NumeroLabel_Click

    lErro = Move_OrcamentoVenda_Memoria(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 84106 '26498

    Call Chama_Tela("OrcamentoVendaLista", colSelecao, objOrcamentoVenda, objEventoNumero)
    
    Exit Sub

Erro_NumeroLabel_Click:

    Select Case gErr

        Case 84106

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177375)

    End Select

    Exit Sub

End Sub


Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    Cliente.Text = objCliente.sNomeReduzido

    'Dispara o Validate de Cliente
    Call Cliente_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNaturezaOp_evSelecao(obj1 As Object)

Dim objNaturezaOp As New ClassNaturezaOp

    Set objNaturezaOp = obj1

    If giFrameAtual = 1 Then

        'Preenche a natureza de Opereração do frame principal
        NaturezaOp.Text = objNaturezaOp.sCodigo
        Call NaturezaOp_Validate(bSGECancelDummy)

    Else
        'Preenche a NatOp do frame de tributação
        NaturezaOpItem.Text = objNaturezaOp.sCodigo
        Call NaturezaOpItem_Validate(bSGECancelDummy)

    End If

    Me.Show

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOrcamentoVenda As ClassOrcamentoVenda

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objOrcamentoVenda = obj1

    lErro = Traz_OrcamentoVenda_Tela(objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 84363 Then gError 84479

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case gErr

        Case 84479
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177376)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As ClassProduto
Dim sProduto As String
Dim lErro As Long

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Verifica se alguma linha está selecionada
    If GridItens.Row < 1 Then Exit Sub

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 94414

    Produto.PromptInclude = False
    Produto.Text = sProduto
    Produto.PromptInclude = True

    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text

    'Faz o Tratamento do produto
    lErro = Produto_Saida_Celula()
    If lErro <> SUCESSO Then gError 94415
    
    Call ComandoSeta_Fechar(Me.Name)
    
    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr
            
        Case 94414
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
        
        Case 94415

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177377)

    End Select

    Exit Sub

End Sub

Private Sub objEventoTiposDeTributacao_evSelecao(obj1 As Object)

Dim objTipoTrib As ClassTipoDeTributacaoMovto

    Set objTipoTrib = obj1

    If giFrameAtualTributacao = 1 Then

        TipoTributacao.Text = objTipoTrib.iTipo
        Call TipoTributacao_Validate(bSGECancelDummy)

    Else

        TipoTributacaoItem.Text = objTipoTrib.iTipo
        Call TipoTributacaoItem_Validate(bSGECancelDummy)

    End If

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1

    'Preenche campo Vendedor
    Vendedor.Text = objVendedor.sNomeReduzido

    iAlterado = 0

    Me.Show

End Sub

Private Sub Opcao_Click()

Dim lErro As Long

On Error GoTo Erro_Opcao_Click

    'Se frame selecionado não for o atual
    If Opcao.SelectedItem.Index <> giFrameAtual Then

        If TabStrip_PodeTrocarTab(giFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'se abriu o tab de tributacao
        If Opcao.SelectedItem.Index = TAB_Tributacao Then
            
            lErro = TributacaoItem_InicializaTab
            If lErro <> SUCESSO Then gError 84308 '27840
            
        '??? Alteração Daniel em 29/10/2002
        ElseIf Opcao.SelectedItem.Index = TAB_Cobranca Then
        
            'Recalcula as parcelas
            Call CobrancaAutomatica_Click
        
        End If

        'Esconde o frame atual, mostra o novo
        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(giFrameAtual).Visible = False

        'Armazena novo valor de giFrameAtual
        giFrameAtual = Opcao.SelectedItem.Index
       
    End If

    Exit Sub

Erro_Opcao_Click:

    Select Case gErr

        Case 84308

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177378)

    End Select

    Exit Sub

End Sub

Public Sub OpcaoTributacao_Click()
Dim lErro As Long
On Error GoTo Erro_OpcaoTributacao_Click

    'Se frame selecionado não for o atual
    If OpcaoTributacao.SelectedItem.Index <> giFrameAtualTributacao Then

        If TabStrip_PodeTrocarTab(giFrameAtualTributacao, OpcaoTributacao, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        FrameTributacao(OpcaoTributacao.SelectedItem.Index).Visible = True
        FrameTributacao(giFrameAtualTributacao).Visible = False
        'Armazena novo valor de giFrameAtualTributacao
        giFrameAtualTributacao = OpcaoTributacao.SelectedItem.Index

        'se abriu o tab de detalhamento
        If OpcaoTributacao.SelectedItem.Index = 2 Then
            lErro = TributacaoItem_InicializaTab
            If lErro <> SUCESSO Then gError 27777
        End If

    End If

    Exit Sub

Erro_OpcaoTributacao_Click:

    Select Case gErr

        Case 27777

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177379)

    End Select

    Exit Sub

End Sub

Private Sub PercAcrescFin_Change()

    iAlterado = REGISTRO_ALTERADO
    giPercAcresFinAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PercAcrescFin_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercAcrescFin_Validate

    If giPercAcresFinAlterado = 0 Then Exit Sub

    If Len(Trim(PercAcrescFin.Text)) > 0 Then
        lErro = Porcentagem_Critica_Negativa(PercAcrescFin)
        If lErro <> SUCESSO Then gError 84091 '26717
    End If

    If Len(Trim(TabelaPreco.Text)) > 0 Then

        lErro = Trata_TabelaPreco()
        If lErro <> SUCESSO Then gError 84092 '46190

    End If

    giPercAcresFinAlterado = 0

    Exit Sub

Erro_PercAcrescFin_Validate:

    Cancel = True


    Select Case gErr

        Case 84091, 84092

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177380)

    End Select

    Exit Sub


End Sub

Private Sub PercentDesc_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PercentDesc_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub PercentDesc_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)


End Sub

Private Sub PercentDesc_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PercentDesc
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PrazoValidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PrecoTotal_Change()

    iAlterado = REGISTRO_ALTERADO


End Sub

Private Sub PrecoTotal_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub PrecoTotal_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub PrecoTotal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PrecoTotal
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PrecoUnitario_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PrecoUnitario_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub PrecoUnitario_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub PrecoUnitario_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PrecoUnitario
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Produto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Produto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True


End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TabelaPreco_Click()

Dim lErro As Long

On Error GoTo Erro_TabelaPreco_Click

    iAlterado = REGISTRO_ALTERADO

    If TabelaPreco.ListIndex = -1 Then Exit Sub

    If objGridItens.iLinhasExistentes = 0 Then Exit Sub

    'Faz o tratamento para a Tabela de Preços escolhida
    lErro = Trata_TabelaPreco()
    If lErro <> SUCESSO Then gError 84013

    Exit Sub

Erro_TabelaPreco_Click:

    Select Case gErr

        Case 84013

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177381)

    End Select

    Exit Sub

End Sub

Private Sub TabelaPreco_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTabelaPreco As New ClassTabelaPreco
Dim iCodigo As Integer

On Error GoTo Erro_TabelaPreco_Validate

    'Verifica se foi preenchida a ComboBox TabelaPreco
    If Len(Trim(TabelaPreco.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o item selecionado na ComboBox TabelaPreco
    If TabelaPreco.Text = TabelaPreco.List(TabelaPreco.ListIndex) Then Exit Sub

    'Verifica se existe o item na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(TabelaPreco, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 84014

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTabelaPreco.iCodigo = iCodigo

        'Tenta ler TabelaPreço com esse código no BD
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 84015 '26539

        If lErro <> SUCESSO Then gError 84016 '26540 'Não encontrou Tabela Preço no BD

        'Encontrou TabelaPreço no BD, coloca no Text da Combo
        TabelaPreco.Text = CStr(objTabelaPreco.iCodigo) & SEPARADOR & objTabelaPreco.sDescricao

        lErro = Trata_TabelaPreco()
        If lErro <> SUCESSO Then gError 84017 '30527

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 84018 '26541

    Exit Sub

Erro_TabelaPreco_Validate:

    Cancel = True

    Select Case gErr

    Case 84014, 84015, 84017

    Case 84016  'Não encontrou Tabela de Preço no BD

        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TABELA_PRECO")

        If vbMsgRes = vbYes Then
            'Preenche o objTabela com o Codigo
            If Len(Trim(TabelaPreco.Text)) > 0 Then objTabelaPreco.iCodigo = CInt(TabelaPreco.Text)
            'Chama a tela de Tabelas de Preço
            Call Chama_Tela("TabelaPrecoCriacao", objTabelaPreco)
        End If

    Case 84018
        Call Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_ENCONTRADA", gErr, TabelaPreco.Text)

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177382)

    End Select

    Exit Sub

End Sub

Private Sub TipoDesconto1_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoDesconto1_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub TipoDesconto1_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub TipoDesconto1_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = TipoDesconto1
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TipoDesconto2_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoDesconto2_GotFocus()

     Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub TipoDesconto2_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub TipoDesconto2_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = TipoDesconto2
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TipoDesconto3_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoDesconto3_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub TipoDesconto3_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub TipoDesconto3_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = TipoDesconto3
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub TipoTributacao_Change()

    iAlterado = REGISTRO_ALTERADO
    giTipoTributacaoAlterado = 1

End Sub

Public Sub TipoTributacao_GotFocus()

Dim iTipoTributacaoAux As Integer

    iTipoTributacaoAux = giTipoTributacaoAlterado
    Call MaskEdBox_TrataGotFocus(TipoTributacao, iAlterado)
    giTipoTributacaoAlterado = iTipoTributacaoAux
    
End Sub

Private Sub TipoTributacao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTipoDeTributacao As New ClassTipoDeTributacaoMovto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TipoTributacao_Validate

    If Len(Trim(TipoTributacao.Text)) = 0 Then
        'Limpa o campo da descrição
        DescTipoTrib.Caption = ""
    End If

    If (giTipoTributacaoAlterado = 1) Then

        objTipoDeTributacao.iTipo = StrParaInt(TipoTributacao.Text)

        If objTipoDeTributacao.iTipo <> 0 Then
            lErro = CF("TipoTributacao_Le", objTipoDeTributacao)
            If lErro <> SUCESSO And lErro <> 27259 Then gError 27663
    
            'Se não encontrou o Tipo da Tributação --> erro
            If lErro = 27259 Then gError 27664
        End If

        DescTipoTrib.Caption = objTipoDeTributacao.sDescricao
        
        Call BotaoGravarTrib

        giTipoTributacaoAlterado = 0

    End If
    
    Exit Sub

Erro_TipoTributacao_Validate:

    Cancel = True


    Select Case gErr

        Case 27663

        Case 27664
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPOTRIBUTACAO", TipoTributacao.Text)

            If vbMsgRes = vbYes Then

                Call Chama_Tela("TipoDeTributacao", objTipoDeTributacao)

            Else
            End If

        Case Else
            vbMsgRes = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177383)

    End Select

    Exit Sub

End Sub

'Por Leo em 02/05/02
Public Sub TipoTributacaoItem_Change()

    giTipoTributacaoItemAlterado = 1
    iAlterado = REGISTRO_ALTERADO


End Sub

'Por Leo em 02/05/02
Public Sub TipoTributacaoItem_GotFocus()

Dim iTipoTributacaoItemAux As Integer

    iTipoTributacaoItemAux = giTipoTributacaoItemAlterado
    
    Call MaskEdBox_TrataGotFocus(TipoTributacaoItem, iAlterado)
    
    giTipoTributacaoItemAlterado = iTipoTributacaoItemAux
    
End Sub

'Por Leo em 02/05/02
Public Sub TipoTributacaoItem_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTributacaoTipo As New ClassTipoDeTributacaoMovto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TipoTributacaoItem_Validate

    'Se trocou o tipo de tributação
    If giTipoTributacaoItemAlterado Then

        objTributacaoTipo.iTipo = StrParaInt(TipoTributacaoItem.Text)
        If objTributacaoTipo.iTipo <> 0 Then

            lErro = CF("TipoTributacao_Le", objTributacaoTipo)
            If lErro <> SUCESSO And lErro <> 27259 Then gError 103006 ' 27711

            'Se não encontrou o Tipo da Tributação --> erro
            If lErro = 27259 Then gError 103007 '58083

            DescTipoTribItem.Caption = objTributacaoTipo.sDescricao

            Call BotaoGravarTribItem_Click
        
        Else
            'Limpa o campo
            DescTipoTribItem.Caption = ""
        
        End If

        giTipoTributacaoItemAlterado = 0

    End If

    Exit Sub

Erro_TipoTributacaoItem_Validate:

    Cancel = True


    Select Case gErr

        Case 103006
        
        Case 103007
        
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPOTRIBUTACAO", TipoTributacaoItem.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("TipoDeTributacao", objTributacaoTipo)
            End If
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177384)

    End Select

    Exit Sub

End Sub

Private Sub UnidadeMed_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UnidadeMed_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UnidadeMed_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub UnidadeMed_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub UnidadeMed_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = UnidadeMed
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 84066 '26583

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case gErr

        Case 84066 '26583

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177385)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissao_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 84067

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case gErr

        Case 84067

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177386)

    End Select

    Exit Sub

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Private Sub Unload(objme As Object)
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

Private Function Inicializa_Grid_Parcelas(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Desconto 1 Tipo")
    objGridInt.colColuna.Add ("Desc. 1 Data")
    objGridInt.colColuna.Add ("Desc. 1 Valor")
    objGridInt.colColuna.Add ("Desc. 1 %")
    objGridInt.colColuna.Add ("Desconto 2 Tipo")
    objGridInt.colColuna.Add ("Desc. 2 Data")
    objGridInt.colColuna.Add ("Desc. 2 Valor")
    objGridInt.colColuna.Add ("Desc. 2 %")
    objGridInt.colColuna.Add ("Desconto 3 Tipo")
    objGridInt.colColuna.Add ("Desc. 3 Data")
    objGridInt.colColuna.Add ("Desc. 3 Valor")
    objGridInt.colColuna.Add ("Desc. 3 %")

    objGridInt.colCampo.Add (DataVencimento.Name)
    objGridInt.colCampo.Add (ValorParcela.Name)
    objGridInt.colCampo.Add (TipoDesconto1.Name)
    objGridInt.colCampo.Add (Desconto1Ate.Name)
    objGridInt.colCampo.Add (Desconto1Valor.Name)
    objGridInt.colCampo.Add (Desconto1Percentual.Name)
    objGridInt.colCampo.Add (TipoDesconto2.Name)
    objGridInt.colCampo.Add (Desconto2Ate.Name)
    objGridInt.colCampo.Add (Desconto2Valor.Name)
    objGridInt.colCampo.Add (Desconto2Percentual.Name)
    objGridInt.colCampo.Add (TipoDesconto3.Name)
    objGridInt.colCampo.Add (Desconto3Ate.Name)
    objGridInt.colCampo.Add (Desconto3Valor.Name)
    objGridInt.colCampo.Add (Desconto3Percentual.Name)


    'Controles que participam do Grid
    iGrid_Vencimento_col = 1
    iGrid_ValorParcela_Col = 2
    iGrid_Desc1Codigo_Col = 3
    iGrid_Desc1Ate_Col = 4
    iGrid_Desc1Valor_Col = 5
    iGrid_Desc1Perc_Col = 6
    iGrid_Desc2Codigo_Col = 7
    iGrid_Desc2Ate_Col = 8
    iGrid_Desc2Valor_Col = 9
    iGrid_Desc2Perc_Col = 10
    iGrid_Desc3Codigo_Col = 11
    iGrid_Desc3Ate_Col = 12
    iGrid_Desc3Valor_Col = 13
    iGrid_Desc3Perc_Col = 14

    'Grid do GridInterno
    objGridInt.objGrid = GridParcelas

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_PARCELAS + 1

    'Habilita a execução da Rotina_Grid_Enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridParcelas.ColWidth(0) = 700

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Parcelas = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

Dim iIncremento As Integer
Dim objUserControl As Object

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Cor")
    objGridInt.colColuna.Add ("Detalhe")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Preço Unitário")
    'precodesc
    Call CF("Orcamento_Inicializa_Grid_Itens1", objGridInt)
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Total")
    objGridInt.colColuna.Add ("Data Entrega")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoProduto.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (PrecoUnitario.Name)
    'precodesc
    Set objUserControl = Me
    Call CF("Orcamento_Inicializa_Grid_Itens2", objGridInt, objUserControl)
    objGridInt.colCampo.Add (PercentDesc.Name)
    objGridInt.colCampo.Add (Desconto.Name)
    objGridInt.colCampo.Add (PrecoTotal.Name)
    objGridInt.colCampo.Add (DataEntrega.Name)

    'Colunas do Grid
    iGrid_ItemProduto_Col = 0
    iGrid_Produto_Col = 1
    iGrid_DescProduto_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_Quantidade_Col = 4
    iGrid_PrecoUnitario_Col = 5
    'precodesc
    Call CF("Orcamento_Inicializa_Grid_Itens3", iIncremento)
    iGrid_PercDesc_Col = 6 + iIncremento
    iGrid_Desconto_Col = 7 + iIncremento
    iGrid_PrecoTotal_Col = 8 + iIncremento
    iGrid_DataEntrega_Col = 9 + iIncremento

    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Private Function Trata_TabelaPreco() As Long

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Trata_TabelaPreco

    If Not gbCarregandoTela Then

        For iLinha = 1 To objGridItens.iLinhasExistentes

            lErro = Trata_TabelaPreco_Item(iLinha)
            If lErro <> SUCESSO Then gError 84019

        Next

        'Calcula o Valor Total da Nota
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 84020

    End If

    Trata_TabelaPreco = SUCESSO

    Exit Function

Erro_Trata_TabelaPreco:

    Trata_TabelaPreco = gErr

    Select Case gErr

        Case 84019 'tratado na rotina chamada

        Case 84020

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177387)

    End Select

    Exit Function

End Function

Private Function Trata_TabelaPreco_Item(iLinha As Integer) As Long
'faz tratamento de tabela de preço para um ítem (produto)

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_Trata_TabelaPreco_Item

    'Verifica se o Produto está preenchido
    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 84021 '39147

    If iPreenchido <> PRODUTO_VAZIO And Len(Trim(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col))) > 0 Then

        objTabelaPrecoItem.sCodProduto = sProduto
        objTabelaPrecoItem.iCodTabela = Codigo_Extrai(TabelaPreco.Text)
        objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa

        'Lê a Tabela preço para filialEmpresa
        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
        If lErro <> SUCESSO And lErro <> 28014 Then gError 84022 '39148

        'Se não encontrar
        If lErro = 28014 Then
            objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
            'Lê a Tabela de Preço a nível de Empresa toda
            lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
            If lErro <> SUCESSO And lErro <> 28014 Then gError 84023 '39149

        End If

        'Se  conseguir ler a Tabela de Preços
        If lErro = SUCESSO Then
            'Calcula o Preco Unitário do Ítem
            lErro = PrecoUnitario_Calcula(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), objTabelaPrecoItem, dPrecoUnitario)
            If lErro <> SUCESSO Then gError 84024 '39150
            'Coloca no Grid
            If dPrecoUnitario > 0 Then
                GridItens.TextMatrix(iLinha, iGrid_PrecoUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
            Else
                GridItens.TextMatrix(iLinha, iGrid_PrecoUnitario_Col) = ""
            End If
            'Calcula o Preco Total do Ítem
            Call PrecoTotal_Calcula(iLinha)

         End If

    End If

    Trata_TabelaPreco_Item = SUCESSO

    Exit Function

Erro_Trata_TabelaPreco_Item:

    Trata_TabelaPreco_Item = gErr

    Select Case gErr

        Case 84021, 84022, 84023, 84024 'tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177388)

    End Select

    Exit Function

End Function

'mario
Function ValorTotal_Calcula() As Long
'Calcula o Valor Total do Pedido

Dim dValorDespesas As Double
Dim dValorProdutos As Double
Dim dValorTotal As Double
Dim dValorFrete As Double
Dim dValorSeguro As Double
Dim dValorIPI As Double
Dim dValorICMSSubst As Double
Dim vbMsgRes As VbMsgBoxResult
Dim dValorAposIR As Double
Dim dValorIRRF As Double
Dim lErro As Long
Dim dValorISS As Double

On Error GoTo Erro_ValorTotal_Calcula

    If Not gbCarregandoTela Then
        'Atualiza os valores de tributação
        lErro = AtualizarTributacao()
        If lErro <> SUCESSO Then gError 101102
    End If

    'Recolhe os valores da tela
    If Len(Trim(ValorProdutos.Caption)) > 0 And IsNumeric(ValorProdutos.Caption) Then dValorProdutos = CDbl(ValorProdutos.Caption)
    If Len(Trim(ValorFrete.Text)) > 0 And IsNumeric(ValorFrete.Text) Then dValorFrete = CDbl(ValorFrete.Text)
    If Len(Trim(ValorIRRF.Text)) > 0 And IsNumeric(ValorIRRF.Text) Then dValorIRRF = CDbl(ValorIRRF.Text)
    If Len(Trim(ValorSeguro.Text)) > 0 And IsNumeric(ValorSeguro.Text) Then dValorSeguro = CDbl(ValorSeguro.Text)
    If Len(Trim(ValorDespesas.Text)) > 0 And IsNumeric(ValorDespesas.Text) Then dValorDespesas = CDbl(ValorDespesas.Text)
    If Len(Trim(ICMSSubstValor1.Caption)) > 0 And IsNumeric(ICMSSubstValor1.Caption) Then dValorICMSSubst = CDbl(ICMSSubstValor1.Caption)
    If Len(Trim(IPIValor1.Caption)) > 0 And IsNumeric(IPIValor1.Caption) Then dValorIPI = CDbl(IPIValor1.Caption)
    If Len(Trim(ISSValor.Text)) > 0 And IsNumeric(ISSValor.Text) And ISSIncluso.Value = vbUnchecked Then dValorISS = CDbl(ISSValor.Text)

    'Calcula o Valor Total
    dValorTotal = dValorProdutos + dValorFrete + dValorSeguro + dValorDespesas + dValorIPI + dValorICMSSubst

    dValorAposIR = dValorTotal - dValorIRRF

    If dValorTotal <> 0 And dValorAposIR < 0 And dValorIRRF > 0 Then
        vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_IR_FONTE_MAIOR_VALOR_TOTAL", dValorIRRF, dValorTotal)
        ValorIRRF.Text = ""

        '????? Validar com Jones a necessidade de chamar esta funcao
        Call ValorIRRF_Validate(bSGECancelDummy)

        'Faz a atualização dos valores da tributação
        lErro = AtualizarTributacao()
        If lErro <> SUCESSO Then gError 101103

    End If
    
    If Not gbCarregandoTela Then
    
        'Faz o cálculo automático das comissões
        lErro = Cobranca_Automatica()
        If lErro <> SUCESSO Then gError 56912
        
    End If

    ValorTotal.Caption = Format(dValorTotal, "Standard")

    ValorTotal_Calcula = SUCESSO

    Exit Function

Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = gErr

    Select Case gErr

        Case 84025, 84026, 84027, 84028 'tratados nas rotinas chamadas

        Case 84249, 84250, 101102, 101103

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177389)

    End Select

    Exit Function

End Function

Private Function PrecoUnitario_Calcula(sUM As String, objTabelaPrecoItem As ClassTabelaPrecoItem, dPrecoUnitario As Double) As Long
'Calcula o Preço unitário do item de acordo com a UM e a tabela de preço

Dim objProduto As New ClassProduto
Dim objUM As New ClassUnidadeDeMedida
Dim objUMEst As New ClassUnidadeDeMedida
Dim dFator As Double
Dim lErro As Long
Dim dPercAcresFin As Double
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_PrecoUnitario_Calcula

    objProduto.sCodigo = objTabelaPrecoItem.sCodProduto
    'Lê o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 84029 '26638

    If lErro = 28030 Then gError 84030 '26639
    'Converte a quantidade para a UM de Venda
    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMVenda, dFator)
    If lErro <> SUCESSO Then gError 84031 '26640

    dPrecoUnitario = objTabelaPrecoItem.dPreco * dFator

    'Recolhe o percentual de acréscimo financeiro
    dPercAcresFin = StrParaDbl(PercAcrescFin.Text) / 100

    'Calcula o Preço unitário
    If dPercAcresFin <> 0 Then
        dPrecoUnitario = dPrecoUnitario * (1 + dPercAcresFin)
    End If

    PrecoUnitario_Calcula = SUCESSO

    Exit Function

Erro_PrecoUnitario_Calcula:

    PrecoUnitario_Calcula = gErr

    Select Case gErr

        Case 84029, 84031

        Case 84030
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objTabelaPrecoItem.sCodProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177390)

    End Select

    Exit Function

End Function

Private Sub ValorDesconto_Change()

    iAlterado = REGISTRO_ALTERADO
    giValorDescontoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorDesconto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorDesconto As Double
Dim dValorProdutos As Double
Dim iIndice As Integer

On Error GoTo Erro_ValorDesconto_Validate

    'Verifica se o valor foi alterado
    If giValorDescontoAlterado = 0 Then Exit Sub

    'Vale o desconto que foi colocado aqui
    giValorDescontoManual = 1

    dValorDesconto = 0

    'Calcula a soma dos valores de produtos
    For iIndice = 1 To objGridItens.iLinhasExistentes
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
            If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorProdutos = dValorProdutos + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
        End If
    Next

    'Verifica se o Valor está preenchido
    If Len(Trim(ValorDesconto.Text)) > 0 Then

        'Faz a Crítica do Valor digitado
        lErro = Valor_NaoNegativo_Critica(ValorDesconto.Text)
        If lErro <> SUCESSO Then gError 84055 '26652

        dValorDesconto = CDbl(ValorDesconto.Text)

        'Coloca o Valor formatado na tela
        ValorDesconto.Text = Format(dValorDesconto, "Standard")

        If dValorDesconto > dValorProdutos Then gError 84056 '26653

        dValorProdutos = dValorProdutos - dValorDesconto

    End If

    ValorProdutos.Caption = Format(dValorProdutos, "Standard")

    'Para tributação
    gobjOrcamentoVenda.dValorDesconto = dValorDesconto

    'Recalcula valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 84057 '51038

    giValorDescontoAlterado = 0

    Exit Sub

Erro_ValorDesconto_Validate:

    Cancel = True

    Select Case gErr

        Case 84055, 84057

        Case 84056
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_MAIOR", gErr, dValorDesconto, dValorProdutos)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177391)

    End Select

    Exit Sub

End Sub

Private Sub ValorDespesas_Change()

    giValorDespesasAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorDespesas_Validate(Cancel As Boolean)

Dim dValorDespesas As Double
Dim lErro As Long

On Error GoTo Erro_ValorDespesas_Validate

    If giValorDespesasAlterado = 0 Then Exit Sub

    'Se  estiver preenchido
    If Len(Trim(ValorDespesas.Text)) > 0 Then

        'Faz a crítica do valor
        lErro = Valor_NaoNegativo_Critica(ValorDespesas.Text)
        If lErro <> SUCESSO Then gError 84108 '35885

        dValorDespesas = CDbl(ValorDespesas.Text)

        'coloca o valor formatado na tela
        ValorDespesas.Text = Format(dValorDespesas, "Standard")

    End If

    'Para tributação
    gobjOrcamentoVenda.dValorOutrasDespesas = dValorDespesas
    
    'Recalcula valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 84109 '51039

    giValorDespesasAlterado = 0

    Exit Sub

Erro_ValorDespesas_Validate:

    Cancel = True

    Select Case gErr

        Case 84108

        Case 84109

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177392)

    End Select

    Exit Sub

End Sub

Private Sub ValorFrete_Change()

    iAlterado = REGISTRO_ALTERADO
    giValorFreteAlterado = 1

End Sub

Private Sub ValorFrete_Validate(Cancel As Boolean)

Dim dValorFrete As Double
Dim lErro As Long

On Error GoTo Erro_ValorFrete_Validate

    If giValorFreteAlterado = 0 Then Exit Sub

    dValorFrete = 0

    If Len(Trim(ValorFrete.Text)) > 0 Then

        'Critica se valor é não negativo
        lErro = Valor_NaoNegativo_Critica(ValorFrete.Text)
        If lErro <> SUCESSO Then gError 84058 '26650

        dValorFrete = CDbl(ValorFrete.Text)

        ValorFrete.Text = Format(dValorFrete, "Standard")

    End If

    'Para tributação
    gobjOrcamentoVenda.dValorFrete = dValorFrete

    'Recalcula valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 84059 '51040

    giValorFreteAlterado = 0

    Exit Sub

Erro_ValorFrete_Validate:

    Cancel = True

    Select Case gErr

        Case 84058, 84059

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177393)

    End Select

    Exit Sub

End Sub

'por Leo em 02/05/02
'mario
Public Sub ValorIRRF_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorIRRF As Double
Dim dValorTotal As Double

On Error GoTo Erro_ValorIRRF_Validate
    
    If giValorIRRFAlterado = 0 Then Exit Sub

    'Verifica se ValorIRRF foi preenchido
    If Len(Trim(ValorIRRF.Text)) > 0 Then

        'Critica o Valor
        lErro = Valor_NaoNegativo_Critica(ValorIRRF.Text)
        If lErro <> SUCESSO Then gError 103000

        dValorIRRF = CDbl(ValorIRRF.Text)

        ValorIRRF.Text = Format(dValorIRRF, "Standard")

        If Len(Trim(ValorTotal.Caption)) > 0 Then dValorTotal = StrParaDbl(ValorTotal.Caption)

        If dValorIRRF > dValorTotal Then gError 103001

    End If

    Call BotaoGravarTrib
    
    giValorIRRFAlterado = 0

    Exit Sub

Erro_ValorIRRF_Validate:

    Cancel = True

    Select Case gErr

        Case 103000

        Case 103001
            lErro = Rotina_Erro(vbOKOnly, "ERRO_IR_FONTE_MAIOR_VALOR_TOTAL", gErr, dValorIRRF, dValorTotal)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177394)

    End Select

    Exit Sub

End Sub

Private Sub ValorParcela_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorParcela_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub ValorParcela_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub ValorParcela_Validate(Cancel As Boolean)

    Dim lErro As Long

    Set objGridParcelas.objControle = ValorParcela
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ValorSeguro_Change()

    iAlterado = REGISTRO_ALTERADO
    giValorSeguroAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorSeguro_Validate(Cancel As Boolean)

Dim dValorSeguro As Double
Dim lErro As Long

On Error GoTo Erro_Valorseguro_Validate

    If giValorSeguroAlterado = 0 Then Exit Sub

    dValorSeguro = 0

    If Len(Trim(ValorSeguro.Text)) > 0 Then

        'Critica se valor é não negativo
        lErro = Valor_NaoNegativo_Critica(ValorSeguro.Text)
        If lErro <> SUCESSO Then gError 84060 '26651

        dValorSeguro = CDbl(ValorSeguro.Text)

        ValorSeguro.Text = Format(dValorSeguro, "Standard")

    End If

    'Para tributação
    gobjOrcamentoVenda.dValorSeguro = dValorSeguro

    'Recalcula valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 84061 '51041

    giValorSeguroAlterado = 0

    Exit Sub

Erro_Valorseguro_Validate:

    Cancel = True

    Select Case gErr

        Case 84060, 84061

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177395)

    End Select

    Exit Sub

End Sub

'mario
Private Function Move_GridItens_Memoria(objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Move Grid Itens para memória

Dim lErro As Long, iIndice As Integer

On Error GoTo Erro_Move_GridItens_Memoria

    For iIndice = 1 To objGridItens.iLinhasExistentes

        lErro = Move_GridItem_Memoria(objOrcamentoVenda, iIndice, GridItens.TextMatrix(iIndice, iGrid_Produto_Col))
        If lErro <> SUCESSO Then gError 84102

        '********************* TRATAMENTO DE GRADE *****************
        Call Move_ItensGrade_Tela(objOrcamentoVenda.colItens(iIndice).colItensRomaneioGrade, gobjOrcamentoVenda.colItens(iIndice).colItensRomaneioGrade)

    Next

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = gErr

    Select Case gErr

        Case 84102

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177396)

    End Select

    Exit Function

End Function

Private Function Move_GridItem_Memoria(objOrcamentoVenda As ClassOrcamentoVenda, iIndice As Integer, sProduto1 As String) As Long
'Recolhe do Grid os dados do item orçamento no parametro

Dim lErro As Long
Dim sProduto As String
Dim objItemOrcamento As New ClassItemOV, objTributacaoOV As New ClassTributacaoOV
Dim iPreenchido As Integer

On Error GoTo Erro_Move_GridItem_Memoria

    Set objItemOrcamento = New ClassItemOV

    'Verifica se o Produto está preenchido
    If Len(Trim(sProduto1)) > 0 Then

        'Formata o produto
        lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 84103 '27682

        objItemOrcamento.sProduto = sProduto
    End If

    If Len(Trim(sProduto1)) = 0 Or iPreenchido = PRODUTO_VAZIO Then gError 84104 '20767

    'Armazena os dados do item
    objItemOrcamento.sUnidadeMed = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
    objItemOrcamento.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
    objItemOrcamento.dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col))
    objItemOrcamento.dPrecoTotal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
    objItemOrcamento.dValorDesconto = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
    objItemOrcamento.dtDataEntrega = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataEntrega_Col))
    objItemOrcamento.dValorDesconto = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
    objItemOrcamento.sDescricao = GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)
    objItemOrcamento.iFilialEmpresa = giFilialEmpresa
    
    If gobjOrcamentoVenda.colItens.Count >= iIndice Then
        Set objItemOrcamento.objTributacaoItemOV = gobjOrcamentoVenda.colItens.Item(iIndice).objTributacaoItemOV
    Else
        Set objItemOrcamento.objTributacaoItemOV = Nothing
    End If
    
    'Adiciona o item na colecao de itens do orçamento de venda
     objOrcamentoVenda.colItens.Add objItemOrcamento

    Move_GridItem_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItem_Memoria:

    Move_GridItem_Memoria = gErr

    Select Case gErr

        Case 84103

        Case 84104
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177397)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula
    'aquii está devolvendo erro em vez de sucesso
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridParcelas
            Case GridParcelas.Name

                lErro = Saida_Celula_GridParcelas(objGridInt)
                If lErro <> SUCESSO Then gError 84132 '26064

            'Se for o GridItens
            Case GridItens.Name

                lErro = Saida_Celula_GridItens(objGridInt)
                If lErro <> SUCESSO Then gError 84133 '26065


        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 84134 '26068

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 84132, 84133, 84134

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177398)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridParcelas(objGridInt As AdmGrid) As Long
'Faz a crítica do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridParcelas

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case iGrid_Vencimento_col
            lErro = Saida_Celula_DataVencimento(objGridInt)
            If lErro <> SUCESSO Then gError 84126 '26784

        Case iGrid_ValorParcela_Col
            lErro = Saida_Celula_ValorParcela(objGridInt)
            If lErro <> SUCESSO Then gError 84127 '26785

        Case iGrid_Desc1Codigo_Col, iGrid_Desc2Codigo_Col, iGrid_Desc3Codigo_Col
            lErro = Saida_Celula_TipoDesconto(objGridInt)
            If lErro <> SUCESSO Then gError 84128 '26786

        Case iGrid_Desc1Ate_Col, iGrid_Desc2Ate_Col, iGrid_Desc3Ate_Col
            lErro = Saida_Celula_DescontoData(objGridInt)
            If lErro <> SUCESSO Then gError 84129 '26787

        Case iGrid_Desc1Valor_Col, iGrid_Desc2Valor_Col, iGrid_Desc3Valor_Col
            lErro = Saida_Celula_DescontoValor(objGridInt)
            If lErro <> SUCESSO Then gError 84130 '26830

        Case iGrid_Desc1Perc_Col, iGrid_Desc2Perc_Col, iGrid_Desc3Perc_Col
            lErro = Saida_Celula_DescontoPerc(objGridInt)
            If lErro <> SUCESSO Then gError 84131 '26788

    End Select

    Saida_Celula_GridParcelas = SUCESSO

    Exit Function

Erro_Saida_Celula_GridParcelas:

    Saida_Celula_GridParcelas = gErr

    Select Case gErr

        Case 84126, 84127, 84128, 84129, 84130, 84131

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177399)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridItens(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridItens

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Se for a de Produto
        Case iGrid_Produto_Col
            lErro = Saida_Celula_Produto(objGridInt)
            If lErro <> SUCESSO Then gError 84135 '26593

        'Se for a de Unidade de Medida
        Case iGrid_UnidadeMed_Col
            lErro = Saida_Celula_UM(objGridInt)
            If lErro <> SUCESSO Then gError 84136 '26594

        'Se for a de Quantidade Pedida
        Case iGrid_Quantidade_Col
            lErro = Saida_Celula_Quantidade(objGridInt)
            If lErro <> SUCESSO Then gError 84137 '26595

        'Se for a de Preço Unitário
        Case iGrid_PrecoUnitario_Col
            lErro = Saida_Celula_PrecoUnitario(objGridInt)
            If lErro <> SUCESSO Then gError 84139 '26596

        'Se for a de Percentual de Desconto
        Case iGrid_PercDesc_Col
            lErro = Saida_Celula_PercentDesc(objGridInt)
            If lErro <> SUCESSO Then gError 84140 '26599

        'Se for a de Data de Entrega
        Case iGrid_DataEntrega_Col
            lErro = Saida_Celula_DataEntrega(objGridInt)
            If lErro <> SUCESSO Then gError 84141 '26601

    End Select

    Saida_Celula_GridItens = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItens:

    Saida_Celula_GridItens = gErr

    Select Case gErr

        Case 84135, 84136, 84137, 84138, 84139, 84140, 84141

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177400)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim dValorTotal As Double
Dim dValor As Double

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Codigo.Text)) = 0 Then gError 84308 '26807
    If Len(Trim(Cliente.Text)) = 0 Then gError 84309 '26808
    If Len(Trim(Filial.Text)) = 0 Then gError 84310 '26809
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 84311 '26810
    If Len(Trim(NaturezaOp.Text)) = 0 Then gError 94496

    dValor = CDbl(ValorTotal.Caption)

    If dValor < 0 Then gError 84312 ''30594

    lErro = Valida_Grid_Itens()
    If lErro <> SUCESSO Then gError 84313 '26812

    lErro = Valida_Grid_Parcelas()
    If lErro <> SUCESSO Then gError 84314 '26816

    lErro = Move_OrcamentoVenda_Memoria(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 84315 '26829

'????? Trocar a numeracao de erro e colocar tratador
    lErro = Valida_Tributacao_Gravacao()
    If lErro <> SUCESSO Then gError 56931
    
    'Grava no BD
    lErro = CF("OrcamentoVenda_Grava", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 84317 '46183

    'Incluído por Luiz Nogueira em 04/06/03
    'Se for para imprimir o orçamento depois da gravação
    If ImprimeOrcamentoGravacao.Value = vbChecked Then
        
        'Dispara função para imprimir orçamento
        lErro = Orcamento_Imprime(Trim(objOrcamentoVenda.lCodigo))
        If lErro <> SUCESSO Then gError 102239
    
    End If

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 56931
        
        Case 84308
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PEDIDO_NAO_PREENCHIDO", gErr)

        Case 84309
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 84310
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_INFORMADA", gErr)

        Case 84311
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)

        Case 84312
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_OV_NEGATIVO", gErr)

        Case 84313, 84314, 84315, 84316, 84317

        Case 94496
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177401)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim lErro As Long
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim objUM As ClassUnidadeDeMedida
Dim sUM As String
Dim iTipo As Integer
Dim sUnidadeMed As String
Dim iIndice As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    'Formata o produto do grid de itens
    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 84147 '31389

    Select Case objControl.Name

        Case Produto.Name
            'Se o produto estiver preenchido desabilita
            If iProdutoPreenchido <> PRODUTO_VAZIO Then
                Produto.Enabled = False
            Else
                Produto.Enabled = True
            End If


        Case UnidadeMed.Name
            'guarda a um go grid nessa coluna
            sUM = GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)

            UnidadeMed.Enabled = True

            'Guardo o valor da Unidade de Medida da Linha
            sUnidadeMed = UnidadeMed.Text

            UnidadeMed.Clear

            If iProdutoPreenchido <> PRODUTO_VAZIO Then

                objProduto.sCodigo = sProdutoFormatado
                'Lê o produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 84148 '26644

                If lErro = 28030 Then gError 84149 '26645

                objClasseUM.iClasse = objProduto.iClasseUM
                'Lê as UMs do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 84150 '26646
                'Carrega a combo de UMs
                For Each objUM In colSiglas
                    UnidadeMed.AddItem objUM.sSigla
                Next

                'Tento selecionar na Combo a Unidade anterior
                If UnidadeMed.ListCount <> 0 Then

                    For iIndice = 0 To UnidadeMed.ListCount - 1

                        If UnidadeMed.List(iIndice) = sUnidadeMed Then
                            UnidadeMed.ListIndex = iIndice
                            Exit For
                        End If
                    Next
                End If

            Else
                UnidadeMed.Enabled = False
            End If

        Case PrecoUnitario.Name, PercentDesc.Name, DataEntrega.Name
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        '******** O TRATAMENTO DE QUANTIDADE FOI DESTACADO PARA TRATAR GRADE
        Case Quantidade.Name
                
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Or Left(GridItens.TextMatrix(iLinha, 0), 1) = "#" Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If


        Case Desconto1Ate.Name, Desconto1Valor.Name, Desconto1Percentual.Name
            'Habilita os campos de desconto em sequencia
            If Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_Desc1Codigo_Col))) = 0 Then
                objControl.Enabled = False
            Else
                iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Desc1Codigo_Col))
                If objControl.Name = Desconto1Ate.Name Then
                    objControl.Enabled = True
                ElseIf objControl.Name = Desconto1Valor.Name And (iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Or iTipo = VALOR_FIXO) Then
                    Desconto1Valor.Enabled = True
                ElseIf objControl.Name = Desconto1Percentual.Name And (iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = Percentual) Then
                    Desconto1Percentual.Enabled = True
                Else
                    objControl.Enabled = False
                End If
            End If

        Case Desconto2Ate.Name, Desconto2Valor.Name, Desconto2Percentual.Name
            'Habilita os campos de desconto em sequencia
            iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Desc2Codigo_Col))
            If Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_Desc2Codigo_Col))) = 0 Then
                objControl.Enabled = False
            Else
                If objControl.Name = Desconto2Ate.Name Then
                    objControl.Enabled = True
                ElseIf objControl.Name = Desconto2Valor.Name And (iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Or iTipo = VALOR_FIXO) Then
                    Desconto2Valor.Enabled = True
                ElseIf objControl.Name = Desconto2Percentual.Name And (iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = Percentual) Then
                    Desconto2Percentual.Enabled = True
                Else
                    objControl.Enabled = False
                End If
            End If

        Case Desconto3Ate.Name, Desconto3Valor.Name, Desconto3Percentual.Name
            'Habilita os campos de desconto em sequencia
            iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Desc3Codigo_Col))
            If Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_Desc3Codigo_Col))) = 0 Then
                objControl.Enabled = False
            Else
                If objControl.Name = Desconto3Ate.Name Then
                    objControl.Enabled = True
                ElseIf objControl.Name = Desconto3Valor.Name And (iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Or iTipo = VALOR_FIXO) Then
                    Desconto3Valor.Enabled = True
                ElseIf objControl.Name = Desconto3Percentual.Name And (iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = Percentual) Then
                    Desconto3Percentual.Enabled = True
                Else
                    objControl.Enabled = False
                End If
            End If


        Case ValorParcela.Name
            'Se o vencimento estiver preenchido, habilita o controle
            If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_col))) = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case TipoDesconto2.Name, TipoDesconto3.Name
            'Habilita os campos de desconto em sequencia
            If Len(Trim(GridParcelas.TextMatrix(iLinha, GridParcelas.Col - 4))) = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 84148, 84150, 84147

        Case 84149
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177402)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto Data que está deixando de ser a corrente

Dim lErro As Long
Dim sProduto As String

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    If Len(Trim(Produto.ClipText)) > 0 Then

        lErro = Produto_Saida_Celula()
        If lErro <> SUCESSO And lErro <> 26658 Then gError 84152
        If lErro = 26658 Then gError 84151
    End If

    'Necessário para o funcionamento da Rotina_Grid_Enable
    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84153

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 84151 To 84153
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177403)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UM(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Unidadede Medida que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UM

    Set objGridInt.objControle = UnidadeMed

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84154 '26627

    Saida_Celula_UM = SUCESSO

    Exit Function

Erro_Saida_Celula_UM:

    Saida_Celula_UM = gErr

    Select Case gErr

        Case 84154
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177404)

    End Select

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidadeque está deixando de ser a corrente

Dim lErro As Long
Dim bQuantidadeIgual As Boolean
Dim iIndice As Integer
Dim dPrecoUnitario As Double
Dim dQuantAnterior As Double

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    bQuantidadeIgual = False

    If Len(Quantidade.Text) > 0 Then

        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 84157 '26665

        Quantidade.Text = Formata_Estoque(Quantidade.Text)

    End If

    'Comparação com quantidade anterior
    dQuantAnterior = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
    If dQuantAnterior = StrParaDbl(Quantidade.Text) Then bQuantidadeIgual = True

    'Passa quantidade para o grid (p/ usar PrecoTotal_Calcula)
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84161 '59727

    'Preço unitário
    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col))

    'Recalcula preço do ítem e valor total da nota
    If dPrecoUnitario > 0 And Not bQuantidadeIgual Then
        Call PrecoTotal_Calcula(GridItens.Row)
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 84162 '51037
    End If

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 84161, 84157, 84162
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177405)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PrecoUnitario(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Preço Unitário que está deixando de ser a corrente

Dim lErro As Long
Dim bPrecoUnitarioIgual As Boolean

On Error GoTo Erro_Saida_Celula_PrecoUnitario

    bPrecoUnitarioIgual = False

    Set objGridInt.objControle = PrecoUnitario

    If Len(Trim(PrecoUnitario.Text)) > 0 Then

        lErro = Valor_Positivo_Critica(PrecoUnitario.Text)
        If lErro <> SUCESSO Then gError 84170  '26684

    End If

    'Comparação com Preço Unitário anterior
    If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col)) = StrParaDbl(PrecoUnitario.Text) Then bPrecoUnitarioIgual = True

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84171 '26685

    If Not bPrecoUnitarioIgual Then

        Call PrecoTotal_Calcula(GridItens.Row)
        
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 84172 '51042

    End If

   Saida_Celula_PrecoUnitario = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoUnitario:

    Saida_Celula_PrecoUnitario = gErr


    Select Case gErr

        Case 84170, 84171, 84172
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177406)

    End Select

    Exit Function

End Function

Function Saida_Celula_PercentDesc(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Percentual Desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentDesc As Double
Dim dPrecoUnitario As Double
Dim dDesconto As Double
Dim dValorTotal As Double
Dim dQuantidade As Double
Dim sValorPercAnterior As String

On Error GoTo Erro_Saida_Celula_PercentDesc

    Set objGridInt.objControle = PercentDesc

    If Len(PercentDesc.Text) > 0 Then
        'Critica a porcentagem
        lErro = Porcentagem_Critica(PercentDesc.Text)
        If lErro <> SUCESSO Then gError 84329 '26694

        dPercentDesc = CDbl(PercentDesc.Text)

        If Format(dPercentDesc, "#0.#0\%") <> GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) Then
            'se for igual a 100% -> erro
            If dPercentDesc = 100 Then gError 84330 '26695

            PercentDesc.Text = Format(dPercentDesc, "Fixed")

        End If

    Else

        dDesconto = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col))
        dValorTotal = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col))

        GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = ""
        GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col) = Format(dValorTotal + dDesconto, "Standard")

    End If

    sValorPercAnterior = GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col)

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84331 '26696
    'Se foi alterada
    If Format(dPercentDesc, "#0.#0\%") <> sValorPercAnterior Then
        
        'Recalcula o preço total
        Call PrecoTotal_Calcula(GridItens.Row)

        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 84333 '51044
        
        'Preenche GridParcelas a partir da Condição de Pagto
        lErro = Cobranca_Automatica()
        If lErro <> SUCESSO Then gError 84332

    End If

    Saida_Celula_PercentDesc = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentDesc:

    Saida_Celula_PercentDesc = gErr

    Select Case gErr

        Case 84329, 84331, 84333
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 84330
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 84332

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177407)

    End Select

    Exit Function

End Function

Function Saida_Celula_DataEntrega(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Data Entrega que está deixando de ser a corrente

Dim lErro As Long
Dim dtDataEntrega As Date
Dim dtDataEmissao As Date

On Error GoTo Erro_Saida_Celula_DataEntrega

    Set objGridInt.objControle = DataEntrega

    If Len(Trim(DataEntrega.ClipText)) > 0 Then
        'Critica a Data informada
        lErro = Data_Critica(DataEntrega.Text)
        If lErro <> SUCESSO Then gError 84173 ' 26697
        'Se data de emissão estiver preenchida
        If Len(Trim(DataEmissao.ClipText)) > 0 Then

            dtDataEntrega = CDate(DataEntrega.Text)
            dtDataEmissao = CDate(DataEmissao.Text)
            'Veerifica se a data de emissão é maior que a data de entrega
            If dtDataEntrega < dtDataEmissao Then gError 84174 '26698

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84175 '26699

    Saida_Celula_DataEntrega = SUCESSO

    Exit Function

Erro_Saida_Celula_DataEntrega:

    Saida_Celula_DataEntrega = gErr

    Select Case gErr

        Case 84173, 84175
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 84174
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_MAIOR_DATAENTREGA", gErr, dtDataEntrega, dtDataEmissao)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177408)

    End Select

    Exit Function

End Function

Function Produto_Saida_Celula() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim iIndice As Integer
Dim sProduto As String
Dim vbMsgRes As VbMsgBoxResult
Dim objItemOV As ClassItemOV
Dim iPossuiGrade As Integer
Dim objRomaneioGrade As ClassRomaneioGrade
Dim colItensRomaneioGrade As New Collection
Dim objItensRomaneio As ClassItemRomaneioGrade
Dim sProdutoPai As String
Dim objGridItens1 As Object

On Error GoTo Erro_Produto_Saida_Celula

'***********  FUNÇÃO ALTERADA PARA TRATAMENTO DE GRADE  ******************
    'Critica o Produto
    lErro = CF("Produto_Critica_Filial2", Produto.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 84176
    
    If lErro = 86295 And Len(Trim(objProduto.sGrade)) = 0 Then
        gError 86296
    ElseIf Len(Trim(objProduto.sGrade)) > 0 Then
        'Sinaliza que o produto possui grade
        iPossuiGrade = MARCADO
    End If
    
    'Se o produto não foi encontrado ==> Pergunta se deseja criar
    If lErro = 51381 Then gError 84177
        
    'Se não for um produto de grade
    If iPossuiGrade = DESMARCADO Then
        
        'Se existir um produto pai de grade no grid
        If Grid_Possui_Grade Then
            
            'Busca, caso exista, o produto pai de grade o prod em questão
            lErro = CF("Produto_Le_PaiGrade", objProduto, sProdutoPai)
            If lErro <> SUCESSO Then gError 86327
            
            'Se o produto tem um pai de grade
            If Len(Trim(sProdutoPai)) > 0 Then
                'Verifica se seu pai aparece no grid
                For iIndice = 1 To gobjOrcamentoVenda.colItens.Count
                    'Se aparecer ==> erro
                    If gobjOrcamentoVenda.colItens(iIndice).sProduto = sProdutoPai Then gError 86328
                
                Next
            
            End If
            
        End If
        
    Else
        'Verifica se há filhos válidos com a grade preenchida
        lErro = CF("Produto_Le_Filhos_Grade", objProduto, colItensRomaneioGrade)
        If lErro <> SUCESSO Then gError 86329
        
        'Se nao existir, erro
        If colItensRomaneioGrade.Count = 0 Then gError 86330
        
        'Para cada filho de grade do produto
        For Each objItensRomaneio In colItensRomaneioGrade
            'Verifica se ele já aparece no grid
            For iIndice = 1 To gobjOrcamentoVenda.colItens.Count
                'Se aparecer ==> Erro
                If gobjOrcamentoVenda.colItens(iIndice).sProduto = objItensRomaneio.sProduto Then gError 86331
            Next
        Next
 
    End If
    

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 84178

        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True

    End If

    'Verifica se já está em outra linha do Grid
    For iIndice = 1 To objGridItens.iLinhasExistentes
        If iIndice <> GridItens.Row Then
            If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text Then gError 84179 '26659
        End If
    Next

    Set objItemOV = New ClassItemOV
    
    objItemOV.iPossuiGrade = iPossuiGrade

    If objItemOV.iPossuiGrade = MARCADO Then
        
        objItemOV.sProduto = objProduto.sCodigo
        objItemOV.sUnidadeMed = objProduto.sSiglaUMVenda
        objItemOV.lCodOrcamento = StrParaDbl(Codigo.Text)
        objItemOV.iItem = GridItens.Row
        objItemOV.lNumIntDoc = 0
        objItemOV.sDescricao = objProduto.sDescricao
                
        Set objRomaneioGrade = New ClassRomaneioGrade
        
        objRomaneioGrade.sNomeTela = Me.Name
        
        Set objRomaneioGrade.objObjetoTela = objItemOV
                    
        Call Chama_Tela_Modal("RomaneioGrade", objRomaneioGrade)
        If giRetornoTela <> vbOK Then gError 86310
        

    End If

    'Unidade de Medida
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMVenda

    'Descricao Produto
    GridItens.TextMatrix(GridItens.Row, iGrid_DescProduto_Col) = objProduto.sDescricao

    'Preço Unitário
    If Len(Trim(TabelaPreco.Text)) > 0 Then

        'Coloca Produto no grid (necessario p/usar Trata_TabelaPreco_Item)
        GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text

        objTabelaPrecoItem.iCodTabela = Codigo_Extrai(TabelaPreco.Text)
        objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
        objTabelaPrecoItem.sCodProduto = objProduto.sCodigo

        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
        If lErro <> SUCESSO And lErro <> 28014 Then gError 84181 '26661

        If lErro = 28014 Then
            objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA

            lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
            If lErro <> SUCESSO And lErro <> 28014 Then gError 84182 '26662

        End If

        If lErro <> 28014 Then

            lErro = PrecoUnitario_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col), objTabelaPrecoItem, dPrecoUnitario)
            If lErro <> SUCESSO Then gError 84183 '26663

            If dPrecoUnitario > 0 Then
                GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
            Else
                GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col) = ""
            End If
            
            'precodesc
            Set objGridItens1 = GridItens
            Call CF("Produto_Saida_Celula_PrecoDesc", objGridItens1, GridItens.Row, iGrid_PrecoUnitario_Col + 1, GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col))

        End If

    End If

    'Acrescenta uma linha no Grid se for o caso
    If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1

        'permite que a tributacao reflita a inclusao de uma linha no grid
        lErro = Tributacao_Inclusao_Item_Grid(GridItens.Row, Produto.Text)
        If lErro <> SUCESSO Then gError 101105
    
        If iPossuiGrade = MARCADO Then
        
            '************** GRADE ************
            gobjOrcamentoVenda.colItens(GridItens.Row).iPossuiGrade = MARCADO

            gobjOrcamentoVenda.colItens(GridItens.Row).iItem = GridItens.Row
                       
            Set gobjOrcamentoVenda.colItens(GridItens.Row).colItensRomaneioGrade = objItemOV.colItensRomaneioGrade
            
            GridItens.TextMatrix(GridItens.Row, 0) = "# " & GridItens.TextMatrix(GridItens.Row, 0)
                   
            Call Atualiza_Grid_Itens(objItemOV)
                        
        End If
    
    End If

    Produto_Saida_Celula = SUCESSO

    Exit Function

Erro_Produto_Saida_Celula:

    Produto_Saida_Celula = gErr

    Select Case gErr

        Case 84178
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, Produto.Text)

        Case 84176, 84181, 84182, 84183, 84184, 84334, 101105, 86310

        Case 84177
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)
            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridItens)

                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridItens)
            End If

        Case 84179
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE", gErr, Produto.Text, Produto.Text, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177409)

    End Select

    Exit Function

End Function

Public Sub PrecoTotal_Calcula(iLinha As Integer)

Dim dPrecoTotal As Double
Dim dPrecoTotalReal As Double
Dim dPrecoUnitario As Double
Dim dQuantidade As Double
Dim dDesconto As Double
Dim dPercentDesc As Double
Dim lTamanho As Long
Dim dValorTotal As Double
Dim iIndice As Integer
Dim dValorDesconto As Double
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long
Dim objGridItens1 As Object

On Error GoTo Erro_PrecoTotal_Calcula

    'Quantidades e preço unitário
    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_PrecoUnitario_Col))
    dQuantidade = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))

    'Cálculo do desconto
    lTamanho = Len(Trim(GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col)))
    If lTamanho > 0 Then
        dPercentDesc = CDbl(Format(GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col), "General Number"))
    Else
        dPercentDesc = 0
    End If

    dPrecoTotal = dPrecoUnitario * (dQuantidade)

    'Se percentual for >0 tira o desconto
    If dPercentDesc > 0 Then dDesconto = dPercentDesc * dPrecoTotal
    dPrecoTotalReal = dPrecoTotal - dDesconto

    'precodesc
    Set objGridItens1 = GridItens
    Call CF("PrecoTotal_Calcula_PrecoDesc", objGridItens1, iLinha, iGrid_PrecoUnitario_Col + 1, Format(dPrecoUnitario * (1 - dPercentDesc), "Standard"))

    'Coloca valor do desconto no Grid
    If dDesconto > 0 Then
        GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = Format(dDesconto, "Standard")
    Else
        GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = ""
    End If

    'Coloca preco total do ítem no grid
    GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col) = Format(dPrecoTotalReal, "Standard")

    'Calcula a soma dos valores de produtos
    For iIndice = 1 To objGridItens.iLinhasExistentes
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
            If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
        End If
    Next

    If gdDesconto > 0 Then
        dValorDesconto = gdDesconto * dValorTotal
    ElseIf Len(Trim(ValorDesconto.Text)) > 0 And IsNumeric(ValorDesconto.Text) Then
        dValorDesconto = CDbl(ValorDesconto.Text)
    End If
    dValorTotal = dValorTotal - dValorDesconto

    'Verifica se o valor de desconto é maior que o valor dos produtos
    If dValorTotal < 0 And dValorDesconto > 0 Then

        vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_VALOR_DESCONTO_MAIOR_PRODUTOS", dValorDesconto, dValorTotal)

        gdDesconto = 0
        ValorDesconto.Text = ""
        giValorDescontoAlterado = 0
        dValorDesconto = 0

        'Para tributação
        gobjOrcamentoVenda.dValorDesconto = dValorDesconto

        'Faz a atualização dos valores da tributação
        lErro = AtualizarTributacao()
        If lErro <> SUCESSO Then gError 101104
        
        'Calcula a soma dos valores de produtos
        dValorTotal = 0
        For iIndice = 1 To objGridItens.iLinhasExistentes
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
                If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
            End If
        Next

    End If

    'Coloca valor total dos produtos na tela
    ValorProdutos.Caption = Format(dValorTotal, "Standard")
    ValorDesconto.Text = Format(dValorDesconto, "Standard")
    
    Call Tributacao_Alteracao_Item_Grid(iLinha)

    Exit Sub

Erro_PrecoTotal_Calcula:

    Select Case gErr

        Case 84195, 84247, 101104

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177410)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_DataVencimento(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Data Vencimento que está deixando de serr a corrente

Dim lErro As Long
Dim dtDataReferencia As Date
Dim dtDataVencimento As Date
Dim sDataVencimento As String
Dim iCriouLinha As Boolean

On Error GoTo Erro_Saida_Celula_DataVencimento

    Set objGridInt.objControle = DataVencimento

    'Verifica se Data de Vencimento esta preenchida
    If Len(Trim(DataVencimento.ClipText)) > 0 Then

        'Critica a data
        lErro = Data_Critica(DataVencimento.Text)
        If lErro <> SUCESSO Then gError 84196 '26726

         dtDataVencimento = CDate(DataVencimento.Text)

        'Se data de Emissao estiver preenchida verificar se a Data de Vencimento é maior que a Data de Emissão
        If Len(Trim(DataReferencia.ClipText)) > 0 Then
            dtDataReferencia = CDate(DataReferencia.Text)
            If dtDataVencimento < dtDataReferencia Then gError 84197 '26728
        End If

        sDataVencimento = Format(dtDataVencimento, "dd/mm/yyyy")

        iCriouLinha = False
        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            iCriouLinha = True
        End If

    End If

    If sDataVencimento <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_col) Then CobrancaAutomatica.Value = vbUnchecked

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84198 '26727

    If iCriouLinha Then
        'Coloca desconto padrao (le em CPRConfig)
        lErro = Preenche_DescontoPadrao(GridParcelas.Row)
        If lErro <> SUCESSO Then gError 84199 '51032
    End If

    Saida_Celula_DataVencimento = SUCESSO

    Exit Function

Erro_Saida_Celula_DataVencimento:

    Saida_Celula_DataVencimento = gErr

    Select Case gErr

        Case 84196, 84198, 84199
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 84197
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR_REFERENCIA", gErr, dtDataVencimento, GridParcelas.Row, dtDataReferencia)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177411)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorParcela(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Valor Parcela que está deixando de ser a corrente

Dim lErro As Long
Dim dColunaSoma As Double
Dim iIndice As Integer
Dim iColDescPerc As Integer
Dim iColTipoDesconto As Integer
Dim lTamanho As Long
Dim dPercentual As Double
Dim dValorParcela As Double
Dim sValorDesconto As String
Dim iTipoDesconto As Integer

On Error GoTo Erro_Saida_Celula_ValorParcela

    Set objGridInt.objControle = ValorParcela

    'Verifica se valor está preenchido
    If Len(ValorParcela.ClipText) > 0 Then

        'Critica se valor é positivo
        lErro = Valor_Positivo_Critica(ValorParcela.Text)
        If lErro <> SUCESSO Then gError 84200 '26724

        ValorParcela.Text = Format(ValorParcela.Text, "Standard")

        If ValorParcela.Text <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col) Then

            CobrancaAutomatica.Value = vbUnchecked

            '***Código para colocar valores de desconto
            dValorParcela = StrParaDbl(ValorParcela.Text)
            If dValorParcela > 0 Then

                'Vai varrer todos os 3 descontos para colocar valores
                For iIndice = 1 To 3

                    Select Case iIndice
                        Case 1
                            iColDescPerc = iGrid_Desc1Perc_Col
                            iColTipoDesconto = iGrid_Desc1Codigo_Col
                        Case 2
                            iColDescPerc = iGrid_Desc2Perc_Col
                            iColTipoDesconto = iGrid_Desc2Codigo_Col
                        Case 3
                            iColDescPerc = iGrid_Desc3Perc_Col
                            iColTipoDesconto = iGrid_Desc3Codigo_Col
                    End Select

                    iTipoDesconto = Codigo_Extrai(GridParcelas.TextMatrix(GridParcelas.Row, iColTipoDesconto))
                    lTamanho = Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iColDescPerc)))

                    'Coloca valor de desconto na tela
                    If (iTipoDesconto = Percentual Or iTipoDesconto = PERC_ANT_DIA Or iTipoDesconto = PERC_ANT_DIA_UTIL) And lTamanho > 0 Then
                        dPercentual = PercentParaDbl(GridParcelas.TextMatrix(GridParcelas.Row, iColDescPerc))
                        sValorDesconto = Format(dPercentual * dValorParcela, "Standard")
                        GridParcelas.TextMatrix(GridParcelas.Row, iColDescPerc - 1) = sValorDesconto
                    End If

                Next

            End If
            '***Fim Código para colocar valores de desconto

        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            'Coloca DescontoPadrao
            lErro = Preenche_DescontoPadrao(GridParcelas.Row)
            If lErro <> SUCESSO Then gError 84201  '51061

        End If

    Else

        '***Código para colocar valores de desconto
        'Limpa Valores de Desconto
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desc1Valor_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desc2Valor_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desc3Valor_Col) = ""
        '***Fim Código para colocar valores de desconto

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84202 '26725

    Saida_Celula_ValorParcela = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorParcela:

    Saida_Celula_ValorParcela = gErr

    Select Case gErr

        Case 84200, 84201, 84202
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177412)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TipoDesconto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Tipo Desconto que está deixando de serr a corrente

Dim lErro As Long
Dim iCodigo As Integer
Dim iTipo As Integer
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_TipoDesconto

    If GridParcelas.Col = iGrid_Desc1Codigo_Col Then
        Set objGridInt.objControle = TipoDesconto1
    ElseIf GridParcelas.Col = iGrid_Desc2Codigo_Col Then
        Set objGridInt.objControle = TipoDesconto2
    ElseIf GridParcelas.Col = iGrid_Desc3Codigo_Col Then
        Set objGridInt.objControle = TipoDesconto3
    End If

    'Verifica se o Tipo foi preenchido
    If Len(Trim(objGridInt.objControle.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If objGridInt.objControle.Text <> objGridInt.objControle.List(objGridInt.objControle.ListIndex) Then

            'Tenta selecioná-lo na combo
            lErro = Combo_Seleciona_Grid(objGridInt.objControle, iCodigo)
            If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then gError 84203 '26729

            'Não foi encontrado
            If lErro = 25085 Then gError 84204 '26730
            If lErro = 25086 Then gError 84205 '26731

        End If

        'Extrai o tipo de Desconto
        iTipo = Codigo_Extrai(objGridInt.objControle.Text)

        If (iTipo = VALOR_ANT_DIA) Or (iTipo = VALOR_ANT_DIA_UTIL) Or (iTipo = VALOR_FIXO) Then
            GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col + 3) = ""
        ElseIf iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = Percentual Then
            '*** Acrescentado + 1 If para contabilizar com colocação de valores de desconto
            If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col + 3))) = 0 Then
                GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col + 2) = ""
            End If
        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    Else

        For iIndice = objGridInt.objGrid.Col To iGrid_Desc3Perc_Col
            GridParcelas.TextMatrix(GridParcelas.Row, iIndice) = ""
        Next

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84206 '26732

    Saida_Celula_TipoDesconto = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoDesconto:

    Saida_Celula_TipoDesconto = gErr

    Select Case gErr

        Case 84203, 84206
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 84204
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 84205
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO1", gErr, objGridInt.objControle.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177413)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescontoData(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Desconto Data que está deixando de ser a corrente

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_DescontoData

    If GridParcelas.Col = iGrid_Desc1Ate_Col Then
        Set objGridInt.objControle = Desconto1Ate
    ElseIf GridParcelas.Col = iGrid_Desc2Ate_Col Then
        Set objGridInt.objControle = Desconto2Ate
    ElseIf GridParcelas.Col = iGrid_Desc3Ate_Col Then
        Set objGridInt.objControle = Desconto3Ate
    End If

    If Len(Trim(objGridInt.objControle.ClipText)) > 0 Then

        lErro = Data_Critica(objGridInt.objControle.Text)
        If lErro <> SUCESSO Then gError 84207 '26792
        'Se a data de vencimento estiver preenchida
        If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_col))) > 0 Then
            'critica se DataDesconto ultrapassa DataVencimento
            If CDate(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_col)) < CDate(objGridInt.objControle.Text) Then gError 84208 '26592
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84209 '26793

    Saida_Celula_DescontoData = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoData:

    Saida_Celula_DescontoData = gErr

    Select Case gErr

        Case 84207, 84209
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 84208
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADESCONTO_MAIOR_DATAVENCIMENTO", gErr, CDate(objGridInt.objControle.Text), CDate(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_col)))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177414)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescontoValor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Desconto Valor que está deixando de serr a corrente

Dim lErro As Long
Dim dColunaSoma As Double

On Error GoTo Erro_Saida_Celula_DescontoValor

    If GridParcelas.Col = iGrid_Desc1Valor_Col Then
        Set objGridInt.objControle = Desconto1Valor
    ElseIf GridParcelas.Col = iGrid_Desc2Valor_Col Then
        Set objGridInt.objControle = Desconto2Valor
    ElseIf GridParcelas.Col = iGrid_Desc3Valor_Col Then
        Set objGridInt.objControle = Desconto3Valor
    End If

    'Verifica se valor está preenchido
    If Len(objGridInt.objControle.ClipText) > 0 Then
        'Critica se valor é positivo
        lErro = Valor_Positivo_Critica(objGridInt.objControle.Text)
        If lErro <> SUCESSO Then gError 84210 '26733

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84211 '26734

    Saida_Celula_DescontoValor = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoValor:

    Saida_Celula_DescontoValor = gErr

    Select Case gErr

        Case 84210, 84211
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177415)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescontoPerc(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Desconto Percentual que está deixando de ser a corrente

Dim lErro As Long
Dim iCodigo As Integer
Dim dPercentual As Double
Dim dValorParcela As Double
Dim sValorDesconto As String

On Error GoTo Erro_Saida_Celula_DescontoPerc

    If GridParcelas.Col = iGrid_Desc1Perc_Col Then
        Set objGridInt.objControle = Desconto1Percentual
    ElseIf GridParcelas.Col = iGrid_Desc2Perc_Col Then
        Set objGridInt.objControle = Desconto2Percentual
    ElseIf GridParcelas.Col = iGrid_Desc3Perc_Col Then
        Set objGridInt.objControle = Desconto3Percentual
    End If

    If Len(Trim(objGridInt.objControle.Text)) > 0 Then

        'Critica porcentagem
        lErro = Porcentagem_Critica(objGridInt.objControle.Text)
        If lErro <> SUCESSO Then gError 84212 '26794

        '***Código para colocar valores de desconto
        dPercentual = CDbl(objGridInt.objControle.Text) / 100
        dValorParcela = StrParaDbl(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col))

        'Coloca Valor do Desconto na tela
        If dValorParcela > 0 Then
            sValorDesconto = Format(dPercentual * dValorParcela, "Standard")
            GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col - 1) = sValorDesconto
        End If

    Else

        'Limpa Valor de Desconto
        GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col - 1) = ""
        '***Fim Código para colocar valores de desconto

    End If


    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84213 '26795

    Saida_Celula_DescontoPerc = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoPerc:

    Saida_Celula_DescontoPerc = gErr

    Select Case gErr

        Case 84212, 84213
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177416)

    End Select

    Exit Function

End Function

Private Function Valida_Grid_Itens() As Long

Dim iIndice As Integer
Dim lErro As Long
Dim dQuantidade As Double

On Error GoTo Erro_Valida_Grid_Itens

    'Verifica se há itens no grid
    If objGridItens.iLinhasExistentes = 0 Then gError 84214 '26813

    'para cada item do grid
    For iIndice = 1 To objGridItens.iLinhasExistentes

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 84215 '51455

        lErro = Valor_Positivo_Critica(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        If lErro <> SUCESSO Then gError 84216 '26814

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col))) = 0 Then gError 84217 '51456

        lErro = Valor_Positivo_Critica(GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col))
        If lErro <> SUCESSO Then gError 84218 '26815

    Next

    Valida_Grid_Itens = SUCESSO

    Exit Function

Erro_Valida_Grid_Itens:

    Valida_Grid_Itens = gErr

    Select Case gErr

        Case 84214
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_OV", gErr)

        Case 84215
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_ITEM_NAO_PREENCHIDA", gErr, iIndice)

        Case 84216, 84218

        Case 84217
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORUNITARIO_ITEM_NAO_PREENCHIDO", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177417)

    End Select

    Exit Function

End Function


Private Function Valida_Grid_Parcelas() As Long
'Valida os dados do Grid de Parcelas

Dim lErro As Long
Dim iIndice As Integer
Dim dSomaParcelas As Double
Dim dValorIRRF As Double, dPISRetido As Double, dCOFINSRetido As Double, dCSLLRetido As Double
Dim dValorTotal As Double
Dim dtDataEmissao As Date
Dim dtDataVencimento As Date
Dim iTamanho As Integer
Dim iTipo As Integer
Dim dPercAcrecFin As Double
Dim iDesconto As Integer
Dim dtDataDesconto As Date

On Error GoTo Erro_Valida_Grid_Parcelas

    'Verifica se alguma parcela foi informada
    If objGridParcelas.iLinhasExistentes = 0 Then gError 84219 '26817

    dSomaParcelas = 0

    'Para cada Parcela do grid de parcelas
    For iIndice = 1 To objGridParcelas.iLinhasExistentes

        dtDataEmissao = StrParaDate(DataEmissao.Text)
        dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col))

        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col))) = 0 Then gError 84220 '26818
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))) = 0 Then gError 84221 '26821

        'Se o tipo de desconto 1 estiver preenchido
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))) > 0 Then
            iDesconto = 1
            iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))
            'Verifica se a data de desconto está preenchdida
            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))) = 0 Then gError 84223 '51066
            'Recolhe o Valor ou Percentual de desconto
            If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col))) = 0 Then gError 84224 '51069
            Else
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Perc_Col))) = 0 Then gError 84225 '51070
            End If
            'Se o tipo de desconto 2 estiver preenchido
            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))) > 0 Then
                iDesconto = 2
                iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))
                'Verifica se a data de desconto está preenchdida
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))) = 0 Then gError 84226 '51067
                'Faz a crítica da ordem das datas de desconto
                If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) < StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col)) Then gError 84227 '51075
                If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col)) Then gError 84228 '51077
                'Recolhe o Valor ou Percentual de desconto
                If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col))) = 0 Then gError 84229 '51071
                Else
                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Perc_Col))) = 0 Then gError 84230 '51072
                End If
                'Se o tipo de desconto 3 estiver preenchido
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))) > 0 Then
                    iDesconto = 3
                    iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))
                    'Verifica se a data de desconto está preenchdida
                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))) = 0 Then gError 84231 '51068
                    'Faz a crítica da ordem das datas de desconto
                    If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col)) < StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) Then gError 84232 '51076
                    If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col)) = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) Then gError 84233 '51078
                    'Recolhe o Valor ou Percentual de desconto
                    If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
                        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col))) = 0 Then gError 84234 '51073
                    Else
                        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Perc_Col))) = 0 Then gError 84235 '51074
                    End If
                    dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))
                    If dtDataDesconto > dtDataVencimento Then gError 84236 '51364
                Else
                    dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))
                    If dtDataDesconto > dtDataVencimento Then gError 84237 '51363
                End If
            Else
                dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))
                If dtDataDesconto > dtDataVencimento Then gError 84238 '51362
            End If
        End If



        If iIndice > 1 Then If CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col)) < CDate(GridParcelas.TextMatrix(iIndice - 1, iGrid_Vencimento_col)) Then gError 84239 '26820

        dSomaParcelas = dSomaParcelas + CDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))

    Next

    dValorTotal = StrParaDbl(ValorTotal.Caption)
    dValorIRRF = StrParaDbl(ValorIRRF.Text)
    If Len(Trim(PISRetido.Text)) <> 0 And IsNumeric(PISRetido.Text) Then dPISRetido = CDbl(PISRetido.Text)
    If Len(Trim(COFINSRetido.Text)) <> 0 And IsNumeric(COFINSRetido.Text) Then dCOFINSRetido = CDbl(COFINSRetido.Text)
    If Len(Trim(CSLLRetido.Text)) <> 0 And IsNumeric(CSLLRetido.Text) Then dCSLLRetido = CDbl(CSLLRetido.Text)
    
    If Format((dValorTotal - (dValorIRRF + dPISRetido + dCOFINSRetido + dCSLLRetido)), "Standard") <> Format(dSomaParcelas, "Standard") Then gError 26822

    Valida_Grid_Parcelas = SUCESSO

    Exit Function

Erro_Valida_Grid_Parcelas:

    Valida_Grid_Parcelas = gErr

    Select Case gErr

        Case 84219
            Call Rotina_Erro(vbOKOnly, "ERRO_FALTA_PARCELA_COBRANCA", gErr)

        Case 84220
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_COBRANCA_NAO_INFORMADA", gErr, iIndice)

        Case 84222
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_COBRANCA_MENOR", gErr, iIndice, dtDataVencimento, dtDataEmissao)
        Case 84221
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_PARCELA_COBRANCA_NAO_INFORMADO", gErr, iIndice)

        Case 84223, 84226, 84231
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_PARCELA_NAO_PREENCHIDA", gErr, iDesconto, iIndice)

        Case 84224, 84225, 84229, 84230, 84234, 84235
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_PARCELA_NAO_PREENCHIDO", gErr, iDesconto, iIndice)

        Case 84227, 84232
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_DESCONTOS_DESORDENADAS", gErr, iIndice)

        Case 84228, 84233
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_DESCONTO_IGUAIS", gErr, iDesconto - 1, iDesconto, iIndice)

        Case 84238, 84237, 84236
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_PARCELA_SUPERIOR_DATA_VENCIMENTO", gErr, dtDataDesconto, iDesconto, iIndice)

        Case 26822
            Call Rotina_Erro(vbOKOnly, "ERRO_SOMA_PARCELAS_COBRANCA_INVALIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177418)

    End Select

    Exit Function

End Function

Private Function Move_OrcamentoVenda_Memoria(objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Move os dados da tela para objOrcamentoVenda

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim dValorTotalParcelas As Double
Dim dValorIRRF As Double
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_OrcamentoVenda_Memoria

    If Len(Trim(Codigo.Text)) > 0 Then objOrcamentoVenda.lCodigo = CLng(Codigo.Text)

    'Verifica se o Cliente foi preenchido
    If Len(Trim(Cliente.ClipText)) > 0 Then

        objCliente.sNomeReduzido = Cliente.Text

        'Lê o Cliente através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 84239 '26779

        If lErro = SUCESSO Then
            'Guarda código do Cliente em objOrcamentoVenda
            objOrcamentoVenda.lCliente = objCliente.lCodigo
        End If
            
        objOrcamentoVenda.sNomeCli = objCliente.sNomeReduzido
            
    End If
    
    'Verifica se vendedor existe
    If Len(Trim(Vendedor.Text)) > 0 Then
        
        objVendedor.sNomeReduzido = Trim(Vendedor.Text)

        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
        If lErro <> SUCESSO And lErro <> 25008 Then gError 94418

        'Não encontrou o vendedor ==> erro
        If lErro = 25008 Then gError 94419

        objOrcamentoVenda.iVendedor = objVendedor.iCodigo

    End If
    
    'Verifica se a Filial está preenchida
    If Len(Trim(Filial.Text)) > 0 Then
        
        'Se o Cliente estiver cadastrado
        If objOrcamentoVenda.lCliente <> 0 Then
        
            'a filial tb deverá estar cadastrada e por isso teremos o código da filial na tela
            objOrcamentoVenda.iFilial = Codigo_Extrai(Filial.Text)
            objOrcamentoVenda.sNomeFilialCli = Nome_Extrai(Filial.Text)
                                                
        Else
            'se não, guardaremos o Texto digitado pelo usuário
            objOrcamentoVenda.sNomeFilialCli = Trim(Filial.Text)
            
        End If
            
    End If

    'Preenche objOrcamentoVenda com dados da tela
    objOrcamentoVenda.dtDataEmissao = MaskedParaDate(DataEmissao)
    objOrcamentoVenda.iTabelaPreco = Codigo_Extrai(TabelaPreco.Text)
    objOrcamentoVenda.sNaturezaOp = Trim(NaturezaOp.Text)
    objOrcamentoVenda.dValorFrete = StrParaDbl(ValorFrete.Text)
    objOrcamentoVenda.dValorSeguro = StrParaDbl(ValorSeguro.Text)
    objOrcamentoVenda.dValorDesconto = StrParaDbl(ValorDesconto.Text)
    objOrcamentoVenda.dValorOutrasDespesas = StrParaDbl(ValorDespesas.Text)
    objOrcamentoVenda.dValorProdutos = StrParaDbl(ValorProdutos.Caption)
    objOrcamentoVenda.dValorTotal = StrParaDbl(ValorTotal.Caption)
    objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa
    objOrcamentoVenda.iPrazoValidade = StrParaInt(PrazoValidade.Text)
    objOrcamentoVenda.dPercAcrescFinanceiro = StrParaDbl(PercAcrescFin.ClipText)
    objOrcamentoVenda.dtDataReferencia = MaskedParaDate(DataReferencia)
    objOrcamentoVenda.dValorOutrasDespesas = StrParaDbl(ValorDespesas.ClipText)
    objOrcamentoVenda.iCobrancaAutomatica = StrParaInt(CobrancaAutomatica.Value)
    objOrcamentoVenda.iCondicaoPagto = Codigo_Extrai(CondicaoPagamento.Text)
    
    'Move Grid Itens para memória
    lErro = Move_GridItens_Memoria(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 84241 '26781

    'Move Tab Cobrança para memória
    Call Move_TabCobranca_Memoria(objOrcamentoVenda)

    objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa

    'Move Tributacao para objOrcamentoVenda
    Set objOrcamentoVenda.objTributacaoOV = gobjOrcamentoVenda.objTributacaoOV

    Move_OrcamentoVenda_Memoria = SUCESSO

    Exit Function

Erro_OrcamentoVenda_Memoria:

    Move_OrcamentoVenda_Memoria = gErr

    Select Case gErr

        Case 84239, 84241

        Case 84240
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)
        'Por Leo em 21/03/02
        Case 94418
        
        Case 94419
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", gErr, objVendedor.sNomeReduzido)
        'Leo até aqui
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177419)

    End Select

    Exit Function

End Function

Private Sub Move_TabCobranca_Memoria(objOrcamentoVenda As ClassOrcamentoVenda)
'Recolhe os dados do tab de cobrança

Dim lTamanho As Long
Dim iIndice As Integer
Dim objParcelaOV As ClassParcelaOV

    'Recolhe os dados da Cobrança
    objOrcamentoVenda.iCobrancaAutomatica = CobrancaAutomatica.Value
    objOrcamentoVenda.dtDataReferencia = MaskedParaDate(DataReferencia)
    objOrcamentoVenda.iCondicaoPagto = Codigo_Extrai(CondicaoPagamento.Text)
    objOrcamentoVenda.dPercAcrescFinanceiro = StrParaDbl(PercAcrescFin.Text) / 100

    If objGridParcelas.iLinhasExistentes = 0 Then Exit Sub

    'Recolhe os Dados do Grid de Parcelas
    For iIndice = 1 To objGridParcelas.iLinhasExistentes

        Set objParcelaOV = New ClassParcelaOV

        objParcelaOV.iNumParcela = iIndice

        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col))) > 0 Then
            objParcelaOV.dtDataVencimento = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col))
        Else
            objParcelaOV.dtDataVencimento = DATA_NULA
        End If

        objParcelaOV.dValor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))
        objParcelaOV.iDesconto1Codigo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))
        objParcelaOV.iDesconto2Codigo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))
        objParcelaOV.iDesconto3Codigo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))) > 0 Then
            objParcelaOV.dtDesconto1Ate = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))
        Else
            objParcelaOV.dtDesconto1Ate = DATA_NULA
        End If
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))) > 0 Then
            objParcelaOV.dtDesconto2Ate = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))
        Else
            objParcelaOV.dtDesconto2Ate = DATA_NULA
        End If
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))) > 0 Then
            objParcelaOV.dtDesconto3Ate = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))
        Else
            objParcelaOV.dtDesconto3Ate = DATA_NULA
        End If

        If objParcelaOV.iDesconto1Codigo = VALOR_FIXO Or objParcelaOV.iDesconto1Codigo = VALOR_ANT_DIA Or objParcelaOV.iDesconto1Codigo = VALOR_ANT_DIA_UTIL Then
            objParcelaOV.dDesconto1Valor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col))
        ElseIf objParcelaOV.iDesconto1Codigo = Percentual Or objParcelaOV.iDesconto1Codigo = PERC_ANT_DIA Or objParcelaOV.iDesconto1Codigo = PERC_ANT_DIA_UTIL Then
            lTamanho = Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Perc_Col)))
            If lTamanho > 0 Then objParcelaOV.dDesconto1Valor = PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Perc_Col))
        End If

        If objParcelaOV.iDesconto2Codigo = VALOR_FIXO Or objParcelaOV.iDesconto2Codigo = VALOR_ANT_DIA Or objParcelaOV.iDesconto2Codigo = VALOR_ANT_DIA_UTIL Then
            objParcelaOV.dDesconto2Valor = CDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col))
        ElseIf objParcelaOV.iDesconto2Codigo = Percentual Or objParcelaOV.iDesconto2Codigo = PERC_ANT_DIA Or objParcelaOV.iDesconto2Codigo = PERC_ANT_DIA_UTIL Then
            lTamanho = Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Perc_Col)))
            If lTamanho > 0 Then objParcelaOV.dDesconto2Valor = PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Perc_Col))
        End If

        If objParcelaOV.iDesconto3Codigo = VALOR_FIXO Or objParcelaOV.iDesconto3Codigo = VALOR_ANT_DIA Or objParcelaOV.iDesconto3Codigo = VALOR_ANT_DIA_UTIL Then
            objParcelaOV.dDesconto3Valor = CDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col))
        ElseIf objParcelaOV.iDesconto3Codigo = Percentual Or objParcelaOV.iDesconto3Codigo = PERC_ANT_DIA Or objParcelaOV.iDesconto3Codigo = PERC_ANT_DIA_UTIL Then
            lTamanho = Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Perc_Col)))
            If lTamanho > 0 Then objParcelaOV.dDesconto3Valor = PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Perc_Col))
        End If

        objOrcamentoVenda.colParcela.Add objParcelaOV

    Next

End Sub

Public Sub NaturezaOpItem_Validate(Cancel As Boolean)

Dim sNatOp As String
Dim lErro As Long
Dim objNaturezaOp As New ClassNaturezaOp
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_NaturezaOpItem_Validate

    If giNatOpItemAlterado = 0 Then Exit Sub

    sNatOp = Trim(NaturezaOpItem.Text)

    If sNatOp <> "" Then

        objNaturezaOp.sCodigo = sNatOp
                
        If objNaturezaOp.sCodigo < NATUREZA_SAIDA_COD_INICIAL Or objNaturezaOp.sCodigo > NATUREZA_SAIDA_COD_FINAL Then gError 103003 '56910
        
        lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
        If lErro <> SUCESSO And lErro <> 17958 Then gError 103004 '43168
        
        'Se não achou a Natureza de Operação --> erro
        If lErro <> SUCESSO Then gError 103005  ' 43169

        LabelDescrNatOpItem.Caption = objNaturezaOp.sDescricao
        
        Call BotaoGravarTribItem_Click
    
    Else
        
        'Limpa a descrição
        LabelDescrNatOpItem.Caption = ""
    
    End If

    giNatOpItemAlterado = 0

    Exit Sub

Erro_NaturezaOpItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 103003
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_SAIDA", gErr)
        
        Case 103004
     
        Case 103005
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_NATUREZA_OPERACAO", NaturezaOpItem.Text)
            If vbMsgRes = vbYes Then
                Call Chama_Tela("NaturezaOperacao", objNaturezaOp)
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177420)

    End Select

End Sub

Public Sub NaturezaOpItem_GotFocus()

Dim iNaturezaOpAux As Integer

    iNaturezaOpAux = giNatOpItemAlterado
    
    Call MaskEdBox_TrataGotFocus(NaturezaOpItem, iAlterado)
    
    giNatOpItemAlterado = iNaturezaOpAux

End Sub

Public Sub TributacaoRecalcular_Click()

Dim lErro As Long

On Error GoTo Erro_TributacaoRecalcular_Click

    giRecalculandoTributacao = 1

    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 103002

    giRecalculandoTributacao = 0
    
    Exit Sub
    
Erro_TributacaoRecalcular_Click:

    Select Case gErr
    
        Case 103002
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177421)
            
    End Select

    Exit Sub

End Sub

Public Sub ValorIRRF_Change()

    iAlterado = REGISTRO_ALTERADO
    giValorIRRFAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ISSIncluso_Click()

Dim lErro As Long

On Error GoTo Erro_ISSIncluso_Click

    iAlterado = REGISTRO_ALTERADO

    Call BotaoGravarTrib

    Exit Sub

Erro_ISSIncluso_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177422)

    End Select

    Exit Sub

End Sub

Function Traz_OrcamentoVenda_Tela(objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Coloca na tela os dados do Orcamento de Venda

Dim lErro As Long
Dim objFilial As AdmFiliais
Dim bCancel As Boolean

On Error GoTo Erro_Traz_OrcamentoVenda_Tela

    gbCarregandoTela = True

    Call Limpa_OrcamentoVenda2

    'Lê os dados do Orcamento de Venda
    lErro = CF("OrcamentoVenda_Le", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 101232 Then gError 84359
    If lErro = 101232 Then gError 84363

    'Lê a parte de Tributação
    lErro = CF("OrcamentoVenda_Le_Tributacao", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 101162 Then gError 84360
    If lErro = 101162 Then gError 101282

    lErro = CF("OrcamentoVenda_Le_Itens_ComTributacao", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 101278 Then gError 84469
    If lErro = 101278 Then gError 101280
    
    lErro = CF("ParcelasOV_Le", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 101284 Then gError 84470
    If lErro = 101284 Then gError 101286
    
    lErro = TributacaoOV_Reset(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 101100
  
    ValorTotal.Caption = Format(objOrcamentoVenda.dValorTotal, "Standard")
    ValorProdutos.Caption = Format(objOrcamentoVenda.dValorProdutos, "Standard")

    PercAcrescFin.Text = ""
    'Coloca os dados do Orcamento na tela

    'Se existe um código para o Cliente
    If objOrcamentoVenda.lCliente <> 0 Then 'Por Leo em 18/04/02
    
        Call Cliente_Formata(objOrcamentoVenda.lCliente)
        Call Filial_Formata(Filial, objOrcamentoVenda.iFilial)
                
    Else 'Trecho por Leo em 18/04/02 ***
        
        'Preenche o Cliente e a Filial com os Nomes Informados.
        Cliente.Text = objOrcamentoVenda.sNomeCli
        Filial.Text = objOrcamentoVenda.sNomeFilialCli
                
    End If 'Leo até aqui ***
    
    giFilialAlterada = 0

    Codigo.Text = objOrcamentoVenda.lCodigo
    NatOpEspelho.Caption = objOrcamentoVenda.sNaturezaOp
    
    ValorFrete.Text = Format(objOrcamentoVenda.dValorFrete, "Standard")
    ValorSeguro.Text = Format(objOrcamentoVenda.dValorSeguro, "Standard")
    ValorDesconto.Text = Format(objOrcamentoVenda.dValorDesconto, "Standard")
    ValorDespesas.Text = Format(objOrcamentoVenda.dValorOutrasDespesas, "Standard")

    NaturezaOp.Text = objOrcamentoVenda.sNaturezaOp 'Por Leo em 30/04/02
    Call NaturezaOp_Validate(bSGECancelDummy)

    giValorFreteAlterado = 0
    giValorSeguroAlterado = 0
    giValorDescontoAlterado = 0
    giValorDespesasAlterado = 0

    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(objOrcamentoVenda.dtDataEmissao, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    
    If objOrcamentoVenda.iPrazoValidade <> 0 Then 'Incluido por Leo em 29/04/02
        PrazoValidade.Text = objOrcamentoVenda.iPrazoValidade
    End If
    
    'Se a tabela de preços estiver preenchida coloca na tela
    If objOrcamentoVenda.iTabelaPreco > 0 Then
        TabelaPreco.Text = objOrcamentoVenda.iTabelaPreco
        Call TabelaPreco_Validate(bSGECancelDummy)
    Else
        TabelaPreco.Text = ""
    End If
    
    'Preenche o campo de vendedores
    If objOrcamentoVenda.iVendedor <> 0 Then
        
        Vendedor.Text = objOrcamentoVenda.iVendedor
        Call Vendedor_Validate(bCancel)
    
    End If
    
    'Carrega o Tab Cobrança
    lErro = Carrega_Tab_Cobranca(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 84361 '51164

    'Carrega o Grid de itens
    lErro = Carrega_Grid_Itens(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 84362 '26569
    
    ValorTotal.Caption = Format(objOrcamentoVenda.dValorTotal, "Standard")
    
    'Carrega o Tab de Tributação
    lErro = Carrega_Tab_Tributacao(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 27640
    
    iAlterado = 0

    gbCarregandoTela = False

    Traz_OrcamentoVenda_Tela = SUCESSO

    Exit Function

Erro_Traz_OrcamentoVenda_Tela:

    gbCarregandoTela = False

    Traz_OrcamentoVenda_Tela = gErr

    Select Case gErr

        Case 84359 To 84363, 84469, 84470, 101100, 101101, 27640, 101282, 101286
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177423)

    End Select

    Exit Function

End Function

Public Sub Form_Activate()

    Dim lErro As Long

On Error GoTo Erro_Form_Activate

'???? Trocar numeracao de erro
    lErro = CargaPosFormLoad
    If lErro <> SUCESSO Then gError 59332
        
    Call TelaIndice_Preenche(Me)

    Exit Sub
     
Erro_Form_Activate:

    Select Case gErr
          
        Case 59332
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177424)
     
    End Select
     
    Exit Sub

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""
'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "OrcamentoVenda"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_OrcamentoVenda_Memoria(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 84404 '51159

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objOrcamentoVenda.lCodigo, 0, "Codigo"
    colCampoValor.Add "Cliente", objOrcamentoVenda.lCliente, 0, "Cliente"
    colCampoValor.Add "Filial", objOrcamentoVenda.iFilial, 0, "Filial"
    colCampoValor.Add "CondicaoPagto", objOrcamentoVenda.iCondicaoPagto, 0, "CondicaoPagto"
    colCampoValor.Add "PercAcrescFinanceiro", objOrcamentoVenda.dPercAcrescFinanceiro, 0, "PercAcrescFinanceiro"
    colCampoValor.Add "DataEmissao", objOrcamentoVenda.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "ValorTotal", objOrcamentoVenda.dValorTotal, 0, "ValorTotal"
    colCampoValor.Add "ValorFrete", objOrcamentoVenda.dValorFrete, 0, "ValorFrete"
    colCampoValor.Add "ValorDesconto", objOrcamentoVenda.dValorDesconto, 0, "ValorDesconto"
    colCampoValor.Add "ValorSeguro", objOrcamentoVenda.dValorSeguro, 0, "ValorSeguro"
    colCampoValor.Add "TabelaPreco", objOrcamentoVenda.iTabelaPreco, 0, "TabelaPreco"
    colCampoValor.Add "Vendedor", objOrcamentoVenda.iVendedor, 0, "Vendedor"
    colCampoValor.Add "NomeCli", objOrcamentoVenda.sNomeCli, STRING_CLIENTE_NOME_REDUZIDO, "NomeCli"
    colCampoValor.Add "NomeFilialCli", objOrcamentoVenda.sNomeFilialCli, STRING_FILIAL_CLIENTE_NOME, "NomeFilialCli"
    colCampoValor.Add "NaturezaOp", objOrcamentoVenda.sNaturezaOp, STRING_NATUREZAOP_CODIGO, "NaturezaOp" 'por Leo em 02/05/02
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 84404

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177425)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda

On Error GoTo Erro_Tela_Preenche

    objOrcamentoVenda.lCodigo = colCampoValor.Item("Codigo").vValor
    objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa

    If objOrcamentoVenda.lCodigo <> 0 Then

        'Mostra os dados do Pedido de Venda na tela
        lErro = Traz_OrcamentoVenda_Tela(objOrcamentoVenda)
        If lErro <> SUCESSO Then gError 84405

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 84405

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177426)

    End Select

    Exit Sub

End Sub

Function Carrega_Tab_Cobranca(objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Coloca os dados do tab de cobrança na tela

Dim objParcelaOV As ClassParcelaOV
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim dValorDesconto As Double

    Call DateParaMasked(DataReferencia, objOrcamentoVenda.dtDataReferencia)
    giDataReferenciaAlterada = 0

    PercAcrescFin.Text = ""

    If objOrcamentoVenda.iCondicaoPagto > 0 Then
        CondicaoPagamento.Text = objOrcamentoVenda.iCondicaoPagto
        Call CondicaoPagamento_Validate(bSGECancelDummy)
    Else
        CondicaoPagamento.Text = ""
    End If

    PercAcrescFin.Text = Format(objOrcamentoVenda.dPercAcrescFinanceiro * 100, "Fixed")

    CobrancaAutomatica.Value = objOrcamentoVenda.iCobrancaAutomatica

    'Limpa o Grid de Parcelas antes de preencher com dados da coleção
    Call Grid_Limpa(objGridParcelas)

    iIndice = 0

    For Each objParcelaOV In objOrcamentoVenda.colParcela

        iIndice = iIndice + 1
        GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col) = Format(objParcelaOV.dtDataVencimento, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col) = Format(objParcelaOV.dValor, "Standard")
        If objParcelaOV.dtDesconto1Ate <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col) = Format(objParcelaOV.dtDesconto1Ate, "dd/mm/yyyy")
        If objParcelaOV.dtDesconto2Ate <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col) = Format(objParcelaOV.dtDesconto2Ate, "dd/mm/yyyy")
        If objParcelaOV.dtDesconto3Ate <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col) = Format(objParcelaOV.dtDesconto3Ate, "dd/mm/yyyy")
        If objParcelaOV.iDesconto1Codigo = VALOR_FIXO Or objParcelaOV.iDesconto1Codigo = VALOR_ANT_DIA Or objParcelaOV.iDesconto1Codigo = VALOR_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col) = Format(objParcelaOV.dDesconto1Valor, "Standard")
        ElseIf objParcelaOV.iDesconto1Codigo = Percentual Or objParcelaOV.iDesconto1Codigo = PERC_ANT_DIA Or objParcelaOV.iDesconto1Codigo = PERC_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc1Perc_Col) = Format(objParcelaOV.dDesconto1Valor, "Percent")
            '*** Inicio código p/ colocar Valor Desconto
            If objParcelaOV.dValor > 0 Then
                dValorDesconto = objParcelaOV.dDesconto1Valor * objParcelaOV.dValor
                GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col) = Format(dValorDesconto, "Standard")
            End If
            '*** Fim
        End If
        If objParcelaOV.iDesconto2Codigo = VALOR_FIXO Or objParcelaOV.iDesconto2Codigo = VALOR_ANT_DIA Or objParcelaOV.iDesconto2Codigo = VALOR_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col) = Format(objParcelaOV.dDesconto2Valor, "Standard")
        ElseIf objParcelaOV.iDesconto2Codigo = Percentual Or objParcelaOV.iDesconto2Codigo = PERC_ANT_DIA Or objParcelaOV.iDesconto2Codigo = PERC_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc2Perc_Col) = Format(objParcelaOV.dDesconto2Valor, "Percent")
            '*** Inicio código p/ colocar Valor Desconto
            If objParcelaOV.dValor > 0 Then
                dValorDesconto = objParcelaOV.dDesconto2Valor * objParcelaOV.dValor
                GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col) = Format(dValorDesconto, "Standard")
            End If
            '*** Fim
        End If
        If objParcelaOV.iDesconto3Codigo = VALOR_FIXO Or objParcelaOV.iDesconto3Codigo = VALOR_ANT_DIA Or objParcelaOV.iDesconto3Codigo = VALOR_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col) = Format(objParcelaOV.dDesconto3Valor, "Standard")
        ElseIf objParcelaOV.iDesconto3Codigo = Percentual Or objParcelaOV.iDesconto3Codigo = PERC_ANT_DIA Or objParcelaOV.iDesconto3Codigo = PERC_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc3Perc_Col) = Format(objParcelaOV.dDesconto3Valor, "Percent")
            '*** Inicio código p/ colocar Valor Desconto
            If objParcelaOV.dValor > 0 Then
                dValorDesconto = objParcelaOV.dDesconto3Valor * objParcelaOV.dValor
                GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col) = Format(dValorDesconto, "Standard")
            End If
            '*** Fim
        End If
        For iIndice2 = 0 To TipoDesconto1.ListCount - 1
            If TipoDesconto1.ItemData(iIndice2) = objParcelaOV.iDesconto1Codigo Then GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col) = TipoDesconto1.List(iIndice2)
            If TipoDesconto2.ItemData(iIndice2) = objParcelaOV.iDesconto2Codigo Then GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col) = TipoDesconto2.List(iIndice2)
            If TipoDesconto3.ItemData(iIndice2) = objParcelaOV.iDesconto3Codigo Then GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col) = TipoDesconto3.List(iIndice2)
        Next

    Next

    objGridParcelas.iLinhasExistentes = iIndice

    Carrega_Tab_Cobranca = SUCESSO

    Exit Function

End Function

Private Function Carrega_Grid_Itens(objOrcamentoVenda As ClassOrcamentoVenda) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim dPercDesc As Double
Dim objGridItens1 As Object

On Error GoTo Erro_Carrega_Grid_Itens

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridItens)

    For iIndice = 1 To objOrcamentoVenda.colItens.Count

        objOrcamentoVenda.colItens(iIndice).iItem = iIndice

        lErro = Mascara_RetornaProdutoEnxuto(objOrcamentoVenda.colItens(iIndice).sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 84406

        'Mascara o produto enxuto
        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True

        'Calcula o percentual de desconto
        If objOrcamentoVenda.colItens(iIndice).dPrecoTotal + objOrcamentoVenda.colItens(iIndice).dValorDesconto > 0 Then
            dPercDesc = objOrcamentoVenda.colItens(iIndice).dValorDesconto / (objOrcamentoVenda.colItens(iIndice).dPrecoTotal + objOrcamentoVenda.colItens(iIndice).dValorDesconto)
        End If

        '****** IF INCLUÍDO PARA TRATAMENTO DE GRADE ***************
        If objOrcamentoVenda.colItens(iIndice).iPossuiGrade = MARCADO Then GridItens.TextMatrix(iIndice, 0) = "# " & GridItens.TextMatrix(iIndice, 0)
        
        'Coloca os dados dos itens na tela
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col) = objOrcamentoVenda.colItens(iIndice).sDescricao
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objOrcamentoVenda.colItens(iIndice).sUnidadeMed
        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objOrcamentoVenda.colItens(iIndice).dQuantidade)
        GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col) = Format(objOrcamentoVenda.colItens(iIndice).dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
        'precodesc
        Set objGridItens1 = GridItens
        Call CF("Carrega_Grid_Itens_PrecoDesc", objGridItens1, iIndice, iGrid_PrecoUnitario_Col + 1, Format(objOrcamentoVenda.colItens(iIndice).dPrecoUnitario * (1 - dPercDesc), gobjFAT.sFormatoPrecoUnitario))
        GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(dPercDesc, "Percent")
        GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(objOrcamentoVenda.colItens(iIndice).dValorDesconto, "Standard")
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col) = Format(objOrcamentoVenda.colItens(iIndice).dPrecoTotal, "Standard")
        If objOrcamentoVenda.colItens(iIndice).dtDataEntrega <> DATA_NULA Then GridItens.TextMatrix(iIndice, iGrid_DataEntrega_Col) = Format(objOrcamentoVenda.colItens(iIndice).dtDataEntrega, "dd/mm/yyyy")

    Next

    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = objOrcamentoVenda.colItens.Count

    Carrega_Grid_Itens = SUCESSO
    
    Exit Function

Erro_Carrega_Grid_Itens:

    Carrega_Grid_Itens = gErr

    Select Case gErr

        Case 84406
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objOrcamentoVenda.colItens(iIndice).sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177427)

    End Select

    Exit Function

End Function

Public Sub Cliente_Formata(lCliente As Long)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Formata

    Cliente.Text = lCliente

    'Busca o Cliente no BD
    lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
    If lErro <> SUCESSO Then gError 84411 '56915

    lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
    If lErro <> SUCESSO Then gError 84412 '56916

    'Preenche ComboBox de Filiais
    Call CF("Filial_Preenche", Filial, colCodigoNome)

    'para fazer valer o que veio do bd
    giValorDescontoManual = 1

    giClienteAlterado = 0
    
    Exit Sub

Erro_Cliente_Formata:

    Select Case gErr

        Case 84411, 84412

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177428)

    End Select

    Exit Sub

End Sub

Public Sub Filial_Formata(objFilial As Object, iFilial As Integer)

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Formata

    objFilial.Text = CStr(iFilial)
    sCliente = Cliente.Text
    objFilialCliente.iCodFilial = iFilial

    'Pesquisa se existe Filial com o código extraído
    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
    If lErro <> SUCESSO And lErro <> 17660 Then gError 84414 '56918

    If lErro = 17660 Then gError 84415 '56919

    'Coloca na tela a Filial lida
    objFilial.Text = iFilial & SEPARADOR & objFilialCliente.sNome

    Exit Sub

Erro_Filial_Formata:

    Select Case gErr

        Case 84414

        Case 84415
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, objFilial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177429)

    End Select

    Exit Sub

End Sub

Private Sub ComboItensTrib_Click()
    
Dim iIndice As Integer, objItemOV As ClassItemOV

    iIndice = ComboItensTrib.ListIndex

    If iIndice <> -1 Then

        'preenche os campos da tela em funcao do item selecionado

        Set objItemOV = gobjOrcamentoVenda.colItens.Item(iIndice + 1)

        LabelValorItem.Caption = Format(objItemOV.dPrecoTotal, "Standard")
        LabelQtdeItem.Caption = CStr(objItemOV.dQuantidade)
        LabelUMItem.Caption = objItemOV.sUnidadeMed

        Call TributacaoItem_TrazerTela(objItemOV.objTributacaoItemOV)

    End If

End Sub

Public Sub TribSobreDesconto_Click()

    'se o frame atual for o de itens
    If FrameItensTrib.Visible = True Then
        
        'exibir o de outros
        FrameOutrosTrib.Visible = True
        FrameItensTrib.Visible = False
    
    End If

    Call TributacaoItem_TrazerTela(gobjOrcamentoVenda.objTributacaoOV.objTributacaoDesconto)

End Sub

Public Sub TribSobreOutrasDesp_Click()

   'se o frame atual for o de itens
    If FrameItensTrib.Visible = True Then
        'exibir o de outros
        FrameOutrosTrib.Visible = True
        FrameItensTrib.Visible = False
    End If

    Call TributacaoItem_TrazerTela(gobjOrcamentoVenda.objTributacaoOV.objTributacaoOutras)


End Sub

Public Sub TribSobreSeguro_Click()

    'se o frame atual for o de itens
    If FrameItensTrib.Visible = True Then
        
        'exibir o de outros
        FrameOutrosTrib.Visible = True
        FrameItensTrib.Visible = False
    
    End If

    Call TributacaoItem_TrazerTela(gobjOrcamentoVenda.objTributacaoOV.objTributacaoSeguro)

End Sub

Public Sub TribSobreFrete_Click()

    'exibir o frame de "outros"
    FrameOutrosTrib.Visible = True
    FrameItensTrib.Visible = False

    Call TributacaoItem_TrazerTela(gobjOrcamentoVenda.objTributacaoOV.objTributacaoFrete)

End Sub

Public Sub TribSobreItem_Click()

    iAlterado = REGISTRO_ALTERADO

    'se houver itens na combo
    If gobjOrcamentoVenda.colItens.Count <> 0 Then
        
        'mostra o frame de itens e esconde o de outros
        FrameItensTrib.Visible = True
        FrameOutrosTrib.Visible = False
        
        'selecionar o 1o item
        ComboItensTrib.ListIndex = 0
        
        Call ComboItensTrib_Click
    
    Else
        
        'senao houver itens na combo selecionar Frete
        TribSobreFrete.Value = True
        
        Call TribSobreFrete_Click
    
    End If

End Sub

Private Sub Vendedor_Change()

    iAlterado = REGISTRO_ALTERADO
    iVendedorAlterado = 1


End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Vendedor_Validate

    'Se Vendedor foi alterado,
    If iVendedorAlterado = 1 Then

        If Len(Trim(Vendedor.Text)) > 0 Then
            
            'Tenta ler o Vendedor (NomeReduzido ou Código)
            lErro = TP_Vendedor_Le(Vendedor, objVendedor)
            If lErro <> SUCESSO Then gError 94417

        End If

        iVendedorAlterado = 0

    End If

    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True
    
    Select Case gErr
        
        Case 94417   'Tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 177430)

    End Select
    
End Sub

Private Sub VendedorLabel_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Trim(Vendedor.Text)) > 0 Then objVendedor.sNomeReduzido = Vendedor.Text
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Vendedor Then
            Call VendedorLabel_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
'        ElseIf Me.ActiveControl Is NatOpInterna Then
'            Call LblNatOpInterna_Click
        ElseIf Me.ActiveControl Is NaturezaOpItem Then
            Call NaturezaItemLabel_Click
        ElseIf Me.ActiveControl Is TipoTributacaoItem Then
            Call LblTipoTribItem_Click
        ElseIf Me.ActiveControl Is TipoTributacao Then
            Call LblTipoTrib_Click
        ElseIf Me.ActiveControl Is Codigo Then
            Call NumeroLabel_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        ElseIf Me.ActiveControl Is NaturezaOp Then 'por Leo em 02/05/02
            Call NaturezaLabel_Click
        End If
          
    End If

End Sub

Private Function Tipo_Cliente(ByVal sCliente As String) As enumTipo

    If Len(Trim(sCliente)) = 0 Then
        Tipo_Cliente = TIPO_VAZIO
    ElseIf Not IsNumeric(sCliente) Then
        Tipo_Cliente = TIPO_STRING
    ElseIf Int(CDbl(sCliente)) <> CDbl(sCliente) Then
        Tipo_Cliente = TIPO_DECIMAL
    ElseIf CDbl(sCliente) <= 0 Then
        Tipo_Cliente = TIPO_NAO_POSITIVO
    ElseIf Len(Trim(sCliente)) > STRING_CGC Then
        Tipo_Cliente = TIPO_OVERFLOW
    ElseIf Len(Trim(sCliente)) > STRING_CPF Then
        Tipo_Cliente = TIPO_CGC
    ElseIf CDbl(sCliente) > NUM_MAX_CLIENTES Then
        Tipo_Cliente = TIPO_CPF
    Else
        Tipo_Cliente = TIPO_CODIGO
    End If

End Function

Public Function Nome_Extrai(sTexto As String) As String
'Função que retira de um texto no formato "Codigo - Nome" apenas o nome.

Dim iPosicao As Integer
Dim sString As String

    iPosicao = InStr(1, sTexto, "-")
    sString = Mid(sTexto, iPosicao + 1)
    
    Nome_Extrai = sString
    
    Exit Function

End Function

Private Function TributacaoOV_Reset(Optional objOrcamentoVenda As ClassOrcamentoVenda) As Long
'cria ou atualiza gobjOrcamentoVenda, com dados correspondentes a objOrcamentoVenda (se este for passado) ou com dados "padrao"

Dim lErro As Long
Dim objTributoDoc As ClassTributoDoc

On Error GoTo Erro_TributacaoOV_Reset

    'se gobjOrcamentoVenda já foi inicializado
    If Not (gobjOrcamentoVenda Is Nothing) Then
        
        Set objTributoDoc = gobjOrcamentoVenda
        
        lErro = objTributoDoc.Desativar
        If lErro <> SUCESSO Then gError 94488
        
        Set gobjOrcamentoVenda = Nothing
    
    End If

    'se o pedido de venda veio preenchido
    If Not (objOrcamentoVenda Is Nothing) Then

        Set gobjOrcamentoVenda = objOrcamentoVenda

    Else
        
        Set gobjOrcamentoVenda = New ClassOrcamentoVenda
        gobjOrcamentoVenda.dtDataEmissao = gdtDataAtual

    End If

    Set objTributoDoc = gobjOrcamentoVenda
    lErro = objTributoDoc.Ativar
    If lErro <> SUCESSO Then gError 94489

    giNaturezaOpAlterada = 0
    giISSAliquotaAlterada = 0
    giISSValorAlterado = 0
    giValorIRRFAlterado = 0
    giTipoTributacaoAlterado = 0
    giAliqIRAlterada = 0
    iPISRetidoAlterado = 0
    iCOFINSRetidoAlterado = 0
    iCSLLRetidoAlterado = 0

    giNatOpItemAlterado = 0
    giTipoTributacaoItemAlterado = 0
    giICMSBaseItemAlterado = 0
    giICMSPercRedBaseItemAlterado = 0
    giICMSAliquotaItemAlterado = 0
    giICMSValorItemAlterado = 0
    giICMSSubstBaseItemAlterado = 0
    giICMSSubstAliquotaItemAlterado = 0
    giICMSSubstValorItemAlterado = 0
    giIPIBaseItemAlterado = 0
    giIPIPercRedBaseItemAlterado = 0
    giIPIAliquotaItemAlterado = 0
    giIPIValorItemAlterado = 0

    TributacaoOV_Reset = SUCESSO

    Exit Function

Erro_TributacaoOV_Reset:

    TributacaoOV_Reset = gErr

    Select Case gErr

        Case 94488, 94489

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177431)

    End Select

    Exit Function

End Function

Private Sub BotaoGravarTrib()

Dim lErro As Long

On Error GoTo Erro_BotaoGravarTrib

    lErro = Tributacao_GravarTela()
    If lErro <> SUCESSO Then gError 94490

    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 94491

    lErro = Carrega_Tab_Tributacao(gobjOrcamentoVenda)
    If lErro <> SUCESSO Then gError 94492

    Exit Sub

Erro_BotaoGravarTrib:

    Select Case gErr

        Case 94490, 94491, 94492

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177432)

    End Select

    Exit Sub

End Sub


Function Tributacao_GravarTela() As Long
'transfere dados de tributacao da tela para gobjOrcamentoVenda
'os dados que estiverem diferentes devem ser marcados como "manuais"

Dim lErro As Long
Dim iIndice As Integer, iTemp As Integer, dTemp As Double, objTributacaoOV As ClassTributacaoOV

On Error GoTo Erro_Tributacao_GravarTela

    Set objTributacaoOV = gobjOrcamentoVenda.objTributacaoOV

    If gobjOrcamentoVenda.sNaturezaOp <> NaturezaOp.Text Then
    
        gobjOrcamentoVenda.sNaturezaOp = NaturezaOp.Text
        gobjOrcamentoVenda.iNaturezaOpManual = VAR_PREENCH_MANUAL
        
    End If
    
    iTemp = StrParaInt(TipoTributacao.Text)
    If iTemp <> objTributacaoOV.iTipoTributacao Then
        objTributacaoOV.iTipoTributacao = iTemp
        objTributacaoOV.iTipoTributacaoManual = VAR_PREENCH_MANUAL
    End If

    'setar dados de ISS
    iTemp = ISSIncluso.Value
    If iTemp <> objTributacaoOV.iISSIncluso Then
        objTributacaoOV.iISSIncluso = iTemp
        objTributacaoOV.iISSInclusoManual = VAR_PREENCH_MANUAL
    End If

    If ISSAliquota.Text <> CStr(objTributacaoOV.dISSAliquota * 100) Then
        dTemp = StrParaDbl(ISSAliquota.Text) / 100
        If objTributacaoOV.dISSAliquota <> dTemp Then
            objTributacaoOV.dISSAliquota = dTemp
            objTributacaoOV.iISSAliquotaManual = VAR_PREENCH_MANUAL
        End If
    End If

    If ISSValor.Text <> CStr(objTributacaoOV.dISSValor) Then
        dTemp = StrParaDbl(ISSValor.Text)
        If objTributacaoOV.dISSValor <> dTemp Then
            objTributacaoOV.dISSValor = dTemp
            objTributacaoOV.iISSValorManual = VAR_PREENCH_MANUAL
        End If
    End If

    'setar dados de IR
    If IRAliquota.Text <> CStr(objTributacaoOV.dIRRFAliquota * 100) Then
        dTemp = StrParaDbl(IRAliquota.Text) / 100
        If objTributacaoOV.dIRRFAliquota <> dTemp Then
            objTributacaoOV.dIRRFAliquota = dTemp
            objTributacaoOV.iIRRFAliquotaManual = VAR_PREENCH_MANUAL
        End If
    End If

    If ValorIRRF.Text <> CStr(objTributacaoOV.dIRRFValor) Then
        dTemp = StrParaDbl(ValorIRRF.Text)
        If objTributacaoOV.dIRRFValor <> dTemp Then
            objTributacaoOV.dIRRFValor = dTemp
            objTributacaoOV.iIRRFValorManual = VAR_PREENCH_MANUAL
        End If
    End If

    If PISRetido.Text <> CStr(objTributacaoOV.dPISRetido) Then
        dTemp = StrParaDbl(PISRetido.Text)
        If objTributacaoOV.dPISRetido <> dTemp Then
            objTributacaoOV.dPISRetido = dTemp
            objTributacaoOV.iPISRetidoManual = VAR_PREENCH_MANUAL
        End If
    End If

    If COFINSRetido.Text <> CStr(objTributacaoOV.dCOFINSRetido) Then
        dTemp = StrParaDbl(COFINSRetido.Text)
        If objTributacaoOV.dCOFINSRetido <> dTemp Then
            objTributacaoOV.dCOFINSRetido = dTemp
            objTributacaoOV.iCOFINSRetidoManual = VAR_PREENCH_MANUAL
        End If
    End If

    If CSLLRetido.Text <> CStr(objTributacaoOV.dCSLLRetido) Then
        dTemp = StrParaDbl(CSLLRetido.Text)
        If objTributacaoOV.dCSLLRetido <> dTemp Then
            objTributacaoOV.dCSLLRetido = dTemp
            objTributacaoOV.iCSLLRetidoManual = VAR_PREENCH_MANUAL
        End If
    End If
    
    Tributacao_GravarTela = SUCESSO

    Exit Function

Erro_Tributacao_GravarTela:

    Tributacao_GravarTela = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177433)

    End Select

    Exit Function

End Function

Function Carrega_Tab_Tributacao(objOrcamentoVenda As ClassOrcamentoVenda) As Long

Dim lErro As Long
Dim objTributacaoOV As ClassTributacaoOV
Dim objTributacaoTipo As New ClassTipoDeTributacaoMovto

On Error GoTo Erro_Carrega_Tab_Tributacao

    giTrazendoTribTela = 1

    Set objTributacaoOV = objOrcamentoVenda.objTributacaoOV

    If NaturezaOp.Text <> objOrcamentoVenda.sNaturezaOp Then
    
        NaturezaOp.Text = objOrcamentoVenda.sNaturezaOp
        Call NaturezaOp_Validate(bSGECancelDummy)
    
    End If
    
    'no frame de "resumo"
    objTributacaoTipo.iTipo = objTributacaoOV.iTipoTributacao
    If objTributacaoTipo.iTipo <> 0 Then

        TipoTributacao.Text = CStr(objTributacaoOV.iTipoTributacao)
        
        lErro = CF("TipoTributacao_Le", objTributacaoTipo)
        If lErro <> SUCESSO Then gError 27716

        DescTipoTrib.Caption = objTributacaoTipo.sDescricao

        'se nao incide ISS
        If objTributacaoTipo.iISSIncide = 0 Then
            ISSValor.Enabled = False
            ISSAliquota.Enabled = False
            ISSIncluso.Enabled = False
        Else
            ISSValor.Enabled = True
            ISSAliquota.Enabled = True
            ISSIncluso.Enabled = True
        End If

        'se nao incide IR
        If objTributacaoTipo.iIRIncide = 0 Then
            ValorIRRF.Enabled = False
            IRAliquota.Enabled = False
        Else
            ValorIRRF.Enabled = True
            IRAliquota.Enabled = True
        End If

        'se nao retem PIS
        If objTributacaoTipo.iPISRetencao = 0 Then
            PISRetido.Enabled = False
        Else
            PISRetido.Enabled = True
        End If
    
        'se nao retem COFINS
        If objTributacaoTipo.iCOFINSRetencao = 0 Then
            COFINSRetido.Enabled = False
        Else
            COFINSRetido.Enabled = True
        End If
    
        'se nao retem CSLL
        If objTributacaoTipo.iCSLLRetencao = 0 Then
            CSLLRetido.Enabled = False
        Else
            CSLLRetido.Enabled = True
        End If
    
    Else
        
        TipoTributacao.Text = ""
        DescTipoTrib.Caption = ""
    
    End If

    IPIBase.Caption = Format(objTributacaoOV.dIPIBase, "Standard")
    IPIValor.Caption = Format(objTributacaoOV.dIPIValor, "Standard")
    ISSBase.Caption = Format(objTributacaoOV.dISSBase, "Standard")
    ISSAliquota.Text = CStr(objTributacaoOV.dISSAliquota * 100)
    ISSValor.Text = CStr(objTributacaoOV.dISSValor)
    ISSIncluso.Value = objTributacaoOV.iISSIncluso
    IRBase.Caption = Format(objTributacaoOV.dIRRFBase, "Standard")
    IRAliquota.Text = CStr(objTributacaoOV.dIRRFAliquota * 100)
    ValorIRRF.Text = CStr(objTributacaoOV.dIRRFValor)
    ICMSBase.Caption = Format(objTributacaoOV.dICMSBase, "Standard")
    ICMSValor.Caption = Format(objTributacaoOV.dICMSValor, "Standard")
    ICMSSubstBase.Caption = Format(objTributacaoOV.dICMSSubstBase, "Standard")
    ICMSSubstValor.Caption = Format(objTributacaoOV.dICMSSubstValor, "Standard")
    PISRetido.Text = CStr(objTributacaoOV.dPISRetido)
    COFINSRetido.Text = CStr(objTributacaoOV.dCOFINSRetido)
    CSLLRetido.Text = CStr(objTributacaoOV.dCSLLRetido)

    'o frame de "detalhamento" vou deixar p/carregar qdo o usuario entrar nele

    giISSAliquotaAlterada = 0
    giISSValorAlterado = 0
    giValorIRRFAlterado = 0
    giTipoTributacaoAlterado = 0
    giAliqIRAlterada = 0
    iPISRetidoAlterado = 0
    iCOFINSRetidoAlterado = 0
    iCSLLRetidoAlterado = 0
    
    giTrazendoTribTela = 0

    Carrega_Tab_Tributacao = SUCESSO

    Exit Function

Erro_Carrega_Tab_Tributacao:

    giTrazendoTribTela = 0

    Carrega_Tab_Tributacao = gErr

    Select Case gErr

        Case 27716

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177434)

    End Select

End Function

Private Sub BotaoGravarTribItem_Click()

Dim lErro As Long, objTributacaoItemOV As ClassTributacaoItemPV, iIndice As Integer

On Error GoTo Erro_BotaoGravarTribItem_Click

    'atualizar dados da colecao p/o item ou complemento corrente

    'se um item estiver selecionado
    If TribSobreItem.Value = True Then
        iIndice = ComboItensTrib.ListIndex
        If iIndice <> -1 Then
            Set objTributacaoItemOV = gobjOrcamentoVenda.colItens.Item(iIndice + 1).objTributacaoItemOV
        Else
            gError 103008 '27668
        End If
    Else
        If TribSobreDesconto.Value = True Then
            Set objTributacaoItemOV = gobjOrcamentoVenda.objTributacaoOV.objTributacaoDesconto
        Else
            If TribSobreFrete.Value = True Then
                Set objTributacaoItemOV = gobjOrcamentoVenda.objTributacaoOV.objTributacaoFrete
            Else
                If TribSobreSeguro.Value = True Then
                    Set objTributacaoItemOV = gobjOrcamentoVenda.objTributacaoOV.objTributacaoSeguro
                Else
                    If TribSobreOutrasDesp.Value = True Then
                        Set objTributacaoItemOV = gobjOrcamentoVenda.objTributacaoOV.objTributacaoOutras
                    End If
                End If
            End If
        End If
    End If

    lErro = TributacaoItem_GravarTela(objTributacaoItemOV)
    If lErro <> SUCESSO Then gError 103009 '27667

    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 103010 ' 27666

    lErro = TributacaoItem_TrazerTela(objTributacaoItemOV)
    If lErro <> SUCESSO Then gError 103011 ' 27712

    Exit Sub

Erro_BotaoGravarTribItem_Click:

    Select Case gErr

        Case 103009, 103010, 103011

        Case 103008
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NENHUM_ITEM_TRIB_SEL", gErr, Error)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177435)

    End Select

    Exit Sub

End Sub

Function TributacaoItem_GravarTela(objTributacaoItemOV As ClassTributacaoItemPV) As Long
'transfere dados de tributacao de um item da tela para objTributacaoItemOV
'os dados que estiverem diferentes devem ser marcados como "manuais"

Dim lErro As Long
Dim iIndice As Integer
Dim iTemp As Integer
Dim dTemp As Double
Dim sTemp As String

On Error GoTo Erro_TributacaoItem_GravarTela

    sTemp = Trim(NaturezaOpItem.Text)
    If Trim(objTributacaoItemOV.sNaturezaOp) <> sTemp Then
        objTributacaoItemOV.sNaturezaOp = sTemp
        objTributacaoItemOV.iNaturezaOpManual = VAR_PREENCH_MANUAL
    End If

    iTemp = StrParaInt(TipoTributacaoItem.Text)
    If iTemp <> objTributacaoItemOV.iTipoTributacao Then
        objTributacaoItemOV.iTipoTributacao = iTemp
        objTributacaoItemOV.iTipoTributacaoManual = VAR_PREENCH_MANUAL
    End If

    'Setar dados de ICMS

    iTemp = ComboICMSTipo.ItemData(ComboICMSTipo.ListIndex)
    If iTemp <> objTributacaoItemOV.iICMSTipo Then
        objTributacaoItemOV.iICMSTipo = iTemp
        objTributacaoItemOV.iICMSTipoManual = VAR_PREENCH_MANUAL
    End If

    If ICMSBaseItem.Text <> CStr(objTributacaoItemOV.dICMSBase) Then
        dTemp = StrParaDbl(ICMSBaseItem.Text)
        objTributacaoItemOV.dICMSBase = dTemp
        objTributacaoItemOV.iICMSBaseManual = VAR_PREENCH_MANUAL
    End If

    If ICMSPercRedBaseItem.Text <> CStr(objTributacaoItemOV.dICMSPercRedBase * 100) Then
        dTemp = StrParaDbl(ICMSPercRedBaseItem.Text) / 100
        objTributacaoItemOV.dICMSPercRedBase = dTemp
        objTributacaoItemOV.iICMSPercRedBaseManual = VAR_PREENCH_MANUAL
    End If

    If ICMSAliquotaItem.Text <> CStr(objTributacaoItemOV.dICMSAliquota * 100) Then
        dTemp = StrParaDbl(ICMSAliquotaItem.Text) / 100
        objTributacaoItemOV.dICMSAliquota = dTemp
        objTributacaoItemOV.iICMSAliquotaManual = VAR_PREENCH_MANUAL
    End If

    If ICMSValorItem.Text <> CStr(objTributacaoItemOV.dICMSValor) Then
        dTemp = StrParaDbl(ICMSValorItem.Text)
        objTributacaoItemOV.dICMSValor = dTemp
        objTributacaoItemOV.iICMSValorManual = VAR_PREENCH_MANUAL
    End If

    'setar dados ICMS Substituicao

    If ICMSSubstBaseItem.Text <> CStr(objTributacaoItemOV.dICMSSubstBase) Then
        dTemp = StrParaDbl(ICMSSubstBaseItem.Text)
        objTributacaoItemOV.dICMSSubstBase = dTemp
        objTributacaoItemOV.iICMSSubstBaseManual = VAR_PREENCH_MANUAL
    End If

    If ICMSSubstAliquotaItem.Text <> CStr(objTributacaoItemOV.dICMSSubstAliquota * 100) Then
        dTemp = StrParaDbl(ICMSSubstAliquotaItem.Text) / 100
        objTributacaoItemOV.dICMSSubstAliquota = dTemp
        objTributacaoItemOV.iICMSSubstAliquotaManual = VAR_PREENCH_MANUAL
    End If

    If ICMSSubstValorItem.Text <> CStr(objTributacaoItemOV.dICMSSubstValor) Then
        dTemp = StrParaDbl(ICMSSubstValorItem.Text)
        objTributacaoItemOV.dICMSSubstValor = dTemp
        objTributacaoItemOV.iICMSSubstValorManual = VAR_PREENCH_MANUAL
    End If

    'setar dados de IPI
    iTemp = ComboIPITipo.ItemData(ComboIPITipo.ListIndex)
    If iTemp <> objTributacaoItemOV.iIPITipo Then
        objTributacaoItemOV.iIPITipo = iTemp
        objTributacaoItemOV.iIPITipoManual = VAR_PREENCH_MANUAL
    End If

    If IPIBaseItem.Text <> CStr(objTributacaoItemOV.dIPIBaseCalculo) Then
        dTemp = StrParaDbl(IPIBaseItem.Text)
        objTributacaoItemOV.dIPIBaseCalculo = dTemp
        objTributacaoItemOV.iIPIBaseManual = VAR_PREENCH_MANUAL
    End If

    If IPIPercRedBaseItem.Text <> CStr(objTributacaoItemOV.dIPIPercRedBase * 100) Then
        dTemp = StrParaDbl(IPIPercRedBaseItem.Text) / 100
        objTributacaoItemOV.dIPIPercRedBase = dTemp
        objTributacaoItemOV.iIPIPercRedBaseManual = VAR_PREENCH_MANUAL
    End If

    If IPIAliquotaItem.Text <> CStr(objTributacaoItemOV.dIPIAliquota * 100) Then
        dTemp = StrParaDbl(IPIAliquotaItem.Text) / 100
        objTributacaoItemOV.dIPIAliquota = dTemp
        objTributacaoItemOV.iIPIAliquotaManual = VAR_PREENCH_MANUAL
    End If

    If IPIValorItem.Text <> CStr(objTributacaoItemOV.dIPIValor) Then
        dTemp = StrParaDbl(IPIValorItem.Text)
        objTributacaoItemOV.dIPIValor = dTemp
        objTributacaoItemOV.iIPIValorManual = VAR_PREENCH_MANUAL
    End If

    TributacaoItem_GravarTela = SUCESSO

    Exit Function

Erro_TributacaoItem_GravarTela:

    TributacaoItem_GravarTela = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177436)

    End Select

    Exit Function

End Function

Function TributacaoItem_TrazerTela(objTributacaoItemOV As ClassTributacaoItemPV) As Long
'Traz para a tela dados de tributacao de um item

Dim iIndice As Integer
Dim objItemOrcamento As ClassItemOV
Dim lErro As Long
Dim objTipoTribIPI As New ClassTipoTribIPI
Dim objTipoTribICMS As New ClassTipoTribICMS
Dim objTributacaoTipo As New ClassTipoDeTributacaoMovto
Dim objNaturezaOp As New ClassNaturezaOp
Dim sNatOp As String

On Error GoTo Erro_TributacaoItem_TrazerTela

    giTrazendoTribItemTela = 1

    NaturezaOpItem.Text = objTributacaoItemOV.sNaturezaOp

    sNatOp = Trim(NaturezaOpItem.Text)

    If sNatOp <> "" Then

        objNaturezaOp.sCodigo = sNatOp
        'Lê a Natureza de Operação
        lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
        If lErro <> SUCESSO And lErro <> 17958 Then gError 103012 '28651

        'Se não achou a Natureza de Operação --> erro
        If lErro <> SUCESSO Then gError 103014 '28652

        LabelDescrNatOpItem.Caption = objNaturezaOp.sDescricao
    Else
        LabelDescrNatOpItem.Caption = ""
    End If

    objTributacaoTipo.iTipo = objTributacaoItemOV.iTipoTributacao
    If objTributacaoTipo.iTipo <> 0 Then

        lErro = CF("TipoTributacao_Le", objTributacaoTipo)
        If lErro <> SUCESSO Then gError 103013 '27711

        TipoTributacaoItem.Text = CStr(objTributacaoItemOV.iTipoTributacao)
        DescTipoTribItem.Caption = objTributacaoTipo.sDescricao

        'Se não incide IPI
        If objTributacaoTipo.iIPIIncide = 0 Then
            
            ComboIPITipo.Enabled = False
            IPIBaseItem.Enabled = False
        Else
            
            ComboIPITipo.Enabled = True
            IPIBaseItem.Enabled = True
        
        End If

        'Se não incide ICMS
        If objTributacaoTipo.iICMSIncide = 0 Then
            
            ComboICMSTipo.Enabled = False
            ICMSBaseItem.Enabled = False
        Else
            
            ComboICMSTipo.Enabled = True
            ICMSBaseItem.Enabled = True
        
        End If

    Else
    
        TipoTributacaoItem.Text = ""
        DescTipoTribItem.Caption = ""
    
    End If

    'Setar dados de ICMS
    Call Combo_Seleciona_ItemData(ComboICMSTipo, objTributacaoItemOV.iICMSTipo)
    
    ICMSBaseItem.Text = CStr(objTributacaoItemOV.dICMSBase)
    ICMSPercRedBaseItem.Text = CStr(objTributacaoItemOV.dICMSPercRedBase * 100)
    ICMSAliquotaItem.Text = CStr(objTributacaoItemOV.dICMSAliquota * 100)
    ICMSValorItem.Text = CStr(objTributacaoItemOV.dICMSValor)

    'setar dados ICMS Substituicao
    ICMSSubstBaseItem.Text = CStr(objTributacaoItemOV.dICMSSubstBase)
    ICMSSubstAliquotaItem.Text = CStr(objTributacaoItemOV.dICMSSubstAliquota * 100)
    ICMSSubstValorItem.Text = CStr(objTributacaoItemOV.dICMSSubstValor)

    For Each objTipoTribICMS In gcolTiposTribICMS
        If objTipoTribICMS.iTipo = objTributacaoItemOV.iICMSTipo Then Exit For
    Next

    'Se permite redução de base habilitar este campo
    If objTipoTribICMS.iPermiteReducaoBase Then
        ICMSPercRedBaseItem.Enabled = True
    Else
        'Desabilita-lo e limpa-lo em caso contrário
        ICMSPercRedBaseItem.Enabled = False
    End If

    'Se permite aliquota habilitar este campo e valor.
    If objTipoTribICMS.iPermiteAliquota Then
        
        ICMSAliquotaItem.Enabled = True
        ICMSValorItem.Enabled = True
    
    Else
        
        'Desabilitar os dois campos e coloca-los com zero
        ICMSAliquotaItem.Enabled = False
        ICMSValorItem.Enabled = False
    
    End If

    'Se permite margem de lucro habilitar campos do frame de substituicao
    If objTipoTribICMS.iPermiteMargLucro Then
        
        ICMSSubstBaseItem.Enabled = True
        ICMSSubstAliquotaItem.Enabled = True
        ICMSSubstValorItem.Enabled = True
    
    Else
        
        'Limpa-los e desabilita-los
        ICMSSubstBaseItem.Enabled = False
        ICMSSubstAliquotaItem.Enabled = False
        ICMSSubstValorItem.Enabled = False
    
    End If

    'Setar dados de IPI
    Call Combo_Seleciona_ItemData(ComboIPITipo, objTributacaoItemOV.iIPITipo)
    
    IPIBaseItem.Text = CStr(objTributacaoItemOV.dIPIBaseCalculo)
    IPIPercRedBaseItem.Text = CStr(objTributacaoItemOV.dIPIPercRedBase * 100)
    IPIAliquotaItem.Text = CStr(objTributacaoItemOV.dIPIAliquota * 100)
    IPIValorItem.Text = CStr(objTributacaoItemOV.dIPIValor)

    For Each objTipoTribIPI In gcolTiposTribIPI
        If objTipoTribIPI.iTipo = objTributacaoItemOV.iIPITipo Then Exit For
    Next

    'Se permite redução de base habilitar este campo
    If objTipoTribIPI.iPermiteReducaoBase Then 'leo voltar aqui
        IPIPercRedBaseItem.Enabled = True
    Else
        
        'desabilita-lo e limpa-lo em caso contrário
        IPIPercRedBaseItem.Enabled = False
    
    End If

    'Se permite alíquota habilitar este campo e valor.
    If objTipoTribIPI.iPermiteAliquota Then
        
        IPIAliquotaItem.Enabled = True
        IPIValorItem.Enabled = True
    
    Else
        'Desabilitar os dois campos e coloca-los com zero
        IPIAliquotaItem.Enabled = False
        IPIValorItem.Enabled = False
    
    End If

    giTrazendoTribItemTela = 0
    giNatOpItemAlterado = 0
    giTipoTributacaoItemAlterado = 0
    giICMSBaseItemAlterado = 0
    giICMSPercRedBaseItemAlterado = 0
    giICMSAliquotaItemAlterado = 0
    giICMSValorItemAlterado = 0
    giICMSSubstBaseItemAlterado = 0
    giICMSSubstAliquotaItemAlterado = 0
    giICMSSubstValorItemAlterado = 0
    giIPIBaseItemAlterado = 0
    giIPIPercRedBaseItemAlterado = 0
    giIPIAliquotaItemAlterado = 0
    giIPIValorItemAlterado = 0

    TributacaoItem_TrazerTela = SUCESSO

    Exit Function

Erro_TributacaoItem_TrazerTela:

    giTrazendoTribItemTela = 0

    TributacaoItem_TrazerTela = gErr

    Select Case gErr

        Case 103013, 103012

        Case 103014
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", objNaturezaOp.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177437)

    End Select

    Exit Function

End Function

Private Function TributacaoItem_InicializaTab() As Long
'deve ser chamada na entrada do tab de detalhamento dentro do tab de tributacao
Dim lErro As Long
Dim objItemOrcamento As ClassItemOV
Dim sItem As String

On Error GoTo Erro_TributacaoItem_InicializaTab

    'preencher o valor de frete, seguro, descontos e outras desp no frameOutros
    LabelValorFrete.Caption = Format(gobjOrcamentoVenda.dValorFrete, "Standard")
    LabelValorDesconto.Caption = Format(gobjOrcamentoVenda.dValorDesconto, "Standard")
    LabelValorSeguro.Caption = Format(gobjOrcamentoVenda.dValorSeguro, "Standard")
    LabelValorOutrasDespesas.Caption = Format(gobjOrcamentoVenda.dValorOutrasDespesas, "Standard")

    'esvaziar a combo de itens
    ComboItensTrib.Clear

    'preencher a combo de itens: com "codigo do produto - descricao"
    For Each objItemOrcamento In gobjOrcamentoVenda.colItens

        lErro = Mascara_MascararProduto(objItemOrcamento.sProduto, sItem)
        If lErro <> SUCESSO Then gError 103033

        sItem = sItem & " - " & objItemOrcamento.sDescricao
        ComboItensTrib.AddItem sItem
    
    Next

    TribSobreItem.Value = True
    Call TribSobreItem_Click

    TributacaoItem_InicializaTab = SUCESSO

    Exit Function

Erro_TributacaoItem_InicializaTab:

    TributacaoItem_InicializaTab = gErr

    Select Case gErr

        Case 103033

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177438)

    End Select

    Exit Function

End Function

Private Sub BotaoGravarTribCarga()

Dim lErro As Long

On Error GoTo Erro_BotaoGravarTribCarga

    lErro = Tributacao_GravarTela()
    If lErro <> SUCESSO Then gError 59289

    'Atualiza os valores de tributação
    lErro = AtualizarTributacao()
    If lErro <> SUCESSO Then gError 59290

    Exit Sub

Erro_BotaoGravarTribCarga:

    Select Case gErr

        Case 59289, 59290

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177439)

    End Select

    Exit Sub

End Sub
 
Private Function TributacaoOV_Terminar() As Long

Dim lErro As Long, objTributoDoc As ClassTributoDoc

On Error GoTo Erro_TributacaoOV_Terminar

    If Not (gobjOrcamentoVenda Is Nothing) Then
        Set objTributoDoc = gobjOrcamentoVenda
        lErro = objTributoDoc.Desativar
        If lErro <> SUCESSO Then gError 27710
        Set gobjOrcamentoVenda = Nothing
    End If

    TributacaoOV_Terminar = SUCESSO

    Exit Function

Erro_TributacaoOV_Terminar:

    TributacaoOV_Terminar = gErr

    Select Case gErr

        Case 27710

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177440)

    End Select

End Function

Private Function CarregaTiposTrib() As Long

Dim lErro As Long, sCodigo As String
Dim objTipoTribICMS As ClassTipoTribICMS
Dim objTipoTribIPI As ClassTipoTribIPI

On Error GoTo Erro_CarregaTiposTrib

    lErro = CF("TiposTribICMS_Le_Todos", gcolTiposTribICMS)
    If lErro <> SUCESSO Then gError 27636

    'Preenche ComboICMSTipo
    For Each objTipoTribICMS In gcolTiposTribICMS

        sCodigo = Space(STRING_TIPO_ICMS_CODIGO - Len(CStr(objTipoTribICMS.iTipo)))
        sCodigo = sCodigo & CStr(objTipoTribICMS.iTipo) & SEPARADOR & objTipoTribICMS.sDescricao
        ComboICMSTipo.AddItem (sCodigo)
        ComboICMSTipo.ItemData(ComboICMSTipo.NewIndex) = objTipoTribICMS.iTipo

    Next

    lErro = CF("TiposTribIPI_Le_Todos", gcolTiposTribIPI)
    If lErro <> SUCESSO Then gError 27637

    'Preenche ComboIPITipo
    For Each objTipoTribIPI In gcolTiposTribIPI

        sCodigo = Space(STRING_TIPO_ICMS_CODIGO - Len(CStr(objTipoTribIPI.iTipo)))
        sCodigo = sCodigo & CStr(objTipoTribIPI.iTipo) & SEPARADOR & objTipoTribIPI.sDescricao
        ComboIPITipo.AddItem (sCodigo)
        ComboIPITipo.ItemData(ComboIPITipo.NewIndex) = objTipoTribIPI.iTipo

    Next

    CarregaTiposTrib = SUCESSO

    Exit Function

Erro_CarregaTiposTrib:

    CarregaTiposTrib = gErr

    Select Case gErr

        Case 27636, 27637

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177441)

    End Select

    Exit Function

End Function

Private Function AtualizarTributacao() As Long

Dim lErro As Long

On Error GoTo Erro_AtualizarTributacao

    If Not (gobjOrcamentoVenda Is Nothing) Then

        'Atualiza os impostos
        lErro = gobjTributacao.AtualizaImpostos(gobjOrcamentoVenda, giRecalculandoTributacao)
        If lErro <> SUCESSO Then gError 27649

        'joga dados do obj atualizado p/a tela
        lErro = Carrega_Tab_Tributacao(gobjOrcamentoVenda)
        If lErro <> SUCESSO Then gError 27661

    End If

    AtualizarTributacao = SUCESSO

    Exit Function

Erro_AtualizarTributacao:

    AtualizarTributacao = gErr

    Select Case gErr

        Case 27649, 27661

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177442)

    End Select

    Exit Function

End Function

Private Function Tributacao_Inclusao_Item_Grid(iLinha As Integer, sProduto As String) As Long
'trata a inclusao de uma linha de item no grid
Dim lErro As Long
Dim objTributoDocItem As ClassTributoDocItem
Dim objItemOV As ClassItemOV
On Error GoTo Erro_Tributacao_Inclusao_Item_Grid

    lErro = Move_GridItem_Memoria(gobjOrcamentoVenda, iLinha, sProduto)
    If lErro <> SUCESSO Then gError 27683

    Set objItemOV = gobjOrcamentoVenda.colItens.Item(iLinha)
    Set objTributoDocItem = objItemOV
    lErro = objTributoDocItem.Ativar(gobjOrcamentoVenda)
    If lErro <> SUCESSO Then gError 27686

    Tributacao_Inclusao_Item_Grid = SUCESSO

    Exit Function

Erro_Tributacao_Inclusao_Item_Grid:

    Tributacao_Inclusao_Item_Grid = gErr

    Select Case gErr

        Case 27683, 27686

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177443)

    End Select

    Exit Function

End Function

Function Tributacao_Remover_Item_Grid(iLinha As Integer) As Long
'trata a exclusao de uma linha de item no grid
Dim objItemOV As ClassItemOV, objTributoDocItem As ClassTributoDocItem

        Set objItemOV = gobjOrcamentoVenda.colItens(iLinha)
        Set objTributoDocItem = objItemOV
        Call objTributoDocItem.Desativar
        Call gobjOrcamentoVenda.RemoverItem(iLinha)

End Function

Function Tributacao_Alteracao_Item_Grid(iIndice As Integer) As Long
'trata a alteracao de uma linha de item no grid

Dim lErro As Long, sProduto As String, iPreenchido As Integer
Dim objItemOV As ClassItemOV

On Error GoTo Erro_Tributacao_Alteracao_Item_Grid

    Set objItemOV = gobjOrcamentoVenda.colItens.Item(iIndice)

    If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Produto_Col))) > 0 Then

        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 27709

        objItemOV.sProduto = sProduto
    
    End If

    objItemOV.sUnidadeMed = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)

    objItemOV.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
    'objItemOV.dQuantCancelada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantCancel_Col))
    objItemOV.dPrecoTotal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
    objItemOV.dValorDesconto = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))

    Tributacao_Alteracao_Item_Grid = SUCESSO

    Exit Function

Erro_Tributacao_Alteracao_Item_Grid:

    Tributacao_Alteracao_Item_Grid = gErr

    Select Case gErr

        Case 27709

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177444)

    End Select

    Exit Function

End Function

Public Function Valida_Tributacao_Gravacao() As Long

Dim lErro As Long
Dim objItemOV As ClassItemOV
Dim iIndice As Integer, dtDataRef As Date

On Error GoTo Erro_Valida_Tributacao_Gravacao

    If gobjOrcamentoVenda.objTributacaoOV.iTipoTributacao = 0 Then gError 56920
    
    dtDataRef = gobjOrcamentoVenda.dtDataEmissao
    
    iIndice = 0
    
    For Each objItemOV In gobjOrcamentoVenda.colItens

        iIndice = iIndice + 1
        If Len(Trim(objItemOV.objTributacaoItemOV.sNaturezaOp)) = 0 Then gError 56921
        If objItemOV.objTributacaoItemOV.iTipoTributacao = 0 Then gError 56922
        If Natop_ErroTamanho(dtDataRef, objItemOV.objTributacaoItemOV.sNaturezaOp) Then Error 32282

    Next

    If Len(Trim(gobjOrcamentoVenda.objTributacaoOV.objTributacaoDesconto.sNaturezaOp)) = 0 Then gError 56923
    If gobjOrcamentoVenda.objTributacaoOV.objTributacaoDesconto.iTipoTributacao = 0 Then gError 56924
    
    If Len(Trim(gobjOrcamentoVenda.objTributacaoOV.objTributacaoFrete.sNaturezaOp)) = 0 Then gError 56925
    If gobjOrcamentoVenda.objTributacaoOV.objTributacaoFrete.iTipoTributacao = 0 Then gError 56926
    
    If Len(Trim(gobjOrcamentoVenda.objTributacaoOV.objTributacaoOutras.sNaturezaOp)) = 0 Then gError 56927
    If gobjOrcamentoVenda.objTributacaoOV.objTributacaoOutras.iTipoTributacao = 0 Then gError 56928
    
    If Len(Trim(gobjOrcamentoVenda.objTributacaoOV.objTributacaoSeguro.sNaturezaOp)) = 0 Then gError 56929
    If gobjOrcamentoVenda.objTributacaoOV.objTributacaoSeguro.iTipoTributacao = 0 Then gError 56930

    If Natop_ErroTamanho(dtDataRef, gobjOrcamentoVenda.sNaturezaOp) Or _
        Natop_ErroTamanho(dtDataRef, gobjOrcamentoVenda.objTributacaoOV.objTributacaoDesconto.sNaturezaOp) Or _
        Natop_ErroTamanho(dtDataRef, gobjOrcamentoVenda.objTributacaoOV.objTributacaoFrete.sNaturezaOp) Or _
        Natop_ErroTamanho(dtDataRef, gobjOrcamentoVenda.objTributacaoOV.objTributacaoOutras.sNaturezaOp) Or _
        Natop_ErroTamanho(dtDataRef, gobjOrcamentoVenda.objTributacaoOV.objTributacaoSeguro.sNaturezaOp) Then Error 32281

    Valida_Tributacao_Gravacao = SUCESSO

    Exit Function
    
Erro_Valida_Tributacao_Gravacao:

    Valida_Tributacao_Gravacao = gErr

    Select Case gErr
    
        Case 56920
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_NAO_PREENCHIDO", gErr)
        
        Case 56921
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_ITEM_TRIBUTACAO_NAO_PREENCHIDA", iIndice)
        
        Case 56922
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_ITEM_NAO_PREENCHIDO", gErr, iIndice)
        
        Case 56923
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_DESCONTO_NAO_PRENCHIDA", gErr)
        
        Case 56924
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_DESCONTO_NAO_PREENCHIDO", gErr)
        
        Case 56925
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_FRETE_NAO_PRENCHIDA", gErr)
        
        Case 56926
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_FRETE_NAO_PREENCHIDO", gErr)

        Case 56927
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_DESPESAS_NAO_PRENCHIDA", gErr)
        
        Case 56928
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_DESPESAS_NAO_PREENCHIDO", gErr)
    
        Case 56929
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_SEGURO_NAO_PRENCHIDA", gErr)
        
        Case 56930
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_SEGURO_NAO_PREENCHIDO", gErr)
        
        Case 32281, 32282
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_TAMANHO_INCORRETO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177445)
    
    End Select

    Exit Function
    
End Function

Public Function CargaPosFormLoad(Optional bTrazendoDoc As Boolean = False) As Long

Dim lErro As Long

On Error GoTo Erro_CargaPosFormLoad

    If (giPosCargaOk = 0) Then
    
        'p/permitir o redesenho da tela
        DoEvents
        
        gbCarregandoTela = True
             
        lErro = TributacaoOV_Reset()
        If lErro <> SUCESSO Then gError 27643
    
        Call BotaoGravarTribCarga
    
        lErro = CarregaTiposTrib()
        If lErro <> SUCESSO Then gError 27638
   
        'Carrega a combo combo de Tabela de Preços
        lErro = Carrega_TabelaPreco()
        If lErro <> SUCESSO Then gError 26481
    
        'Carrega a combo de Condição de Pagamento
        lErro = Carrega_CondicaoPagamento()
        If lErro <> SUCESSO Then gError 26490
        
        PrecoUnitario.Format = gobjFAT.sFormatoPrecoUnitario
        
        Quantidade.Format = FORMATO_ESTOQUE
    
        'Preenche Data Referencia e Data de Emissão coma Data Atual
        DataReferencia.PromptInclude = False
        DataReferencia.Text = Format(gdtDataAtual, "dd/mm/yy")
        DataReferencia.PromptInclude = True
        giDataReferenciaAlterada = 0
    
        Set objGridItens = New AdmGrid
        Set objGridParcelas = New AdmGrid
        
        Set objEventoCliente = New AdmEvento
        Set objEventoNumero = New AdmEvento
        Set objEventoCondPagto = New AdmEvento
        Set objEventoProduto = New AdmEvento
        Set objEventoVendedor = New AdmEvento
        Set objEventoNaturezaOp = New AdmEvento
        Set objEventoTiposDeTributacao = New AdmEvento

        'Faz as Inicializações dos Grids
        lErro = Inicializa_Grid_Itens(objGridItens)
        If lErro <> SUCESSO Then gError 26493
    
        lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
        If lErro <> SUCESSO Then gError 26636
    
        lErro = Inicializa_Grid_Parcelas(objGridParcelas)
        If lErro <> SUCESSO Then gError 26496
    
        gbCarregandoTela = False
     
        iAlterado = 0
        
        giPosCargaOk = 1

        Call ValorTotal_Calcula
    
    End If
    
    CargaPosFormLoad = SUCESSO
    
    Exit Function
     
Erro_CargaPosFormLoad:

    gbCarregandoTela = False
    
    CargaPosFormLoad = gErr
    
    Select Case gErr
          '????? Trocar a numeracao de erro
        Case 46531, 26481, 26483, 26485, 26487, 26490, 26491
        Case 26493, 26636, 26495, 26496, 26497, 26635
        Case 27638, 46177, 27643, 96126
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177446)
     
    End Select
     
    Exit Function

End Function

Private Sub objEventoCondPagto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCondicaoPagto As ClassCondicaoPagto
Dim vbMsgRes As VbMsgBoxResult
Dim dPercAcresFin As Double

On Error GoTo Erro_objEventoCondPagto_evSelecao

    Set objCondicaoPagto = obj1

    'Preenche campo CondicaoPagamento
    CondicaoPagamento.Text = CStr(objCondicaoPagto.iCodigo) & SEPARADOR & objCondicaoPagto.sDescReduzida

    'Altera PercAcrescFin
    If Len(Trim(PercAcrescFin.ClipText)) > 0 Then
        
        dPercAcresFin = StrParaDbl(PercAcrescFin.Text) / 100

        If dPercAcresFin <> objCondicaoPagto.dAcrescimoFinanceiro Then
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUICAO_PERC_ACRESCIMO_FINANCEIRO")
            If vbMsgRes = vbYes Then
                PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
                Call PercAcrescFin_Validate(bSGECancelDummy)
            End If
        End If
    Else
        PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
        Call PercAcrescFin_Validate(bSGECancelDummy)
    End If

    If Len(Trim(ValorTotal.Caption)) > 0 Then
        'Se DataReferencia estiver preenchida e Valor for positivo
        If Len(Trim(DataReferencia.ClipText)) > 0 And CDbl(ValorTotal.Caption) > 0 Then

            'Preenche GridParcelas a partir da Condição de Pagto
            lErro = Cobranca_Automatica()
            If lErro <> SUCESSO Then gError 26500

        End If
    End If

    Me.Show

    Exit Sub

Erro_objEventoCondPagto_evSelecao:

    Select Case gErr

        Case 26500

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177447)

     End Select

     Exit Sub

End Sub

'Incluído por Luiz Nogueira em 04/06/03
Public Sub BotaoImprimir_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoImprimir_Click

    'Se o código do orçamento não foi informado => erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 102238
    
    'Dispara função para imprimir orçamento
    lErro = Orcamento_Imprime(Trim(Codigo.Text))
    If lErro <> SUCESSO Then gError 102239
    
    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 102239
        
        Case 102238
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PEDIDO_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 177448)

    End Select

    Exit Sub

End Sub

'Incluído por Luiz Nogueira em 04/06/03
Private Function Orcamento_Imprime(ByVal lOrcamento As Long) As Long

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objOrcamentoVenda As New ClassOrcamentoVenda

On Error GoTo Erro_Orcamento_Imprime

    'Transforma o ponteiro do mouse em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Guarda no obj o código do orçamento passado como parâmetro
    objOrcamentoVenda.lCodigo = lOrcamento
    
    'Guarda a FilialEmpresa ativa como filial do orçamento
    objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa
    
    'Lê os dados do orçamento para verificar se o mesmo existe no BD
    lErro = CF("OrcamentoVenda_Le", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 101232 Then gError 102235

    'Se não encontrou => erro, pois não é possível imprimir um orçamento inexistente
    If lErro = 101232 Then gError 102236
    
    'Dispara a impressão do relatório
    lErro = objRelatorio.ExecutarDireto("Orçamento de Vendas", "OrcamentoVenda >= @NORCVENDINIC E OrcamentoVenda <= @NORCVENDFIM", 1, "OrcVenda", "NORCVENDINIC", Trim(Codigo.Text), "NORCVENDFIM", Trim(Codigo.Text), "NEXIBIRORC", 1)
    If lErro <> SUCESSO Then gError 102237

    'Transforma o ponteiro do mouse em seta (padrão)
    GL_objMDIForm.MousePointer = vbDefault
    
    Orcamento_Imprime = SUCESSO
    
    Exit Function

Erro_Orcamento_Imprime:

    Orcamento_Imprime = gErr
    
    Select Case gErr
    
        Case 102235, 102237
        
        Case 102236
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO", gErr, objOrcamentoVenda.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177449)
    
    End Select
    
    'Transforma o ponteiro do mouse em seta (padrão)
    GL_objMDIForm.MousePointer = vbDefault

End Function


Private Sub BotaoGrade_Click()

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim lErro  As Long
Dim objRomaneioGrade As ClassRomaneioGrade
Dim objItemOV As ClassItemOV

On Error GoTo Erro_BotaoGrade_Click

    'Se a linha selecionada for uma existente
    If GridItens.Row > 0 And GridItens.Row <= objGridItens.iLinhasExistentes Then
    
        'Recolhe o item  corr
        Set objItemOV = gobjOrcamentoVenda.colItens(GridItens.Row)
        
        If objItemOV.iPossuiGrade = MARCADO Then
            
            Set objRomaneioGrade = New ClassRomaneioGrade
            
            objRomaneioGrade.sNomeTela = Me.Name
            Set objRomaneioGrade.objObjetoTela = objItemOV
                        
            Call Chama_Tela_Modal("RomaneioGrade", objRomaneioGrade)
            If lErro <> SUCESSO Then gError 86360
        
            Call Atualiza_Grid_Itens(objItemOV)
                    
            If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col))) > 0 Then
                Call PrecoTotal_Calcula(GridItens.Row)
                lErro = ValorTotal_Calcula()
                If lErro <> SUCESSO Then gError 84162
            End If
        End If
    
    End If
    
    Exit Sub

Erro_BotaoGrade_Click:

    Select Case gErr
    
        Case 84162, 86360
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177450)
            
    End Select
    
    Exit Sub

End Sub

Sub Atualiza_Grid_Itens(objItemOV As ClassItemOV)

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim dQuantidade As Double
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
    
    For Each objItemRomaneioGrade In objItemOV.colItensRomaneioGrade
            
        dQuantidade = dQuantidade + objItemRomaneioGrade.dQuantidade
        
    Next

    GridItens.TextMatrix(objItemOV.iItem, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)

    objItemOV.dQuantidade = dQuantidade
    
    Exit Sub

End Sub

Function Grid_Possui_Grade() As Boolean

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim iIndice As Integer

    For iIndice = 1 To gobjOrcamentoVenda.colItens.Count
        If gobjOrcamentoVenda.colItens(iIndice).iPossuiGrade = MARCADO Then
            Grid_Possui_Grade = True
            Exit Function
        End If
    Next
    
    Grid_Possui_Grade = False
        
    Exit Function
    
End Function


Function Move_ItensGrade_Tela(colItensRomaneio As Collection, colItensRomaneioTela As Collection) As Long

Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objItemRomaneioGradeTela As ClassItemRomaneioGrade

    'Para cada Item de Romaneio vindo da tela ( Aqueles que já tem quantidade)
    For Each objItemRomaneioGradeTela In colItensRomaneioTela
                    
        Set objItemRomaneioGrade = New ClassItemRomaneioGrade
            
        objItemRomaneioGrade.sProduto = objItemRomaneioGradeTela.sProduto
        objItemRomaneioGrade.dQuantOP = objItemRomaneioGradeTela.dQuantOP
        objItemRomaneioGrade.dQuantSC = objItemRomaneioGradeTela.dQuantSC
        objItemRomaneioGrade.sDescricao = objItemRomaneioGradeTela.sDescricao
        objItemRomaneioGrade.dQuantAFaturar = objItemRomaneioGradeTela.dQuantAFaturar
        objItemRomaneioGrade.dQuantFaturada = objItemRomaneioGradeTela.dQuantFaturada
        objItemRomaneioGrade.dQuantidade = objItemRomaneioGradeTela.dQuantidade
        objItemRomaneioGrade.dQuantReservada = objItemRomaneioGradeTela.dQuantReservada
        objItemRomaneioGrade.sUMEstoque = objItemRomaneioGradeTela.sUMEstoque
        objItemRomaneioGrade.dQuantCancelada = objItemRomaneioGradeTela.dQuantCancelada
                            
        colItensRomaneio.Add objItemRomaneioGrade
    Next

    Exit Function

End Function

Public Sub PISRetido_Change()

    iAlterado = REGISTRO_ALTERADO
    iPISRetidoAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub PISRetido_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValor As Double
Dim dValorTotal As Double

On Error GoTo Erro_PISRetido_Validate
    
    If iPISRetidoAlterado = 0 Then Exit Sub

    'Verifica se foi preenchido
    If Len(Trim(PISRetido.Text)) > 0 Then

        'Critica o Valor
        lErro = Valor_NaoNegativo_Critica(PISRetido.Text)
        If lErro <> SUCESSO Then gError 26654

        dValor = CDbl(PISRetido.Text)

        PISRetido.Text = Format(dValor, "Standard")

        If Len(Trim(ValorTotal.Caption)) > 0 Then dValorTotal = CDbl(ValorTotal.Caption)

        If dValor > dValorTotal Then gError 26655

    End If

    Call BotaoGravarTrib
    
    iPISRetidoAlterado = 0

    Exit Sub

Erro_PISRetido_Validate:

    Cancel = True

    Select Case gErr

        Case 26654

        Case 26655
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PIS_RETIDO_MAIOR_VALOR_TOTAL", gErr, dValor, dValorTotal)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177451)

    End Select

    Exit Sub

End Sub

Public Sub COFINSRetido_Change()

    iAlterado = REGISTRO_ALTERADO
    iCOFINSRetidoAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub COFINSRetido_Validate(Cancel As Boolean)
    
Dim lErro As Long
Dim dValor As Double
Dim dValorTotal As Double

On Error GoTo Erro_COFINSRetido_Validate
    
    If iCOFINSRetidoAlterado = 0 Then Exit Sub

    'Verifica se foi preenchido
    If Len(Trim(COFINSRetido.Text)) > 0 Then

        'Critica o Valor
        lErro = Valor_NaoNegativo_Critica(COFINSRetido.Text)
        If lErro <> SUCESSO Then gError 26654

        dValor = CDbl(COFINSRetido.Text)

        COFINSRetido.Text = Format(dValor, "Standard")

        If Len(Trim(ValorTotal.Caption)) > 0 Then dValorTotal = CDbl(ValorTotal.Caption)

        If dValor > dValorTotal Then gError 26655

    End If

    Call BotaoGravarTrib
    
    iCOFINSRetidoAlterado = 0

    Exit Sub

Erro_COFINSRetido_Validate:

    Cancel = True

    Select Case gErr

        Case 26654

        Case 26655
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COFINS_RETIDO_MAIOR_VALOR_TOTAL", gErr, dValor, dValorTotal)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177452)

    End Select

    Exit Sub

End Sub

Public Sub CSLLRetido_Change()

    iAlterado = REGISTRO_ALTERADO
    iCSLLRetidoAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub CSLLRetido_Validate(Cancel As Boolean)
    
Dim lErro As Long
Dim dValor As Double
Dim dValorTotal As Double

On Error GoTo Erro_CSLLRetido_Validate
    
    If iCSLLRetidoAlterado = 0 Then Exit Sub

    'Verifica se foi preenchido
    If Len(Trim(CSLLRetido.Text)) > 0 Then

        'Critica o Valor
        lErro = Valor_NaoNegativo_Critica(CSLLRetido.Text)
        If lErro <> SUCESSO Then gError 26654

        dValor = CDbl(CSLLRetido.Text)

        CSLLRetido.Text = Format(dValor, "Standard")

        If Len(Trim(ValorTotal.Caption)) > 0 Then dValorTotal = CDbl(ValorTotal.Caption)

        If dValor > dValorTotal Then gError 26655

    End If

    Call BotaoGravarTrib
    
    iCSLLRetidoAlterado = 0

    Exit Sub

Erro_CSLLRetido_Validate:

    Cancel = True

    Select Case gErr

        Case 26654

        Case 26655
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CSLL_RETIDO_MAIOR_VALOR_TOTAL", gErr, dValor, dValorTotal)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177453)

    End Select

    Exit Sub

End Sub



