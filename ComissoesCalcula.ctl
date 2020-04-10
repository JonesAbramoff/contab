VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ComissoesCalculaOcx 
   ClientHeight    =   5085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9780
   ScaleHeight     =   5085
   ScaleWidth      =   9780
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4575
      Index           =   2
      Left            =   -960
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   9240
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4560
         Index           =   1
         Left            =   6255
         TabIndex        =   1
         Top             =   3825
         Width           =   9165
         Begin VB.CheckBox FaturaIntegral 
            Caption         =   "Só libera o pedido integralmente"
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
            Left            =   225
            TabIndex        =   25
            Top             =   3795
            Width           =   3300
         End
         Begin VB.Frame Frame2 
            Caption         =   "Dados do Cliente"
            Height          =   900
            Index           =   6
            Left            =   210
            TabIndex        =   20
            Top             =   1605
            Width           =   8865
            Begin VB.ComboBox Filial 
               Height          =   315
               Left            =   5475
               TabIndex        =   21
               Top             =   345
               Width           =   2145
            End
            Begin MSMask.MaskEdBox Cliente 
               Height          =   300
               Left            =   1980
               TabIndex        =   22
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
               TabIndex        =   24
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
               TabIndex        =   23
               Top             =   405
               Width           =   660
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Preços"
            Height          =   900
            Index           =   2
            Left            =   210
            TabIndex        =   13
            Top             =   2610
            Width           =   8865
            Begin VB.ComboBox CondicaoPagamento 
               Height          =   315
               Left            =   4485
               Sorted          =   -1  'True
               TabIndex        =   15
               Top             =   345
               Width           =   1815
            End
            Begin VB.ComboBox TabelaPreco 
               Height          =   315
               Left            =   1320
               TabIndex        =   14
               Top             =   345
               Width           =   1875
            End
            Begin MSMask.MaskEdBox PercAcrescFin 
               Height          =   315
               Left            =   7995
               TabIndex        =   16
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
               TabIndex        =   19
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
               TabIndex        =   18
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
               TabIndex        =   17
               Top             =   405
               Width           =   1215
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Identificação"
            Height          =   1320
            Index           =   0
            Left            =   180
            TabIndex        =   2
            Top             =   195
            Width           =   8865
            Begin VB.CommandButton BotaoProxNum 
               Height          =   285
               Left            =   2940
               Picture         =   "ComissoesCalcula.ctx":0000
               Style           =   1  'Graphical
               TabIndex        =   4
               ToolTipText     =   "Numeração Automática"
               Top             =   330
               Width           =   300
            End
            Begin VB.ComboBox FilialFaturamento 
               Height          =   315
               ItemData        =   "ComissoesCalcula.ctx":00EA
               Left            =   5475
               List            =   "ComissoesCalcula.ctx":00EC
               TabIndex        =   3
               Top             =   780
               Width           =   2145
            End
            Begin MSMask.MaskEdBox NaturezaOp 
               Height          =   300
               Left            =   2115
               TabIndex        =   5
               Top             =   765
               Width           =   585
               _ExtentX        =   1032
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
            Begin MSMask.MaskEdBox Codigo 
               Height          =   300
               Left            =   2115
               TabIndex        =   6
               Top             =   315
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownEmissao 
               Height          =   300
               Left            =   6525
               TabIndex        =   7
               TabStop         =   0   'False
               Top             =   315
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataEmissao 
               Height          =   300
               Left            =   5475
               TabIndex        =   8
               Top             =   315
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
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
               Left            =   330
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   12
               Top             =   795
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
               Left            =   4665
               TabIndex        =   11
               Top             =   360
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
               Left            =   1335
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   10
               Top             =   345
               Width           =   720
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Filial Faturamento:"
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
               Index           =   17
               Left            =   3870
               TabIndex        =   9
               Top             =   795
               Width           =   1575
            End
         End
      End
      Begin VB.CommandButton BotaoEstoqueProd 
         Caption         =   "Estoque - Produto"
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
         Left            =   7260
         TabIndex        =   56
         Top             =   3990
         Width           =   1785
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
         Left            =   5550
         TabIndex        =   55
         Top             =   3990
         Width           =   1365
      End
      Begin VB.Frame Frame2 
         Caption         =   "Itens"
         Height          =   2685
         Index           =   3
         Left            =   180
         TabIndex        =   41
         Top             =   0
         Width           =   8865
         Begin VB.TextBox DescricaoProduto 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   3990
            MaxLength       =   50
            TabIndex        =   43
            Top             =   660
            Width           =   1305
         End
         Begin VB.ComboBox UnidadeMed 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "ComissoesCalcula.ctx":00EE
            Left            =   1575
            List            =   "ComissoesCalcula.ctx":00F0
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   240
            Width           =   720
         End
         Begin MSMask.MaskEdBox QuantCancelada 
            Height          =   225
            Left            =   7140
            TabIndex        =   44
            Top             =   360
            Width           =   1485
            _ExtentX        =   2619
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
         Begin MSMask.MaskEdBox QuantFaturada 
            Height          =   225
            Left            =   6975
            TabIndex        =   45
            Top             =   720
            Width           =   1500
            _ExtentX        =   2646
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantReservadaPV 
            Height          =   225
            Left            =   5400
            TabIndex        =   46
            Top             =   720
            Width           =   1500
            _ExtentX        =   2646
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataEntrega 
            Height          =   225
            Left            =   2640
            TabIndex        =   47
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
            Left            =   1410
            TabIndex        =   48
            Top             =   630
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
            Left            =   255
            TabIndex        =   49
            Top             =   660
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
            TabIndex        =   50
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
            TabIndex        =   51
            Top             =   300
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
         Begin MSMask.MaskEdBox Produto 
            Height          =   225
            Left            =   315
            TabIndex        =   52
            Top             =   330
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrecoTotal 
            Height          =   225
            Left            =   5670
            TabIndex        =   53
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
            Left            =   165
            TabIndex        =   54
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
         Begin MSMask.MaskEdBox PrecoBase 
            Height          =   225
            Left            =   0
            TabIndex        =   79
            Top             =   0
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
      End
      Begin VB.Frame Frame2 
         Caption         =   "Valores"
         Height          =   1290
         Index           =   4
         Left            =   180
         TabIndex        =   26
         Top             =   2640
         Width           =   8865
         Begin MSMask.MaskEdBox ValorFrete 
            Height          =   285
            Left            =   1695
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
         Begin MSMask.MaskEdBox ValorDesconto 
            Height          =   285
            Left            =   390
            TabIndex        =   28
            Top             =   420
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
            TabIndex        =   29
            Top             =   915
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
            TabIndex        =   30
            Top             =   915
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label IPIValor 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6030
            TabIndex        =   81
            Top             =   870
            Width           =   1125
         End
         Begin VB.Label ISSValor 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   390
            TabIndex        =   40
            Top             =   900
            Width           =   1125
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
            Left            =   510
            TabIndex        =   39
            Top             =   210
            Width           =   7695
         End
         Begin VB.Label ICMSSubstValor 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6000
            TabIndex        =   38
            Top             =   420
            Width           =   1125
         End
         Begin VB.Label ICMSSubstBase 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4320
            TabIndex        =   37
            Top             =   420
            Width           =   1500
         End
         Begin VB.Label ICMSValor 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3000
            TabIndex        =   36
            Top             =   420
            Width           =   1125
         End
         Begin VB.Label ICMSBase 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1695
            TabIndex        =   35
            Top             =   420
            Width           =   1125
         End
         Begin VB.Label ValorProdutos 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7305
            TabIndex        =   34
            Top             =   405
            Width           =   1125
         End
         Begin VB.Label IPIValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6000
            TabIndex        =   33
            Top             =   915
            Width           =   1125
         End
         Begin VB.Label ValorTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7320
            TabIndex        =   32
            Top             =   900
            Width           =   1125
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "ISS               Frete             Seguro              Despesas               IPI                Total"
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
            Left            =   825
            TabIndex        =   31
            Top             =   705
            Width           =   7230
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4605
      Index           =   5
      Left            =   480
      TabIndex        =   57
      Top             =   345
      Visible         =   0   'False
      Width           =   9180
      Begin VB.CheckBox ComissaoAutomatica 
         Caption         =   "Calcula comissão automaticamente"
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
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   525
         TabIndex        =   78
         Top             =   135
         Value           =   1  'Checked
         Width           =   3360
      End
      Begin VB.Frame SSFrame4 
         Caption         =   "Comissões"
         Height          =   4125
         Index           =   0
         Left            =   60
         TabIndex        =   58
         Top             =   390
         Width           =   9060
         Begin VB.CommandButton BotaoCalculaComissoes 
            Caption         =   "Calcula Comissoes"
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
            Height          =   645
            Left            =   7440
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   3300
            Width           =   1500
         End
         Begin VB.Frame SSFrame4 
            Caption         =   "Totais - Comissões"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   1
            Left            =   120
            TabIndex        =   61
            Top             =   3195
            Width           =   6975
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Percentual:"
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
               Index           =   19
               Left            =   2880
               TabIndex        =   67
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label1 
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
               Height          =   255
               Index           =   20
               Left            =   5040
               TabIndex        =   66
               Top             =   360
               Width           =   615
            End
            Begin VB.Label TotalPercentualComissao 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   3960
               TabIndex        =   65
               Top             =   345
               Width           =   825
            End
            Begin VB.Label TotalValorComissao 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   5640
               TabIndex        =   64
               Top             =   345
               Width           =   1155
            End
            Begin VB.Label Label1 
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
               Height          =   195
               Index           =   1
               Left            =   360
               TabIndex        =   63
               Top             =   360
               Width           =   1005
            End
            Begin VB.Label TotalValorBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   1440
               TabIndex        =   62
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.CommandButton BotaoVendedores 
            Caption         =   "Vendedores"
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
            Height          =   645
            Left            =   7440
            Picture         =   "ComissoesCalcula.ctx":00F2
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   3300
            Width           =   1500
         End
         Begin VB.ComboBox DiretoIndireto 
            Height          =   315
            Left            =   5070
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   915
            Width           =   1335
         End
         Begin MSMask.MaskEdBox Vendedor 
            Height          =   180
            Left            =   420
            TabIndex        =   69
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   318
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercentualBaixa 
            Height          =   180
            Left            =   7020
            TabIndex        =   70
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   318
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
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
         Begin MSFlexGridLib.MSFlexGrid GridComissoes 
            Height          =   1950
            Left            =   150
            TabIndex        =   71
            Top             =   360
            Width           =   8805
            _ExtentX        =   15531
            _ExtentY        =   3440
            _Version        =   393216
            Rows            =   11
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox ValorBaixa 
            Height          =   180
            Left            =   7875
            TabIndex        =   72
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   318
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
         Begin MSMask.MaskEdBox ValorComissao 
            Height          =   180
            Left            =   3825
            TabIndex        =   73
            Top             =   165
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   318
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
         Begin MSMask.MaskEdBox ValorBase 
            Height          =   180
            Left            =   2700
            TabIndex        =   74
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   318
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
         Begin MSMask.MaskEdBox PercentualComissao 
            Height          =   180
            Left            =   1815
            TabIndex        =   75
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   318
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
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
         Begin MSMask.MaskEdBox ValorEmissao 
            Height          =   180
            Left            =   5880
            TabIndex        =   76
            Top             =   180
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   318
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
         Begin MSMask.MaskEdBox PercentualEmissao 
            Height          =   180
            Left            =   5025
            TabIndex        =   77
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   318
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
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
      End
   End
   Begin MSMask.MaskEdBox VendedorCupom 
      Height          =   315
      Left            =   4440
      TabIndex        =   80
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   5
      Mask            =   "#####"
      PromptChar      =   " "
   End
   Begin VB.Label ICMSSubstValor1 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5235
      TabIndex        =   82
      Top             =   3675
      Width           =   1125
   End
End
Attribute VB_Name = "ComissoesCalculaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Public colComissoes As Collection
Public gbCarregandoTela As Boolean
Public gbLimpandoTela As Boolean

Event Unload()
Dim m_Caption As String

Dim iAlterado As Integer

'inicia objeto associado a GridComissoes
Dim objTabComissoes As New ClassTabComissoes

' ****************** COMISSOES - 02/04/02 ******************
Public objComissoesRegrasCalc As Object 'Declara a classe que executará o cálculo das comissões
Public objMnemonicoComissCalc As ClassMnemonicoComissCalc 'Declara a classe que executará o cálculo dos mnemônicos de comissões
Public objMnemonicoComissCalcAux As ClassMnemonicoComissCalcAux 'Declara a classe que executará o cálculo dos mnemônicos de comissões customizados para o cliente
Public iComissoesAlterada As Integer 'Indica que foi alterado pelo menos um campo na tela que seja utilizado para calcular as comissões '*** 11/04/02 - Luiz G.F.Nogueira ***
'***********************************************************

Public objGridComissoes As AdmGrid

Public objGridItens As AdmGrid

'Grid Itens
Public iGrid_Item_Col As Integer
Public iGrid_ProdutoAlmox_Col As Integer
Public iGrid_UMEstoque_Col As Integer
Public iGrid_Almoxarifado_Col As Integer
Public iGrid_QuantReservar_Col As Integer
Public iGrid_QuantReserv_Col As Integer
Public iGrid_Validade_Col As Integer
Public iGrid_Responsavel_Col As Integer
Public iGrid_ItemProduto_Col As Integer
Public iGrid_Produto_Col As Integer
Public iGrid_DescProduto_Col As Integer
Public iGrid_UnidadeMed_Col As Integer
Public iGrid_QuantPedida_Col As Integer
Public iGrid_QuantCancel_Col As Integer
Public iGrid_ValorUnitario_Col As Integer

'4 - Marcio - 08/2000 - incluido no GridItens a coluna preço base
Public iGrid_PrecoBase_Col As Integer

Public iGrid_PercDesc_Col As Integer
Public iGrid_Desconto_Col As Integer
Public iGrid_PrecoTotal_Col As Integer
Public iGrid_DataEntrega_Col As Integer
Public iGrid_QuantReservada_Col As Integer
Public iGrid_QuantFaturada_Col As Integer

Private iGrid_Vendedor_Col As Integer
Private iGrid_PercentualComissao_Col As Integer
Private iGrid_ValorBase_Col As Integer
Private iGrid_ValorComissao_Col As Integer
Private iGrid_PercentualEmissao_Col As Integer
Private iGrid_ValorEmissao_Col As Integer
Private iGrid_PercentualBaixa_Col As Integer
Private iGrid_ValorBaixa_Col As Integer
Private iGrid_DiretoIndireto_Col As Integer

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True

    Set objGridItens = New AdmGrid
    Set objGridComissoes = New AdmGrid

    'Faz as Inicializações dos Grids
    lErro = Inicializa_Grid_Itens(objGridItens)
    If lErro <> SUCESSO Then gError 26493
    
    'se a empresa utiliza regras para calculo de comissoes
    If gobjCRFAT.iUsaComissoesRegras = USA_REGRAS Then
        
        'instancia a classe q executa as regras
        Set objComissoesRegrasCalc = CreateObject("RotinasContab.ClassComissoesRegrasCalc")
        
        'instancia a classe q calcula os mnemonicos
        Set objMnemonicoComissCalc = New ClassMnemonicoComissCalc
        Set objMnemonicoComissCalcAux = New ClassMnemonicoComissCalcAux
        
        'Instancia objTela das classes que calculam os mnemônicos
        Set objMnemonicoComissCalc.objTela = Me
        Set objMnemonicoComissCalcAux.objTela = Me
    
    End If

    Set objTabComissoes.objTela = Me

    lErro = objTabComissoes.Inicializa_Grid_Comissoes(objGridComissoes)
    If lErro <> SUCESSO Then gError 26495
    
    'alterado por cyntia
    objGridComissoes.iLinhasVisiveis = 5
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridComissoes)
    
    iComissoesAlterada = 0
    
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 26636
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154376)

    End Select

    Exit Sub

End Sub

'******************************************
'4 eventos do controle do Grid de Comissoes: DiretoIndireto
'******************************************

Public Sub DiretoIndireto_Change()

    '*** 11/04/02 - Luiz G.F.Nogueira ***
    'Desmarca o cálculo automático de comissões
    ComissaoAutomatica.Value = vbUnchecked
    '************************************

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DiretoIndireto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Public Sub DiretoIndireto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Public Sub DiretoIndireto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = DiretoIndireto
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objGridItens = Nothing
    Set objGridComissoes = Nothing
    
    'se foi instanciada a classe de execucao de regras de comissoes => libera
    Set objComissoesRegrasCalc = Nothing
        
    'se foi instanciada a classe que calcula os mnemonicos => libera
    Set objMnemonicoComissCalc = Nothing
    Set objMnemonicoComissCalcAux = Nothing
    Set objTabComissoes = Nothing

End Sub

Private Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quant. Pedida")
    objGridInt.colColuna.Add ("Quant. Cancelada")
    objGridInt.colColuna.Add ("Preço Base")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Total")
    objGridInt.colColuna.Add ("Data Entrega")
    objGridInt.colColuna.Add ("Quant Reservada")
    objGridInt.colColuna.Add ("Quant Faturada")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoProduto.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (QuantCancelada.Name)
    
    '4 - Marcio - 08/2000 - incluido no GridItens a coluna preço base
    objGridInt.colCampo.Add (PrecoBase.Name)
    
    objGridInt.colCampo.Add (PrecoUnitario.Name)
    objGridInt.colCampo.Add (PercentDesc.Name)
    objGridInt.colCampo.Add (Desconto.Name)
    objGridInt.colCampo.Add (PrecoTotal.Name)
    objGridInt.colCampo.Add (DataEntrega.Name)
    objGridInt.colCampo.Add (QuantReservadaPV.Name)
    objGridInt.colCampo.Add (QuantFaturada.Name)

    'Colunas do Grid
    iGrid_ItemProduto_Col = 0
    iGrid_Produto_Col = 1
    iGrid_DescProduto_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_QuantPedida_Col = 4
    iGrid_QuantCancel_Col = 5
        
    '4 - Marcio - 08/2000 - incluido no GridItens a coluna preço base
    iGrid_PrecoBase_Col = 6
    
    iGrid_ValorUnitario_Col = 7
    iGrid_PercDesc_Col = 8
    iGrid_Desconto_Col = 9
    iGrid_PrecoTotal_Col = 10
    iGrid_DataEntrega_Col = 11
    iGrid_QuantReservada_Col = 12
    iGrid_QuantFaturada_Col = 13
        
    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 2 + 1 '??? alterei

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 1

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_NAO_EXECUTAR_ROTINA_ENABLE '??? alterei

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Function Trata_Parametros(ByVal iFilialFaturamento As Integer, ByVal sNomeRedCliente As String, ByVal iFilialCli As Integer, ByVal sProduto As String, ByVal dQuantidade As Double, ByVal sUM As String, ByVal dPrecoUnitario As Double, colComissoesAux As Collection, Optional dValorDesconto As Double = 0, Optional dPercentualDesc As Double = 0, Optional iCodVendedor As Integer = 0, Optional sOrigem As String) As Long

Dim lErro As Long, iTipoTela As Integer
Dim sProdutoEnxuto As String, objComissao As ClassComissaoPedVendas
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim objVendedor As New ClassVendedor
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_Trata_Parametros

    'Alterado por Luiz Nogueira em 19/04/04
    'se a chamada da tela foi feita pelo loja, muda o nome pra indicar q está calculando comissoes do loja
    If sOrigem = MODULO_LOJA Then Caption = "ComissoesCalculaLoja"
    
    'limpar os grids de itens e de comissoes
    Call Grid_Limpa(objGridItens)
    Call Grid_Limpa(objGridComissoes)
    
    'carregar os campos de cliente, filial de cliente e filialfaturamento
    Cliente.Text = sNomeRedCliente
    Filial.Text = CStr(iFilialCli)
    objFilialCliente.iCodFilial = iFilialCli
    FilialFaturamento.Text = CStr(iFilialFaturamento)
    
    'coloca na tela o código do vendedor
    VendedorCupom.PromptInclude = False
    VendedorCupom.Text = iCodVendedor
    VendedorCupom.PromptInclude = True
    
    'preencher valor total (vou deixar em branco frete, seguro, ipi e despesas)
    ValorTotal.Caption = CStr(Round((dQuantidade * dPrecoUnitario) - dValorDesconto, 2))
    
    'carregar uma linha no grid de itens
    lErro = Mascara_RetornaProdutoEnxuto(sProduto, sProdutoEnxuto)
    If lErro <> SUCESSO Then gError 26649

    'Mascara o produto enxuto
    Produto.PromptInclude = False
    Produto.Text = sProdutoEnxuto
    Produto.PromptInclude = True

    'Coloca os dados dos itens na tela
    GridItens.TextMatrix(1, iGrid_Produto_Col) = Produto.Text
    
    GridItens.TextMatrix(1, iGrid_UnidadeMed_Col) = sUM
    GridItens.TextMatrix(1, iGrid_QuantPedida_Col) = Formata_Estoque(dQuantidade)
    GridItens.TextMatrix(1, iGrid_QuantCancel_Col) = Formata_Estoque(0)
    GridItens.TextMatrix(1, iGrid_ValorUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
            
    '4 - Marcio - 08/2000 - incluido no GridItens a coluna preço base
    GridItens.TextMatrix(1, iGrid_PrecoBase_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
    
    GridItens.TextMatrix(1, iGrid_PercDesc_Col) = Format(dPercentualDesc, "Percent")
    GridItens.TextMatrix(1, iGrid_Desconto_Col) = Format(dValorDesconto, "Standard")
    GridItens.TextMatrix(1, iGrid_PrecoTotal_Col) = Format(ValorTotal.Caption, "Standard")
    GridItens.TextMatrix(1, iGrid_QuantReservada_Col) = Formata_Estoque(0)
    GridItens.TextMatrix(1, iGrid_QuantFaturada_Col) = Formata_Estoque(0)
    
    objGridItens.iLinhasExistentes = 1
    iComissoesAlterada = REGISTRO_ALTERADO
    
    'se a empresa nao utiliza as regras p/ calculo de comissoes
    If Not gobjCRFAT.iUsaComissoesRegras = USA_REGRAS Then
    
        'significa que está carregando cupons oriundos do loja
        If iCodVendedor <> 0 Then
    
            objVendedor.iCodigo = iCodVendedor
    
            'Lê o Vendedor
            lErro = CF("Vendedor_Le", objVendedor)
            If lErro <> SUCESSO And lErro <> 12582 Then gError 126302
    
            'Se não achou o nome do Vendedor --> erro
            If lErro <> SUCESSO Then gError 126303
        
            GridComissoes.TextMatrix(1, 1) = objVendedor.sNomeReduzido
    
            GridComissoes.TextMatrix(1, 2) = Format(objVendedor.dPercComissao, "Percent")
            
            'percentual de comissao na emissao que no loja vai ser sempre 100%
            GridComissoes.TextMatrix(1, 5) = Format(1, "Percent")
        
            objGridComissoes.iLinhasExistentes = 1
    
        Else
        
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Trim(Cliente.Text), objFilialCliente)
            If lErro <> SUCESSO And lErro <> 17660 Then gError 126306
            
            If lErro = 17660 Then gError 126307
        
            'calcular as comissoes no modelo antigo...
            lErro = objTabComissoes.Comissao_Automatica_FilialCli_Exibe(objFilialCliente)
            If lErro <> SUCESSO Then gError 126308
        
        End If
    
    End If
    
    lErro = objTabComissoes.Comissoes_Calcula
    If lErro <> SUCESSO Then gError 126304
    
    'move as comissoes do grid p/a colecao
    iTipoTela = PEDIDO_DE_VENDA
    lErro = objTabComissoes.Move_TabComissoes_Memoria(objPedidoVenda, iTipoTela)
    If lErro <> SUCESSO Then gError 126305
        
    For Each objComissao In objPedidoVenda.colComissoes
        colComissoesAux.Add objComissao
    Next
    
    Trata_Parametros = SUCESSO
     
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = gErr
     
    Select Case gErr
          
        Case 126302, 126304, 126305, 126306, 126308
        
        Case 126303
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", gErr, objVendedor.iCodigo)
          
        Case 126307
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA1", gErr, Cliente.Text, objFilialCliente.iCodFilial)
          
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154377)
     
    End Select
     
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Cálculo Automático de Comissões"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ComissoesCalcula"
    
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

Private Sub IPIValor1_change()
    IPIValor.Caption = IPIValor1.Caption
End Sub

Private Sub ICMSSubstValor1_change()
    ICMSSubstValor.Caption = ICMSSubstValor1.Caption
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

'***** fim do trecho a ser copiado ******

Public Sub Fecha_Tela()
    Unload Me
End Sub
