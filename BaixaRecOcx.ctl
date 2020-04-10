VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl BaixaRecOcx 
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   KeyPreview      =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   16995
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8235
      Index           =   1
      Left            =   90
      TabIndex        =   21
      Top             =   750
      Width           =   16620
      Begin VB.Frame Frame9 
         Caption         =   "Filtros"
         Height          =   6105
         Left            =   30
         TabIndex        =   80
         Top             =   1005
         Width           =   10590
         Begin VB.Frame Frame12 
            Caption         =   "Produtos (Exibir somente documentos ligados a NFs que contém um dos produtos filtrados abaixo)"
            Height          =   1230
            Left            =   240
            TabIndex        =   161
            Top             =   4290
            Width           =   10215
            Begin MSMask.MaskEdBox ProdutoInicial 
               Height          =   315
               Left            =   615
               TabIndex        =   19
               Top             =   300
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   556
               _Version        =   393216
               AllowPrompt     =   -1  'True
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ProdutoFinal 
               Height          =   315
               Left            =   615
               TabIndex        =   20
               Top             =   765
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   556
               _Version        =   393216
               AllowPrompt     =   -1  'True
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin VB.Label DescProdFim 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   2160
               TabIndex        =   165
               Top             =   765
               Width           =   7800
            End
            Begin VB.Label DescProdInic 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   2160
               TabIndex        =   164
               Top             =   300
               Width           =   7770
            End
            Begin VB.Label LabelProdutoDe 
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
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   225
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   163
               Top             =   330
               Width           =   360
            End
            Begin VB.Label LabelProdutoAte 
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
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   195
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   162
               Top             =   810
               Width           =   435
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Cobrança Bancária"
            Height          =   735
            Left            =   5535
            TabIndex        =   158
            Top             =   3315
            Width           =   4920
            Begin VB.ComboBox CobradorFiltro 
               Height          =   315
               Left            =   1455
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   255
               Width           =   3195
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Cobrador:"
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
               Left            =   570
               TabIndex        =   160
               Top             =   285
               Width           =   840
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Vendedor"
            Height          =   720
            Left            =   225
            TabIndex        =   155
            Top             =   3315
            Width           =   4785
            Begin MSMask.MaskEdBox Vendedor 
               Height          =   315
               Left            =   1200
               TabIndex        =   17
               Top             =   240
               Width           =   3510
               _ExtentX        =   6191
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
               Left            =   210
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   156
               Top             =   300
               Width           =   885
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Outros"
            Height          =   1065
            Left            =   5535
            TabIndex        =   152
            Top             =   1980
            Width           =   4920
            Begin VB.ComboBox FormaPagto 
               Height          =   315
               Left            =   1455
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   660
               Width           =   3165
            End
            Begin MSMask.MaskEdBox CodigoViagem 
               Height          =   315
               Left            =   1440
               TabIndex        =   15
               Top             =   240
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   6
               Mask            =   "999999"
               PromptChar      =   " "
            End
            Begin VB.Label LabelFormaPagto 
               Caption         =   "Forma Pagto: "
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
               Left            =   315
               TabIndex        =   154
               Top             =   705
               Width           =   1995
            End
            Begin VB.Label LabelViagem 
               Caption         =   "Viagem:"
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
               Left            =   630
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   153
               Top             =   300
               Width           =   750
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Tipo de Documento"
            Height          =   1065
            Left            =   225
            TabIndex        =   142
            Top             =   1980
            Width           =   4785
            Begin VB.ComboBox TipoDocSeleciona 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "BaixaRecOcx.ctx":0000
               Left            =   1200
               List            =   "BaixaRecOcx.ctx":0002
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   615
               Width           =   3510
            End
            Begin VB.OptionButton TipoDocTodos 
               Caption         =   "Todos"
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
               TabIndex        =   12
               Top             =   330
               Value           =   -1  'True
               Width           =   1005
            End
            Begin VB.OptionButton TipoDocApenas 
               Caption         =   "Apenas:"
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
               TabIndex        =   13
               Top             =   645
               Width           =   1050
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Nº do Título"
            Height          =   1410
            Left            =   7935
            TabIndex        =   87
            Top             =   345
            Width           =   2520
            Begin MSMask.MaskEdBox TituloInic 
               Height          =   300
               Left            =   1185
               TabIndex        =   10
               Top             =   435
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Mask            =   "99999999"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox TituloFim 
               Height          =   300
               Left            =   1185
               TabIndex        =   11
               Top             =   960
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Mask            =   "99999999"
               PromptChar      =   " "
            End
            Begin VB.Label Label22 
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
               Height          =   255
               Left            =   765
               TabIndex        =   98
               Top             =   990
               Width           =   375
            End
            Begin VB.Label Label21 
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
               Height          =   255
               Left            =   810
               TabIndex        =   99
               Top             =   480
               Width           =   375
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Data de Vencimento"
            Height          =   1410
            Left            =   4005
            TabIndex        =   84
            Top             =   330
            Width           =   3165
            Begin MSComCtl2.UpDown UpDownVencInic 
               Height          =   300
               Left            =   1680
               TabIndex        =   7
               TabStop         =   0   'False
               Top             =   480
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox VencInic 
               Height          =   300
               Left            =   615
               TabIndex        =   6
               Top             =   465
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownVencFim 
               Height          =   300
               Left            =   1680
               TabIndex        =   9
               TabStop         =   0   'False
               Top             =   990
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox VencFim 
               Height          =   300
               Left            =   600
               TabIndex        =   8
               Top             =   990
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label Label20 
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
               Height          =   255
               Left            =   195
               TabIndex        =   100
               Top             =   1020
               Width           =   375
            End
            Begin VB.Label Label17 
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
               Height          =   255
               Left            =   240
               TabIndex        =   101
               Top             =   510
               Width           =   375
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Data de Emissão"
            Height          =   1410
            Left            =   225
            TabIndex        =   81
            Top             =   345
            Width           =   3165
            Begin MSComCtl2.UpDown UpDownEmissaoInic 
               Height          =   300
               Left            =   1755
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   450
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox EmissaoInic 
               Height          =   300
               Left            =   690
               TabIndex        =   2
               Top             =   450
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownEmissaoFim 
               Height          =   300
               Left            =   1755
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   960
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox EmissaoFim 
               Height          =   300
               Left            =   705
               TabIndex        =   4
               Top             =   960
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label Label16 
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
               Height          =   255
               Left            =   270
               TabIndex        =   102
               Top             =   990
               Width           =   375
            End
            Begin VB.Label Label11 
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
               Height          =   255
               Left            =   315
               TabIndex        =   103
               Top             =   480
               Width           =   375
            End
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Cliente"
         Height          =   750
         Left            =   30
         TabIndex        =   77
         Top             =   150
         Width           =   10590
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   6780
            TabIndex        =   1
            Top             =   270
            Width           =   3390
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   300
            Left            =   930
            TabIndex        =   0
            Top             =   270
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin VB.Label LabelCli 
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
            Height          =   255
            Left            =   135
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   104
            Top             =   315
            Width           =   675
         End
         Begin VB.Label LabelFilial 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6165
            TabIndex        =   105
            Top             =   315
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8145
      Index           =   3
      Left            =   195
      TabIndex        =   51
      Top             =   840
      Visible         =   0   'False
      Width           =   16635
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4920
         TabIndex        =   141
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
         Left            =   6675
         TabIndex        =   57
         Top             =   315
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
         Height          =   270
         Left            =   6675
         TabIndex        =   55
         Top             =   0
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6690
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   840
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
         Height          =   270
         Left            =   8130
         TabIndex        =   56
         Top             =   0
         Width           =   1245
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4800
         TabIndex        =   65
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
         TabIndex        =   67
         Top             =   2565
         Width           =   870
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   66
         Top             =   2175
         Width           =   1770
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2985
         Left            =   9795
         TabIndex        =   69
         Top             =   1515
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1695
         Left            =   195
         TabIndex        =   92
         Top             =   3450
         Width           =   6045
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
            TabIndex        =   115
            Top             =   975
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
            Left            =   1125
            TabIndex        =   116
            Top             =   300
            Width           =   570
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Left            =   1845
            TabIndex        =   117
            Top             =   285
            Width           =   3975
         End
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Left            =   1845
            TabIndex        =   118
            Top             =   960
            Visible         =   0   'False
            Width           =   3975
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
         Left            =   3450
         TabIndex        =   60
         Top             =   960
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin MSMask.MaskEdBox CTBConta 
         Height          =   225
         Left            =   525
         TabIndex        =   61
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
         TabIndex        =   64
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
         Left            =   2280
         TabIndex        =   63
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
         TabIndex        =   62
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
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   525
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil 
         Height          =   300
         Left            =   570
         TabIndex        =   54
         Top             =   525
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
         TabIndex        =   53
         Top             =   135
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
         Left            =   3810
         TabIndex        =   52
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
         Left            =   0
         TabIndex        =   68
         Top             =   1185
         Width           =   9420
         _ExtentX        =   16616
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
         Left            =   9795
         TabIndex        =   70
         Top             =   1515
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
         Left            =   9780
         TabIndex        =   71
         Top             =   1515
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
         Left            =   6690
         TabIndex        =   58
         Top             =   630
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
         TabIndex        =   119
         Top             =   165
         Width           =   720
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   120
         Top             =   120
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
         TabIndex        =   121
         Top             =   600
         Width           =   735
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   122
         Top             =   570
         Width           =   1185
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   123
         Top             =   555
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
         TabIndex        =   124
         Top             =   585
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
         TabIndex        =   125
         Top             =   945
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
         Left            =   9810
         TabIndex        =   126
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
         Left            =   9810
         TabIndex        =   127
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
         Left            =   9810
         TabIndex        =   128
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
         TabIndex        =   129
         Top             =   3045
         Width           =   615
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   130
         Top             =   3030
         Width           =   1155
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   131
         Top             =   3030
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
         TabIndex        =   132
         Top             =   555
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
         TabIndex        =   133
         Top             =   165
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
         TabIndex        =   134
         Top             =   165
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8250
      Index           =   2
      Left            =   30
      TabIndex        =   22
      Top             =   735
      Visible         =   0   'False
      Width           =   16725
      Begin VB.Frame FrameRecebimento 
         Caption         =   "Dados da Perda"
         Height          =   1590
         Index           =   3
         Left            =   75
         TabIndex        =   82
         Top             =   6660
         Visible         =   0   'False
         Width           =   16590
         Begin MSMask.MaskEdBox HistoricoPerda 
            Height          =   300
            Left            =   1620
            TabIndex        =   50
            Top             =   645
            Width           =   12000
            _ExtentX        =   21167
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
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
            Left            =   705
            TabIndex        =   110
            Top             =   690
            Width           =   810
         End
      End
      Begin VB.Frame FrameRecebimento 
         Caption         =   "Adiantamentos de Clientes"
         Height          =   1620
         Index           =   1
         Left            =   75
         TabIndex        =   85
         Top             =   6630
         Visible         =   0   'False
         Width           =   16590
         Begin MSMask.MaskEdBox Doc 
            Height          =   225
            Left            =   3960
            TabIndex        =   138
            Top             =   645
            Width           =   1005
            _ExtentX        =   1773
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
         Begin MSMask.MaskEdBox Seq 
            Height          =   225
            Left            =   5325
            TabIndex        =   139
            Top             =   735
            Width           =   1005
            _ExtentX        =   1773
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
         Begin MSMask.MaskEdBox Hist 
            Height          =   225
            Left            =   3705
            TabIndex        =   140
            Top             =   960
            Width           =   2925
            _ExtentX        =   5159
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
         Begin VB.CheckBox SelecionarRA 
            Height          =   225
            Left            =   7890
            TabIndex        =   48
            Top             =   345
            Width           =   510
         End
         Begin MSMask.MaskEdBox ContaCorrenteRA 
            Height          =   225
            Left            =   1635
            TabIndex        =   43
            Top             =   285
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   2
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MeioPagtoRA 
            Height          =   225
            Left            =   2820
            TabIndex        =   44
            Top             =   330
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   2
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox DataMovimentoRA 
            Height          =   225
            Left            =   285
            TabIndex        =   42
            Top             =   300
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox ValorRA 
            Height          =   225
            Left            =   4350
            TabIndex        =   45
            Top             =   315
            Width           =   1005
            _ExtentX        =   1773
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
         Begin MSMask.MaskEdBox SaldoRA 
            Height          =   225
            Left            =   5415
            TabIndex        =   46
            Top             =   300
            Width           =   1005
            _ExtentX        =   1773
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
         Begin MSMask.MaskEdBox FilialRA 
            Height          =   225
            Left            =   6465
            TabIndex        =   47
            Top             =   315
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
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
            Mask            =   "9999"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridRecebAntecipados 
            Height          =   1275
            Left            =   45
            TabIndex        =   49
            Top             =   210
            Width           =   16470
            _ExtentX        =   29051
            _ExtentY        =   2249
            _Version        =   393216
            Rows            =   5
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.Frame FrameRecebimento 
         Caption         =   "Dados do Recebimento"
         Height          =   1620
         Index           =   0
         Left            =   75
         TabIndex        =   97
         Top             =   6630
         Width           =   16590
         Begin VB.Frame FrameTipoMeioPagto 
            BorderStyle     =   0  'None
            Height          =   585
            Index           =   0
            Left            =   6750
            TabIndex        =   78
            Top             =   240
            Width           =   1770
         End
         Begin VB.Frame FrameTipoMeioPagto 
            BorderStyle     =   0  'None
            Height          =   585
            Index           =   1
            Left            =   6780
            TabIndex        =   79
            Top             =   240
            Width           =   1770
         End
         Begin VB.ComboBox ContaCorrente 
            Height          =   315
            Left            =   1125
            Sorted          =   -1  'True
            TabIndex        =   40
            Top             =   315
            Width           =   1695
         End
         Begin MSMask.MaskEdBox Historico 
            Height          =   300
            Left            =   1125
            TabIndex        =   41
            Top             =   1200
            Width           =   12000
            _ExtentX        =   21167
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin VB.Label Label3 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   465
            TabIndex        =   106
            Top             =   345
            Width           =   570
         End
         Begin VB.Label Label13 
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
            Left            =   210
            TabIndex        =   107
            Top             =   1230
            Width           =   825
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
            Left            =   525
            TabIndex        =   108
            Top             =   795
            Width           =   510
         End
         Begin VB.Label ValorReceber 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1125
            TabIndex        =   109
            Top             =   765
            Width           =   1680
         End
      End
      Begin VB.Frame FrameRecebimento 
         Caption         =   "Débitos"
         Height          =   1620
         Index           =   2
         Left            =   75
         TabIndex        =   83
         Top             =   6630
         Visible         =   0   'False
         Width           =   16575
         Begin VB.CheckBox SelecionarDB 
            Height          =   225
            Left            =   6555
            TabIndex        =   146
            Top             =   180
            Width           =   390
         End
         Begin VB.TextBox OBSDB 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   4200
            TabIndex        =   144
            Top             =   480
            Width           =   2070
         End
         Begin MSMask.MaskEdBox SaldoDB 
            Height          =   225
            Left            =   4170
            TabIndex        =   143
            Top             =   300
            Width           =   825
            _ExtentX        =   1455
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
         Begin MSMask.MaskEdBox DataEmissaoDB 
            Height          =   240
            Left            =   0
            TabIndex        =   145
            Top             =   30
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox ValorDB 
            Height          =   225
            Left            =   3000
            TabIndex        =   147
            Top             =   105
            Width           =   825
            _ExtentX        =   1455
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
         Begin MSMask.MaskEdBox NumTituloDB 
            Height          =   225
            Left            =   2250
            TabIndex        =   148
            Top             =   90
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   397
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
            Mask            =   "999999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TipoDocDB 
            Height          =   225
            Left            =   1560
            TabIndex        =   149
            Top             =   90
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   50
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox FilialDB 
            Height          =   225
            Left            =   5265
            TabIndex        =   150
            Top             =   0
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
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
            Mask            =   "9999"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridDebitos 
            Height          =   1320
            Left            =   30
            TabIndex        =   151
            Top             =   225
            Width           =   16485
            _ExtentX        =   29078
            _ExtentY        =   2328
            _Version        =   393216
            Rows            =   5
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.Frame Cobranca 
         Caption         =   "Parcelas em Aberto"
         Height          =   5685
         Left            =   75
         TabIndex        =   96
         Top             =   405
         Width           =   16590
         Begin MSMask.MaskEdBox NossoNumero 
            Height          =   225
            Left            =   4230
            TabIndex        =   159
            Top             =   1740
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
            Mask            =   "9999999999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Emissao 
            Height          =   225
            Left            =   4140
            TabIndex        =   157
            Top             =   1155
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox FilialEmpresa 
            Height          =   225
            Left            =   6075
            TabIndex        =   89
            Top             =   1035
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
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
            Mask            =   "9999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialClienteItem 
            Height          =   225
            Left            =   4725
            TabIndex        =   135
            Top             =   705
            Width           =   1245
            _ExtentX        =   2196
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
         Begin MSMask.MaskEdBox ClienteItem 
            Height          =   225
            Left            =   3120
            TabIndex        =   136
            Top             =   585
            Width           =   2805
            _ExtentX        =   4948
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
         Begin VB.CheckBox Selecionar 
            Height          =   225
            Left            =   7365
            TabIndex        =   32
            Top             =   375
            Width           =   525
         End
         Begin MSMask.MaskEdBox DataVencimento 
            Height          =   225
            Left            =   870
            TabIndex        =   27
            Top             =   705
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Saldo 
            Height          =   225
            Left            =   3210
            TabIndex        =   29
            Top             =   1020
            Width           =   1005
            _ExtentX        =   1773
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
         Begin MSMask.MaskEdBox ValorBaixar 
            Height          =   225
            Left            =   4395
            TabIndex        =   30
            Top             =   1035
            Width           =   1005
            _ExtentX        =   1773
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
         Begin MSMask.MaskEdBox Numero 
            Height          =   225
            Left            =   2010
            TabIndex        =   26
            Top             =   450
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
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
            Mask            =   "99999999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Tipo 
            Height          =   225
            Left            =   3060
            TabIndex        =   86
            Top             =   1110
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Parcela 
            Height          =   225
            Left            =   2115
            TabIndex        =   28
            Top             =   915
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
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
            Mask            =   "99"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDesconto 
            Height          =   225
            Left            =   4725
            TabIndex        =   34
            Top             =   345
            Width           =   855
            _ExtentX        =   1508
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
         Begin MSMask.MaskEdBox ValorMulta 
            Height          =   225
            Left            =   6195
            TabIndex        =   31
            Top             =   405
            Width           =   855
            _ExtentX        =   1508
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
         Begin MSMask.MaskEdBox ValorJuros 
            Height          =   225
            Left            =   7365
            TabIndex        =   33
            Top             =   720
            Width           =   855
            _ExtentX        =   1508
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
         Begin MSMask.MaskEdBox Cobrador 
            Height          =   225
            Left            =   1365
            TabIndex        =   25
            Top             =   1635
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
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
            Mask            =   "9999"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridParcelas 
            Height          =   1230
            Left            =   45
            TabIndex        =   35
            Top             =   195
            Width           =   16500
            _ExtentX        =   29104
            _ExtentY        =   2170
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox ValorParcela 
            Height          =   225
            Left            =   4500
            TabIndex        =   88
            Top             =   675
            Width           =   1260
            _ExtentX        =   2223
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
         Begin MSMask.MaskEdBox ValorAReceber 
            Height          =   225
            Left            =   3120
            TabIndex        =   90
            Top             =   705
            Width           =   1275
            _ExtentX        =   2249
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
      End
      Begin VB.Frame Frame10 
         Caption         =   "Tipo de Baixa"
         Height          =   495
         Left            =   75
         TabIndex        =   76
         Top             =   6120
         Width           =   16590
         Begin VB.OptionButton Recebimento 
            Caption         =   "Débito / Devolução"
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
            Left            =   4795
            TabIndex        =   38
            Top             =   210
            Width           =   2055
         End
         Begin VB.OptionButton Recebimento 
            Caption         =   "Adiantamento"
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
            Left            =   2660
            TabIndex        =   37
            Top             =   210
            Width           =   1575
         End
         Begin VB.OptionButton Recebimento 
            Caption         =   "Recebimento"
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
            Left            =   630
            TabIndex        =   36
            Top             =   210
            Value           =   -1  'True
            Width           =   1470
         End
         Begin VB.OptionButton Recebimento 
            Caption         =   "Perda"
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
            Index           =   3
            Left            =   7410
            TabIndex        =   39
            Top             =   210
            Width           =   825
         End
      End
      Begin MSComCtl2.UpDown UpDownDataBaixa 
         Height          =   300
         Left            =   2550
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   90
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataBaixa 
         Height          =   300
         Left            =   1485
         TabIndex        =   23
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataCredito 
         Height          =   300
         Left            =   5415
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   90
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataCredito 
         Height          =   300
         Left            =   4335
         TabIndex        =   24
         Top             =   90
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Data da Baixa:"
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
         Left            =   135
         TabIndex        =   111
         Top             =   150
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Data Crédito:"
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
         Left            =   3120
         TabIndex        =   112
         Top             =   150
         Width           =   1140
      End
      Begin VB.Label TotalBaixar 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6765
         TabIndex        =   113
         Top             =   120
         Width           =   1560
      End
      Begin VB.Label Label5 
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
         Left            =   6180
         TabIndex        =   114
         Top             =   150
         Width           =   510
      End
   End
   Begin VB.CheckBox ImprimirRecibo 
      Caption         =   "Imprimir Recibo ao Gravar"
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
      Left            =   12165
      TabIndex        =   137
      Top             =   210
      Width           =   2880
   End
   Begin VB.PictureBox Picture4 
      Height          =   555
      Left            =   15195
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   90
      Width           =   1695
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "BaixaRecOcx.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "BaixaRecOcx.ctx":0182
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "BaixaRecOcx.ctx":06B4
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   8700
      Left            =   0
      TabIndex        =   95
      Top             =   420
      Width           =   16920
      _ExtentX        =   29845
      _ExtentY        =   15346
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Títulos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Parcelas"
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
Attribute VB_Name = "BaixaRecOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTBaixaRec
Attribute objCT.VB_VarHelpID = -1

Private Sub UserControl_Initialize()
    Set objCT = New CTBaixaRec
    Set objCT.objUserControl = Me
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub DataBaixa_GotFocus()
     Call objCT.DataBaixa_GotFocus
End Sub

Private Sub DataCredito_GotFocus()
     Call objCT.DataCredito_GotFocus
End Sub

Private Sub DataEmissaoDB_GotFocus()
     Call objCT.DataEmissaoDB_GotFocus
End Sub

Private Sub DataEmissaoDB_KeyPress(KeyAscii As Integer)
     Call objCT.DataEmissaoDB_KeyPress(KeyAscii)
End Sub

Private Sub DataEmissaoDB_Validate(Cancel As Boolean)
     Call objCT.DataEmissaoDB_Validate(Cancel)
End Sub

Private Sub EmissaoFim_GotFocus()
     Call objCT.EmissaoFim_GotFocus
End Sub

Private Sub EmissaoInic_GotFocus()
     Call objCT.EmissaoInic_GotFocus
End Sub

Private Sub HistoricoPerda_Change()
     Call objCT.HistoricoPerda_Change
End Sub

Private Sub LabelCli_Click()
     Call objCT.LabelCli_Click
End Sub

Private Sub TipoDocDB_GotFocus()
     Call objCT.TipoDocDB_GotFocus
End Sub

Private Sub TipoDocDB_KeyPress(KeyAscii As Integer)
     Call objCT.TipoDocDB_KeyPress(KeyAscii)
End Sub

Private Sub TipoDocDB_Validate(Cancel As Boolean)
     Call objCT.TipoDocDB_Validate(Cancel)
End Sub

Private Sub NumTituloDB_GotFocus()
     Call objCT.NumTituloDB_GotFocus
End Sub

Private Sub NumTituloDB_KeyPress(KeyAscii As Integer)
     Call objCT.NumTituloDB_KeyPress(KeyAscii)
End Sub

Private Sub NumTituloDB_Validate(Cancel As Boolean)
     Call objCT.NumTituloDB_Validate(Cancel)
End Sub

Private Sub TituloFim_GotFocus()
     Call objCT.TituloFim_GotFocus
End Sub

Private Sub TituloInic_GotFocus()
     Call objCT.TituloInic_GotFocus
End Sub

Private Sub ValorDB_GotFocus()
     Call objCT.ValorDB_GotFocus
End Sub

Private Sub ValorDB_KeyPress(KeyAscii As Integer)
     Call objCT.ValorDB_KeyPress(KeyAscii)
End Sub

Private Sub ValorDB_Validate(Cancel As Boolean)
     Call objCT.ValorDB_Validate(Cancel)
End Sub

Private Sub SaldoDB_GotFocus()
     Call objCT.SaldoDB_GotFocus
End Sub

Private Sub SaldoDB_KeyPress(KeyAscii As Integer)
     Call objCT.SaldoDB_KeyPress(KeyAscii)
End Sub

Private Sub SaldoDB_Validate(Cancel As Boolean)
     Call objCT.SaldoDB_Validate(Cancel)
End Sub

Private Sub SelecionarDB_GotFocus()
     Call objCT.SelecionarDB_GotFocus
End Sub

Private Sub SelecionarDB_KeyPress(KeyAscii As Integer)
     Call objCT.SelecionarDB_KeyPress(KeyAscii)
End Sub

Private Sub SelecionarDB_Validate(Cancel As Boolean)
     Call objCT.SelecionarDB_Validate(Cancel)
End Sub

Private Sub FilialDB_GotFocus()
     Call objCT.FilialDB_GotFocus
End Sub

Private Sub FilialDB_KeyPress(KeyAscii As Integer)
     Call objCT.FilialDB_KeyPress(KeyAscii)
End Sub

Private Sub FilialDB_Validate(Cancel As Boolean)
     Call objCT.FilialDB_Validate(Cancel)
End Sub

Private Sub ContaCorrenteRA_GotFocus()
     Call objCT.ContaCorrenteRA_GotFocus
End Sub

Private Sub ContaCorrenteRA_KeyPress(KeyAscii As Integer)
     Call objCT.ContaCorrenteRA_KeyPress(KeyAscii)
End Sub

Private Sub ContaCorrenteRA_Validate(Cancel As Boolean)
     Call objCT.ContaCorrenteRA_Validate(Cancel)
End Sub

Private Sub ContaCorrente_Click()
     Call objCT.ContaCorrente_Click
End Sub

Private Sub ContaCorrente_Validate(Cancel As Boolean)
     Call objCT.ContaCorrente_Validate(Cancel)
End Sub

Private Sub DataBaixa_Change()
     Call objCT.DataBaixa_Change
End Sub

Private Sub DataBaixa_Validate(Cancel As Boolean)
     Call objCT.DataBaixa_Validate(Cancel)
End Sub

Private Sub DataMovimentoRA_GotFocus()
     Call objCT.DataMovimentoRA_GotFocus
End Sub

Private Sub DataMovimentoRA_KeyPress(KeyAscii As Integer)
     Call objCT.DataMovimentoRA_KeyPress(KeyAscii)
End Sub

Private Sub DataMovimentoRA_Validate(Cancel As Boolean)
     Call objCT.DataMovimentoRA_Validate(Cancel)
End Sub

Private Sub DataVencimento_GotFocus()
     Call objCT.DataVencimento_GotFocus
End Sub

Private Sub DataVencimento_KeyPress(KeyAscii As Integer)
     Call objCT.DataVencimento_KeyPress(KeyAscii)
End Sub

Private Sub DataVencimento_Validate(Cancel As Boolean)
     Call objCT.DataVencimento_Validate(Cancel)
End Sub

Private Sub EmissaoFim_Change()
     Call objCT.EmissaoFim_Change
End Sub

Private Sub EmissaoFim_Validate(Cancel As Boolean)
     Call objCT.EmissaoFim_Validate(Cancel)
End Sub

Private Sub EmissaoInic_Change()
     Call objCT.EmissaoInic_Change
End Sub

Private Sub EmissaoInic_Validate(Cancel As Boolean)
     Call objCT.EmissaoInic_Validate(Cancel)
End Sub

Private Sub Filial_Click()
     Call objCT.Filial_Click
End Sub

Private Sub Filial_Validate(Cancel As Boolean)
     Call objCT.Filial_Validate(Cancel)
End Sub

Private Sub FilialEmpresa_GotFocus()
     Call objCT.FilialEmpresa_GotFocus
End Sub

Private Sub FilialEmpresa_KeyPress(KeyAscii As Integer)
     Call objCT.FilialEmpresa_KeyPress(KeyAscii)
End Sub

Private Sub FilialEmpresa_Validate(Cancel As Boolean)
     Call objCT.FilialEmpresa_Validate(Cancel)
End Sub

Private Sub FilialRA_GotFocus()
     Call objCT.FilialRA_GotFocus
End Sub

Private Sub FilialRA_KeyPress(KeyAscii As Integer)
     Call objCT.FilialRA_KeyPress(KeyAscii)
End Sub

Private Sub FilialRA_Validate(Cancel As Boolean)
     Call objCT.FilialRA_Validate(Cancel)
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub Cliente_Change()
     Call objCT.Cliente_Change
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)
     Call objCT.Cliente_Validate(Cancel)
End Sub

Private Sub GridParcelas_Click()
     Call objCT.GridParcelas_Click
End Sub

Private Sub GridParcelas_GotFocus()
     Call objCT.GridParcelas_GotFocus
End Sub

Private Sub GridParcelas_EnterCell()
     Call objCT.GridParcelas_EnterCell
End Sub

Private Sub GridParcelas_LeaveCell()
     Call objCT.GridParcelas_LeaveCell
End Sub

Private Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridParcelas_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridParcelas_KeyPress(KeyAscii As Integer)
     Call objCT.GridParcelas_KeyPress(KeyAscii)
End Sub

Private Sub GridParcelas_Validate(Cancel As Boolean)
     Call objCT.GridParcelas_Validate(Cancel)
End Sub

Private Sub GridParcelas_RowColChange()
     Call objCT.GridParcelas_RowColChange
End Sub

Private Sub GridParcelas_Scroll()
     Call objCT.GridParcelas_Scroll
End Sub

Private Sub GridRecebAntecipados_Click()
     Call objCT.GridRecebAntecipados_Click
End Sub

Private Sub GridRecebAntecipados_GotFocus()
     Call objCT.GridRecebAntecipados_GotFocus
End Sub

Private Sub GridRecebAntecipados_EnterCell()
     Call objCT.GridRecebAntecipados_EnterCell
End Sub

Private Sub GridRecebAntecipados_LeaveCell()
     Call objCT.GridRecebAntecipados_LeaveCell
End Sub

Private Sub GridRecebAntecipados_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridRecebAntecipados_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridRecebAntecipados_KeyPress(KeyAscii As Integer)
     Call objCT.GridRecebAntecipados_KeyPress(KeyAscii)
End Sub

Private Sub GridRecebAntecipados_Validate(Cancel As Boolean)
     Call objCT.GridRecebAntecipados_Validate(Cancel)
End Sub

Private Sub GridRecebAntecipados_RowColChange()
     Call objCT.GridRecebAntecipados_RowColChange
End Sub

Private Sub GridRecebAntecipados_Scroll()
     Call objCT.GridRecebAntecipados_Scroll
End Sub

Private Sub GridDebitos_Click()
     Call objCT.GridDebitos_Click
End Sub

Private Sub GridDebitos_GotFocus()
     Call objCT.GridDebitos_GotFocus
End Sub

Private Sub GridDebitos_EnterCell()
     Call objCT.GridDebitos_EnterCell
End Sub

Private Sub GridDebitos_LeaveCell()
     Call objCT.GridDebitos_LeaveCell
End Sub

Private Sub GridDebitos_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridDebitos_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridDebitos_KeyPress(KeyAscii As Integer)
     Call objCT.GridDebitos_KeyPress(KeyAscii)
End Sub

Private Sub GridDebitos_Validate(Cancel As Boolean)
     Call objCT.GridDebitos_Validate(Cancel)
End Sub

Private Sub GridDebitos_RowColChange()
     Call objCT.GridDebitos_RowColChange
End Sub

Private Sub GridDebitos_Scroll()
     Call objCT.GridDebitos_Scroll
End Sub

Private Sub Historico_Change()
     Call objCT.Historico_Change
End Sub

Private Sub MeioPagtoRA_GotFocus()
     Call objCT.MeioPagtoRA_GotFocus
End Sub

Private Sub MeioPagtoRA_KeyPress(KeyAscii As Integer)
     Call objCT.MeioPagtoRA_KeyPress(KeyAscii)
End Sub

Private Sub MeioPagtoRA_Validate(Cancel As Boolean)
     Call objCT.MeioPagtoRA_Validate(Cancel)
End Sub

Private Sub Numero_GotFocus()
     Call objCT.Numero_GotFocus
End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
     Call objCT.Numero_KeyPress(KeyAscii)
End Sub

Private Sub Numero_Validate(Cancel As Boolean)
     Call objCT.Numero_Validate(Cancel)
End Sub

Private Sub Opcao_Click()
     Call objCT.Opcao_Click
End Sub

Private Sub Recebimento_Click(Index As Integer)
     Call objCT.Recebimento_Click(Index)
End Sub

Private Sub Parcela_GotFocus()
     Call objCT.Parcela_GotFocus
End Sub

Private Sub Parcela_KeyPress(KeyAscii As Integer)
     Call objCT.Parcela_KeyPress(KeyAscii)
End Sub

Private Sub Parcela_Validate(Cancel As Boolean)
     Call objCT.Parcela_Validate(Cancel)
End Sub

Private Sub Saldo_GotFocus()
     Call objCT.Saldo_GotFocus
End Sub

Private Sub Saldo_KeyPress(KeyAscii As Integer)
     Call objCT.Saldo_KeyPress(KeyAscii)
End Sub

Private Sub Saldo_Validate(Cancel As Boolean)
     Call objCT.Saldo_Validate(Cancel)
End Sub

Private Sub SaldoRA_GotFocus()
     Call objCT.SaldoRA_GotFocus
End Sub

Private Sub SaldoRA_KeyPress(KeyAscii As Integer)
     Call objCT.SaldoRA_KeyPress(KeyAscii)
End Sub

Private Sub SaldoRA_Validate(Cancel As Boolean)
     Call objCT.SaldoRA_Validate(Cancel)
End Sub

Private Sub Selecionar_Click()
     Call objCT.Selecionar_Click
End Sub

Private Sub SelecionarRA_Click()
     Call objCT.SelecionarRA_Click
End Sub

Private Sub SelecionarDB_Click()
     Call objCT.SelecionarDB_Click
End Sub

Private Sub Tipo_GotFocus()
     Call objCT.Tipo_GotFocus
End Sub

Private Sub Tipo_KeyPress(KeyAscii As Integer)
     Call objCT.Tipo_KeyPress(KeyAscii)
End Sub

Private Sub Tipo_Validate(Cancel As Boolean)
     Call objCT.Tipo_Validate(Cancel)
End Sub

Private Sub TituloFim_Change()
     Call objCT.TituloFim_Change
End Sub

Private Sub TituloFim_Validate(Cancel As Boolean)
     Call objCT.TituloFim_Validate(Cancel)
End Sub

Private Sub TituloInic_Change()
     Call objCT.TituloInic_Change
End Sub

Private Sub UpDownDataBaixa_DownClick()
     Call objCT.UpDownDataBaixa_DownClick
End Sub

Private Sub UpDownDataBaixa_UpClick()
     Call objCT.UpDownDataBaixa_UpClick
End Sub

Private Sub UpDownEmissaoFim_DownClick()
     Call objCT.UpDownEmissaoFim_DownClick
End Sub

Private Sub UpDownEmissaoFim_UpClick()
     Call objCT.UpDownEmissaoFim_UpClick
End Sub

Private Sub UpDownEmissaoInic_DownClick()
     Call objCT.UpDownEmissaoInic_DownClick
End Sub

Private Sub UpDownEmissaoInic_UpClick()
     Call objCT.UpDownEmissaoInic_UpClick
End Sub

Private Sub UpDownVencFim_DownClick()
     Call objCT.UpDownVencFim_DownClick
End Sub

Private Sub UpDownVencFim_UpClick()
     Call objCT.UpDownVencFim_UpClick
End Sub

Private Sub UpDownVencInic_DownClick()
     Call objCT.UpDownVencInic_DownClick
End Sub

Private Sub UpDownVencInic_UpClick()
     Call objCT.UpDownVencInic_UpClick
End Sub

Private Sub ValorAReceber_GotFocus()
     Call objCT.ValorAReceber_GotFocus
End Sub

Private Sub ValorAReceber_KeyPress(KeyAscii As Integer)
     Call objCT.ValorAReceber_KeyPress(KeyAscii)
End Sub

Private Sub ValorAReceber_Validate(Cancel As Boolean)
     Call objCT.ValorAReceber_Validate(Cancel)
End Sub

Private Sub ValorBaixar_Change()
     Call objCT.ValorBaixar_Change
End Sub

Private Sub ValorBaixar_GotFocus()
     Call objCT.ValorBaixar_GotFocus
End Sub

Private Sub ValorBaixar_KeyPress(KeyAscii As Integer)
     Call objCT.ValorBaixar_KeyPress(KeyAscii)
End Sub

Private Sub ValorBaixar_Validate(Cancel As Boolean)
     Call objCT.ValorBaixar_Validate(Cancel)
End Sub

Private Sub ValorDesconto_Change()
     Call objCT.ValorDesconto_Change
End Sub

Private Sub ValorDesconto_GotFocus()
     Call objCT.ValorDesconto_GotFocus
End Sub

Private Sub ValorDesconto_KeyPress(KeyAscii As Integer)
     Call objCT.ValorDesconto_KeyPress(KeyAscii)
End Sub

Private Sub ValorDesconto_Validate(Cancel As Boolean)
     Call objCT.ValorDesconto_Validate(Cancel)
End Sub

Private Sub ValorJuros_Change()
     Call objCT.ValorJuros_Change
End Sub

Private Sub ValorMulta_Change()
     Call objCT.ValorMulta_Change
End Sub

Private Sub ValorMulta_GotFocus()
     Call objCT.ValorMulta_GotFocus
End Sub

Private Sub ValorMulta_KeyPress(KeyAscii As Integer)
     Call objCT.ValorMulta_KeyPress(KeyAscii)
End Sub

Private Sub ValorMulta_Validate(Cancel As Boolean)
     Call objCT.ValorMulta_Validate(Cancel)
End Sub

Private Sub ValorJuros_GotFocus()
     Call objCT.ValorJuros_GotFocus
End Sub

Private Sub ValorJuros_KeyPress(KeyAscii As Integer)
     Call objCT.ValorJuros_KeyPress(KeyAscii)
End Sub

Private Sub ValorJuros_Validate(Cancel As Boolean)
     Call objCT.ValorJuros_Validate(Cancel)
End Sub

Private Sub Selecionar_GotFocus()
     Call objCT.Selecionar_GotFocus
End Sub

Private Sub Selecionar_KeyPress(KeyAscii As Integer)
     Call objCT.Selecionar_KeyPress(KeyAscii)
End Sub

Private Sub Selecionar_Validate(Cancel As Boolean)
     Call objCT.Selecionar_Validate(Cancel)
End Sub

Private Sub SelecionarRA_GotFocus()
     Call objCT.SelecionarRA_GotFocus
End Sub

Private Sub SelecionarRA_KeyPress(KeyAscii As Integer)
     Call objCT.SelecionarRA_KeyPress(KeyAscii)
End Sub

Private Sub SelecionarRA_Validate(Cancel As Boolean)
     Call objCT.SelecionarRA_Validate(Cancel)
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Function Trata_Parametros(Optional objBaixaReceber As ClassBaixaReceber) As Long
     Trata_Parametros = objCT.Trata_Parametros(objBaixaReceber)
End Function

Private Sub ValorRA_GotFocus()
     Call objCT.ValorRA_GotFocus
End Sub

Private Sub ValorRA_KeyPress(KeyAscii As Integer)
     Call objCT.ValorRA_KeyPress(KeyAscii)
End Sub

Private Sub ValorRA_Validate(Cancel As Boolean)
     Call objCT.ValorRA_Validate(Cancel)
End Sub

Private Sub ValorParcela_GotFocus()
     Call objCT.ValorParcela_GotFocus
End Sub

Private Sub ValorParcela_KeyPress(KeyAscii As Integer)
     Call objCT.ValorParcela_KeyPress(KeyAscii)
End Sub

Private Sub ValorParcela_Validate(Cancel As Boolean)
     Call objCT.ValorParcela_Validate(Cancel)
End Sub

Private Sub VencFim_Change()
     Call objCT.VencFim_Change
End Sub

Private Sub VencFim_GotFocus()
     Call objCT.VencFim_GotFocus
End Sub

Private Sub VencFim_Validate(Cancel As Boolean)
     Call objCT.VencFim_Validate(Cancel)
End Sub

Private Sub VencInic_Change()
     Call objCT.VencInic_Change
End Sub

Private Sub VencInic_GotFocus()
     Call objCT.VencInic_GotFocus
End Sub

Private Sub VencInic_Validate(Cancel As Boolean)
     Call objCT.VencInic_Validate(Cancel)
End Sub

Private Sub DataCredito_Change()
     Call objCT.DataCredito_Change
End Sub

Private Sub DataCredito_Validate(Cancel As Boolean)
     Call objCT.DataCredito_Validate(Cancel)
End Sub

Private Sub UpDownDataCredito_DownClick()
     Call objCT.UpDownDataCredito_DownClick
End Sub

Private Sub UpDownDataCredito_UpClick()
     Call objCT.UpDownDataCredito_UpClick
End Sub

Private Sub Cobrador_GotFocus()
     Call objCT.Cobrador_GotFocus
End Sub

Private Sub Cobrador_KeyPress(KeyAscii As Integer)
     Call objCT.Cobrador_KeyPress(KeyAscii)
End Sub

Private Sub Cobrador_Validate(Cancel As Boolean)
     Call objCT.Cobrador_Validate(Cancel)
End Sub

Private Sub CTBBotaoModeloPadrao_Click()
     Call objCT.CTBBotaoModeloPadrao_Click
End Sub

Private Sub CTBModelo_Click()
     Call objCT.CTBModelo_Click
End Sub

Private Sub CTBGridContabil_Click()
     Call objCT.CTBGridContabil_Click
End Sub

Private Sub CTBGridContabil_EnterCell()
     Call objCT.CTBGridContabil_EnterCell
End Sub

Private Sub CTBGridContabil_GotFocus()
     Call objCT.CTBGridContabil_GotFocus
End Sub

Private Sub CTBGridContabil_KeyPress(KeyAscii As Integer)
     Call objCT.CTBGridContabil_KeyPress(KeyAscii)
End Sub

Private Sub CTBGridContabil_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.CTBGridContabil_KeyDown(KeyCode, Shift)
End Sub

Private Sub CTBGridContabil_LeaveCell()
     Call objCT.CTBGridContabil_LeaveCell
End Sub

Private Sub CTBGridContabil_Validate(Cancel As Boolean)
     Call objCT.CTBGridContabil_Validate(Cancel)
End Sub

Private Sub CTBGridContabil_RowColChange()
     Call objCT.CTBGridContabil_RowColChange
End Sub

Private Sub CTBGridContabil_Scroll()
     Call objCT.CTBGridContabil_Scroll
End Sub

Private Sub CTBConta_Change()
     Call objCT.CTBConta_Change
End Sub

Private Sub CTBConta_GotFocus()
     Call objCT.CTBConta_GotFocus
End Sub

Private Sub CTBConta_KeyPress(KeyAscii As Integer)
     Call objCT.CTBConta_KeyPress(KeyAscii)
End Sub

Private Sub CTBConta_Validate(Cancel As Boolean)
     Call objCT.CTBConta_Validate(Cancel)
End Sub

Private Sub CTBCcl_Change()
     Call objCT.CTBCcl_Change
End Sub

Private Sub CTBCcl_GotFocus()
     Call objCT.CTBCcl_GotFocus
End Sub

Private Sub CTBCcl_KeyPress(KeyAscii As Integer)
     Call objCT.CTBCcl_KeyPress(KeyAscii)
End Sub

Private Sub CTBCcl_Validate(Cancel As Boolean)
     Call objCT.CTBCcl_Validate(Cancel)
End Sub

Private Sub CTBCredito_Change()
     Call objCT.CTBCredito_Change
End Sub

Private Sub CTBCredito_GotFocus()
     Call objCT.CTBCredito_GotFocus
End Sub

Private Sub CTBCredito_KeyPress(KeyAscii As Integer)
     Call objCT.CTBCredito_KeyPress(KeyAscii)
End Sub

Private Sub CTBCredito_Validate(Cancel As Boolean)
     Call objCT.CTBCredito_Validate(Cancel)
End Sub

Private Sub CTBDebito_Change()
     Call objCT.CTBDebito_Change
End Sub

Private Sub CTBDebito_GotFocus()
     Call objCT.CTBDebito_GotFocus
End Sub

Private Sub CTBDebito_KeyPress(KeyAscii As Integer)
     Call objCT.CTBDebito_KeyPress(KeyAscii)
End Sub

Private Sub CTBDebito_Validate(Cancel As Boolean)
     Call objCT.CTBDebito_Validate(Cancel)
End Sub

Private Sub CTBSeqContraPartida_Change()
     Call objCT.CTBSeqContraPartida_Change
End Sub

Private Sub CTBSeqContraPartida_GotFocus()
     Call objCT.CTBSeqContraPartida_GotFocus
End Sub

Private Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)
     Call objCT.CTBSeqContraPartida_KeyPress(KeyAscii)
End Sub

Private Sub CTBSeqContraPartida_Validate(Cancel As Boolean)
     Call objCT.CTBSeqContraPartida_Validate(Cancel)
End Sub

Private Sub CTBHistorico_Change()
     Call objCT.CTBHistorico_Change
End Sub

Private Sub CTBHistorico_GotFocus()
     Call objCT.CTBHistorico_GotFocus
End Sub

Private Sub CTBHistorico_KeyPress(KeyAscii As Integer)
     Call objCT.CTBHistorico_KeyPress(KeyAscii)
End Sub

Private Sub CTBHistorico_Validate(Cancel As Boolean)
     Call objCT.CTBHistorico_Validate(Cancel)
End Sub

Private Sub CTBLancAutomatico_Click()
     Call objCT.CTBLancAutomatico_Click
End Sub

Private Sub CTBAglutina_Click()
     Call objCT.CTBAglutina_Click
End Sub

Private Sub CTBAglutina_GotFocus()
     Call objCT.CTBAglutina_GotFocus
End Sub

Private Sub CTBAglutina_KeyPress(KeyAscii As Integer)
     Call objCT.CTBAglutina_KeyPress(KeyAscii)
End Sub

Private Sub CTBAglutina_Validate(Cancel As Boolean)
     Call objCT.CTBAglutina_Validate(Cancel)
End Sub

Private Sub CTBTvwContas_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.CTBTvwContas_NodeClick(Node)
End Sub

Private Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)
     Call objCT.CTBTvwContas_Expand(Node)
End Sub

Private Sub CTBTvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.CTBTvwCcls_NodeClick(Node)
End Sub

Private Sub CTBListHistoricos_DblClick()
     Call objCT.CTBListHistoricos_DblClick
End Sub

Private Sub CTBBotaoLimparGrid_Click()
     Call objCT.CTBBotaoLimparGrid_Click
End Sub

Private Sub CTBLote_Change()
     Call objCT.CTBLote_Change
End Sub

Private Sub CTBLote_GotFocus()
     Call objCT.CTBLote_GotFocus
End Sub

Private Sub CTBLote_Validate(Cancel As Boolean)
     Call objCT.CTBLote_Validate(Cancel)
End Sub

Private Sub CTBDataContabil_Change()
     Call objCT.CTBDataContabil_Change
End Sub

Private Sub CTBDataContabil_GotFocus()
     Call objCT.CTBDataContabil_GotFocus
End Sub

Private Sub CTBDataContabil_Validate(Cancel As Boolean)
     Call objCT.CTBDataContabil_Validate(Cancel)
End Sub

Private Sub CTBDocumento_Change()
     Call objCT.CTBDocumento_Change
End Sub

Private Sub CTBDocumento_GotFocus()
     Call objCT.CTBDocumento_GotFocus
End Sub

Private Sub CTBBotaoImprimir_Click()
     Call objCT.CTBBotaoImprimir_Click
End Sub

Private Sub CTBUpDown_DownClick()
     Call objCT.CTBUpDown_DownClick
End Sub

Private Sub CTBUpDown_UpClick()
     Call objCT.CTBUpDown_UpClick
End Sub

Private Sub CTBLabelDoc_Click()
     Call objCT.CTBLabelDoc_Click
End Sub

Private Sub CTBLabelLote_Click()
     Call objCT.CTBLabelLote_Click
End Sub

Private Sub Label22_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label22, Source, X, Y)
End Sub
Private Sub Label22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label22, Button, Shift, X, Y)
End Sub
Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub
Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub
Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub
Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
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
Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub
Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub
Private Sub LabelCli_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCli, Source, X, Y)
End Sub
Private Sub LabelCli_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCli, Button, Shift, X, Y)
End Sub
Private Sub LabelFilial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilial, Source, X, Y)
End Sub
Private Sub LabelFilial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilial, Button, Shift, X, Y)
End Sub
Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub
Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub
Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub
Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub
Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub
Private Sub ValorReceber_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorReceber, Source, X, Y)
End Sub
Private Sub ValorReceber_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorReceber, Button, Shift, X, Y)
End Sub
Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub
Private Sub Label18_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label18, Source, X, Y)
End Sub
Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label18, Button, Shift, X, Y)
End Sub
Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub
Private Sub TotalBaixar_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalBaixar, Source, X, Y)
End Sub
Private Sub TotalBaixar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalBaixar, Button, Shift, X, Y)
End Sub
Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub
Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
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
Private Sub Opcao_BeforeClick(Cancel As Integer)
     Call objCT.Opcao_BeforeClick(Cancel)
End Sub

Private Sub GridParcelas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Faz com que apareca um PopupMenu o botao direito do mouse acionado sobre o grid

    Call objCT.GridParcelas_MouseDown(Button, Shift, X, Y)

End Sub
Private Sub GridRecebAntecipados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Faz com que apareca um PopupMenu o botao direito do mouse acionado sobre o grid

    Call objCT.GridRecebAntecipados_MouseDown(Button, Shift, X, Y)

End Sub
Private Sub GridDebitos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Faz com que apareca um PopupMenu o botao direito do mouse acionado sobre o grid

    Call objCT.GridDebitos_MouseDown(Button, Shift, X, Y)

End Sub
Private Sub Cliente_Preenche()
     Call objCT.Cliente_Preenche
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

Private Sub TipoDocTodos_Click()
     Call objCT.TipoDocTodos_Click
End Sub

Private Sub TipoDocApenas_Click()
     Call objCT.TipoDocApenas_Click
End Sub

Private Sub TipoDocSeleciona_Change()
     Call objCT.TipoDocSeleciona_Change
End Sub

Private Sub TipoDocSeleciona_Click()
     Call objCT.TipoDocSeleciona_Change
End Sub

Private Sub LabelViagem_Click()
    Call objCT.LabelViagem_Click
End Sub

Private Sub CodigoViagem_GotFocus()
     Call objCT.CodigoViagem_GotFocus
End Sub

Private Sub CodigoViagem_Change()
     Call objCT.CodigoViagem_Change
End Sub

Private Sub CodigoViagem_Validate(Cancel As Boolean)
     Call objCT.CodigoViagem_Validate(Cancel)
End Sub

Private Sub FormaPagto_Change()
     Call objCT.FormaPagto_Change
End Sub

Private Sub FormaPagto_Click()
     Call objCT.FormaPagto_Change
End Sub

Private Sub Vendedor_Change()
     Call objCT.Vendedor_Change
End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)
     Call objCT.Vendedor_Validate(Cancel)
End Sub

Private Sub VendedorLabel_Click()
     Call objCT.VendedorLabel_Click
End Sub

Private Sub CobradorFiltro_Change()
     Call objCT.CobradorFiltro_Change
End Sub

Private Sub CobradorFiltro_Click()
     Call objCT.CobradorFiltro_Change
End Sub

Private Sub NossoNumero_GotFocus()
     Call objCT.NossoNumero_GotFocus
End Sub

Private Sub NossoNumero_KeyPress(KeyAscii As Integer)
     Call objCT.NossoNumero_KeyPress(KeyAscii)
End Sub

Private Sub NossoNumero_Validate(Cancel As Boolean)
     Call objCT.NossoNumero_Validate(Cancel)
End Sub

Private Sub LabelProdutoAte_Click()
    Call objCT.LabelProdutoAte_Click
End Sub

Private Sub LabelProdutoDe_Click()
    Call objCT.LabelProdutoDe_Click
End Sub

Private Sub ProdutoFinal_Validate(Cancel As Boolean)
    Call objCT.ProdutoFinal_Validate(Cancel)
End Sub

Private Sub ProdutoInicial_Validate(Cancel As Boolean)
    Call objCT.ProdutoInicial_Validate(Cancel)
End Sub

Private Sub ProdutoFinal_Change()
     Call objCT.ProdutoFinal_Change
End Sub

Private Sub ProdutoInicial_Change()
     Call objCT.ProdutoInicial_Change
End Sub
