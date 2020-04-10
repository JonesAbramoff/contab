VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl BaixaPagtosCancelarOcx 
   ClientHeight    =   6270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   LockControls    =   -1  'True
   ScaleHeight     =   6270
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5235
      Index           =   1
      Left            =   210
      TabIndex        =   1
      Top             =   828
      Width           =   9165
      Begin VB.Frame Frame9 
         Caption         =   "Filtros"
         Height          =   2715
         Left            =   420
         TabIndex        =   8
         Top             =   1950
         Width           =   8355
         Begin VB.Frame Frame6 
            Caption         =   "Nº do Título"
            Height          =   1815
            Left            =   5790
            TabIndex        =   23
            Top             =   480
            Width           =   2175
            Begin MSMask.MaskEdBox TituloInic 
               Height          =   300
               Left            =   675
               TabIndex        =   24
               Top             =   495
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   9
               Mask            =   "#########"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox TituloFim 
               Height          =   300
               Left            =   690
               TabIndex        =   25
               Top             =   1080
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   9
               Mask            =   "#########"
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
               Left            =   225
               TabIndex        =   27
               Top             =   1125
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
               Left            =   300
               TabIndex        =   26
               Top             =   540
               Width           =   375
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Data de Vencimento"
            Height          =   1815
            Left            =   3084
            TabIndex        =   16
            Top             =   480
            Width           =   2175
            Begin MSComCtl2.UpDown UpDownVencInic 
               Height          =   300
               Left            =   1695
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   510
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox VencInic 
               Height          =   300
               Left            =   630
               TabIndex        =   18
               Top             =   510
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
               Left            =   1695
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   1110
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox VencFim 
               Height          =   300
               Left            =   615
               TabIndex        =   20
               Top             =   1110
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
               Left            =   210
               TabIndex        =   22
               Top             =   1140
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
               TabIndex        =   21
               Top             =   540
               Width           =   375
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Data da Baixa"
            Height          =   1815
            Left            =   384
            TabIndex        =   9
            Top             =   480
            Width           =   2175
            Begin MSComCtl2.UpDown UpDownBaixaInic 
               Height          =   300
               Left            =   1692
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   516
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox BaixaInic 
               Height          =   300
               Left            =   660
               TabIndex        =   11
               Top             =   516
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownBaixaFim 
               Height          =   300
               Left            =   1692
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   1116
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox BaixaFim 
               Height          =   300
               Left            =   645
               TabIndex        =   13
               Top             =   1110
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
               Left            =   195
               TabIndex        =   15
               Top             =   1170
               Width           =   360
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
               Left            =   240
               TabIndex        =   14
               Top             =   525
               Width           =   315
            End
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Fornecedor"
         Height          =   1005
         Left            =   405
         TabIndex        =   3
         Top             =   396
         Width           =   8385
         Begin VB.ComboBox Filial 
            Height          =   288
            Left            =   5388
            TabIndex        =   4
            Top             =   432
            Width           =   2235
         End
         Begin MSMask.MaskEdBox Fornecedor 
            Height          =   300
            Left            =   1716
            TabIndex        =   5
            Top             =   420
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin VB.Label FornecLabel 
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
            ForeColor       =   &H00000000&
            Height          =   192
            Left            =   612
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   7
            Top             =   468
            Width           =   1032
         End
         Begin VB.Label Label12 
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
            ForeColor       =   &H00000000&
            Height          =   192
            Left            =   4836
            TabIndex        =   6
            Top             =   468
            Width           =   492
         End
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   555
      Left            =   7680
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   90
      Width           =   1695
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "BaixaPagtosCancelarOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   100
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   615
         Picture         =   "BaixaPagtosCancelarOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   99
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "BaixaPagtosCancelarOcx.ctx":068C
         Style           =   1  'Graphical
         TabIndex        =   98
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5295
      Index           =   2
      Left            =   165
      TabIndex        =   2
      Top             =   804
      Visible         =   0   'False
      Width           =   9105
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   1716
         Index           =   1
         Left            =   684
         TabIndex        =   81
         Top             =   3276
         Width           =   7788
         Begin VB.Frame FrameBaixa 
            Caption         =   "Dados da Baixa"
            Height          =   1704
            Left            =   72
            TabIndex        =   82
            Top             =   0
            Width           =   7620
            Begin VB.Frame Frame7 
               Caption         =   "Valores"
               Height          =   1365
               Left            =   4950
               TabIndex        =   83
               Top             =   210
               Width           =   2475
               Begin VB.Label ValorPago 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   1110
                  TabIndex        =   87
                  Top             =   330
                  Width           =   945
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  Caption         =   "Pago:"
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
                  Left            =   555
                  TabIndex        =   86
                  Top             =   383
                  Width           =   510
               End
               Begin VB.Label ValorBaixado 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   300
                  Left            =   1110
                  TabIndex        =   85
                  Top             =   900
                  Width           =   945
               End
               Begin VB.Label Label31 
                  AutoSize        =   -1  'True
                  Caption         =   "Baixado:"
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
                  Left            =   315
                  TabIndex        =   84
                  Top             =   953
                  Width           =   750
               End
            End
            Begin VB.Label Juros 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   3300
               TabIndex        =   95
               Top             =   465
               Width           =   945
            End
            Begin VB.Label DataBaixa 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   930
               TabIndex        =   94
               Top             =   465
               Width           =   1125
            End
            Begin VB.Label Desconto 
               BorderStyle     =   1  'Fixed Single
               Height          =   330
               Left            =   3300
               TabIndex        =   93
               Top             =   1080
               Width           =   945
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
               Height          =   285
               Left            =   2370
               TabIndex        =   92
               Top             =   1103
               Width           =   915
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
               Height          =   195
               Left            =   435
               TabIndex        =   91
               Top             =   518
               Width           =   480
            End
            Begin VB.Label Multa 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   930
               TabIndex        =   90
               Top             =   1095
               Width           =   945
            End
            Begin VB.Label Label10 
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
               Left            =   2760
               TabIndex        =   89
               Top             =   518
               Width           =   525
            End
            Begin VB.Label Label9 
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
               Left            =   375
               TabIndex        =   88
               Top             =   1148
               Width           =   540
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Parcelas Baixadas "
         Height          =   2670
         Left            =   540
         TabIndex        =   28
         Top             =   96
         Width           =   7995
         Begin MSMask.MaskEdBox DataEmissao 
            Height          =   288
            Left            =   4824
            TabIndex        =   96
            Top             =   396
            Width           =   912
            _ExtentX        =   1588
            _ExtentY        =   529
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.TextBox Tipo 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   5796
            TabIndex        =   80
            Top             =   396
            Width           =   924
         End
         Begin VB.TextBox TipoBaixa 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   3612
            TabIndex        =   73
            Top             =   1416
            Width           =   2808
         End
         Begin VB.TextBox Sequencial 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   2592
            TabIndex        =   74
            Top             =   1416
            Width           =   696
         End
         Begin VB.TextBox Parcela 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   1650
            TabIndex        =   79
            Top             =   1416
            Width           =   696
         End
         Begin VB.TextBox Numero 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   570
            TabIndex        =   78
            Top             =   1416
            Width           =   996
         End
         Begin VB.TextBox FilialFornec 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   3108
            TabIndex        =   77
            Top             =   390
            Width           =   1665
         End
         Begin VB.CheckBox Seleciona 
            Height          =   300
            Left            =   288
            TabIndex        =   76
            Top             =   390
            Width           =   840
         End
         Begin VB.TextBox Fornec 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   1296
            TabIndex        =   75
            Top             =   390
            Width           =   1815
         End
         Begin MSFlexGridLib.MSFlexGrid GridParcelas 
            Height          =   2160
            Left            =   210
            TabIndex        =   29
            Top             =   330
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   3810
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   1995
         Index           =   2
         Left            =   684
         TabIndex        =   31
         Top             =   3168
         Width           =   7800
         Begin VB.Frame FramePagamento 
            Caption         =   "Adiantamento à Fornecedor"
            Height          =   1650
            Index           =   1
            Left            =   120
            TabIndex        =   60
            Top             =   156
            Visible         =   0   'False
            Width           =   7515
            Begin VB.Label Label5 
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
               Height          =   285
               Left            =   720
               TabIndex        =   72
               Top             =   758
               Width           =   1245
            End
            Begin VB.Label Label6 
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
               Left            =   3525
               TabIndex        =   71
               Top             =   293
               Width           =   1350
            End
            Begin VB.Label Label7 
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
               Left            =   3840
               TabIndex        =   70
               Top             =   1238
               Width           =   1035
            End
            Begin VB.Label Label2 
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
               Height          =   285
               Left            =   1260
               TabIndex        =   69
               Top             =   278
               Width           =   705
            End
            Begin VB.Label Label23 
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
               Left            =   4365
               TabIndex        =   68
               Top             =   773
               Width           =   510
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
               Height          =   288
               Left            =   648
               TabIndex        =   67
               Top             =   1224
               Width           =   1320
            End
            Begin VB.Label DataMovimento 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "DataMovto"
               Height          =   300
               Left            =   2070
               TabIndex        =   66
               Top             =   750
               Width           =   1095
            End
            Begin VB.Label MeioPagtoDescricao 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "MeioPagto"
               Height          =   300
               Left            =   4980
               TabIndex        =   65
               Top             =   1215
               Width           =   1860
            End
            Begin VB.Label ValorPA 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "ValorPagtoAnt"
               Height          =   300
               Left            =   4980
               TabIndex        =   64
               Top             =   750
               Width           =   1860
            End
            Begin VB.Label FilialEmpresaPA 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "FilEmpr"
               Height          =   300
               Left            =   2070
               TabIndex        =   63
               Top             =   1215
               Width           =   525
            End
            Begin VB.Label CCIntNomeReduzido 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "CCorrente"
               Height          =   300
               Left            =   4980
               TabIndex        =   62
               Top             =   270
               Width           =   1845
            End
            Begin VB.Label NumeroMP 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Numero"
               Height          =   300
               Left            =   2070
               TabIndex        =   61
               Top             =   270
               Width           =   720
            End
         End
         Begin VB.Frame FramePagamento 
            Caption         =   "Crédito"
            Height          =   1650
            Index           =   2
            Left            =   120
            TabIndex        =   32
            Top             =   144
            Visible         =   0   'False
            Width           =   7515
            Begin VB.Label SaldoCredito 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Saldo"
               Height          =   300
               Left            =   3720
               TabIndex        =   44
               Top             =   1020
               Width           =   1080
            End
            Begin VB.Label NumTitulo 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Numero"
               Height          =   300
               Left            =   1590
               TabIndex        =   43
               Top             =   390
               Width           =   720
            End
            Begin VB.Label FilialEmpresaCR 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "FilEmpr"
               Height          =   300
               Left            =   6120
               TabIndex        =   42
               Top             =   390
               Width           =   525
            End
            Begin VB.Label ValorCredito 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Valor"
               Height          =   300
               Left            =   6120
               TabIndex        =   41
               Top             =   1020
               Width           =   1080
            End
            Begin VB.Label SiglaDocumentoCR 
               AutoSize        =   -1  'True
               BorderStyle     =   1  'Fixed Single
               Caption         =   "SiglaDoc"
               Height          =   300
               Left            =   3720
               TabIndex        =   40
               Top             =   390
               Width           =   705
            End
            Begin VB.Label DataEmissaoCred 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "DataEmissao"
               Height          =   300
               Left            =   1590
               TabIndex        =   39
               Top             =   1020
               Width           =   1095
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
               Height          =   252
               Left            =   4740
               TabIndex        =   38
               Top             =   420
               Width           =   1284
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
               Left            =   3060
               TabIndex        =   37
               Top             =   1050
               Width           =   555
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
               Left            =   5505
               TabIndex        =   36
               Top             =   1050
               Width           =   510
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
               Left            =   780
               TabIndex        =   35
               Top             =   420
               Width           =   705
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
               Left            =   3165
               TabIndex        =   34
               Top             =   420
               Width           =   450
            End
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
               Left            =   240
               TabIndex        =   33
               Top             =   1050
               Width           =   1245
            End
         End
         Begin VB.Frame FramePagamento 
            Caption         =   "Dados do Pagamento"
            Height          =   1650
            Index           =   3
            Left            =   120
            TabIndex        =   45
            Top             =   150
            Width           =   7515
            Begin VB.Frame Frame10 
               Caption         =   "Meio Pagamento"
               Height          =   1335
               Left            =   5910
               TabIndex        =   46
               Top             =   210
               Width           =   1455
               Begin VB.OptionButton TipoMeioPagto 
                  Caption         =   "Borderô"
                  CausesValidation=   0   'False
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
                  Index           =   1
                  Left            =   210
                  TabIndex        =   49
                  Top             =   630
                  Width           =   1035
               End
               Begin VB.OptionButton TipoMeioPagto 
                  Caption         =   "Cheque"
                  CausesValidation=   0   'False
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
                  Left            =   210
                  TabIndex        =   48
                  Top             =   285
                  Value           =   -1  'True
                  Width           =   1035
               End
               Begin VB.OptionButton TipoMeioPagto 
                  Caption         =   "Dinheiro"
                  CausesValidation=   0   'False
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
                  Left            =   210
                  TabIndex        =   47
                  Top             =   990
                  Width           =   1035
               End
            End
            Begin VB.Label Portador 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   3510
               TabIndex        =   59
               Top             =   345
               Width           =   2145
            End
            Begin VB.Label NumOuSequencial 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   1050
               TabIndex        =   58
               Top             =   345
               Width           =   735
            End
            Begin VB.Label Historico 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   1080
               TabIndex        =   57
               Top             =   1200
               Width           =   4575
            End
            Begin VB.Label ValorPagoPagto 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   4065
               TabIndex        =   56
               Top             =   765
               Width           =   1590
            End
            Begin VB.Label ContaCorrente 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   870
               TabIndex        =   55
               Top             =   765
               Width           =   2310
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
               Left            =   210
               TabIndex        =   54
               Top             =   390
               Width           =   750
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
               Left            =   3510
               TabIndex        =   53
               Top             =   795
               Width           =   495
            End
            Begin VB.Label Label14 
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
               Left            =   2070
               TabIndex        =   52
               Top             =   375
               Width           =   1365
            End
            Begin VB.Label Label13 
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
               Left            =   210
               TabIndex        =   51
               Top             =   1230
               Width           =   810
            End
            Begin VB.Label Label4 
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
               Left            =   210
               TabIndex        =   50
               Top             =   795
               Width           =   555
            End
         End
      End
      Begin MSComctlLib.TabStrip TabBaixaPag 
         Height          =   2355
         Left            =   570
         TabIndex        =   30
         Top             =   2835
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   4154
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dados da Baixa"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dados do Pagamento"
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
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5748
      Left            =   108
      TabIndex        =   0
      Top             =   432
      Width           =   9312
      _ExtentX        =   16431
      _ExtentY        =   10134
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Parcelas"
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
Attribute VB_Name = "BaixaPagtosCancelarOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

'Definicoes do Grid de Parcelas
Public objGridParcelas As AdmGrid

Dim iGrid_Seleciona_Col As Integer
Dim iGrid_Fornecedor_Col As Integer
Dim iGrid_Filial_Col As Integer
Dim iGrid_DataBaixa_Col As Integer
Dim iGrid_Tipo_Col As Integer
Dim iGrid_Numero_Col As Integer
Dim iGrid_Parcela_Col As Integer
Dim iGrid_Sequencial_Col As Integer
Dim iGrid_TipoBaixa_Col As Integer

'Definicoes dos TABS da Tela
Private Const TAB_SELECAO = 1
Private Const TAB_Parcelas = 2

Private Const TAB_BAIXA = 1

'Definicoes de Constantes
Const NUM_MAX_PARCELAS_CANCEL = 200

Private Const FORNECEDOR_NOME_RED As String = "Fornecedor_NomeRed"
Private Const FILIAL_NOME_RED As String = "FilialForn_Nome"
Private Const PARCELAS As String = "Parcelas"

'Guarda se o Tab de Selecao foi alterado
Public giSelecaoAlterado As Integer
    
'Browser
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1

'Property Variables:
Dim m_Caption As String
Event Unload()

Public iAlterado As Integer
Public iFornecedorAlterado As Integer
Dim iFrameAtual As Integer
Dim iFramePagBaixaAtual As Integer
Dim gcolInfoParcPag As Collection

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Cancelamento de Baixa de Parcelas a Pagar"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "BaixaPagtosCancelar"
    
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

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 95331
    
    'Limpa a Tela
    Call Limpa_Tela_BaixaPagtosCancelar
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 95331 'Tratado na Rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143253)

    End Select

End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela(Me)
    
    Call Limpa_Tela_BaixaPagtosCancelar
    
    'Zera iAlterado
    iAlterado = 0
    
End Sub

Private Function Limpa_Tela_BaixaPagtosCancelar()
'Limpa a tela

    'Limpa os comandos da tela
    Call Limpa_Tela(Me)
    
    'Limpa os Grid
    Call Grid_Limpa(objGridParcelas)
    
    'Limpa os TABS do TAB de Parcelas
    Desconto.Caption = ""
    ValorPago.Caption = ""
    Multa.Caption = ""
    ValorBaixado.Caption = ""
    Juros.Caption = ""
    DataBaixa.Caption = ""
    
    ContaCorrente.Caption = ""
    ValorPagoPagto.Caption = ""
    Historico.Caption = ""
    NumOuSequencial.Caption = ""
    
    DataEmissaoCred.Caption = ""
    NumTitulo.Caption = ""
    SaldoCredito.Caption = ""
    SiglaDocumentoCR.Caption = ""
    ValorCredito.Caption = ""
    FilialEmpresaCR.Caption = ""
    
    DataMovimento.Caption = ""
    ValorPA.Caption = ""
    FilialEmpresaPA.Caption = ""
    CCIntNomeReduzido.Caption = ""
    NumeroMP.Caption = ""
    MeioPagtoDescricao.Caption = ""
    
    Filial.Clear

    Exit Function

End Function

Private Sub TabBaixaPag_Click()

    'Se o Frame atual não corresponde ao TAB clicado
    If TabBaixaPag.SelectedItem.Index <> iFramePagBaixaAtual Then
    
        If TabStrip_PodeTrocarTab(iFramePagBaixaAtual, TabBaixaPag, Me) <> SUCESSO Then Exit Sub
        
        'Torna Frame selecionado visível
        Frame3(TabBaixaPag.SelectedItem.Index).Visible = True
        
        'Torna Frame atual invisível
        Frame3(iFramePagBaixaAtual).Visible = False
        
        'Armazena novo valor de iFramePagBaixaAtual
        iFramePagBaixaAtual = TabBaixaPag.SelectedItem.Index
    
    End If

End Sub

Sub Opcao_Click()

Dim lErro As Long

On Error GoTo Erro_Opcao_Click

    'Se Frame atual não corresponde ao Tab clicado
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub
        
        'Torna Frame selecionado visível
        Frame1(Opcao.SelectedItem.Index).Visible = True
        
        'Torna Frame atual invisível
        Frame1(iFrameAtual).Visible = False
        
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
        
        'Se Frame selecionado foi o de seleção
        If Opcao.SelectedItem.Index = TAB_SELECAO Then

            giSelecaoAlterado = 0

        'Se Frame selecionado foi o de Parcelas
        ElseIf Opcao.SelectedItem.Index = TAB_Parcelas Then

            If giSelecaoAlterado <> 0 Then
                
                lErro = Carrega_Tab_Parcelas()
                If lErro <> SUCESSO Then gError 95251

                giSelecaoAlterado = 0

            End If

        End If

    End If
    
    Exit Sub

Erro_Opcao_Click:

    Select Case gErr
    
        Case 95251
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB:", gErr, Error, 143254)
            
    End Select
    
    Exit Sub

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

'***** fim do trecho a ser copiado ******

Private Function Inicializa_Grid_Parcelas(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Parcelas

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Cancelar")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Emissão")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Parc.")
    objGridInt.colColuna.Add ("Seq.")
    objGridInt.colColuna.Add ("Tipo da Baixa")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (Seleciona.Name)
    objGridInt.colCampo.Add (Fornec.Name)
    objGridInt.colCampo.Add (FilialFornec.Name)
    objGridInt.colCampo.Add (DataEmissao.Name)
    objGridInt.colCampo.Add (Tipo.Name)
    objGridInt.colCampo.Add (Numero.Name)
    objGridInt.colCampo.Add (Parcela.Name)
    objGridInt.colCampo.Add (Sequencial.Name)
    objGridInt.colCampo.Add (TipoBaixa.Name)
    
    'Colunas do Grid
    iGrid_Seleciona_Col = 1
    iGrid_Fornecedor_Col = 2
    iGrid_Filial_Col = 3
    iGrid_DataBaixa_Col = 4
    iGrid_Tipo_Col = 5
    iGrid_Numero_Col = 6
    iGrid_Parcela_Col = 7
    iGrid_Sequencial_Col = 8
    iGrid_TipoBaixa_Col = 9
    
    'Grid do GridInterno
    objGridInt.objGrid = GridParcelas

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_PARCELAS_CANCEL + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridParcelas.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)
    
    giSelecaoAlterado = 0

    Inicializa_Grid_Parcelas = SUCESSO

    Exit Function

End Function

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    If iFornecedorAlterado = 0 Then Exit Sub
    
    'Se Fornecedor está preenchido
    If Len(Trim(Fornecedor.Text)) > 0 Then

        'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
        If lErro <> SUCESSO Then gError 95220

        'Lê coleção de códigos, nomes de Filiais do Fornecedor
        lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
        If lErro <> SUCESSO Then gError 95221

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)
        
        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)
        
    'Se Fornecedor não está preenchido
    ElseIf Len(Trim(Fornecedor.Text)) = 0 Then

        'Limpa a Combo de Filiais
        Filial.Clear

    End If

    iFornecedorAlterado = 0
    
    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 95220, 95221

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143255)

    End Select

    Exit Sub

End Sub

Private Sub FornecLabel_Click()
'Chamada do Browse de Fornecedores

Dim colSelecao As Collection
Dim objFornecedor As New ClassFornecedor

    'Passa o Fornecedor que está na tela para o Obj
    objFornecedor.sNomeReduzido = Trim(Fornecedor.Text)

    'Chama a tela com a lista de Fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

    Exit Sub

End Sub

Private Sub Fornecedor_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFornecedorAlterado = REGISTRO_ALTERADO
    giSelecaoAlterado = REGISTRO_ALTERADO

    Call Fornecedor_Preenche

End Sub

Public Sub Filial_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    giSelecaoAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Filial_Click()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim iCodigo As Integer
Dim sNomeRed As String

On Error GoTo Erro_Filial_Validate

    'Verifica se foi preenchida a ComboBox Filial
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Filial
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 95222

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        'Verifica se foi preenchido o Fornecedor
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 95223

        'Lê o Fornecedor que está na tela
        sNomeRed = Trim(Fornecedor.Text)

        'Passa o Código da Filial que está na tela para o Obj
        objFilialFornecedor.iCodFilial = iCodigo

        'Lê Filial no BD a partir do NomeReduzido do Fornecedor e Código da Filial
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sNomeRed, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 95224

        'Se não existe a Filial
        If lErro = 18272 Then gError 95225

        'Encontrou Filial no BD, coloca no Text da Combo
        Filial.Text = CStr(objFilialFornecedor.iCodFilial) & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 95226

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 95222, 95224

        Case 95223
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 95225
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIAL_FORNECEDOR")

            If vbMsgRes = vbYes Then
                'Chama a tela de Filiais
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            Else
                'Segura o foco
                Filial.SetFocus
            End If

        Case 95226
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143256)

    End Select

    Exit Sub

End Sub

Public Sub BaixaInic_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    giSelecaoAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub BaixaInic_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(BaixaInic, iAlterado)

End Sub

Public Sub BaixaInic_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_BaixaInic_Validate

    'Se a data está preenchida
    If Len(BaixaInic.ClipText) > 0 Then

        'Verifica se a data é válida
        lErro = Data_Critica(BaixaInic.Text)
        If lErro <> SUCESSO Then gError 95227

    
        'Se a data está preenchida
        If Len(BaixaFim.ClipText) > 0 Then
        
            'Verifica se a BaixaFim é menor que a BaixaInic
            If CDate(BaixaFim.Text) < CDate(BaixaInic.Text) Then gError 95242
        
        End If
    
    End If


    Exit Sub

Erro_BaixaInic_Validate:

    Cancel = True

    Select Case gErr

        Case 95227
        
        Case 95242
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143257)

    End Select

    Exit Sub

End Sub

Public Sub BaixaFim_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    giSelecaoAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub BaixaFim_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(BaixaFim, iAlterado)

End Sub

Public Sub BaixaFim_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_BaixaFim_Validate

    'Se a data está preenchida
    If Len(BaixaFim.ClipText) > 0 Then

        'Verifica se a data é válida
        lErro = Data_Critica(BaixaFim.Text)
        If lErro <> SUCESSO Then gError 95228

        'Se a data está preenchida
        If Len(BaixaInic.ClipText) > 0 Then

            'Verifica se a BaixaFim é menor que a BaixaInic
            If CDate(BaixaFim.Text) < CDate(BaixaInic.Text) Then gError 95229

        End If

    End If

    Exit Sub

Erro_BaixaFim_Validate:

    Cancel = True

    Select Case gErr

        Case 95228

        Case 95229
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143258)

    End Select

    Exit Sub

End Sub

Public Sub VencFim_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    giSelecaoAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VencFim_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(VencFim, iAlterado)

End Sub

Public Sub VencFim_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_VencFim_Validate

    'Se a data VencFim está preenchida
    If Len(VencFim.ClipText) > 0 Then

        'Verifica se a data VencFim é válida
        lErro = Data_Critica(VencFim.Text)
        If lErro <> SUCESSO Then gError 95230

        'Se a data vencInica está preenchida
        If Len(VencInic.ClipText) > 0 Then

            'Verifica se a VencFim é maior ou igual a VencInic
            If CDate(VencFim.Text) < CDate(VencInic.Text) Then gError 95231

        End If

    End If

    Exit Sub

Erro_VencFim_Validate:

    Cancel = True

    Select Case gErr

        Case 95230

        Case 95231
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143259)

    End Select

    Exit Sub

End Sub

Public Sub VencInic_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    giSelecaoAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub VencInic_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(VencInic, iAlterado)
    
End Sub

Public Sub VencInic_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_VencInic_Validate

    'Se a data VencInic está preenchida
    If Len(VencInic.ClipText) > 0 Then

        'Verifica se a data VencInic é válida
        lErro = Data_Critica(VencInic.Text)
        If lErro <> SUCESSO Then gError 95232

        'Se a data está preenchida
        If Len(VencFim.ClipText) > 0 Then
    
            'Verifica se a VencFim é maior ou igual a VencInic
            If CDate(VencFim.Text) < CDate(VencInic.Text) Then gError 95243
    
        End If
    
    End If
    
Exit Sub

Erro_VencInic_Validate:

    Cancel = True

    Select Case gErr

        Case 95232
        
        Case 95243
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143260)

    End Select

    Exit Sub

End Sub

Public Sub TituloFim_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    giSelecaoAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub TituloFim_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TituloFim, iAlterado)

End Sub

Public Sub TituloFim_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TituloFim_Validate

    'Se TituloFim e TituloInic estão preenchidos
    If Len(Trim(TituloFim.Text)) And Len(Trim(TituloInic.Text)) > 0 Then

        'Verifica se TituloFim é maior ou igual que TituloInic
        If CLng(TituloFim.Text) < CLng(TituloInic.Text) Then gError 95233

    End If

    Exit Sub

Erro_TituloFim_Validate:

    Cancel = True

    Select Case gErr

        Case 95233
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULOINIC_MAIOR_TITULOFIM", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143261)

    End Select

    Exit Sub

End Sub

Public Sub TituloInic_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    giSelecaoAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub TituloInic_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(TituloInic, iAlterado)
    
End Sub

Public Sub TituloInic_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TituloInic_Validate

    'Se TituloFim e TituloInic estão preenchidos
    If Len(Trim(TituloFim.Text)) And Len(Trim(TituloInic.Text)) > 0 Then

        'Verifica se TituloFim é maior ou igual que TituloInic
        If CLng(TituloFim.Text) < CLng(TituloInic.Text) Then gError 95243

    End If

    Exit Sub

Erro_TituloInic_Validate:

    Cancel = True

    Select Case gErr

        Case 95243
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULOINIC_MAIOR_TITULOFIM", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143262)

    End Select

    Exit Sub

End Sub

Public Sub UpDownBaixaFim_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBaixaFim_DownClick

    'Diminui a data BaixaFim em 1 dia
    lErro = Data_Up_Down_Click(BaixaFim, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 95234

    Exit Sub

Erro_UpDownBaixaFim_DownClick:

    Select Case gErr

        Case 95234

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143263)

    End Select

    Exit Sub

End Sub

Public Sub UpDownBaixaFim_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBaixaFim_UpClick

    'Aumenta a data BaixaFim em 1 dia
    lErro = Data_Up_Down_Click(BaixaFim, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 95235

    Exit Sub

Erro_UpDownBaixaFim_UpClick:

    Select Case gErr

        Case 95235

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143264)

    End Select

    Exit Sub

End Sub

Public Sub UpDownBaixaInic_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBaixaInic_DownClick

    'Diminui a data BaixaInic em 1 dia
    lErro = Data_Up_Down_Click(BaixaInic, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 95236

    Exit Sub

Erro_UpDownBaixaInic_DownClick:

    Select Case gErr

        Case 95236

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143265)

    End Select

    Exit Sub

End Sub

Public Sub UpDownBaixaInic_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBaixaInic_UpClick

    'Aumenta a data BaixaInic em 1 dia
    lErro = Data_Up_Down_Click(BaixaInic, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 95237

    Exit Sub

Erro_UpDownBaixaInic_UpClick:

    Select Case gErr

        Case 95237

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143266)

    End Select

    Exit Sub

End Sub

Public Sub UpDownVencFim_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencFim_DownClick

    'Diminui a data VencFim em 1 dia
    lErro = Data_Up_Down_Click(VencFim, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 95238

    Exit Sub

Erro_UpDownVencFim_DownClick:

    Select Case gErr

        Case 95238

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143267)

    End Select

    Exit Sub

End Sub

Public Sub UpDownVencFim_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencFim_UpClick

    'Aumenta a data VencFim em 1 dia
    lErro = Data_Up_Down_Click(VencFim, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 95239

    Exit Sub

Erro_UpDownVencFim_UpClick:

    Select Case gErr

        Case 95239

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143268)

    End Select

    Exit Sub

End Sub

Public Sub UpDownVencInic_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencInic_DownClick

    'Diminui a data VencInic em 1 dia
    lErro = Data_Up_Down_Click(VencInic, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 95240

    Exit Sub

Erro_UpDownVencInic_DownClick:

    Select Case gErr

        Case 95240

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143269)

    End Select

    Exit Sub

End Sub

Public Sub UpDownVencInic_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencInic_UpClick

    'Aumenta a data VencInic em 1 dia
    lErro = Data_Up_Down_Click(VencInic, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 95241

    Exit Sub

Erro_UpDownVencInic_UpClick:

    Select Case gErr

        Case 95241

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143270)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()
   
Dim lErro As Long
Dim iSubTipo As Integer

On Error GoTo Erro_Form_Load

    Set objGridParcelas = New AdmGrid
    
    iFrameAtual = TAB_SELECAO
    iFramePagBaixaAtual = TAB_BAIXA
     
    'Inicializa o Grid de Parcelas
    lErro = Inicializa_Grid_Parcelas(objGridParcelas)
    If lErro <> SUCESSO Then gError 95250
    
    If gobjCP.iContabSemDet <> 0 Then gError 81870
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr

        Case 95250
        
        Case 81870
            Call Rotina_Erro(vbOKOnly, "ERRO_CANC_BXPAG_MANUAL", gErr, Error)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143271)
    
    End Select
    
    Exit Sub

End Sub
    

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 15884

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 15884
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143272)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objBaixaPag As ClassBaixaPagar) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143273)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Function Carrega_Tab_Parcelas() As Long
'Carrega os dados das parcelas, créditos e pagamentos antecipados para tela

Dim lErro As Long
Dim iCodFilialFornecedor As Integer
Dim iFilialForn As Integer
Dim lCodForn As Long
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim colCreditos As New Collection
Dim colPagtoAntecipado As New Collection
Dim dtBaixaInic As Date
Dim dtBaixaFim As Date
Dim dtVencInic As Date
Dim dtVencFim As Date
Dim lTituloInic As Long
Dim lTituloFim As Long

On Error GoTo Erro_Carrega_Tab_Parcelas

    'Verifica se Fornecedor e a filial estão preenchidos
    'If Len(Trim(Fornecedor.Text)) = 0 Then gError 95252
    'If Len(Trim(Filial.Text)) = 0 Then gError 95253

    If giSelecaoAlterado <> 0 Then

        'Limpa o grid Parcelas
        Call Grid_Limpa(objGridParcelas)
        
        If Len(Trim(Fornecedor.Text)) Then
            'Lê os dados do Fornecedor
            lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilialFornecedor)
            If lErro <> SUCESSO Then gError 95254

            'PassaFornecedor que está na tela para o Obj
            lCodForn = objFornecedor.lCodigo
            objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
    
            'Passa a Filial que está na tela para o Obj
            iFilialForn = Codigo_Extrai(Filial.Text)
            objFilialFornecedor.iCodFilial = iFilialForn

        End If

        dtBaixaInic = MaskedParaDate(BaixaInic)
        dtBaixaFim = MaskedParaDate(BaixaFim)
        dtVencInic = MaskedParaDate(VencInic)
        dtVencFim = MaskedParaDate(VencFim)

        'Se as datas baixaFim e baixaInic estão preenchidas
        If dtBaixaInic <> DATA_NULA And dtBaixaFim <> DATA_NULA Then
            'Verifica se a baixaFim é maior ou igual a baixaInic
            If dtBaixaFim < dtBaixaInic Then gError 95255
        End If

        'Se as datas VencFim e VencInic estão preenchidas
        If dtVencInic <> DATA_NULA And dtVencFim <> DATA_NULA Then
            'Verifica se a baixaFim é maior ou igual a baixaInic
            If dtVencFim < dtVencInic Then gError 95256
        End If

        'Lê TituloInic e TituloFim que estão na tela
        lTituloInic = StrParaLong(TituloInic.Text)
        lTituloFim = StrParaLong(TituloFim.Text)
                
        If (lTituloInic <> 0 And lTituloFim <> 0) Then
            'Verifica se TituloFim é maior ou igual que TituloInic
            If lTituloFim < lTituloInic Then gError 95257
        End If

        'Limpa gcolInfoParcPag antes de carregar as novas parcelas
        Set gcolInfoParcPag = Nothing
        Set gcolInfoParcPag = New Collection

        'Preenche a Coleção de Parcelas
        lErro = ParcelasPagarBaixadas_Le_Cancelamento(lCodForn, iFilialForn, dtBaixaInic, dtBaixaFim, dtVencInic, dtVencFim, lTituloInic, lTituloFim, gcolInfoParcPag)
        If lErro <> SUCESSO Then gError 95258
       
       'Verifica o número máximo de parcelas
        If gcolInfoParcPag.Count > NUM_MAX_PARCELAS_CANCEL Then gError 95259

        'Preenche o GridParcelas
        Call Grid_Parcelas_Preenche(gcolInfoParcPag)

    End If

    Carrega_Tab_Parcelas = SUCESSO

    Exit Function

Erro_Carrega_Tab_Parcelas:

    Carrega_Tab_Parcelas = gErr

    Select Case gErr

        Case 95252
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 95253
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 95255
            Call Rotina_Erro(vbOKOnly, "ERRO_DATABAIXA_INICIAL_MAIOR", gErr)

        Case 95256
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_INICIAL_MAIOR", gErr)

        Case 95257
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULOINIC_MAIOR_TITULOFIM", gErr)

        Case 95258, 95254

        Case 95259
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAS_SUPERIOR_NUM_MAX_PARCELAS_CANCEL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143274)

    End Select

    Exit Function

End Function

Public Function ParcelasPagarBaixadas_Le_Cancelamento(lCodForn As Long, iFilialForn As Integer, dtBaixaInic As Date, dtBaixaFim As Date, dtVencInic As Date, dtVencFim As Date, lTituloInic As Long, lTituloFim As Long, colInfoParcPag As Collection) As Long
'preenche a colecao com informacoes de parcelas a pagar a partir dos criterios informados
'datas nao preenchidas devem ser passadas como DATA_NULA
'numeros de titulo nao preenchidos devem ser passados como zero
'IMPORTANTE: nao estou preenchendo a razao social e o nome reduzido do fornecedor nos objetos armazenados na colecao

Dim lErro As Long
Dim sSelect As String
Dim iStatusAberto As Integer
Dim lComando As Long
Dim sNomeRedPortador As String

'buffers para receber registros das parcelas

Dim lNumIntParc As Long
Dim lNumIntDoc As Long
Dim lNumConta As Long
Dim iMotivo As Integer
Dim dValorMulta As Double
Dim dValorJuros As Double
Dim dValorDesconto As Double
Dim dValorBaixado As Double
Dim dtDataBaixa As Date
Dim iSequencial As Integer
Dim iNumParcela As Integer
Dim sSiglaDocumento As String
Dim lNumTitulo As Long
Dim dtEmissaoTitulo As Date
Dim iStatusBaixa As Integer
Dim iFilialEmpresa As Integer
Dim lFornecedor As Long
Dim iFilial As Integer
Dim iStatusParcela As Integer
Dim lNumIntBaixa As Long

On Error GoTo Erro_ParcelasPagarBaixadas_Le_Cancelamento

    iStatusAberto = STATUS_ABERTO
    'lFornecedor = lCodForn
    'iFilial = iFilialForn
    iFilialEmpresa = giFilialEmpresa

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 95260

    sSiglaDocumento = String(STRING_SIGLA_DOCUMENTO, 0)
    
    'Prepara a parte Lógica do SELECT
    Call ParcelasPagarBaixadas_Le_Cancelamento1(sSelect, lCodForn, iFilialForn, dtBaixaInic, dtBaixaFim, dtVencInic, dtVencFim, lTituloInic, lTituloFim)

    'Atribui Constantes para preparacao do SELECT
    iStatusBaixa = STATUS_EXCLUIDO
    
    'executa a preparacao da parte fixa do SELECT
    lErro = ParcelasPagarBaixadas_Le_Cancelamento2(lComando, sSelect, lNumIntDoc, lNumIntParc, lNumConta, iMotivo, dValorMulta, dValorJuros, dValorDesconto, dValorBaixado, dtDataBaixa, iSequencial, iNumParcela, sSiglaDocumento, lNumTitulo, dtEmissaoTitulo, lNumIntBaixa, lFornecedor, iFilial)
    If lErro <> SUCESSO Then gError 95283
        
    'complementa a passagem dos parametros que variam de acordo com a selecao do usuario
    'e executa o SELECT p/obtencao das parcelas
    lErro = ParcelasPagarBaixadas_Le_Cancelamento3(lComando, dtBaixaInic, dtBaixaFim, dtVencInic, dtVencFim, lTituloInic, lTituloFim, iStatusBaixa, iFilialEmpresa, lCodForn, iFilialForn)
    If lErro <> SUCESSO Then gError 95293
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 95294

    If lErro = AD_SQL_SEM_DADOS Then gError 95295

    Do While lErro = AD_SQL_SUCESSO
                
        ' inclui a parcela lida na colecao
        Call ParcelasPagarBaixadas_Le_Cancelamento4(colInfoParcPag, lNumIntDoc, lNumIntParc, lNumConta, iMotivo, dValorMulta, dValorJuros, dValorDesconto, dValorBaixado, dtDataBaixa, iSequencial, iNumParcela, sSiglaDocumento, lNumTitulo, dtEmissaoTitulo, lFornecedor, iFilial, lNumIntBaixa)
            
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 95297

    Loop

    lErro = Comando_Fechar(lComando)
    
    ParcelasPagarBaixadas_Le_Cancelamento = SUCESSO
    
    Exit Function
    
Erro_ParcelasPagarBaixadas_Le_Cancelamento:

    ParcelasPagarBaixadas_Le_Cancelamento = gErr
    
    Select Case gErr

        Case 95283, 95284, 95292, 95293, 95296

        Case 95260, 95261
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 95294, 95297
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_PAG", gErr)

        Case 95295
            Call Rotina_Erro(vbOKOnly, "ERRO_SEM_PARCELAS_PAG_SEL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143275)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Sub ParcelasPagarBaixadas_Le_Cancelamento1(sSelect As String, ByVal lForn As Long, ByVal iFilial As Integer, ByVal dtBaixaInic As Date, ByVal dtBaixaFim As Date, ByVal dtVencInic As Date, ByVal dtVencFim As Date, ByVal lTituloInic As Long, ByVal lTituloFim As Long)
'monta o SELECT para obtencao das parcelas dinamicamente.
'OBS -> O Select tem UNION !!!

Dim sFromN As String
Dim sFromB As String
Dim sWhereN As String
Dim sWhereB As String
Dim sFieldsN As String
Dim sFieldsB As String

'    sFieldsB = "BaixasPag.NumIntBaixa, BaixasPag.NumIntDoc, BaixasParcPag.NumIntParcela, BaixasPag.NumMovCta, BaixasPag.Motivo, BaixasParcPag.ValorMulta, BaixasParcPag.ValorJuros, BaixasParcPag.ValorDesconto, BaixasParcPag.ValorBaixado, BaixasPag.Data, BaixasParcPag.Sequencial, ParcelasPagBaixadas.NumParcela, TitulosPagBaixados.SiglaDocumento, TitulosPagBaixados.NumTitulo, TitulosPagBaixados.DataEmissao "
'    sFieldsN = "BaixasPag.NumIntBaixa, BaixasPag.NumIntDoc, BaixasParcPag.NumIntParcela, BaixasPag.NumMovCta, BaixasPag.Motivo, BaixasParcPag.ValorMulta, BaixasParcPag.ValorJuros, BaixasParcPag.ValorDesconto, BaixasParcPag.ValorBaixado, BaixasPag.Data, BaixasParcPag.Sequencial, ParcelasPag.NumParcela, TitulosPag.SiglaDocumento, TitulosPag.NumTitulo, TitulosPag.DataEmissao "

    sFieldsB = "BaixasPag.NumIntBaixa, BaixasPag.NumIntDoc, BaixasParcPag.NumIntParcela, BaixasPag.NumMovCta, BaixasPag.Motivo, BaixasParcPag.ValorMulta, BaixasParcPag.ValorJuros, BaixasParcPag.ValorDesconto, BaixasParcPag.ValorBaixado, BaixasPag.Data, BaixasParcPag.Sequencial, ParcelasPagBaixadas.NumParcela, TitulosPagBaixados.SiglaDocumento, TitulosPagBaixados.NumTitulo, TitulosPagBaixados.DataEmissao, TitulosPagBaixados.Fornecedor , TitulosPagBaixados.Filial "
    sFieldsN = "BaixasPag.NumIntBaixa, BaixasPag.NumIntDoc, BaixasParcPag.NumIntParcela, BaixasPag.NumMovCta, BaixasPag.Motivo, BaixasParcPag.ValorMulta, BaixasParcPag.ValorJuros, BaixasParcPag.ValorDesconto, BaixasParcPag.ValorBaixado, BaixasPag.Data, BaixasParcPag.Sequencial, ParcelasPag.NumParcela, TitulosPag.SiglaDocumento, TitulosPag.NumTitulo, TitulosPag.DataEmissao, TitulosPag.Fornecedor , TitulosPag.Filial "

    sFromB = "FROM BaixasPag, BaixasParcPag, ParcelasPagBaixadas, TitulosPagBaixados "
    sFromN = "FROM BaixasPag, BaixasParcPag, ParcelasPag, TitulosPag "

    sWhereB = "WHERE BaixasPag.NumIntBaixa = BaixasParcPag.NumIntBaixa AND BaixasParcPag.NumIntParcela = ParcelasPagBaixadas.NumIntDoc AND ParcelasPagBaixadas.NumIntTitulo = TitulosPagBaixados.NumIntDoc AND BaixasParcPag.Status <> ? AND TitulosPagBaixados.FilialEmpresa = ? "
    sWhereN = "WHERE BaixasPag.NumIntBaixa = BaixasParcPag.NumIntBaixa AND BaixasParcPag.NumIntParcela = ParcelasPag.NumIntDoc AND ParcelasPag.NumIntTitulo = TitulosPag.NumIntDoc AND BaixasParcPag.Status <> ? AND TitulosPag.FilialEmpresa = ? "

'    sWhereB = "WHERE BaixasPag.NumIntBaixa = BaixasParcPag.NumIntBaixa AND BaixasParcPag.NumIntParcela = ParcelasPagBaixadas.NumIntDoc AND ParcelasPagBaixadas.NumIntTitulo = TitulosPagBaixados.NumIntDoc AND BaixasParcPag.Status <> ? AND TitulosPagBaixados.FilialEmpresa = ? AND TitulosPagBaixados.Fornecedor = ? AND TitulosPagBaixados.Filial = ?"
'    sWhereN = "WHERE BaixasPag.NumIntBaixa = BaixasParcPag.NumIntBaixa AND BaixasParcPag.NumIntParcela = ParcelasPag.NumIntDoc AND ParcelasPag.NumIntTitulo = TitulosPag.NumIntDoc AND BaixasParcPag.Status <> ? AND TitulosPag.FilialEmpresa = ? AND TitulosPag.Fornecedor = ? AND TitulosPag.Filial = ?"

    If lForn <> 0 Then
       sWhereB = sWhereB & " AND TitulosPagBaixados.Fornecedor = ?"
       sWhereN = sWhereN & " AND TitulosPag.Fornecedor = ?"
    End If
    
    If iFilial <> 0 Then
       sWhereB = sWhereB & " AND TitulosPagBaixados.Filial = ?"
       sWhereN = sWhereN & " AND TitulosPag.Filial >= ?"
    End If

    'Se titulo inicial preenchido
    If (lTituloInic <> 0) Then
       sWhereB = sWhereB & " AND TitulosPagBaixados.NumTitulo >= ?"
       sWhereN = sWhereN & " AND TitulosPag.NumTitulo >= ?"
    End If

    'Se titulo final preenchido
    If (lTituloFim <> 0) Then
       sWhereB = sWhereB & " AND TitulosPagBaixados.NumTitulo <= ?"
       sWhereN = sWhereN & " AND TitulosPag.NumTitulo <= ?"
    End If

    'Se o limite inicial de data de Baixa de titulo estiver preenchido
    If (dtBaixaInic <> DATA_NULA) Then
       sWhereB = sWhereB & " AND BaixasPag.Data >= ?"
       sWhereN = sWhereN & " AND BaixasPag.Data >= ?"
    End If

    'Se o limite final de data de Baixa de titulo estiver preenchido
    If (dtBaixaFim <> DATA_NULA) Then
       sWhereB = sWhereB & " AND BaixasPag.Data <= ?"
       sWhereN = sWhereN & " AND BaixasPag.Data <= ?"
    End If

    'Se o limite inicial de data de vencimento de parcela estiver preenchido
    If (dtVencInic <> DATA_NULA) Then
       sWhereB = sWhereB & " AND ParcelasPagBaixadas.DataVencimento >= ?"
       sWhereN = sWhereN & " AND ParcelasPag.DataVencimento >= ?"
    End If

    'Se o limite final de data de vencimento de parcela estiver preenchido
    If (dtVencFim <> DATA_NULA) Then
       sWhereB = sWhereB & " AND ParcelasPagBaixadas.DataVencimento <= ?"
       sWhereN = sWhereN & " AND ParcelasPag.DataVencimento <= ?"
    End If
  
    sSelect = "SELECT " & sFieldsB & sFromB & sWhereB & " UNION " & "SELECT " & sFieldsN & sFromN & sWhereN & " ORDER BY NumTitulo, NumParcela, Sequencial"

End Sub

Public Function ParcelasPagarBaixadas_Le_Cancelamento2(lComando As Long, sSelect As String, vlNumIntDoc As Variant, vlNumIntParc As Variant, vlNumConta As Variant, viMotivo As Variant, vdValorMulta As Variant, vdValorJuros As Variant, vdValorDesconto As Variant, vdValorBaixado As Variant, vdtDataBaixa As Variant, viSequencial As Variant, viNumParcela As Variant, vsSiglaDocumento As Variant, viNumTitulo As Variant, vdtEmissaoTitulo As Variant, vlNumIntBaixa As Variant, vlFornecedor As Variant, viFilial As Variant) As Long
'Isola a Preparacao da Parte Fixa do SELECT

Dim iRet As Integer
Dim lErro As Long

On Error GoTo Erro_ParcelasPagarBaixadas_Le_Cancelamento2

    iRet = Comando_PrepararInt(lComando, sSelect)
    If (iRet <> AD_SQL_SUCESSO) Then gError 95270
    
    iRet = Comando_BindVarInt(lComando, vlNumIntBaixa)
    If (iRet <> AD_SQL_SUCESSO) Then gError 92661
    
    iRet = Comando_BindVarInt(lComando, vlNumIntDoc)
    If (iRet <> AD_SQL_SUCESSO) Then gError 95364

    iRet = Comando_BindVarInt(lComando, vlNumIntParc)
    If (iRet <> AD_SQL_SUCESSO) Then gError 95329
    
    iRet = Comando_BindVarInt(lComando, vlNumConta)
    If (iRet <> AD_SQL_SUCESSO) Then gError 95312
    
    iRet = Comando_BindVarInt(lComando, viMotivo)
    If (iRet <> AD_SQL_SUCESSO) Then gError 95307

    iRet = Comando_BindVarInt(lComando, vdValorMulta)
    If (iRet <> AD_SQL_SUCESSO) Then gError 95308

    iRet = Comando_BindVarInt(lComando, vdValorJuros)
    If (iRet <> AD_SQL_SUCESSO) Then gError 95309
    
    iRet = Comando_BindVarInt(lComando, vdValorDesconto)
    If (iRet <> AD_SQL_SUCESSO) Then gError 95310

    iRet = Comando_BindVarInt(lComando, vdValorBaixado)
    If (iRet <> AD_SQL_SUCESSO) Then gError 95311
    
    iRet = Comando_BindVarInt(lComando, vdtDataBaixa)
    If (iRet <> AD_SQL_SUCESSO) Then gError 95271

    iRet = Comando_BindVarInt(lComando, viSequencial)
    If (iRet <> AD_SQL_SUCESSO) Then gError 95272

    iRet = Comando_BindVarInt(lComando, viNumParcela)
    If (iRet <> AD_SQL_SUCESSO) Then gError 95273
    
    iRet = Comando_BindVarInt(lComando, vsSiglaDocumento)
    If (iRet <> AD_SQL_SUCESSO) Then gError 95274

    iRet = Comando_BindVarInt(lComando, viNumTitulo)
    If (iRet <> AD_SQL_SUCESSO) Then gError 95275

    iRet = Comando_BindVarInt(lComando, vdtEmissaoTitulo)
    If (iRet <> AD_SQL_SUCESSO) Then gError 95276
    
    iRet = Comando_BindVarInt(lComando, vlFornecedor)
    If (iRet <> AD_SQL_SUCESSO) Then gError 95275

    iRet = Comando_BindVarInt(lComando, viFilial)
    If (iRet <> AD_SQL_SUCESSO) Then gError 95276
    
    ParcelasPagarBaixadas_Le_Cancelamento2 = SUCESSO

    Exit Function

Erro_ParcelasPagarBaixadas_Le_Cancelamento2:

    ParcelasPagarBaixadas_Le_Cancelamento2 = gErr

    Select Case gErr

        Case 92661, 95270 To 95282, 95307 To 95312, 95364
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_PAG", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143276)

    End Select

    Exit Function

End Function

Public Function ParcelasPagarBaixadas_Le_Cancelamento3(lComando As Long, dtBaixaInic As Date, dtBaixaFim As Date, dtVencInic As Date, dtVencFim As Date, lTituloInic As Long, lTituloFim As Long, iStatusBaixa As Integer, iFilialEmpresa As Integer, lFornecedor As Long, iFilial As Integer) As Long
'complementa a passagem dos parametros que variam de acordo com a selecao do usuario e executa o SELECT p/obtencao das parcelas

Dim lErro As Long
Dim iRet As Integer
Dim vFilialEmpresa As Variant
Dim vTituloInic As Variant
Dim vTituloFim As Variant
Dim vBaixaInic As Variant
Dim vBaixaFim As Variant
Dim vVencInic As Variant
Dim vVencFim As Variant

Dim viStatusBaixa As Variant, viFilialEmpresa As Variant
Dim vlFornecedor As Variant, viFilial As Variant

On Error GoTo Erro_ParcelasPagarBaixadas_Le_Cancelamento3
    
    viStatusBaixa = iStatusBaixa
    viFilialEmpresa = iFilialEmpresa
    vlFornecedor = lFornecedor
    viFilial = iFilial
    
    iRet = Comando_BindVarInt(lComando, viStatusBaixa)
    If (iRet <> AD_SQL_SUCESSO) Then gError 95277

    iRet = Comando_BindVarInt(lComando, viFilialEmpresa)
    If (iRet <> AD_SQL_SUCESSO) Then gError 95278

    If vlFornecedor <> 0 Then
        iRet = Comando_BindVarInt(lComando, vlFornecedor)
        If (iRet <> AD_SQL_SUCESSO) Then gError 95279
    End If

    If viFilial <> 0 Then
        iRet = Comando_BindVarInt(lComando, viFilial)
        If (iRet <> AD_SQL_SUCESSO) Then gError 95280
    End If
    
    'Se titulo inicial preenchido
    If (lTituloInic <> 0) Then
        vTituloInic = lTituloInic
        iRet = Comando_BindVarInt(lComando, vTituloInic)
        If (iRet <> AD_SQL_SUCESSO) Then gError 95285
    End If

    'Se titulo final preenchido
    If (lTituloFim <> 0) Then
        vTituloFim = lTituloFim
        iRet = Comando_BindVarInt(lComando, vTituloFim)
        If (iRet <> AD_SQL_SUCESSO) Then gError 95286
    End If

    'Se o limite inicial de data de Baixa de titulo estiver preenchido
    If (dtBaixaInic <> DATA_NULA) Then
        vBaixaInic = dtBaixaInic
        iRet = Comando_BindVarInt(lComando, vBaixaInic)
        If (iRet <> AD_SQL_SUCESSO) Then gError 95287
    End If

    'Se o limite final de data de Baixa de titulo estiver preenchido
    If (dtBaixaFim <> DATA_NULA) Then
        vBaixaFim = dtBaixaFim
        iRet = Comando_BindVarInt(lComando, vBaixaFim)
        If (iRet <> AD_SQL_SUCESSO) Then gError 95288
    End If

    'Se o limite inicial de data de vencimento de parcela estiver preenchido
    If (dtVencInic <> DATA_NULA) Then
        vVencInic = dtVencInic
        iRet = Comando_BindVarInt(lComando, vVencInic)
        If (iRet <> AD_SQL_SUCESSO) Then gError 95289
    End If

    'Se o limite final de data de vencimento de parcela estiver preenchido
    If (dtVencFim <> DATA_NULA) Then
        vVencFim = dtVencFim
        iRet = Comando_BindVarInt(lComando, vVencFim)
        If (iRet <> AD_SQL_SUCESSO) Then gError 95290
    End If
    
    iRet = Comando_BindVarInt(lComando, viStatusBaixa)
    If (iRet <> AD_SQL_SUCESSO) Then gError 95277

    iRet = Comando_BindVarInt(lComando, viFilialEmpresa)
    If (iRet <> AD_SQL_SUCESSO) Then gError 95278

    If vlFornecedor <> 0 Then
        iRet = Comando_BindVarInt(lComando, vlFornecedor)
        If (iRet <> AD_SQL_SUCESSO) Then gError 95279
    End If

    If viFilial <> 0 Then
        iRet = Comando_BindVarInt(lComando, viFilial)
        If (iRet <> AD_SQL_SUCESSO) Then gError 95280
    End If
    
    'Se titulo inicial preenchido
    If (lTituloInic <> 0) Then
        vTituloInic = lTituloInic
        iRet = Comando_BindVarInt(lComando, vTituloInic)
        If (iRet <> AD_SQL_SUCESSO) Then gError 95300
    End If

    'Se titulo final preenchido
    If (lTituloFim <> 0) Then
        vTituloFim = lTituloFim
        iRet = Comando_BindVarInt(lComando, vTituloFim)
        If (iRet <> AD_SQL_SUCESSO) Then gError 95301
    End If

    'Se o limite inicial de data de Baixa de titulo estiver preenchido
    If (dtBaixaInic <> DATA_NULA) Then
        vBaixaInic = dtBaixaInic
        iRet = Comando_BindVarInt(lComando, vBaixaInic)
        If (iRet <> AD_SQL_SUCESSO) Then gError 95302
    End If

    'Se o limite final de data de Baixa de titulo estiver preenchido
    If (dtBaixaFim <> DATA_NULA) Then
        vBaixaFim = dtBaixaFim
        iRet = Comando_BindVarInt(lComando, vBaixaFim)
        If (iRet <> AD_SQL_SUCESSO) Then gError 95303
    End If

    'Se o limite inicial de data de vencimento de parcela estiver preenchido
    If (dtVencInic <> DATA_NULA) Then
        vVencInic = dtVencInic
        iRet = Comando_BindVarInt(lComando, vVencInic)
        If (iRet <> AD_SQL_SUCESSO) Then gError 95304
    End If

    'Se o limite final de data de vencimento de parcela estiver preenchido
    If (dtVencFim <> DATA_NULA) Then
        vVencFim = dtVencFim
        iRet = Comando_BindVarInt(lComando, vVencFim)
        If (iRet <> AD_SQL_SUCESSO) Then gError 95305
    End If
    
    iRet = Comando_ExecutarInt(lComando)
    If (iRet <> AD_SQL_SUCESSO) Then gError 95306

    ParcelasPagarBaixadas_Le_Cancelamento3 = SUCESSO

    Exit Function
    
Erro_ParcelasPagarBaixadas_Le_Cancelamento3:

    ParcelasPagarBaixadas_Le_Cancelamento3 = gErr

    Select Case gErr

        Case 95284 To 95291, 95300 To 95306
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_PAG", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143277)

    End Select

    Exit Function

End Function

Public Sub ParcelasPagarBaixadas_Le_Cancelamento4(colInfoParcPag As Collection, lNumIntDoc As Long, lNumIntParc As Long, lNumConta As Long, iMotivo As Integer, dValorMulta As Double, dValorJuros As Double, dValorDesconto As Double, dValorBaixado As Double, dtDataBaixa As Date, iSequencial As Integer, iNumParcela As Integer, sSiglaDocumento As String, lNumTitulo As Long, dtEmissaoTitulo As Date, lFornecedor As Long, iFilial As Integer, lNumIntBaixa As Long)
' inclui a parcela lida na colecao

Dim objInfoParcPag As ClassInfoParcPag

        Set objInfoParcPag = New ClassInfoParcPag

        objInfoParcPag.iFilialForn = iFilial
        objInfoParcPag.dtDataEmissao = dtEmissaoTitulo
        objInfoParcPag.iNumParcela = iNumParcela
        objInfoParcPag.lFornecedor = lFornecedor
        objInfoParcPag.lNumTitulo = lNumTitulo
        objInfoParcPag.dtDataVencimento = dtDataBaixa
        objInfoParcPag.sSiglaDocumento = sSiglaDocumento
        objInfoParcPag.iSequencial = iSequencial
        objInfoParcPag.iMotivo = iMotivo
        objInfoParcPag.dValorDesconto = dValorDesconto
        objInfoParcPag.dValorMulta = dValorMulta
        objInfoParcPag.dValorJuros = dValorJuros
        objInfoParcPag.dValor = dValorBaixado
        objInfoParcPag.lNumMovCta = lNumConta
        objInfoParcPag.lNumIntParc = lNumIntParc
        objInfoParcPag.lNumIntDoc = lNumIntDoc
        objInfoParcPag.lNumIntBaixa = lNumIntBaixa

        colInfoParcPag.Add objInfoParcPag

End Sub


Public Sub GridParcelas_Click()

Dim iExecutaEntradaCelula As Integer
Dim objParcPagBaixa As New ClassBaixaParcPagar
Dim objBaixaPagar As New ClassBaixaPagar
Dim lErro As Long

On Error GoTo Erro_GridParcelas_Click
    
    If objGridParcelas.iLinhasExistentes > 0 Then
    
        Call Grid_Click(objGridParcelas, iExecutaEntradaCelula)
    
        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
        End If
        
        lErro = Traz_Dados_Baixa(GridParcelas.Row)
        If lErro <> SUCESSO Then gError 95365
            
    End If
            
    Exit Sub

Erro_GridParcelas_Click:
    
    Select Case gErr
    
        Case 95365
            '???
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143278)
    
    End Select
    
    Exit Sub
            
End Sub

Public Sub GridParcelas_GotFocus()
    Call Grid_Recebe_Foco(objGridParcelas)
End Sub

Public Sub GridParcelas_EnterCell()
    Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
End Sub

Public Sub GridParcelas_LeaveCell()
    Call Saida_Celula(objGridParcelas)
End Sub

Public Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridParcelas)
End Sub

Public Sub GridParcelas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Public Sub GridParcelas_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridParcelas)
End Sub

Public Sub GridParcelas_RowColChange()
    
    Call Grid_RowColChange(objGridParcelas)
    
End Sub

Public Sub GridParcelas_Scroll()
    Call Grid_Scroll(objGridParcelas)
End Sub

Public Sub Fornec_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub FilialFornec_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataEmissao_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Tipo_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Numero_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Parcela_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Sequencial_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoBaixa_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub


Public Sub DataEmissao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)

End Sub

Public Sub Fornec_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub FilialFornec_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Tipo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Numero_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Parcela_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Sequencial_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub TipoBaixa_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub Grid_Parcelas_Preenche(gcolInfoParcPag As Collection)

Dim lErro As Long
Dim iIndice As Integer
Dim objInfoParcPag As New ClassInfoParcPag
Dim objFornecedor As ClassFornecedor
Dim objFilialForn As ClassFilialFornecedor

On Error GoTo Erro_Grid_Parcelas_Preenche

    'Se o número de parcelas for maior que o número de linhas do Grid
    If gcolInfoParcPag.Count + 1 > GridParcelas.Rows Then
    
        'Altera o número de linhas do Grid de acordo com o número de parcelas
        GridParcelas.Rows = gcolInfoParcPag.Count + 1
        
        'Chama rotina de Inicialização do Grid
        Call Grid_Inicializa(objGridParcelas)

    End If

    iIndice = 0

    'Percorre todas as Parcelas da Coleção
    For Each objInfoParcPag In gcolInfoParcPag

        iIndice = iIndice + 1

        'Passa para a tela os dados da Parcela em questão
        GridParcelas.TextMatrix(iIndice, iGrid_Filial_Col) = Filial.Text
        If objInfoParcPag.dtDataEmissao <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_DataBaixa_Col) = Format(objInfoParcPag.dtDataEmissao, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iIndice, iGrid_Parcela_Col) = objInfoParcPag.iNumParcela
        GridParcelas.TextMatrix(iIndice, iGrid_Fornecedor_Col) = Fornecedor.Text
        GridParcelas.TextMatrix(iIndice, iGrid_Numero_Col) = objInfoParcPag.lNumTitulo
        If objInfoParcPag.dtDataVencimento <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_DataBaixa_Col) = Format(objInfoParcPag.dtDataVencimento, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iIndice, iGrid_Tipo_Col) = objInfoParcPag.sSiglaDocumento
        GridParcelas.TextMatrix(iIndice, iGrid_Sequencial_Col) = objInfoParcPag.iSequencial
        
        '#################################################
        'Inserido por Wagner
        Set objFornecedor = New ClassFornecedor
        Set objFilialForn = New ClassFilialFornecedor
        
        objFornecedor.lCodigo = objInfoParcPag.lFornecedor
        
        objFilialForn.lCodFornecedor = objFornecedor.lCodigo
        objFilialForn.iCodFilial = objInfoParcPag.iFilialForn
            
        'le o nome reduzido do cliente
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 And lErro <> 12732 Then gError 181431
            
        'le o nome reduzido da filial  cliente
        lErro = CF("FilialFornecedor_Le", objFilialForn)
        If lErro <> SUCESSO And lErro <> 12929 Then gError 181432
            
        GridParcelas.TextMatrix(iIndice, iGrid_Fornecedor_Col) = objFornecedor.lCodigo & SEPARADOR & objFornecedor.sNomeReduzido
        GridParcelas.TextMatrix(iIndice, iGrid_Filial_Col) = objFilialForn.iCodFilial & SEPARADOR & objFilialForn.sNome
        '#################################################
        
        Select Case objInfoParcPag.iMotivo
        
            Case MOTIVO_PAGAMENTO
                GridParcelas.TextMatrix(iIndice, iGrid_TipoBaixa_Col) = "Pagamento"
            Case MOTIVO_PAGTO_ANTECIPADO
                GridParcelas.TextMatrix(iIndice, iGrid_TipoBaixa_Col) = "Adiantamento para Fornecedor"
            Case MOTIVO_CREDITO_FORNECEDOR
                GridParcelas.TextMatrix(iIndice, iGrid_TipoBaixa_Col) = "Devolução / Crédito com Fornecedor"
                
        End Select
        
    Next
    
    'Passa para o Obj o número de Parcelas passadas pela Coleção
    objGridParcelas.iLinhasExistentes = gcolInfoParcPag.Count

    Call Grid_Refresh_Checkbox(objGridParcelas)
    
    Exit Sub

Erro_Grid_Parcelas_Preenche:

    Select Case gErr
    
        Case 181431, 181432

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143279)

    End Select

    Exit Sub

End Sub

Private Function Traz_Dados_Baixa(iIndice As Integer) As Long
'Mostra na tela os dados da baixa

Dim lErro As Long
Dim iIndiceFrame As Integer

On Error GoTo Erro_Traz_Dados_Baixa

    'Coloca os dados da Baixa na tela
    Desconto.Caption = Format(gcolInfoParcPag.Item(iIndice).dValorDesconto, "Standard")
    ValorPago.Caption = Format(gcolInfoParcPag.Item(iIndice).dValor - gcolInfoParcPag.Item(iIndice).dValorDesconto + gcolInfoParcPag.Item(iIndice).dValorMulta + gcolInfoParcPag.Item(iIndice).dValorJuros, "Standard")
    Multa.Caption = Format(gcolInfoParcPag.Item(iIndice).dValorMulta, "Standard")
    ValorBaixado.Caption = Format(gcolInfoParcPag.Item(iIndice).dValor, "Standard")
    Juros.Caption = Format(gcolInfoParcPag.Item(iIndice).dValorJuros, "Standard")
    DataBaixa.Caption = Format(gcolInfoParcPag.Item(iIndice).dtDataVencimento, "dd/mm/yyyy")

    'Torna os frames invisiveis a fim de só tornar visivel o frame correspondente
    For iIndiceFrame = 1 To 3
        FramePagamento(iIndiceFrame).Visible = False
    Next
        
    If gcolInfoParcPag.Item(iIndice).iMotivo = MOTIVO_PAGAMENTO Then
        
        'Torna visivel o frame
        FramePagamento(3).Visible = True
        
        'Traz os dados do pagamento
        lErro = Traz_Dados_Pagamento(iIndice)
        If lErro <> SUCESSO Then gError 95312

    ElseIf gcolInfoParcPag.Item(iIndice).iMotivo = MOTIVO_PAGTO_ANTECIPADO Then

        'Torna visivel o frame
        FramePagamento(1).Visible = True
        
        'Traz os dados do pagamento antecipado
        lErro = Traz_Dados_Pagamento_Antecipado(iIndice)
        If lErro <> SUCESSO Then gError 95313

    ElseIf gcolInfoParcPag.Item(iIndice).iMotivo = MOTIVO_CREDITO_FORNECEDOR Then

        'Torna visivel o frame
        FramePagamento(2).Visible = True
        
        'Traz os dados do crédito
        lErro = Traz_Dados_Credito_Fornecedor(iIndice)
        If lErro <> SUCESSO Then gError 95314

    End If
    
    Traz_Dados_Baixa = SUCESSO

    Exit Function

Erro_Traz_Dados_Baixa:

    Traz_Dados_Baixa = gErr

    Select Case gErr

        Case 95312, 95314, 95313

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143280)

    End Select

    Exit Function

End Function

Private Function Traz_Dados_Pagamento(iIndice As Integer) As Long
'Mostra os dados de pagamento na tela

Dim lErro As Long
Dim objMovCCI As New ClassMovContaCorrente
Dim objContaCorrente As New ClassContasCorrentesInternas
Dim objBorderoPagto As New ClassBorderoPagto
Dim objPortador As New ClassPortador

On Error GoTo Erro_Traz_Dados_Pagamento

    objMovCCI.lNumMovto = gcolInfoParcPag.Item(iIndice).lNumMovCta

    'Lê o Movimento
    lErro = CF("MovContaCorrente_Le", objMovCCI)
    If lErro <> SUCESSO And lErro <> 11893 Then gError 95316

    'Se não encontrou Movimento --> erro
    If lErro = 11893 Then gError 95317

    'Lê a Conta Corrente
    lErro = CF("ContaCorrenteInt_Le", objMovCCI.iCodConta, objContaCorrente)
    If lErro <> SUCESSO And lErro <> 11807 Then gError 95318

    'Se não encontrou a conta Corrente --> erro
    If lErro <> SUCESSO Then gError 95315
    
    'Coloca os dados na tela
    ContaCorrente.Caption = objContaCorrente.sNomeReduzido
    ValorPagoPagto.Caption = Format(objMovCCI.dValor, "Standard")
    Historico.Caption = objMovCCI.sHistorico
    NumOuSequencial.Caption = IIf(objMovCCI.lNumero <> 0, CStr(objMovCCI.lNumero), "")

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

    Traz_Dados_Pagamento = gErr

    Select Case gErr

        Case 95316, 95318

        Case 95317
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_CADASTRADO3", gErr, objMovCCI.lNumMovto)

        Case 95315
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", gErr, objContaCorrente.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143281)

    End Select

    Exit Function

End Function

Private Function Traz_Dados_Credito_Fornecedor(iIndice As Integer) As Long

Dim lErro As Long
Dim objCreditoPag As New ClassCreditoPagar

On Error GoTo Erro_Traz_Dados_Credito_Fornecedor

    objCreditoPag.lNumIntDoc = gcolInfoParcPag.Item(iIndice).lNumIntDoc

    'Lê o Crédito Pagar
    lErro = CF("CreditoPagar_Le", objCreditoPag)
    If lErro <> AD_SQL_SUCESSO And 17071 Then gError 95321
    If lErro <> SUCESSO Then gError 95322

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

    Traz_Dados_Credito_Fornecedor = gErr

    Select Case gErr

        Case 95321

        Case 95322
            Call Rotina_Erro(vbOKOnly, "ERRO_CREDITO_PAG_FORN_INEXISTENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143282)

    End Select

    Exit Function

End Function

Private Function Traz_Dados_Pagamento_Antecipado(iIndice) As Long

Dim lErro As Long
Dim objMovCCI As New ClassMovContaCorrente
Dim objCCI As New ClassContasCorrentesInternas
Dim objAntecipPag As New ClassAntecipPag

On Error GoTo Erro_Traz_Dados_Pagamento_Antecipado

    objMovCCI.lNumMovto = gcolInfoParcPag.Item(iIndice).lNumMovCta

    'Lê o movimento  da Baixa
    lErro = CF("MovContaCorrente_Le", objMovCCI)
    If lErro <> SUCESSO And lErro <> 11893 Then gError 95323
    If lErro = 11893 Then gError 95324 'Não encontrou

    'Lê a Conta Corrente
    lErro = CF("ContaCorrenteInt_Le", objMovCCI.iCodConta, objCCI)
    If lErro <> SUCESSO And lErro <> 11807 Then gError 95325
    If lErro <> SUCESSO Then gError 95326 'Não encontrou

    objAntecipPag.lNumMovto = objMovCCI.lNumMovto

    lErro = CF("AntecipPag_Le_NumMovto", objAntecipPag)
    If lErro <> AD_SQL_SUCESSO And lErro <> 42845 Then gError 95327
    If lErro = 42845 Then gError 95328

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

    Traz_Dados_Pagamento_Antecipado = gErr

    Select Case gErr

        Case 95323, 95325, 95327

        Case 95324
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_CADASTRADO3", gErr, objMovCCI.lNumMovto)

        Case 95326
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", gErr, objCCI.iCodigo)

        Case 95328
            Call Rotina_Erro(vbOKOnly, "ERRO_PAGTO_ANTECIPADO_INEXISTENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143283)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim iIndice As Integer
Dim colParcSel As New Collection
Dim lErro As Long
Dim objMovContaCorrente As ClassMovContaCorrente
Dim colInfoParcPag As Collection
Dim objInfoParcPag As ClassInfoParcPag
Dim objInfoParcPag1 As ClassInfoParcPag
Dim iJaIncluido As Integer
Dim colInfoParcPag1 As New Collection
Dim sParcelas As String
Dim vbMsgRes As VbMsgBoxResult
Dim objTela As Object
'Dim objContabil As ClassContabil 'Comentado por Leo em 12/12/01

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Preenche a colecao com as parcelas a baixar
    For iIndice = 1 To gcolInfoParcPag.Count
        
        'Se estiver selecionado para cancelar => adiciona na colecao
        If StrParaInt(GridParcelas.TextMatrix(iIndice, iGrid_Seleciona_Col)) = vbChecked Then
            
            'se a contabilidade estiver ativa,
            'tenta descobrir as baixas manuais pois a transação (ou seja todas as parcelas envolvidas na transacao) precisa ser cancelada como um todo em função da contabilidade associada
            If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
            
                Set objInfoParcPag = gcolInfoParcPag.Item(iIndice)
                
                If objInfoParcPag.iMotivo = MOTIVO_PAGAMENTO Then
                
                    Set objMovContaCorrente = New ClassMovContaCorrente
                
                    objMovContaCorrente.lNumMovto = objInfoParcPag.lNumMovCta

                    'Le o movimento de conta corrente passado como parametro
                    lErro = CF("MovContaCorrente_Le", objMovContaCorrente)
                    If lErro <> SUCESSO And lErro <> 11893 Then gError 92662
    
                    'se o movimento de conta corrente não estiver cadastrado ==> erro
                    If lErro = 11893 Then gError 92663
    
                    'se foi baixado por pagamento em dinheiro ==> baixa manual ==> tem que pesquisar as parcelas que foram baixadas junto para que também elas tenham sua baixa cancelada já que a contabilização é uma só para a transação
                    If objMovContaCorrente.iTipo = MOVCCI_PAGTO_TITULO_POR_DINHEIRO Then
                                    
                        Set colInfoParcPag = New Collection
                                    
                        'le as parcelas referentes a baixa (lNumIntBaixa) passada como parametro com excecao de lNumIntParceçla/iSequencial e coloca-as na coleção
                        'se encontrar alguma parcela com a baixa já excluida(cancelada) ===> erro.
                        lErro = ParcelasPagarBaixadas_Le(objInfoParcPag.lNumIntBaixa, objInfoParcPag.lNumIntParc, objInfoParcPag.iSequencial, objInfoParcPag.lNumTitulo, objInfoParcPag.iNumParcela, colInfoParcPag)
                        If lErro <> SUCESSO Then gError 92669
                                    
                        For Each objInfoParcPag1 In colInfoParcPag
                            colParcSel.Add objInfoParcPag1
                        Next
                                    
                        If colInfoParcPag.Count > 0 Then colInfoParcPag1.Add objInfoParcPag
                                    
                    End If
                    
                'se foi baixado utilizando pagamento antecipado ou credito junto ao fornecedor ==> baixa manual ==> tem que pesquisar as parcelas que foram baixadas junto para que também elas tenham sua baixa cancelada já que a contabilização é uma só para a transação
                ElseIf objInfoParcPag.iMotivo = MOTIVO_PAGTO_ANTECIPADO Or objInfoParcPag.iMotivo = MOTIVO_CREDITO_FORNECEDOR Then
            
                        Set colInfoParcPag = New Collection
                                    
                        'le as parcelas referentes a baixa (lNumIntBaixa) passada como parametro com excecao de lNumIntParceçla/iSequencial e coloca-as na coleção
                        'se encontrar alguma parcela com a baixa já excluida(cancelada) ===> erro.
                        lErro = ParcelasPagarBaixadas_Le(objInfoParcPag.lNumIntBaixa, objInfoParcPag.lNumIntParc, objInfoParcPag.iSequencial, objInfoParcPag.lNumTitulo, objInfoParcPag.iNumParcela, colInfoParcPag)
                        If lErro <> SUCESSO Then gError 92669
                                    
                        For Each objInfoParcPag1 In colInfoParcPag
                            colParcSel.Add objInfoParcPag1
                        Next
                        
                        If colInfoParcPag.Count > 0 Then colInfoParcPag1.Add objInfoParcPag
            
                Else
                    gError 92670
            
                End If
                
                iJaIncluido = 0
                For Each objInfoParcPag1 In colParcSel
                    If objInfoParcPag.lNumIntParc = objInfoParcPag1.lNumIntParc And objInfoParcPag.iSequencial = objInfoParcPag1.iSequencial Then
                        iJaIncluido = 1
                        Exit For
                    End If
                Next
                
                If iJaIncluido = 0 Then colParcSel.Add objInfoParcPag
                
            Else
                
                colParcSel.Add gcolInfoParcPag.Item(iIndice)
            
            End If
            
        End If
            
    Next
    
    Set objTela = Me
    
    'se estão sendo canceladas parcelas que não foram selecionadas mas que foram agregadas por terem sido baixadas na mesma transação de outra parcela para a qual foi pedida o cancelamento devido a contabilidade ==> avisa
    If colInfoParcPag1.Count > 0 Then
    
        For Each objInfoParcPag1 In colInfoParcPag1
            sParcelas = sParcelas & CRLF & "Titulo = " & CStr(objInfoParcPag1.lNumTitulo) & ", Parcela = " & CStr(objInfoParcPag1.iNumParcela) & ", Sequencial = " & CStr(objInfoParcPag1.iSequencial)
        Next
    
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CANCELA_PARCPAG_BAIXAMANUAL", sParcelas)

        If vbMsgRes = vbYes Then
            lErro = CF("BaixaPagtosCancelar_Grava", colParcSel, objTela) 'Parâmetros Alterados por Leo em 12/12/01
            If lErro <> SUCESSO Then gError 92681
        End If
    
    Else
    
        lErro = CF("BaixaPagtosCancelar_Grava", colParcSel, objTela) 'Parâmetros Alterados por Leo em 12/12/01
        If lErro <> SUCESSO Then gError 95363
    
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 92662, 92669, 92681, 95363
    
        Case 92663
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVCTA_BAIXAPAG_NAO_CADASTRADA", gErr, objInfoParcPag.lNumMovCta, objInfoParcPag.lNumIntBaixa)
    
        Case 92670
            Call Rotina_Erro(vbOKOnly, "ERRO_MOTIVO_BAIXA_INEXISTENTE", gErr, objInfoParcPag.lNumTitulo, objInfoParcPag.iNumParcela, objInfoParcPag.iSequencial, objInfoParcPag.iMotivo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB:", gErr, Error, 143284)
    
    End Select
    
    Exit Function
    
End Function

Public Function ParcelasPagarBaixadas_Le(ByVal lNumIntBaixa As Long, ByVal lNumIntParcela As Long, ByVal iSequencial As Integer, ByVal lNumTitulo As Long, ByVal iNumParcela As Integer, colInfoParcPag As Collection) As Long
'le as parcelas referentes a baixa (lNumIntBaixa) passada como parametro com excecao de lNumIntParceçla/iSequencial e coloca-as na coleção
'se encontrar alguma parcela com a baixa já excluida(cancelada) ===> erro.

Dim sFromN As String, lErro As Long
Dim sFromB As String
Dim sWhereN As String
Dim sWhereB As String
Dim sFieldsN As String
Dim sFieldsB As String, sSelect As String
Dim tInfoParcPag As typeInfoParcPag
Dim iStatus As Integer, lComando As Long
Dim objInfoParcPag As ClassInfoParcPag

On Error GoTo Erro_ParcelasPagarBaixadas_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 92664

    sFieldsB = "BaixasParcPag.Status, BaixasPag.NumIntBaixa, BaixasPag.NumIntDoc, BaixasParcPag.NumIntParcela, BaixasPag.NumMovCta, BaixasPag.Motivo, BaixasParcPag.ValorMulta, BaixasParcPag.ValorJuros, BaixasParcPag.ValorDesconto, BaixasParcPag.ValorBaixado, BaixasPag.Data, BaixasParcPag.Sequencial, ParcelasPagBaixadas.NumParcela, TitulosPagBaixados.SiglaDocumento, TitulosPagBaixados.NumTitulo, TitulosPagBaixados.DataEmissao "
    sFieldsN = "BaixasParcPag.Status, BaixasPag.NumIntBaixa, BaixasPag.NumIntDoc, BaixasParcPag.NumIntParcela, BaixasPag.NumMovCta, BaixasPag.Motivo, BaixasParcPag.ValorMulta, BaixasParcPag.ValorJuros, BaixasParcPag.ValorDesconto, BaixasParcPag.ValorBaixado, BaixasPag.Data, BaixasParcPag.Sequencial, ParcelasPag.NumParcela, TitulosPag.SiglaDocumento, TitulosPag.NumTitulo, TitulosPag.DataEmissao "

    sFromB = "FROM BaixasPag, BaixasParcPag, ParcelasPagBaixadas, TitulosPagBaixados "
    sFromN = "FROM BaixasPag, BaixasParcPag, ParcelasPag, TitulosPag "

    sWhereB = "WHERE BaixasPag.NumIntBaixa = BaixasParcPag.NumIntBaixa AND BaixasParcPag.NumIntParcela = ParcelasPagBaixadas.NumIntDoc AND ParcelasPagBaixadas.NumIntTitulo = TitulosPagBaixados.NumIntDoc AND BaixasPag.NumIntBaixa = ? AND (BaixasParcPag.NumIntParcela <> ? OR BaixasParcPag.Sequencial <> ?)"
    sWhereN = "WHERE BaixasPag.NumIntBaixa = BaixasParcPag.NumIntBaixa AND BaixasParcPag.NumIntParcela = ParcelasPag.NumIntDoc AND ParcelasPag.NumIntTitulo = TitulosPag.NumIntDoc AND BaixasPag.NumIntBaixa =  ? AND (BaixasParcPag.NumIntParcela <> ? OR BaixasParcPag.Sequencial <> ?)"
  
    sSelect = "SELECT " & sFieldsB & sFromB & sWhereB & " UNION " & "SELECT " & sFieldsN & sFromN & sWhereN & " ORDER BY NumTitulo, NumParcela, Sequencial"

    With tInfoParcPag

        .sSiglaDocumento = String(STRING_SIGLA_DOCUMENTO, 0)

        lErro = Comando_Executar(lComando, sSelect, iStatus, .lNumIntBaixa, .lNumIntDoc, .lNumIntParc, .lNumMovCta, .iMotivo, .dValorMulta, .dValorJuros, .dValorDesconto, .dValor, .dtDataVencimento, .iSequencial, .iNumParcela, .sSiglaDocumento, .lNumTitulo, .dtDataEmissao, lNumIntBaixa, lNumIntParcela, iSequencial, lNumIntBaixa, lNumIntParcela, iSequencial)
        If lErro <> AD_SQL_SUCESSO Then gError 92665
        
    End With
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92666
        
    Do While lErro = AD_SQL_SUCESSO

        If iStatus = STATUS_EXCLUIDO Then gError 92668

        Set objInfoParcPag = New ClassInfoParcPag

        With tInfoParcPag

            objInfoParcPag.dtDataEmissao = .dtDataEmissao
            objInfoParcPag.iNumParcela = .iNumParcela
            objInfoParcPag.lNumTitulo = .lNumTitulo
            objInfoParcPag.iSequencial = .iSequencial
            objInfoParcPag.dtDataVencimento = .dtDataVencimento
            objInfoParcPag.sSiglaDocumento = .sSiglaDocumento
            objInfoParcPag.iMotivo = .iMotivo
            objInfoParcPag.dValorDesconto = .dValorDesconto
            objInfoParcPag.dValorMulta = .dValorMulta
            objInfoParcPag.dValorJuros = .dValorJuros
            objInfoParcPag.dValor = .dValor
            objInfoParcPag.lNumMovCta = .lNumMovCta
            objInfoParcPag.lNumIntParc = .lNumIntParc
            objInfoParcPag.lNumIntDoc = .lNumIntDoc
            objInfoParcPag.lNumIntBaixa = .lNumIntBaixa

        End With

        colInfoParcPag.Add objInfoParcPag

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92667

    Loop
    
    Call Comando_Fechar(lComando)
    
    ParcelasPagarBaixadas_Le = SUCESSO

    Exit Function
    
Erro_ParcelasPagarBaixadas_Le:

    ParcelasPagarBaixadas_Le = gErr

    Select Case gErr

        Case 92664
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92665, 92666, 92667
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_PAG", gErr)

        Case 92668
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCPAG_JA_CANCELADO", gErr, lNumTitulo, iNumParcela, iSequencial)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143285)

    End Select

    Exit Function
    
End Function

Private Sub Fornecedor_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Preenche
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134051

    Exit Sub

Erro_Fornecedor_Preenche:

    Select Case gErr

        Case 134051

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143286)

    End Select
    
    Exit Sub

End Sub

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor, Optional objContexto As Object) As Long

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor, sContaTela As String
Dim objFilial As New ClassFilialFornecedor, objConta As New ClassContasCorrentesInternas
Dim objInfoParcPag As ClassInfoParcPag
Dim colParcSel As Collection
Dim sMsg As String

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case FORNECEDOR_NOME_RED

            'Preenche NomeReduzido com o fornecedor da tela
            If Len(Trim(Fornecedor.Text)) > 0 Then

                objMnemonicoValor.colValor.Add Fornecedor.Text

            Else

                objMnemonicoValor.colValor.Add ""

            End If

        Case FILIAL_NOME_RED

            If Len(Filial.Text) > 0 Then

                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then gError 193794

                objMnemonicoValor.colValor.Add objFilial.sNome

            Else

                objMnemonicoValor.colValor.Add ""

            End If


        Case PARCELAS
        
                Set colParcSel = objContexto
        
                For Each objInfoParcPag In colParcSel
        
                    sMsg = sMsg & "Tít: " & CStr(objInfoParcPag.lNumTitulo) & " Parc: " & CStr(objInfoParcPag.iNumParcela) & " Seq: " & CStr(objInfoParcPag.iSequencial) & " - "
        
                Next

                objMnemonicoValor.colValor.Add sMsg


        Case Else
            gError 193795

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 193794

        Case 193795
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193796)

    End Select

    Exit Function

End Function

