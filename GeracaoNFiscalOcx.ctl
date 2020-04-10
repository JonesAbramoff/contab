VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl GeracaoNFiscalOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5355
      Index           =   2
      Left            =   165
      TabIndex        =   30
      Top             =   465
      Visible         =   0   'False
      Width           =   9180
      Begin VB.Frame Frame14 
         Caption         =   "Opções"
         Height          =   945
         Left            =   0
         TabIndex        =   128
         Top             =   4425
         Width           =   9150
         Begin VB.CommandButton BotaoNFiscalImprime 
            Caption         =   "Gerar Nota Fiscal e Imprimir"
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
            Left            =   5850
            TabIndex        =   28
            Top             =   540
            Width           =   3225
         End
         Begin VB.CommandButton BotaoNFiscal 
            Caption         =   "Gerar Nota Fiscal"
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
            Left            =   75
            TabIndex        =   26
            Top             =   540
            Width           =   3225
         End
         Begin VB.CommandButton BotaoNFiscalFatura 
            Caption         =   "Gerar Nota Fiscal Fatura"
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
            Left            =   90
            TabIndex        =   25
            Top             =   210
            Width           =   3225
         End
         Begin VB.CommandButton BotaoNFiscalFaturaImprime 
            Caption         =   "Gerar Nota Fiscal Fatura e Imprimir"
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
            Left            =   5850
            TabIndex        =   27
            Top             =   210
            Width           =   3225
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Pedidos"
         Height          =   3705
         Left            =   0
         TabIndex        =   112
         Top             =   660
         Width           =   9150
         Begin VB.CommandButton BotaoExportar 
            Caption         =   "Exportar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   7395
            TabIndex        =   24
            Top             =   3030
            Visible         =   0   'False
            Width           =   1680
         End
         Begin VB.CommandButton BotaoMarcarTodos 
            Caption         =   "Marcar Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   60
            Picture         =   "GeracaoNFiscalOcx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   3030
            Width           =   1680
         End
         Begin VB.CommandButton BotaoDesmarcarTodos 
            Caption         =   "Desmarcar Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1893
            Picture         =   "GeracaoNFiscalOcx.ctx":101A
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   3030
            Width           =   1680
         End
         Begin VB.CommandButton BotaoPedido 
            Caption         =   "Editar Pedido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3720
            Picture         =   "GeracaoNFiscalOcx.ctx":21FC
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   3030
            Width           =   1680
         End
         Begin VB.CommandButton BotaoImprimirPI 
            Caption         =   "Pedido Interno"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   5565
            Picture         =   "GeracaoNFiscalOcx.ctx":2E7A
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Imprimir Pedido Interno"
            Top             =   3030
            Visible         =   0   'False
            Width           =   1680
         End
         Begin VB.TextBox Filial 
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            Enabled         =   0   'False
            Height          =   225
            Left            =   3915
            TabIndex        =   127
            Text            =   "Filial"
            Top             =   1005
            Width           =   540
         End
         Begin VB.TextBox FilialEmpresa 
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            Enabled         =   0   'False
            Height          =   225
            Left            =   6540
            TabIndex        =   124
            Text            =   "Filial"
            Top             =   2010
            Width           =   1560
         End
         Begin VB.TextBox Cidade 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   6450
            TabIndex        =   123
            Text            =   "Cidade"
            Top             =   1365
            Width           =   1680
         End
         Begin VB.TextBox Estado 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   6345
            TabIndex        =   122
            Text            =   "Estado"
            Top             =   1080
            Width           =   1665
         End
         Begin VB.TextBox NomeReduzido 
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            Enabled         =   0   'False
            Height          =   225
            Left            =   2475
            TabIndex        =   121
            Text            =   "Nome Reduzido"
            Top             =   720
            Width           =   2040
         End
         Begin VB.TextBox DataEntrega 
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            Enabled         =   0   'False
            Height          =   225
            Left            =   6225
            TabIndex        =   120
            Text            =   "Entrega"
            Top             =   705
            Width           =   1095
         End
         Begin VB.CheckBox GeraNFiscal 
            DragMode        =   1  'Automatic
            Height          =   210
            Left            =   150
            TabIndex        =   119
            Top             =   735
            Width           =   816
         End
         Begin VB.TextBox Pedido 
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            Enabled         =   0   'False
            Height          =   225
            Left            =   570
            TabIndex        =   118
            Text            =   "Pedido"
            Top             =   735
            Width           =   852
         End
         Begin VB.TextBox Cliente 
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            Enabled         =   0   'False
            Height          =   225
            Left            =   1440
            TabIndex        =   117
            Text            =   "Cliente"
            Top             =   735
            Width           =   1080
         End
         Begin VB.TextBox DataEmissao 
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            Enabled         =   0   'False
            Height          =   225
            Left            =   4935
            TabIndex        =   116
            Text            =   "Emissão"
            Top             =   750
            Width           =   1050
         End
         Begin VB.ComboBox Ordenados 
            Height          =   315
            ItemData        =   "GeracaoNFiscalOcx.ctx":2F7C
            Left            =   2790
            List            =   "GeracaoNFiscalOcx.ctx":2F7E
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   150
            Width           =   3480
         End
         Begin VB.TextBox Bairro 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   6435
            TabIndex        =   115
            Text            =   "Bairro"
            Top             =   1695
            Width           =   1665
         End
         Begin VB.TextBox Transportadora 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   2730
            TabIndex        =   114
            Text            =   "Transportadora"
            Top             =   1545
            Width           =   2055
         End
         Begin VB.TextBox Motivo 
            BorderStyle     =   0  'None
            DragMode        =   1  'Automatic
            Enabled         =   0   'False
            Height          =   225
            Left            =   2475
            TabIndex        =   113
            Text            =   "Motivo"
            Top             =   1995
            Width           =   3690
         End
         Begin MSMask.MaskEdBox Valor 
            Height          =   240
            Left            =   7305
            TabIndex        =   125
            Top             =   780
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridPedido 
            Height          =   1845
            Left            =   45
            TabIndex        =   19
            Top             =   510
            Width           =   9075
            _ExtentX        =   16007
            _ExtentY        =   3254
            _Version        =   393216
            Rows            =   10
            Cols            =   7
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Ordenados por:"
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
            Left            =   1395
            TabIndex        =   126
            Top             =   195
            Width           =   1320
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Dados"
         Height          =   630
         Left            =   0
         TabIndex        =   105
         Top             =   30
         Width           =   9150
         Begin VB.ComboBox Serie 
            Height          =   315
            Left            =   3480
            TabIndex        =   15
            Top             =   210
            Width           =   975
         End
         Begin MSComCtl2.UpDown UpDownSaida 
            Height          =   300
            Left            =   2565
            TabIndex        =   106
            TabStop         =   0   'False
            Top             =   210
            Width           =   225
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataSaida 
            Height          =   300
            Left            =   1470
            TabIndex        =   14
            Top             =   210
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown CTBUpDown 
            Height          =   300
            Left            =   6855
            TabIndex        =   107
            TabStop         =   0   'False
            Top             =   225
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox CTBDataContabil 
            Height          =   300
            Left            =   5775
            TabIndex        =   16
            Top             =   225
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
            Left            =   8460
            TabIndex        =   17
            Top             =   225
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Série:"
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
            Left            =   2940
            TabIndex        =   111
            Top             =   270
            Width           =   510
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Data de Saída:"
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
            TabIndex        =   110
            Top             =   240
            Width           =   1305
         End
         Begin VB.Label LabelDataContabil 
            AutoSize        =   -1  'True
            Caption         =   "Data Contabil:"
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
            Left            =   4530
            TabIndex        =   109
            Top             =   270
            Width           =   1230
         End
         Begin VB.Label CTBLabelLote 
            AutoSize        =   -1  'True
            Caption         =   "Lote Contábil:"
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
            Left            =   7260
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   108
            Top             =   270
            Width           =   1200
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Lixo"
         Height          =   285
         Left            =   8565
         TabIndex        =   59
         Top             =   105
         Visible         =   0   'False
         Width           =   510
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Height          =   4395
            Index           =   8
            Left            =   1020
            TabIndex        =   61
            Top             =   0
            Visible         =   0   'False
            Width           =   9240
            Begin VB.CheckBox CTBAglutina 
               Height          =   210
               Left            =   4470
               TabIndex        =   74
               Top             =   2565
               Width           =   870
            End
            Begin VB.TextBox CTBHistorico 
               BorderStyle     =   0  'None
               Height          =   225
               Left            =   4245
               MaxLength       =   150
               TabIndex        =   73
               Top             =   2175
               Width           =   1770
            End
            Begin VB.ListBox CTBListHistoricos 
               Height          =   2790
               Left            =   6330
               TabIndex        =   72
               Top             =   1560
               Visible         =   0   'False
               Width           =   2625
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
               TabIndex        =   71
               Top             =   630
               Width           =   1245
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
               TabIndex        =   70
               Top             =   120
               Width           =   1245
            End
            Begin VB.ComboBox CTBModelo 
               Height          =   315
               Left            =   7740
               Style           =   2  'Dropdown List
               TabIndex        =   69
               Top             =   810
               Width           =   1260
            End
            Begin VB.Frame CTBFrame7 
               Caption         =   "Descrição do Elemento Selecionado"
               Height          =   1050
               Left            =   195
               TabIndex        =   64
               Top             =   3330
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
                  Height          =   195
                  Left            =   240
                  TabIndex        =   68
                  Top             =   660
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
                  Height          =   195
                  Left            =   1125
                  TabIndex        =   67
                  Top             =   315
                  Width           =   570
               End
               Begin VB.Label CTBContaDescricao 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   1845
                  TabIndex        =   66
                  Top             =   285
                  Width           =   3720
               End
               Begin VB.Label CTBCclDescricao 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   1845
                  TabIndex        =   65
                  Top             =   645
                  Visible         =   0   'False
                  Width           =   3720
               End
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
               Left            =   7710
               TabIndex        =   63
               Top             =   135
               Width           =   1245
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
               Left            =   3480
               TabIndex        =   62
               Top             =   930
               Value           =   1  'Checked
               Width           =   2745
            End
            Begin MSMask.MaskEdBox CTBSeqContraPartida 
               Height          =   225
               Left            =   4800
               TabIndex        =   75
               Top             =   1560
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
               TabIndex        =   76
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
               TabIndex        =   77
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
               TabIndex        =   78
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
               TabIndex        =   79
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
            Begin MSComCtl2.UpDown CTBUpDown3 
               Height          =   300
               Left            =   1635
               TabIndex        =   80
               TabStop         =   0   'False
               Top             =   540
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox CTBDataContabil3 
               Height          =   300
               Left            =   570
               TabIndex        =   81
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
            Begin MSMask.MaskEdBox CTBLote3 
               Height          =   300
               Left            =   5580
               TabIndex        =   82
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
               Left            =   1845
               TabIndex        =   83
               Top             =   3030
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   5
               Mask            =   "#####"
               PromptChar      =   " "
            End
            Begin MSComctlLib.TreeView CTBTvwCcls 
               Height          =   2790
               Left            =   6330
               TabIndex        =   84
               Top             =   1560
               Visible         =   0   'False
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   4921
               _Version        =   393217
               LabelEdit       =   1
               LineStyle       =   1
               Style           =   7
               BorderStyle     =   1
               Appearance      =   1
            End
            Begin MSComctlLib.TreeView CTBTvwContas 
               Height          =   2790
               Left            =   6330
               TabIndex        =   85
               Top             =   1560
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   4921
               _Version        =   393217
               LabelEdit       =   1
               LineStyle       =   1
               Style           =   7
               BorderStyle     =   1
               Appearance      =   1
            End
            Begin MSFlexGridLib.MSFlexGrid CTBGridContabil 
               Height          =   1860
               Left            =   0
               TabIndex        =   86
               Top             =   1170
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
               Height          =   255
               Left            =   3600
               TabIndex        =   103
               Top             =   3120
               Width           =   720
            End
            Begin VB.Label CTBOrigem 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   4305
               TabIndex        =   102
               Top             =   3075
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
               TabIndex        =   101
               Top             =   600
               Width           =   735
            End
            Begin VB.Label CTBPeriodo 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   5010
               TabIndex        =   100
               Top             =   570
               Width           =   1185
            End
            Begin VB.Label CTBExercicio 
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   2910
               TabIndex        =   99
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
               TabIndex        =   98
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
               TabIndex        =   97
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
               Left            =   6345
               TabIndex        =   96
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
               TabIndex        =   95
               Top             =   1305
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
               Left            =   6360
               TabIndex        =   94
               Top             =   1290
               Visible         =   0   'False
               Width           =   2490
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
               Height          =   195
               Left            =   7755
               TabIndex        =   93
               Top             =   585
               Width           =   690
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
               Height          =   225
               Left            =   1800
               TabIndex        =   92
               Top             =   3045
               Width           =   615
            End
            Begin VB.Label CTBTotalDebito 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   3705
               TabIndex        =   91
               Top             =   3030
               Width           =   1155
            End
            Begin VB.Label CTBTotalCredito 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   2460
               TabIndex        =   90
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
               TabIndex        =   89
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
               Height          =   195
               Left            =   750
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   88
               Top             =   3075
               Width           =   1035
            End
            Begin VB.Label CTBLabelLote3 
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
               Height          =   195
               Left            =   5100
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   87
               Top             =   165
               Width           =   450
            End
         End
         Begin MSFlexGridLib.MSFlexGrid GridAlocacao 
            Height          =   1860
            Left            =   585
            TabIndex        =   60
            Top             =   0
            Visible         =   0   'False
            Width           =   7290
            _ExtentX        =   12859
            _ExtentY        =   3281
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   1575
            Left            =   630
            TabIndex        =   104
            Top             =   0
            Visible         =   0   'False
            Width           =   8880
            _ExtentX        =   15663
            _ExtentY        =   2778
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5370
      Index           =   1
      Left            =   180
      TabIndex        =   29
      Top             =   495
      Width           =   9225
      Begin VB.Frame Frame2 
         Caption         =   "Filtros"
         Height          =   5250
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   9120
         Begin VB.Frame Frame8 
            Caption         =   "Vendedores"
            Height          =   1140
            Left            =   405
            TabIndex        =   49
            Top             =   2535
            Width           =   3105
            Begin MSMask.MaskEdBox VendedorInicial 
               Height          =   300
               Left            =   780
               TabIndex        =   7
               Top             =   285
               Width           =   2190
               _ExtentX        =   3863
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox VendedorFinal 
               Height          =   300
               Left            =   780
               TabIndex        =   8
               Top             =   690
               Width           =   2190
               _ExtentX        =   3863
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin VB.Label LabelVendedorDe 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   375
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   51
               Top             =   330
               Width           =   315
            End
            Begin VB.Label LabelVendedorAte 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   315
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   50
               Top             =   750
               Width           =   360
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Regiões de Vendas"
            Height          =   2340
            Left            =   3900
            TabIndex        =   46
            Top             =   1335
            Width           =   4785
            Begin VB.ListBox ListRegioes 
               Height          =   1860
               Left            =   60
               Style           =   1  'Checkbox
               TabIndex        =   9
               Top             =   300
               Width           =   3030
            End
            Begin VB.CommandButton BotaoMarcar 
               Caption         =   "Marcar Todas"
               Height          =   555
               Left            =   3120
               Picture         =   "GeracaoNFiscalOcx.ctx":2F80
               Style           =   1  'Graphical
               TabIndex        =   48
               Top             =   285
               Width           =   1530
            End
            Begin VB.CommandButton BotaoDesmarcar 
               Caption         =   "Desmarcar Todas"
               Height          =   555
               Left            =   3120
               Picture         =   "GeracaoNFiscalOcx.ctx":3F9A
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   1035
               Width           =   1530
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Data Emissão"
            Height          =   1155
            Left            =   405
            TabIndex        =   40
            Top             =   1335
            Width           =   3120
            Begin MSMask.MaskEdBox DataEmissaoDe 
               Height          =   300
               Left            =   795
               TabIndex        =   5
               Top             =   300
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownEmissaoDe 
               Height          =   300
               Left            =   1950
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   300
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataEmissaoAte 
               Height          =   300
               Left            =   795
               TabIndex        =   6
               Top             =   690
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownEmissaoAte 
               Height          =   300
               Left            =   1950
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   690
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin VB.Label Label1 
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
               Index           =   1
               Left            =   300
               TabIndex        =   43
               Top             =   750
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
               Index           =   0
               Left            =   345
               TabIndex        =   41
               Top             =   360
               Width           =   315
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Pedidos"
            Height          =   690
            Left            =   405
            TabIndex        =   34
            Top             =   495
            Width           =   3135
            Begin MSMask.MaskEdBox PedidoInicial 
               Height          =   300
               Left            =   810
               TabIndex        =   1
               Top             =   285
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox PedidoFinal 
               Height          =   300
               Left            =   2160
               TabIndex        =   2
               Top             =   285
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label LabelPedidoDe 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   405
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   35
               Top             =   345
               Width           =   315
            End
            Begin VB.Label LabelPedidoAte 
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
               Left            =   1695
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   36
               Top             =   345
               Width           =   360
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Clientes"
            Height          =   705
            Left            =   3885
            TabIndex        =   37
            Top             =   495
            Width           =   4785
            Begin MSMask.MaskEdBox ClienteDe 
               Height          =   300
               Left            =   585
               TabIndex        =   3
               Top             =   270
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ClienteAte 
               Height          =   300
               Left            =   2550
               TabIndex        =   4
               Top             =   270
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label LabelClienteAte 
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
               Left            =   2160
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   39
               Top             =   330
               Width           =   360
            End
            Begin VB.Label LabelClienteDe 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   195
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   38
               Top             =   315
               Width           =   315
            End
         End
         Begin VB.CheckBox ExibeTodos 
            Caption         =   "Exibe Todos os Pedidos"
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
            Left            =   3285
            TabIndex        =   0
            Top             =   210
            Width           =   2430
         End
         Begin VB.Frame Frame5 
            Caption         =   "Entrega"
            Height          =   1425
            Left            =   405
            TabIndex        =   45
            Top             =   3720
            Width           =   8280
            Begin VB.Frame Frame10 
               Caption         =   "Viagem"
               Height          =   660
               Left            =   2865
               TabIndex        =   57
               Top             =   195
               Width           =   2700
               Begin MSMask.MaskEdBox CodigoViagem 
                  Height          =   315
                  Left            =   1110
                  TabIndex        =   12
                  Top             =   240
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   556
                  _Version        =   393216
                  MaxLength       =   6
                  Mask            =   "999999"
                  PromptChar      =   " "
               End
               Begin VB.Label LabelViagem 
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
                  Height          =   255
                  Left            =   360
                  MousePointer    =   14  'Arrow and Question
                  TabIndex        =   58
                  Top             =   270
                  Width           =   750
               End
            End
            Begin VB.Frame Frame9 
               Caption         =   "Data"
               Height          =   1140
               Left            =   270
               TabIndex        =   52
               Top             =   195
               Width           =   2220
               Begin MSMask.MaskEdBox DataEntregaDe 
                  Height          =   300
                  Left            =   495
                  TabIndex        =   10
                  Top             =   300
                  Width           =   1170
                  _ExtentX        =   2064
                  _ExtentY        =   529
                  _Version        =   393216
                  MaxLength       =   8
                  Format          =   "dd/mm/yyyy"
                  Mask            =   "##/##/##"
                  PromptChar      =   " "
               End
               Begin MSComCtl2.UpDown UpDownEntregaDe 
                  Height          =   300
                  Left            =   1635
                  TabIndex        =   53
                  TabStop         =   0   'False
                  Top             =   300
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   529
                  _Version        =   393216
                  Enabled         =   -1  'True
               End
               Begin MSMask.MaskEdBox DataEntregaAte 
                  Height          =   300
                  Left            =   495
                  TabIndex        =   11
                  Top             =   720
                  Width           =   1170
                  _ExtentX        =   2064
                  _ExtentY        =   529
                  _Version        =   393216
                  MaxLength       =   8
                  Format          =   "dd/mm/yyyy"
                  Mask            =   "##/##/##"
                  PromptChar      =   " "
               End
               Begin MSComCtl2.UpDown UpDownEntregaAte 
                  Height          =   300
                  Left            =   1665
                  TabIndex        =   55
                  TabStop         =   0   'False
                  Top             =   720
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   529
                  _Version        =   393216
                  Enabled         =   -1  'True
               End
               Begin VB.Label Label1 
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
                  Index           =   3
                  Left            =   60
                  TabIndex        =   56
                  Top             =   780
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
                  Index           =   2
                  Left            =   90
                  TabIndex        =   54
                  Top             =   360
                  Width           =   315
               End
            End
            Begin VB.CheckBox optImprimirRomaneio 
               Caption         =   "Imprimir Romaneio de Entrega ao Gravar"
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
               Left            =   2880
               TabIndex        =   13
               Top             =   1035
               Width           =   3840
            End
         End
      End
   End
   Begin VB.CommandButton BotaoFechar 
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
      Left            =   8190
      Picture         =   "GeracaoNFiscalOcx.ctx":517C
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Fechar"
      Top             =   60
      Width           =   1230
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5805
      Left            =   45
      TabIndex        =   32
      Top             =   135
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   10239
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pedidos"
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
Attribute VB_Name = "GeracaoNFiscalOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'???? Alguns campos da nota fiscal não estão sendo preenchidos: DataEntrada,FilialPedido,NumPedidoVenda e ValorProdutos. Rotina NFiscal_ExtrairPV no ClassFATGrava.
'???? Alterar a rotina de leitura dos pedidos só incluindo as condições no SELECT se a opção de exibir todos os pedidos não estiver "setada".
'???? Chamei a tela de GeracaoNFiscal e depois apertei Gerar NFiscal. Voou tudo e travou o VB. Fecha e abrir os comando caso haja erro na geração de alguma NFiscal.
'???? Depois de gerar a NF não está atualizando o grid de pedidos. E quando eu coloco para atualizar aparece a msg de que não existem mais pedidos… Não devo exibir msg.
'???? O erro 44295 não está sendo tartado em Processa_Gravacao_NFiscal
'???? O grid de pedido não cabe na tela.
'???? Sem comentário:BotaoNFiscal , BotaoNFiscalFatura, NFiscalImprime, BotaoNFiscalImprime, ClienteDe_Validate

'???? Como a data de entrega é escolhida pela menor data de entrega dos itens, se algum item não possuir data de entrega a data da tela fica em branco, pois a data nula é menor que qq outra data digitda pelo usuário
'???? Para abrir a tela de pedido não deveria apenas ser necessário que um pedido estivesse com o foco?

Option Explicit

''OBSERVACAO:
''O frame de contabilidade foi incluido nesta tela apenas p/evitar duplicidade de codigo p/tratamento de campos de lote e data contabil. Por isso ele fica invisivel.
''Os grids de alocacao e itens ficam invisiveis e servem apenas p/auxiliar a contabilizacao
'
''Property Variables:
'Dim m_Caption As String
'
''variaveis auxiliares para criacao da contabilizacao
'Private gobjContabAutomatica As ClassContabAutomatica
'Private gobjNFiscal As ClassNFiscal
'Private gobjPedidoVenda As ClassPedidoDeVenda
'Private giExercicio As Integer, giPeriodo As Integer
'Private gcolAlmoxFilial As New Collection
'
''Associados a contabilidade
'Public objContabil As New ClassContabil
'Public WithEvents objEventoLote As AdmEvento
'Public WithEvents objEventoDoc As AdmEvento
'Public objGrid1 As AdmGrid
'
'Event Unload()
'
'Public iAlterado As Integer
'Dim iTabPrincipalAlterado As Integer
'Dim iFrameAtual As Integer
'Dim iClienteAlterado As Integer
'Dim gobjGeracaoNFiscal As New ClassGeracaoNFiscal
'
'Dim objGrid As AdmGrid
'Dim iGrid_GeraNFiscal_Col As Integer
'Dim iGrid_Pedido_Col As Integer
'Dim iGrid_Cliente_Col As Integer
'Dim iGrid_NomeRed_Col As Integer
'Dim iGrid_Filial_Col As Integer
'Dim iGrid_Estado_Col As Integer
'Dim iGrid_Cidade_Col As Integer
'Dim iGrid_Bairro_Col As Integer
'Dim iGrid_TransPortadora_Col As Integer
'Dim iGrid_Emissao_Col As Integer
'Dim iGrid_Entrega_Col As Integer
'Dim iGrid_Valor_Col As Integer
'Dim iGrid_FilialEmpresa_Col As Integer
'Dim iGrid_Motivo_Col As Integer
'
''Eventos de Browse
'Private WithEvents objEventoPedidoDe As AdmEvento
'Private WithEvents objEventoPedidoAte As AdmEvento
'Private WithEvents objEventoClienteDe As AdmEvento
'Private WithEvents objEventoClienteAte As AdmEvento
'
'Dim asOrdenacao(3) As String
'Dim asOrdenacaoString(3) As String
'
''Constantes públicas dos tabs
'Private Const TAB_Selecao = 1
'Private Const TAB_Pedidos = 2
'
''mnemonicos
'Private Const CODIGO1 As String = "Codigo"
'Private Const NATUREZA_OP As String = "Natureza_OP"
'Private Const CLIENTE1 As String = "Cliente"
'Private Const FILIAL1 As String = "Filial"
'Private Const Serie1 As String = "Serie"
'Private Const NFISCAL1 As String = "Nota_Fiscal"
'Private Const DATA_EMISSAO As String = "Data_Emissao"
'Private Const DATA_SAIDA As String = "Data_Saida"
'Private Const PRODUTO1 As String = "Produto_Codigo"
'Private Const UNIDADE_MED As String = "Unidade_Medida"
'Private Const QUANTIDADE1 As String = "Quantidade"
'Private Const PRECO_UNITARIO As String = "Preco_Unitario"
'Private Const PRECO_TOTAL As String = "Preco_Total"
'Private Const DESCONTO1 As String = "Desconto"
'Private Const DESCRICAO_ITEM As String = "Descricao_Item"
''#######Valores####################
'Private Const ICMS As String = "ICMS_Valor"
'Private Const ICMSSUBST As String = "ICMSSubst_Valor"
'Private Const VALOR_PRODUTOS As String = "Valor_Produtos"
'Private Const VALOR_FRETE As String = "Valor_Frete"
'Private Const VALOR_SEGURO As String = "Valor_Seguro"
'Private Const VALOR_DESPESAS As String = "Valor_Despesas"
'Private Const IPI As String = "Valor_IPI"
'Private Const VALOR_DESCONTO As String = "Valor_Desconto"
'Private Const VALOR_TOTAL As String = "Valor_Total"
''###########Almoxarifado############
'Private Const PRODUTO_ALMOX As String = "Produto_Almox"
'Private Const ALMOX1 As String = "Almoxarifado"
'Private Const QUANT_ALOCADA As String = "Quant_Alocada"
'Private Const UNIDADE_MED_EST As String = "Unidade_Med_Est"
''###########Tributação##############
'Private Const ISS_VALOR As String = "ISS_Valor"
'Private Const ISS_INCLUSO As String = "ISS_Incluso"
'Private Const VALOR_IRRF As String = "Valor_IRRF"
'Private Const CTACONTABILEST1 As String = "ContaContabilEst"
'Private Const QUANT_ESTOQUE As String = "Quant_Estoque" '??? ERRADO: nao está no bd
''fim da contabilidade
'
''ver conceito de filial que fatura vs filial do pedido e de manter todo o historico p/poder "reimprimir" uma NF
'
''ver p/setar sIPICodProduto e numintdoc (do item da nf) a nivel de item trib nf
''incluir coluna no grid p/transportadora
'
''a principio nao poderá entrar como EMPRESA_TODA
'
''incluir outros controles p/:
'    'transportadora, status (pulada por falta de estoque, pulada por bloqueio,...)
'
''trocar sNaturezaOpEntrada p/sNaturezaOpInterna em ClassTributacaoNF, no bd, type, etc.
''natop, deveria estar em tributacaoPV e nao em PV (ou nos dois)
'
''Será que com a checkbox de "todos os pedidos" marcada devo incluir mesmo os pedidos bloqueados ?
''':acho que nao
'''
'''
''' se o flag de "só fatura tudo" estiver setado
'''    Se tem "Bloqueio Parcial" ou "Nao Reserva" pular o pedido deixando-o marcado como "faltou estoque p/poder faturar" (vamos guardar a data da ultima tentativa frustrada)
'''
'''0. Se tem bloqueio total ou outro bloqueio que nao seja de estoque (Credito, Endereco,...) pular o pedido.
'''
'''1. Para os Pedidos que tem SoFaturaTudo,
'''    Se tem "Bloqueio Parcial" ou "Nao Reserva" pular o pedido deixando-o marcado como "faltou estoque p/poder faturar" (vamos guardar a data da ultima tentativa frustrada)
'''    Senao, tenta gerar a NF atendendo a todos os itens completamente, se nao conseguir pular o pedido deixando-o marcado como "faltou estoque p/poder faturar" (vamos guardar a data da ultima tentativa frustrada)
'''
'''2. Para os Pedidos que não têm SoFaturaTudo,
'''2.1. Para os ítens do tipo "estoque (sem reserva)",  tenta  "tirar" QAR no almoxarifado default. Se nao conseguir abrir dialogo.
'''2.2. Para os itens "reserva+estoque", tenta "tirar" a qtde reservada de acordo com a reserva. Se a qtde reservada for zero, pular o item.
'''Se conseguir, libera as reservas, senao erro (nao fatura o pedido, (vamos guardar a data da ultima tentativa frustrada)).
'''2.3. Se houver algum item OK p/faturamento pode fatura-lo. No futuro podemos incluir algum criterio p/faturar apenas qdo o volume a ser faturado for "expressivo" em relacao ao total que poderia ser faturado se nao houvesse falta de estoque.
'''
'''Observacao:
'''     a geracao da NF já deveria poder ler no cadastro de pedidos, diretamente, se o pedido é "faturável" ou nao.
'        'numa versao melhorada poderia manter controle p/só faturar algo "significativo" em relacao ao pedido
'''
'Private Sub BotaoDesmarcarTodos_Click()
''Desmarca todos os pedidos do Grid
'
'Dim iLinha As Integer
'
'    'Percorre todas as linhas do Grid
'    For iLinha = 1 To objGrid.iLinhasExistentes
'
'        'Desmarca na tela o pedido em questão
'        GridPedido.TextMatrix(iLinha, iGrid_GeraNFiscal_Col) = S_DESMARCADO
'
'        'Desmarca no Obj o pedido em questão
'        gobjGeracaoNFiscal.colNFiscalInfo.Item(iLinha).iMarcada = S_DESMARCADO
'
'    Next
'
'    'Atualiza na tela a checkbox desmarcada
'    Call Grid_Refresh_Checkbox(objGrid)
'
'End Sub
'
'Private Sub BotaoFechar_Click()
'
'    'Fecha a tela
'    Unload Me
'
'End Sub
'
'Private Sub BotaoMarcarTodos_Click()
''Marca todos os pedidos do Grid
'
'Dim iLinha As Integer
'Dim objNFiscalInfo As ClassNFiscalInfo
'
'    'Percorre todas as linhas do Grid
'    For iLinha = 1 To objGrid.iLinhasExistentes
'
'        'Marca na tela o pedido em questão
'        GridPedido.TextMatrix(iLinha, iGrid_GeraNFiscal_Col) = S_MARCADO
'
'        gobjGeracaoNFiscal.colNFiscalInfo.Item(iLinha).iMarcada = S_MARCADO
'
'    Next
'
'    'Atualiza na tela a checkbox marcada
'    Call Grid_Refresh_Checkbox(objGrid)
'
'End Sub
'
'Private Function Inicializa_Grid_Pedido(objGridInt As AdmGrid) As Long
''Inicializa o Grid
'
'    'Form do Grid
'    Set objGridInt.objForm = Me
'
'    'Títulos das colunas
'    objGridInt.colColuna.Add ("  ")
'    objGridInt.colColuna.Add ("Gera NF")
'    objGridInt.colColuna.Add ("Pedido")
'    objGridInt.colColuna.Add ("Cliente")
'    objGridInt.colColuna.Add ("Nome")
'    objGridInt.colColuna.Add ("Filial")
'    objGridInt.colColuna.Add ("Estado Entrega")
'    objGridInt.colColuna.Add ("Cidade Entrega")
'    objGridInt.colColuna.Add ("Bairro Entrega")
'    objGridInt.colColuna.Add ("Transportadora")
'    objGridInt.colColuna.Add ("Emissão")
'    objGridInt.colColuna.Add ("Entrega")
'    objGridInt.colColuna.Add ("Valor")
'    objGridInt.colColuna.Add ("Filial Empresa")
'    objGridInt.colColuna.Add ("Erro na Geração")
'
'    'Controles que participam do Grid
'    objGridInt.colCampo.Add (GeraNFiscal.Name)
'    objGridInt.colCampo.Add (Pedido.Name)
'    objGridInt.colCampo.Add (Cliente.Name)
'    objGridInt.colCampo.Add (NomeReduzido.Name)
'    objGridInt.colCampo.Add (Filial.Name)
'    objGridInt.colCampo.Add (Estado.Name)
'    objGridInt.colCampo.Add (Cidade.Name)
'    objGridInt.colCampo.Add (Bairro.Name)
'    objGridInt.colCampo.Add (Transportadora.Name)
'    objGridInt.colCampo.Add (DataEmissao.Name)
'    objGridInt.colCampo.Add (DataEntrega.Name)
'    objGridInt.colCampo.Add (Valor.Name)
'    objGridInt.colCampo.Add (FilialEmpresa.Name)
'    objGridInt.colCampo.Add (Motivo.Name)
'
'    'Colunas do Grid
'    iGrid_GeraNFiscal_Col = 1
'    iGrid_Pedido_Col = 2
'    iGrid_Cliente_Col = 3
'    iGrid_NomeRed_Col = 4
'    iGrid_Filial_Col = 5
'    iGrid_Estado_Col = 6
'    iGrid_Cidade_Col = 7
'    iGrid_Bairro_Col = 8
'    iGrid_TransPortadora_Col = 9
'    iGrid_Emissao_Col = 10
'    iGrid_Entrega_Col = 11
'    iGrid_Valor_Col = 12
'    iGrid_FilialEmpresa_Col = 13
'    iGrid_Motivo_Col = 14
'
'    'Grid do GridInterno
'    objGridInt.objGrid = GridPedido
'
'    'Linhas visíveis do grid
'    objGridInt.iLinhasVisiveis = 5
'
'    'Todas as linhas do grid
'    objGridInt.objGrid.Rows = objGridInt.iLinhasVisiveis + 1
'
'    'Largura da primeira coluna
'    GridPedido.ColWidth(0) = 400
'
'    'Largura automática para as outras colunas
'    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
'    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
'    '???? Estava permitindo excluir linhas do grid de pedidos
'    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
'
'    'Chama função que inicializa o Grid
'    Call Grid_Inicializa(objGridInt)
'
'    GridPedido.Width = 8400
'
'    Inicializa_Grid_Pedido = SUCESSO
'
'    Exit Function
'
'End Function
'
'Private Sub BotaoNFiscal_Click()
'
'Dim lErro As Long
'Dim objNFiscalInfo As ClassNFiscalInfo
'Dim iIndice As Integer
'
'On Error GoTo Erro_BotaoNFiscal_Click
'
'    GL_objMDIForm.MousePointer = vbHourglass
'
'    'Atribui o tipo de NotaFiscal Interna de Saída de Venda a nota a
'    'ser gerada
'    gobjGeracaoNFiscal.iTipoNFiscal = DOCINFO_NFISVPV
'
'    'Verifica se a série está preenchida
'    If Len(Trim(Serie.Text)) = 0 Then Error 51354
'
'    'Recolhe a série e a Data de Saída
'    gobjGeracaoNFiscal.sSerie = Serie.Text
'    gobjGeracaoNFiscal.dtDataSaida = StrParaDate(DataSaida.Text)
'    gobjGeracaoNFiscal.iImprime = 0
'
'    lErro = GeracaoNF_Prepara_CTB
'    If lErro <> SUCESSO Then Error 59382
'
'    'Chama a rotina que gera as notas ficais a partir dos pedidos selecionados
'    lErro = CF("GeracaoNFiscal_GerarNFs", gobjGeracaoNFiscal)
'    If lErro <> SUCESSO Then Error 44197
'
'    For iIndice = gobjGeracaoNFiscal.colNFiscalInfo.Count To 1 Step -1
'        Set objNFiscalInfo = gobjGeracaoNFiscal.colNFiscalInfo(iIndice)
'
'        If objNFiscalInfo.iMarcada = MARCADO And objNFiscalInfo.iMotivoNaoGerada = 0 Then
'            gobjGeracaoNFiscal.colNFiscalInfo.Remove iIndice
'        End If
'    Next
'
'    'Recarrega o grid de Pedidos excluíndo aqueles que já geraram NFs
'    Call Grid_Limpa(objGrid)
'    lErro = Grid_Pedido_Preenche(gobjGeracaoNFiscal.colNFiscalInfo)
'    If lErro = 51429 Then Error 51431
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Exit Sub
'
'Erro_BotaoNFiscal_Click:
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Select Case Err
'
'        Case 44197, 51427, 59382
'
'        Case 51354
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", Err)
'
'        Case 51431
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160852)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub BotaoNFiscalFatura_Click()
'
'Dim lErro As Long
'Dim objNFiscalInfo As ClassNFiscalInfo
'Dim iIndice As Integer
'
'On Error GoTo Erro_BotaoNFiscalFatura_Click
'
'    GL_objMDIForm.MousePointer = vbHourglass
'
'    'Atribui o tipo de Nota Fiscal Interna de Saida de Fatura de Venda
'    'para as notas a serem geradas
'    gobjGeracaoNFiscal.iTipoNFiscal = DOCINFO_NFISFVPV
'
'    'Verifica se a série está preenchida
'    If Len(Trim(Serie.Text)) = 0 Then Error 51437
'
'    'Recolhe a série e a data de saída da tela
'    gobjGeracaoNFiscal.sSerie = Serie.Text
'    gobjGeracaoNFiscal.dtDataSaida = StrParaDate(DataSaida.Text)
'    gobjGeracaoNFiscal.iImprime = 0
'
'    lErro = GeracaoNF_Prepara_CTB
'    If lErro <> SUCESSO Then Error 59383
'
'    'Chama a rotina que gera as NFs a partir de pedidos
'    lErro = CF("GeracaoNFiscal_GerarNFs", gobjGeracaoNFiscal)
'    If lErro <> SUCESSO Then Error 44198
'
'    For iIndice = gobjGeracaoNFiscal.colNFiscalInfo.Count To 1 Step -1
'        Set objNFiscalInfo = gobjGeracaoNFiscal.colNFiscalInfo(iIndice)
'
'        If objNFiscalInfo.iMarcada = MARCADO And objNFiscalInfo.iMotivoNaoGerada = 0 Then
'            gobjGeracaoNFiscal.colNFiscalInfo.Remove iIndice
'        End If
'    Next
'
'
'    'Recarrega o grid de Pedidos excluíndo aqueles que já geraram NFs
'    Call Grid_Limpa(objGrid)
'    lErro = Grid_Pedido_Preenche(gobjGeracaoNFiscal.colNFiscalInfo)
'    If lErro <> SUCESSO Then Error 51432
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Exit Sub
'
'Erro_BotaoNFiscalFatura_Click:
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Select Case Err
'
'        Case 44198, 51425, 59383
'
'        Case 51432
'
'        Case 51437
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", Err)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160853)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub BotaoNFiscalFaturaImprime_Click()
'
'Dim lErro As Long
'Dim objNFiscalInfo As ClassNFiscalInfo
'Dim iIndice As Integer
'
'On Error GoTo Erro_BotaoNFiscalFaturaImprime_Click
'
'    GL_objMDIForm.MousePointer = vbHourglass
'
'    'Atribui o tipo de NotaFiscal Interna de Saída Fatura de Venda a nota a
'    'ser gerada
'    gobjGeracaoNFiscal.iTipoNFiscal = DOCINFO_NFISFVPV
'
'    'Verifica se a série está preenchida
'    If Len(Trim(Serie.Text)) = 0 Then Error 51438
'
'    'Recolhe a série e a Data de Saída
'    gobjGeracaoNFiscal.sSerie = Serie.Text
'    gobjGeracaoNFiscal.dtDataSaida = StrParaDate(DataSaida.Text)
'    gobjGeracaoNFiscal.iImprime = 1
'
'    lErro = GeracaoNF_Prepara_CTB
'    If lErro <> SUCESSO Then Error 59384
'
'    'Chama a rotina que gera as notas ficais a partir dos pedidos selecionados
'    lErro = CF("GeracaoNFiscal_GerarNFs", gobjGeracaoNFiscal)
'    If lErro <> SUCESSO Then Error 44199
'
'    For iIndice = gobjGeracaoNFiscal.colNFiscalInfo.Count To 1 Step -1
'        Set objNFiscalInfo = gobjGeracaoNFiscal.colNFiscalInfo(iIndice)
'
'        If objNFiscalInfo.iMarcada = MARCADO And objNFiscalInfo.iMotivoNaoGerada = 0 Then
'            gobjGeracaoNFiscal.colNFiscalInfo.Remove iIndice
'        End If
'    Next
'
'    'Recarrega o grid de Pedidos excluíndo aqueles que já geraram NFs
'    Call Grid_Limpa(objGrid)
'    lErro = Grid_Pedido_Preenche(gobjGeracaoNFiscal.colNFiscalInfo)
'    If lErro <> SUCESSO Then Error 51440
'
'    Call NotaFiscal_Imprime
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Exit Sub
'
'Erro_BotaoNFiscalFaturaImprime_Click:
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Select Case Err
'
'        Case 44199, 51439, 51440, 59384
'
'        Case 51438
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", Err)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160854)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub BotaoNFiscalImprime_Click()
'
'Dim lErro As Long
'Dim objNFiscalInfo As ClassNFiscalInfo
'Dim iIndice As Integer
'
'On Error GoTo Erro_BotaoNFiscal_Click
'
'    GL_objMDIForm.MousePointer = vbHourglass
'
'    'Atribui o tipo de NotaFiscal Interna de Saída de Venda a nota a
'    'ser gerada
'    gobjGeracaoNFiscal.iTipoNFiscal = DOCINFO_NFISVPV
'
'    'Verifica se a série está preenchida
'    If Len(Trim(Serie.Text)) = 0 Then Error 51441
'
'    'Recolhe a série e a Data de Saída
'    gobjGeracaoNFiscal.sSerie = Serie.Text
'    gobjGeracaoNFiscal.dtDataSaida = StrParaDate(DataSaida.Text)
'    gobjGeracaoNFiscal.iImprime = 1
'
'    lErro = GeracaoNF_Prepara_CTB
'    If lErro <> SUCESSO Then Error 59385
'
'    'Chama a rotina que gera as notas ficais a partir dos pedidos selecionados
'    lErro = CF("GeracaoNFiscal_GerarNFs", gobjGeracaoNFiscal)
'    If lErro <> SUCESSO Then Error 44200
'
'    For iIndice = gobjGeracaoNFiscal.colNFiscalInfo.Count To 1 Step -1
'        Set objNFiscalInfo = gobjGeracaoNFiscal.colNFiscalInfo(iIndice)
'
'        If objNFiscalInfo.iMarcada = MARCADO And objNFiscalInfo.iMotivoNaoGerada = 0 Then
'            gobjGeracaoNFiscal.colNFiscalInfo.Remove iIndice
'        End If
'    Next
'
'
'    'Recarrega o grid de Pedidos excluíndo aqueles que já geraram NFs
'    Call Grid_Limpa(objGrid)
'    lErro = Grid_Pedido_Preenche(gobjGeracaoNFiscal.colNFiscalInfo)
'    If lErro <> SUCESSO Then Error 51442
'
'    'Exibe os pedidos desselecionados
'    Call BotaoDesmarcarTodos_Click
'
'    Call NotaFiscal_Imprime
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Exit Sub
'
'Erro_BotaoNFiscal_Click:
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Select Case Err
'
'        Case 44200, 51442, 59385
'
'        Case 51441
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", Err)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160855)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub BotaoPedido_Click()
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim iAchou As Integer
'Dim objNFiscalInfo As New ClassNFiscalInfo
'Dim objPedidoDeVenda As New ClassPedidoDeVenda
'
'On Error GoTo Erro_BotaoPedido_Click
'
'    If objGrid.iLinhasExistentes = 0 Then Exit Sub
'
'    'Se nenhuma linha válida estiver com o foco
'    If GridPedido.Row < 1 Or GridPedido.Row > objGrid.iLinhasExistentes Then Error 51453
'
'    'Passa a linha do Grid para o Obj
'    Set objNFiscalInfo = gobjGeracaoNFiscal.colNFiscalInfo.Item(GridPedido.Row)
'
'    'Passa os dados do NFiscal para o Obj
'    objPedidoDeVenda.iFilialEmpresa = objNFiscalInfo.iFilialEmpresa
'    objPedidoDeVenda.lCodigo = objNFiscalInfo.lPedido
'
'    If objPedidoDeVenda.iFilialEmpresa <> giFilialEmpresa Then Error 51454
'
'    'Chama a tela de Pedidos de Venda
'    Call Chama_Tela("PedidoVenda", objPedidoDeVenda)
'
'    Exit Sub
'
'Erro_BotaoPedido_Click:
'
'    Select Case Err
'
'        Case 51453
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NENHUM_PEDIDO_SELECIONADO", Err)
'
'        Case 51454
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALPEDIDO_DIFERENTE_FILIALEMPRESA", Err, objPedidoDeVenda.lCodigo, objPedidoDeVenda.iFilialEmpresa, giFilialEmpresa)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160856)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub ClienteAte_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    iTabPrincipalAlterado = REGISTRO_ALTERADO
'    iClienteAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'
'Private Sub ClienteAte_GotFocus()
'Dim iTabAux As Integer
'Dim iClienteAux As Integer
'
'    iTabAux = iTabPrincipalAlterado
'    iClienteAux = iClienteAlterado
'
'    Call MaskEdBox_TrataGotFocus(ClienteAte, iAlterado)
'    iTabPrincipalAlterado = iTabAux
'    iClienteAlterado = iClienteAux
'
'End Sub
'
'Private Sub ClienteAte_Validate(Cancel As Boolean)
''Verifica se o Cliente De é maior que o Cliente Até
''Verifica a integridade do cliente com o BD
'
'Dim lErro As Long
'Dim objClienteAte As New ClassCliente
'Dim iCodFilial As Integer
'Dim iCria As Integer
'Dim colCodigoNome As AdmColCodigoNome
'
'On Error GoTo Erro_ClienteAte_Validate
'
'    If iClienteAlterado = 1 Then
'
'        If Len(Trim(ClienteAte.Text)) > 0 Then
'
'            'Se o Cliente De estiver preenchido
'            If Len(Trim(ClienteDe.Text)) > 0 Then
'                'Verifica se o Cliente De é maior que o Cliente Até ----->>> Erro
'                If LCodigo_Extrai(ClienteDe.Text) > LCodigo_Extrai(ClienteAte.Text) Then Error 58014
'
'            End If
'
'            objClienteAte.lCodigo = ClienteAte.Text
'
'            'Le o Cliente para testar sua integridade com o BD
'            lErro = CF("Cliente_Le", objClienteAte)
'            If lErro <> SUCESSO And lErro <> 12293 Then Error 58015
'
'            'Se não encontrou ----> erro
'            If lErro = 12293 Then Error 58016
'
'        End If
'
'        iClienteAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_ClienteAte_Validate:
'
'    Cancel = True
'
'    Select Case Err
'
'    Case 58014
'        Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTEDE_MAIOR_CLIENTEATE", Err)
'
'    Case 58015 'Tratados nas rotinas chamadas
'
'    Case 58016
'        Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objClienteAte.lCodigo)
'
'    Case Else
'        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160857)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub ClienteDe_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    iTabPrincipalAlterado = REGISTRO_ALTERADO
'    iClienteAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub ClienteDe_GotFocus()
'Dim iTabAux As Integer
'Dim iClienteAux As Integer
'
'    iTabAux = iTabPrincipalAlterado
'    iClienteAux = iClienteAlterado
'
'    Call MaskEdBox_TrataGotFocus(ClienteDe, iAlterado)
'    iTabPrincipalAlterado = iTabAux
'    iClienteAlterado = iClienteAux
'
'End Sub
'
'Private Sub ClienteDe_Validate(Cancel As Boolean)
''Verifica se o Cliente De é maior que o Cliente Até
''Verifica a integridade do cliente com o BD
'
'Dim lErro As Long
'Dim objClienteDe As New ClassCliente
'Dim iCodFilial As Integer
'Dim iCria As Integer
'Dim colCodigoNome As AdmColCodigoNome
'
'On Error GoTo Erro_ClienteDe_Validate
'
'    'Se o ClienteDE não foi alterado --> sai
'    If iClienteAlterado = 0 Then Exit Sub
'    'Se algum clientefoi informado
'    If Len(Trim(ClienteDe.Text)) > 0 Then
'        'Se o cliente até estiver preenchido
'        If Len(Trim(ClienteAte.Text)) > 0 Then
'            'Verifica se o cliente De é menor que o cliente até
'            If LCodigo_Extrai(ClienteDe.Text) > LCodigo_Extrai(ClienteAte.Text) Then Error 58011
'        End If
'
'        objClienteDe.lCodigo = StrParaLong(ClienteDe.Text)
'        'Lê o cliente informado
'        lErro = CF("Cliente_Le", objClienteDe)
'        If lErro <> SUCESSO And lErro <> 12293 Then Error 58012
'        If lErro = 12293 Then Error 58013 'Não encontrou
'
'    End If
'    'Zera flag de alteração de cliente de.
'    iClienteAlterado = 0
'
'    Exit Sub
'
'Erro_ClienteDe_Validate:
'
'    Cancel = True
'
'    Select Case Err
'
'    Case 58011
'        Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTEDE_MAIOR_CLIENTEATE", Err)
'
'    Case 58012 'Tratados nas rotinas chamadas
'
'    Case 58013
'        Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objClienteDe.lCodigo)
'
'    Case Else
'        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160858)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub CTBUpDown_DownClick()
'
'Dim lErro As Long
'Dim sData As String
'
'On Error GoTo Erro_CTBUpDown_DownClick
'
'    lErro = Data_Up_Down_Click(CTBDataContabil, DIMINUI_DATA)
'    If lErro <> SUCESSO Then gError 71544
'
'    Exit Sub
'
'Erro_CTBUpDown_DownClick:
'
'    Select Case gErr
'
'        Case 71544
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160859)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub CTBUpDown_UpClick()
'
'Dim lErro As Long
'Dim sData As String
'
'On Error GoTo Erro_CTBUpDown_UpClick
'
'    lErro = Data_Up_Down_Click(CTBDataContabil, AUMENTA_DATA)
'    If lErro <> SUCESSO Then gError 71545
'
'    Exit Sub
'
'Erro_CTBUpDown_UpClick:
'
'    Select Case gErr
'
'        Case 71545
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160860)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub DataEmissaoAte_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    iTabPrincipalAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub DataEmissaoAte_GotFocus()
'Dim iTabAux As Integer
'
'    iTabAux = iTabPrincipalAlterado
'    Call MaskEdBox_TrataGotFocus(DataEmissaoAte, iAlterado)
'    iTabPrincipalAlterado = iTabAux
'
'End Sub
'
'Private Sub DataEmissaoAte_Validate(Cancel As Boolean)
''Critica a Data
'
'Dim lErro As Long
'
'On Error GoTo Erro_DataEmissaoAte_Validate
'
'    'Se a DataEmissaoAte está preenchida
'    If Len(DataEmissaoAte.ClipText) = 0 Then Exit Sub
'
'    'Verifica se a DataEmissaoAte é válida
'    lErro = Data_Critica(DataEmissaoAte.Text)
'    If lErro <> SUCESSO Then Error 28458
'
'    'Verifica se a data de emissao de está preenchida
'    If Len(DataEmissaoDe.ClipText) = 0 Then Exit Sub
'
'    'Verifica se a data emissão de é maior que a Data de emissão até
'    If CDate(DataEmissaoDe.Text) > CDate(DataEmissaoAte.Text) Then Error 58020
'
'    Exit Sub
'
'Erro_DataEmissaoAte_Validate:
'
'    Cancel = True
'
'    Select Case Err
'
'        Case 28458
'
'        Case 58020
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160861)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub DataEmissaoDe_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    iTabPrincipalAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub DataEmissaoDe_GotFocus()
'Dim iTabAux As Integer
'
'    iTabAux = iTabPrincipalAlterado
'    Call MaskEdBox_TrataGotFocus(DataEmissaoDe, iAlterado)
'    iTabPrincipalAlterado = iTabAux
'
'End Sub
'
'Private Sub DataEmissaoDe_Validate(Cancel As Boolean)
''Critica a Data
'
'Dim lErro As Long
'
'On Error GoTo Erro_DataEmissaoDe_Validate
'
'    'Se a DataEmissaoDe está preenchida
'    If Len(DataEmissaoDe.ClipText) = 0 Then Exit Sub
'
'    'Verifica se a DataEmissaoDe é válida
'    lErro = Data_Critica(DataEmissaoDe.Text)
'    If lErro <> SUCESSO Then Error 28457
'
'    If Len(Trim(DataEmissaoAte.ClipText)) = 0 Then Exit Sub
'
'    If CDate(DataEmissaoDe.Text) > CDate(DataEmissaoAte.Text) Then Error 31385
'
'    Exit Sub
'
'Erro_DataEmissaoDe_Validate:
'
'    Cancel = True
'
'    Select Case Err
'
'        Case 28457
'
'        Case 31385
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160862)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub DataEntregaAte_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    iTabPrincipalAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub DataEntregaAte_GotFocus()
'Dim iTabAux As Integer
'
'    iTabAux = iTabPrincipalAlterado
'    Call MaskEdBox_TrataGotFocus(DataEntregaAte, iAlterado)
'    iTabPrincipalAlterado = iTabAux
'
'End Sub
'
'Private Sub DataEntregaAte_Validate(Cancel As Boolean)
''Critica a Data
'
'Dim lErro As Long
'
'On Error GoTo Erro_DataEntregaAte_Validate
'
'    'Se a DataEntregaAte está preenchida
'    If Len(DataEntregaAte.ClipText) = 0 Then Exit Sub
'
'    'Verifica se a DataEntregaAte é válida
'    lErro = Data_Critica(DataEntregaAte.Text)
'    If lErro <> SUCESSO Then Error 28464
'
'    If Len(Trim(DataEntregaDe.ClipText)) = 0 Then Exit Sub
'
'    If CDate(DataEntregaDe.Text) > CDate(DataEntregaAte.Text) Then Error 31386
'
'    Exit Sub
'
'Erro_DataEntregaAte_Validate:
'
'    Cancel = True
'
'    Select Case Err
'
'        Case 28464
'
'        Case 31386
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160863)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub DataEntregaDe_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    iTabPrincipalAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub DataEntregaDe_GotFocus()
'Dim iTabAux As Integer
'
'    iTabAux = iTabPrincipalAlterado
'    Call MaskEdBox_TrataGotFocus(DataEntregaDe, iAlterado)
'    iTabPrincipalAlterado = iTabAux
'
'End Sub
'
'Private Sub DataEntregaDe_Validate(Cancel As Boolean)
''Critica a Data
'
'Dim lErro As Long
'
'On Error GoTo Erro_DataEntregaDe_Validate
'
'    'Se a DataEntregaDe está preenchida
'    If Len(Trim(DataEntregaDe.ClipText)) = 0 Then Exit Sub
'
'    'Verifica se a DataEntregaDe é válida
'    lErro = Data_Critica(DataEntregaDe.Text)
'    If lErro <> SUCESSO Then Error 28463
'
'    If Len(Trim(DataEntregaAte.ClipText)) = 0 Then Exit Sub
'
'    If CDate(DataEntregaDe.Text) > CDate(DataEntregaAte.Text) Then Error 58021
'
'    Exit Sub
'
'Erro_DataEntregaDe_Validate:
'
'    Cancel = True
'
'    Select Case Err
'
'        Case 28464
'
'        Case 58021
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160864)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub DataSaida_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub DataSaida_GotFocus()
'
'    Call MaskEdBox_TrataGotFocus(DataSaida, iAlterado)
'
'End Sub
'
'Private Sub ExibeTodos_Click()
''???? Quando a checkbox de "exibi todos os pedidos"  está selecionada todas as condições ( DataDe, DataAte, ClienteDe, ClienteAte…) são desabilitadas. OK. Mas só que os labels que chamam os browses esqueceram de ser desabilitados.
'
'    iTabPrincipalAlterado = REGISTRO_ALTERADO
'
'    'Limpa os campos da tela
'    PedidoInicial.Text = ""
'    PedidoFinal.Text = ""
'    ClienteDe.Text = ""
'    ClienteAte.Text = ""
'    DataEmissaoDe.PromptInclude = False
'    DataEmissaoDe.Text = ""
'    DataEmissaoDe.PromptInclude = True
'    DataEmissaoAte.PromptInclude = False
'    DataEmissaoAte.Text = ""
'    DataEmissaoAte.PromptInclude = True
'    DataEntregaDe.PromptInclude = False
'    DataEntregaDe.Text = ""
'    DataEntregaDe.PromptInclude = True
'    DataEntregaAte.PromptInclude = False
'    DataEntregaAte.Text = ""
'    DataEntregaAte.PromptInclude = True
'
'    'Se marcar ExibeTodos, exibe todos os pedidos
'    If ExibeTodos.Value = 1 Then
'        PedidoInicial.Enabled = False
'        PedidoFinal.Enabled = False
'        ClienteDe.Enabled = False
'        ClienteAte.Enabled = False
'        DataEmissaoDe.Enabled = False
'        DataEmissaoAte.Enabled = False
'        DataEntregaDe.Enabled = False
'        DataEntregaAte.Enabled = False
'        UpDownEmissaoDe.Enabled = False
'        UpDownEmissaoAte.Enabled = False
'        UpDownEntregaDe.Enabled = False
'        UpDownEntregaAte.Enabled = False
'        '??? Esqueceu de desabilitar as Labels com browse
'        LabelPedidoAte.Enabled = False
'        LabelPedidoDe.Enabled = False
'        LabelClienteAte.Enabled = False
'        LabelClienteDe.Enabled = False
'        Label1(0).Enabled = False
'        Label1(1).Enabled = False
'        Label1(2).Enabled = False
'        Label1(3).Enabled = False
'    Else
'        PedidoInicial.Enabled = True
'        PedidoFinal.Enabled = True
'        ClienteDe.Enabled = True
'        ClienteAte.Enabled = True
'        DataEmissaoDe.Enabled = True
'        DataEmissaoAte.Enabled = True
'        DataEntregaDe.Enabled = True
'        DataEntregaAte.Enabled = True
'        UpDownEmissaoDe.Enabled = True
'        UpDownEmissaoAte.Enabled = True
'        UpDownEntregaDe.Enabled = True
'        UpDownEntregaAte.Enabled = True
'        '??? Esqueceu de habilitar as Labels com browse
'        LabelPedidoAte.Enabled = True
'        LabelPedidoDe.Enabled = True
'        LabelClienteAte.Enabled = True
'        LabelClienteDe.Enabled = True
'        Label1(0).Enabled = True
'        Label1(1).Enabled = True
'        Label1(2).Enabled = True
'        Label1(3).Enabled = True
'    End If
'
'    Exit Sub
'
'End Sub
'
'Public Sub Form_Load()
'
'Dim lErro As Long
'Dim iIndice As Integer
'
'On Error GoTo Erro_Form_Load
'
'    Set objEventoPedidoDe = New AdmEvento
'    Set objEventoPedidoAte = New AdmEvento
'    Set objEventoClienteDe = New AdmEvento
'    Set objEventoClienteAte = New AdmEvento
'
'    asOrdenacao(0) = " FilialEmpresaPV, CodigoPV"
'    asOrdenacao(1) = " NomeCliente, FilialEmpresaPV, CodigoPV"
'    asOrdenacao(2) = " EmissaoPedido , FilialEmpresaPV, CodigoPV"
'    asOrdenacao(3) = " SiglaEstadoEntrega, CidadeEntrega,  BairroEntrega "
'
'    asOrdenacaoString(0) = "Filial da Empresa + Pedido"
'    asOrdenacaoString(1) = "Cliente + Filial da Empresa + Pedido"
'    asOrdenacaoString(2) = "Data de Emissão do Pedido + Filial da Empresa + Pedido"
'    asOrdenacaoString(3) = "Estado + Cidade + Bairro"
'
'    iFrameAtual = 1
'
'    DataSaida.PromptInclude = False
'    DataSaida.Text = Format(gdtDataAtual, "dd/mm/yy")
'    DataSaida.PromptInclude = True
'
'    Set objGrid = New AdmGrid
'
'    'Executa a Inicialização do grid Pedido
'    lErro = Inicializa_Grid_Pedido(objGrid)
'    If lErro <> SUCESSO Then Error 28480
'
'    'Carrega as Séries
'    lErro = Carrega_Serie()
'    If lErro <> SUCESSO Then Error 19164
'
'    'Carrega a Combobox Ordenados
'    For iIndice = 0 To 3
'        Ordenados.AddItem asOrdenacaoString(iIndice)
'    Next
'
'    Ordenados.ListIndex = 0
'
'    Set objEventoLote = New AdmEvento
'    Set objGrid1 = New AdmGrid
'
'    If (gcolModulo.ATIVO(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
'
'        'Inicialização da parte de contabilidade
'        lErro = objContabil.Contabil_Inicializa_Contabilidade3(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_FATURAMENTO)
'        If lErro <> SUCESSO Then Error 59408
'
'        lErro = objContabil.Contabil_Gera_Cabecalho_Automatico
'        If lErro <> SUCESSO Then Error 59411
'
'    Else
'
'        CTBDataContabil.Enabled = False
'        LabelDataContabil.Enabled = False
'
'    End If
'
'    iAlterado = 0
'
'    lErro_Chama_Tela = SUCESSO
'
'    Exit Sub
'
'Erro_Form_Load:
'
'    lErro_Chama_Tela = Err
'
'    Select Case Err
'
'        Case 28480, 19164, 59408, 59411
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160865)
'
'    End Select
'
'    iAlterado = 0
'
'    Exit Sub
'
'End Sub
'
'Private Function Carrega_Serie() As Long
''Carrega a combo de Série
'
'Dim lErro As Long
'Dim colSerie As New colSerie
'Dim objSerie As ClassSerie
'
'On Error GoTo Erro_Carrega_Serie
'
'    'Lê as séries
'    lErro = CF("Series_Le", colSerie)
'    If lErro <> SUCESSO Then Error 19165
'
'    'Carrega na combo
'    For Each objSerie In colSerie
'        Serie.AddItem objSerie.sSerie
'    Next
'
'    Carrega_Serie = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_Serie:
'
'    Carrega_Serie = Err
'
'    Select Case Err
'
'        Case 19165
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160866)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Sub Form_Unload(Cancel As Integer)
'
'    Set objGrid = Nothing
'
'    Set objEventoPedidoDe = Nothing
'    Set objEventoPedidoAte = Nothing
'    Set objEventoClienteDe = Nothing
'    Set objEventoClienteAte = Nothing
'
'    'variaveis auxiliares à contabilizacao
'    Set gobjContabAutomatica = Nothing
'    Set gobjNFiscal = Nothing
'    Set gobjPedidoVenda = Nothing
'    Set gcolAlmoxFilial = Nothing
'
'    Set objEventoLote = Nothing
'    Set objGrid1 = Nothing
'    Set objContabil = Nothing
'
'    Set gobjGeracaoNFiscal = Nothing
'
'End Sub
'
'Private Sub GeraNFiscal_Click()
'
'Dim iClick As Integer
'
'    iAlterado = REGISTRO_ALTERADO
'
'    'Verifica se é alguma linha válida
'    If GridPedido.Row > objGrid.iLinhasExistentes Then Exit Sub
'
'    'Verifica se está selecionando ou desselecionando
'    If Len(Trim(GridPedido.TextMatrix(GridPedido.Row, iGrid_GeraNFiscal_Col))) > 0 Then
'        iClick = StrParaInt(GridPedido.TextMatrix(GridPedido.Row, iGrid_GeraNFiscal_Col)) = MARCADO
'    End If
'
'    If iClick = True Then
'        gobjGeracaoNFiscal.colNFiscalInfo(GridPedido.Row).iMarcada = MARCADO
'    Else
'        gobjGeracaoNFiscal.colNFiscalInfo(GridPedido.Row).iMarcada = DESMARCADO
'    End If
'
'    Exit Sub
'
'End Sub
'
'Private Sub GeraNFiscal_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGrid)
'
'End Sub
'
'Private Sub GeraNFiscal_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
'
'End Sub
'
'Private Sub GeraNFiscal_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGrid.objControle = GeraNFiscal
'    lErro = Grid_Campo_Libera_Foco(objGrid)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Private Sub GridPedido_Click()
'
'Dim iExecutaEntradaCelula As Integer
'
'    Call Grid_Click(objGrid, iExecutaEntradaCelula)
'
'    If iExecutaEntradaCelula = 1 Then
'        Call Grid_Entrada_Celula(objGrid, iAlterado)
'    End If
'
'End Sub
'
'Private Sub GridPedido_EnterCell()
'
'    Call Grid_Entrada_Celula(objGrid, iAlterado)
'
'End Sub
'
'Private Sub GridPedido_GotFocus()
'
'    Call Grid_Recebe_Foco(objGrid)
'
'End Sub
'
'Private Sub GridPedido_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    Call Grid_Trata_Tecla1(KeyCode, objGrid)
'
'End Sub
'
'Private Sub GridPedido_KeyPress(KeyAscii As Integer)
'
'Dim iExecutaEntradaCelula As Integer
'
'    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)
'
'    If iExecutaEntradaCelula = 1 Then
'        Call Grid_Entrada_Celula(objGrid, iAlterado)
'    End If
'
'End Sub
'
'Private Sub GridPedido_LeaveCell()
'
'    Call Saida_Celula(objGrid)
'
'End Sub
'
'Private Sub GridPedido_Validate(Cancel As Boolean)
'
'    Call Grid_Libera_Foco(objGrid)
'
'End Sub
'
'Private Sub GridPedido_RowColChange()
'
'    Call Grid_RowColChange(objGrid)
'
'End Sub
'
'Private Sub GridPedido_Scroll()
'
'    Call Grid_Scroll(objGrid)
'
'End Sub
'
'Private Sub LabelClienteAte_Click()
'
'Dim objCliente As New ClassCliente
'Dim colSelecao As Collection
'
'    'Preenche ClienteAte com o cliente da tela
'    If Len(Trim(ClienteAte.Text)) > 0 Then objCliente.lCodigo = CLng(ClienteAte.Text)
'
'    'Chama Tela ClientesLista
'    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteAte)
'
'End Sub
'
'Private Sub LabelClienteDe_Click()
'
'Dim objCliente As New ClassCliente
'Dim colSelecao As Collection
'
'    'Preenche ClienteDe com o cliente da tela
'    If Len(Trim(ClienteDe.Text)) > 0 Then objCliente.lCodigo = CLng(ClienteDe.Text)
'
'    'Chama Tela ClientesLista
'    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteDe)
'
'End Sub
'
'Private Sub LabelPedidoAte_Click()
'
'Dim objPedidoDeVenda As New ClassPedidoDeVenda
'Dim colSelecao As Collection
'
'    'Preenche PedidoAte com o pedido da tela
'    If Len(Trim(PedidoFinal.Text)) > 0 Then objPedidoDeVenda.lCodigo = CLng(PedidoFinal.Text)
'
'    objPedidoDeVenda.iFilialEmpresa = giFilialEmpresa
'
'    'Chama Tela PedidoVendaLista
'    Call Chama_Tela("PedidoVendaLista", colSelecao, objPedidoDeVenda, objEventoPedidoAte)
'
'End Sub
'
'Private Sub LabelPedidoDe_Click()
'
'Dim objPedidoDeVenda As New ClassPedidoDeVenda
'Dim colSelecao As Collection
'
'    'Preenche PedidoDe com o pedido da tela
'    If Len(Trim(PedidoInicial.Text)) > 0 Then objPedidoDeVenda.lCodigo = CLng(PedidoInicial.Text)
'
'    objPedidoDeVenda.iFilialEmpresa = giFilialEmpresa
'
'    'Chama Tela PedidoVendaLista
'    Call Chama_Tela("PedidoVendaLista", colSelecao, objPedidoDeVenda, objEventoPedidoDe)
'
'End Sub
'
'Private Sub objEventoClienteAte_evSelecao(obj1 As Object)
'
'Dim lErro As Long
'Dim objCliente As ClassCliente
'Dim bCancel As Boolean
'
'On Error GoTo Erro_objEventoClienteAte_evSelecao
'
'    If Not ClienteAte.Enabled Then Exit Sub
'
'    Set objCliente = obj1
'
'    ClienteAte.Text = CStr(objCliente.lCodigo)
'
'    'Chama o Validate de ClienteAte
'    Call ClienteAte_Validate(bCancel)
'
'    iAlterado = 0
'
'    Me.Show
'
'    Exit Sub
'
'Erro_objEventoClienteAte_evSelecao:
'
'    Select Case Err
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160867)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub objEventoClienteDe_evSelecao(obj1 As Object)
'
'Dim lErro As Long
'Dim objCliente As ClassCliente
'
'On Error GoTo Erro_objEventoClienteDe_evSelecao
'
'    Set objCliente = obj1
'
'    'Coloca o código do cliente em cliente de
'    ClienteDe.Text = CStr(objCliente.lCodigo)
'
'    iAlterado = 0
'
'    Me.Show
'
'    Exit Sub
'
'Erro_objEventoClienteDe_evSelecao:
'
'    Select Case Err
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160868)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub objEventoPedidoAte_evSelecao(obj1 As Object)
'
'Dim lErro As Long
'Dim objPedidoDeVenda As ClassPedidoDeVenda
'Dim bCancel As Boolean
'
'On Error GoTo Erro_objEventoPedidoAte_evSelecao
'
'    Set objPedidoDeVenda = obj1
'
'    PedidoFinal.Text = CStr(objPedidoDeVenda.lCodigo)
'
'    'Chama o Validate de PedidoFinal
'    Call PedidoFinal_Validate(bCancel)
'
'    iAlterado = 0
'
'    Me.Show
'
'    Exit Sub
'
'Erro_objEventoPedidoAte_evSelecao:
'
'    Select Case Err
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160869)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub objEventoPedidoDe_evSelecao(obj1 As Object)
'
'Dim lErro As Long
'Dim objPedidoDeVenda As ClassPedidoDeVenda
'
'On Error GoTo Erro_objEventoPedidoDe_evSelecao
'
'    Set objPedidoDeVenda = obj1
'
'    PedidoInicial.Text = CStr(objPedidoDeVenda.lCodigo)
'
'    iAlterado = 0
'
'    Me.Show
'
'    Exit Sub
'
'Erro_objEventoPedidoDe_evSelecao:
'
'    Select Case Err
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160870)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub Ordenados_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub Ordenados_Click()
'
'Dim lErro As Long
'Dim iIndice As Integer
'
'On Error GoTo Erro_Ordenados_Click
'
'    If Ordenados.ListIndex = -1 Then Exit Sub
'
'    'Verifica se a coleção de NFiscal está vazia
'    If gobjGeracaoNFiscal.colNFiscalInfo.Count = 0 Then Exit Sub
'
'    'Passa a Ordenaçao escolhida para o Obj
'    gobjGeracaoNFiscal.sOrdenacao = asOrdenacao(Ordenados.ListIndex)
'
'    'Verifica se tem seleção e Preenche o Grid se Tiver
'    lErro = Traz_Pedidos_Selecionados()
'    If lErro <> SUCESSO And lErro <> 51429 Then Error 58024
'    If lErro = 51429 Then Error 51430
'
'    Exit Sub
'
'Erro_Ordenados_Click:
'
'    Select Case Err
'
'        Case 28476, 58024
'
'        Case 51430
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SEM_PEDIDOS_VENDA_ENCONTRADOS", Err)
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160871)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Function Traz_Pedidos_Selecionados() As Long
'
'Dim lErro As Long
'
'On Error GoTo Erro_Traz_Pedidos_Selecionados
'
'    'Limpa a coleção de NFiscais
'    Set gobjGeracaoNFiscal = New ClassGeracaoNFiscal
'
'    '???? O grid não estáva sendo limpo quando as opções do tab seleção eram alteradas. E quando ele não achava ninguém no BD com as novas características a msg de que não encontrou ninguém era exibida mas apareciam registros no grid.
'    'Limpa o GridPedido
'    Call Grid_Limpa(objGrid)
'
'    lErro = Move_TabSelecao_Memoria
'    If lErro <> SUCESSO Then Error 51448
'
'    'Verifica se foi feita alguma seleção
'    If ExibeTodos.Value = 0 And Len(Trim(PedidoInicial.Text)) = 0 And Len(Trim(PedidoFinal.Text)) = 0 And Len(Trim(ClienteDe.Text)) = 0 And Len(Trim(ClienteAte.Text)) = 0 And _
'        Len(Trim(DataEmissaoDe.ClipText)) = 0 And Len(Trim(DataEmissaoAte.ClipText)) = 0 And Len(Trim(DataEntregaDe.ClipText)) = 0 And Len(Trim(DataEntregaAte.ClipText)) = 0 Then Exit Function
'
'    '???? Marquei a opção de exibir todos e ele carregou o grid no 2º tab. Voltei ao  tab seleção e desmarquei a opção de exibir todos. Ao coltar p\ o 2º tab o grid está vazio, mas a coleção de pedidos ainda está carregada com os pedidos anteriormente selecionados
'    'Preenche a Coleção de NFiscais
'    lErro = CF("GeracaoNFiscal_ObterPedidos", gobjGeracaoNFiscal)
'    If lErro <> SUCESSO And lErro <> 58166 Then Error 58023
'    If lErro = 58166 Then Error 51428
'
'    'Preenche o GridPedido
'    Call Grid_Pedido_Preenche(gobjGeracaoNFiscal.colNFiscalInfo)
'
'    Traz_Pedidos_Selecionados = SUCESSO
'
'    Exit Function
'
'Erro_Traz_Pedidos_Selecionados:
'
'    Traz_Pedidos_Selecionados = Err
'
'    Select Case Err
'
'    Case 58023, 51428, 51448 'Tratado na rotina chamada
'
'    Case Else
'        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160872)
'
'    End Select
'
'    Exit Function
'
'End Function
'
''Private Sub Ordenados_Validate(Cancel As Boolean)
''
''Dim lErro As Long
''Dim iIndice As Integer
''Dim iCodigo As Integer
''
''On Error GoTo Erro_Ordenados_Validate
''
''    'Verifica se a coleção de NFiscal está vazia
''    If gobjGeracaoNFiscal.colNFiscalInfo.Count <> 0 Then
''
''        'Verifica se foi preenchida a ComboBox Ordenados
''        If Len(Trim(Ordenados.Text)) = 0 Then Exit Sub
''
''        'Verifica se está preenchida com o ítem selecionado na ComboBox Ordenados
''        If Ordenados.Text = Ordenados.List(Ordenados.ListIndex) Then Exit Sub
''
''        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
''        lErro = Combo_Seleciona(Ordenados, iCodigo)
''        If lErro <> SUCESSO And lErro <> 6731 Then Error 28477
''
''        'Não existe o ítem com a STRING na List da ComboBox
''        If lErro = 6731 Then Error 28478
''
''        'Passa a Ordenaçao escolhida para o Obj
''        For iIndice = 0 To 4
''
''            If Ordenados.Text = asOrdenacaoString(iIndice) Then gobjGeracaoNFiscal.sOrdenacao = asOrdenacao(iIndice)
''
''        Next
''
''        'Limpa a coleção de NFiscais
''        If Not (gobjGeracaoNFiscal.colNFiscalInfo Is Nothing) Then
''
''            Do While gobjGeracaoNFiscal.colNFiscalInfo.Count <> 0
''
''                gobjGeracaoNFiscal.colNFiscalInfo.Remove (1)
''
''            Loop
''
''        End If
''
''        'Preenche a Coleção de NFiscais
''        lErro = CF("GeracaoNFiscal_ObterPedidos",gobjGeracaoNFiscal)
''        If lErro <> SUCESSO Then Error 28479
''
''        'Limpa o GridPedido
''        Call Grid_Limpa(objGrid)
''        objGrid.iLinhasExistentes = 0
''
''        'Preenche o GridPedido
''        Call Grid_Pedido_Preenche(gobjGeracaoNFiscal.colNFiscalInfo)
''
''    End If
''
''    Exit Sub
''
''Erro_Ordenados_Validate:
'
''    Cancel = True
'
''
''    Select Case Err
''
''        Case 28477, 28479
''
''        Case 28478
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDENACAO_NAO_ENCONTRADA", Err)
''
''        Case Else
''             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160873)
''
''    End Select
''
''    Exit Sub
''
''End Sub
'
'Private Sub PedidoFinal_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    iTabPrincipalAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub PedidoFinal_GotFocus()
'Dim iTabAux As Integer
'
'    iTabAux = iTabPrincipalAlterado
'    Call MaskEdBox_TrataGotFocus(PedidoFinal, iAlterado)
'    iTabPrincipalAlterado = iTabAux
'
'End Sub
'
'Private Sub PedidoFinal_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objPedidoVenda As New ClassPedidoDeVenda
'
'On Error GoTo Erro_PedidoFinal_Validate
'
'    If Len(Trim(PedidoFinal.Text)) > 0 Then
'
'        'Critica para ver se é um Long
'        lErro = Long_Critica(PedidoFinal.Text)
'        If lErro <> SUCESSO Then Error 58007
'
'        'Se o Pedido Final estiver preenchido então
'        If Len(Trim(PedidoInicial.Text)) > 0 Then
'            'Verifica se o Pedido Inicial é maior que o Pedido Final ---- Erro
'            If CLng(PedidoInicial.Text) > CLng(PedidoFinal.Text) Then Error 58008
'        End If
'
'        objPedidoVenda.lCodigo = CLng(PedidoFinal.Text)
'        objPedidoVenda.iFilialEmpresa = giFilialEmpresa
'
'        'Verifica se o Pedido está cadastrado no BD
'        lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
'        If lErro <> SUCESSO And lErro <> 26509 Then Error 58009
'
'        'Pedido não está cadastrado
'        If lErro <> SUCESSO Then Error 58010
'
'    End If
'
'    Exit Sub
'
'Erro_PedidoFinal_Validate:
'
'    Cancel = True
'
'    Select Case Err
'
'        Case 58007, 58009
'
'        Case 58008
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOINICIAL_MAIOR_PEDIDOFINAL", Err)
'
'        Case 58010
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO", Err, objPedidoVenda.lCodigo)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160874)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub PedidoInicial_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    iTabPrincipalAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub PedidoInicial_GotFocus()
'Dim iTabAux As Integer
'
'    iTabAux = iTabPrincipalAlterado
'    Call MaskEdBox_TrataGotFocus(PedidoInicial, iAlterado)
'    iTabPrincipalAlterado = iTabAux
'
'End Sub
'
'Private Sub PedidoInicial_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objPedidoVenda As New ClassPedidoDeVenda
'
'On Error GoTo Erro_PedidoInicial_Validate
'
'    If Len(Trim(PedidoInicial.Text)) > 0 Then
'
'        'Critica para ver se é um Long
'        lErro = Long_Critica(PedidoInicial.Text)
'        If lErro <> SUCESSO Then Error 58003
'
'        'Se o Pedido Final estiver preenchido então
'        If Len(Trim(PedidoFinal.Text)) > 0 Then
'            'Verifica se o Pedido Inicial é maior que o Pedido Final ---- Erro
'            If CLng(PedidoInicial.Text) > CLng(PedidoFinal.Text) Then Error 58004
'        End If
'
'        objPedidoVenda.lCodigo = CLng(PedidoInicial.Text)
'        objPedidoVenda.iFilialEmpresa = giFilialEmpresa
'
'        'Verifica se o Pedido está cadastrado no BD
'        lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
'        If lErro <> SUCESSO And lErro <> 26509 Then Error 58005
'
'        'Pedido não está cadastrado
'        If lErro <> SUCESSO Then Error 58006
'
'    End If
'
'    Exit Sub
'
'Erro_PedidoInicial_Validate:
'
'    Cancel = True
'
'    Select Case Err
'
'        Case 58003, 58005
'
'        Case 58004
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOINICIAL_MAIOR_PEDIDOFINAL", Err)
'
'        Case 58006
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO", Err, objPedidoVenda.lCodigo)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160875)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub Serie_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_Serie_Validate
'
'    'Verifica se foi preenchida
'    If Len(Trim(Serie.Text)) = 0 Then Exit Sub
'
'    'Verifica se foi selecionada
'    If Serie.Text = Serie.List(Serie.ListIndex) Then Exit Sub
'
'    'Tenta selecionar a serie
'    lErro = Combo_Item_Igual(Serie)
'    If lErro <> SUCESSO And lErro <> 12253 Then Error 44194
'
'    'Se não conseguir --> Erro
'    If lErro <> SUCESSO Then Error 44195
'
'    Exit Sub
'
'Erro_Serie_Validate:
'
'    Cancel = True
'
'
'    Select Case Err
'
'        Case 44194
'
'        Case 44195
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", Err, Serie.Text)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160876)
'
'    End Select
'
'End Sub
'
'
'Private Sub TabStrip1_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_TabStrip1_Click
'
'    'Se Frame atual não corresponde ao Tab clicado
'    If TabStrip1.SelectedItem.Index <> iFrameAtual Then
'
'        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
'
'        'Torna Frame de Pedido visível
'        Frame1(TabStrip1.SelectedItem.Index).Visible = True
'        'Torna Frame atual invisível
'        Frame1(iFrameAtual).Visible = False
'        'Armazena novo valor de iFrameAtual
'        iFrameAtual = TabStrip1.SelectedItem.Index
'
'        'Se Frame selecionado foi o de Pedido
'        If TabStrip1.SelectedItem.Index = TAB_Pedidos Then
'            If iTabPrincipalAlterado = REGISTRO_ALTERADO Then
'                lErro = Trata_TabPedidos()
'                If lErro <> SUCESSO And lErro <> 51429 Then Error 31382
'                If lErro <> SUCESSO Then Error 51433
'            End If
'
'        End If
'
'        Select Case iFrameAtual
'
'            Case TAB_Selecao
'                Parent.HelpContextID = IDH_GERACAO_NFISCAL_SELECAO
'
'            Case TAB_Pedidos
'                Parent.HelpContextID = IDH_GERACAO_NFISCAL_PEDIDOS
'
'        End Select
'
'    End If
'
'    Exit Sub
'
'Erro_TabStrip1_Click:
'
'    Select Case Err
'
'        Case 31382
'
'        Case 51433
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SEM_PEDIDOS_VENDA_ENCONTRADOS", Err)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160877)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub UpDownEmissaoAte_DownClick()
'
'Dim lErro As Long
'
'On Error GoTo Erro_UpDownEmissaoAte_DownClick
'
'    'Diminui a DataEmissaoAte em 1 dia
'    lErro = Data_Up_Down_Click(DataEmissaoAte, DIMINUI_DATA)
'    If lErro <> SUCESSO Then Error 28462
'
'    Exit Sub
'
'Erro_UpDownEmissaoAte_DownClick:
'
'    Select Case Err
'
'        Case 28462
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160878)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub UpDownEmissaoAte_UpClick()
'
'Dim lErro As Long
'
'On Error GoTo Erro_UpDownEmissaoAte_UpClick
'
'    'Aumenta a DataEmissaoAte em 1 dia
'    lErro = Data_Up_Down_Click(DataEmissaoAte, AUMENTA_DATA)
'    If lErro <> SUCESSO Then Error 28461
'
'    Exit Sub
'
'Erro_UpDownEmissaoAte_UpClick:
'
'    Select Case Err
'
'        Case 28461
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160879)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub UpDownEmissaoDe_DownClick()
'
'Dim lErro As Long
'
'On Error GoTo Erro_UpDownEmissaoDe_DownClick
'
'    'Diminui a DataEmissaoDe em 1 dia
'    lErro = Data_Up_Down_Click(DataEmissaoDe, DIMINUI_DATA)
'    If lErro <> SUCESSO Then Error 28459
'
'    Exit Sub
'
'Erro_UpDownEmissaoDe_DownClick:
'
'    Select Case Err
'
'        Case 28459
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160880)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub UpDownEmissaoDe_UpClick()
'
'Dim lErro As Long
'
'On Error GoTo Erro_UpDownEmissaoDe_UpClick
'
'    'Aumenta a DataEmissaoDe em 1 dia
'    lErro = Data_Up_Down_Click(DataEmissaoDe, AUMENTA_DATA)
'    If lErro <> SUCESSO Then Error 28460
'
'    Exit Sub
'
'Erro_UpDownEmissaoDe_UpClick:
'
'    Select Case Err
'
'        Case 28460
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160881)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub UpDownEntregaAte_DownClick()
'
'Dim lErro As Long
'
'On Error GoTo Erro_UpDownEntregaAte_DownClick
'
'    'Diminui a DataEntregaAte em 1 dia
'    lErro = Data_Up_Down_Click(DataEntregaAte, DIMINUI_DATA)
'    If lErro <> SUCESSO Then Error 28467
'
'    Exit Sub
'
'Erro_UpDownEntregaAte_DownClick:
'
'    Select Case Err
'
'        Case 28467
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160882)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub UpDownEntregaAte_UpClick()
'
'Dim lErro As Long
'
'On Error GoTo Erro_UpDownEntregaAte_UpClick
'
'    'Aumenta a DataEntregaAte em 1 dia
'    lErro = Data_Up_Down_Click(DataEntregaAte, AUMENTA_DATA)
'    If lErro <> SUCESSO Then Error 28468
'
'    Exit Sub
'
'Erro_UpDownEntregaAte_UpClick:
'
'    Select Case Err
'
'        Case 28468
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160883)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub UpDownEntregaDe_DownClick()
'
'Dim lErro As Long
'
'On Error GoTo Erro_UpDownEntregaDe_DownClick
'
'    'Diminui a DataEntregaDe em 1 dia
'    lErro = Data_Up_Down_Click(DataEntregaDe, DIMINUI_DATA)
'    If lErro <> SUCESSO Then Error 28465
'
'    Exit Sub
'
'Erro_UpDownEntregaDe_DownClick:
'
'    Select Case Err
'
'        Case 28465
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160884)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub UpDownEntregaDe_UpClick()
'
'Dim lErro As Long
'
'On Error GoTo Erro_UpDownEntregaDe_UpClick
'
'    'Aumenta a DataEntregaDe em 1 dia
'    lErro = Data_Up_Down_Click(DataEntregaDe, AUMENTA_DATA)
'    If lErro <> SUCESSO Then Error 28466
'
'    Exit Sub
'
'Erro_UpDownEntregaDe_UpClick:
'
'    Select Case Err
'
'        Case 28466
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160885)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Function Grid_Pedido_Preenche(colNFiscalInfo As Collection) As Long
''Preenche o Grid Pedido com os dados de colNFiscalInfo
'
'Dim lErro As Long
'Dim iLinha As Integer
'Dim iIndice As Integer
'Dim objNFiscalInfo As ClassNFiscalInfo
'Dim objTransportadora As New ClassTransportadora
'Dim objFilialEmpresa As New AdmFiliais
'Dim colCodigoNome As New AdmColCodigoNome
'Dim objCodigoNome As New AdmCodigoNome
'Dim bAchou As Boolean
'
'On Error GoTo Erro_Grid_Pedido_Preenche
'
'    'Se o número de NFiscal for maior que o número de linhas do Grid
'    If colNFiscalInfo.Count + 1 > GridPedido.Rows Then
'
'        If colNFiscalInfo.Count > NUM_MAXIMO_PARCELAS Then Error 19167
'
'        'Altera o número de linhas do Grid de acordo com o número de NFiscal
'        GridPedido.Rows = colNFiscalInfo.Count + 1
'
'        'Chama rotina de Inicialização do Grid
'        Call Grid_Inicializa(objGrid)
'
'    End If
'
'    iLinha = 0
'
'    'Percorre todas as NFiscais da Coleção
'    For Each objNFiscalInfo In colNFiscalInfo
'
'        iLinha = iLinha + 1
'
'        'Passa para a tela os dados da NFiscal em questão
'        GridPedido.TextMatrix(iLinha, iGrid_GeraNFiscal_Col) = objNFiscalInfo.iMarcada
'        GridPedido.TextMatrix(iLinha, iGrid_Pedido_Col) = objNFiscalInfo.lPedido
'        GridPedido.TextMatrix(iLinha, iGrid_Cliente_Col) = objNFiscalInfo.lCliente
'        GridPedido.TextMatrix(iLinha, iGrid_NomeRed_Col) = objNFiscalInfo.sClienteNomeReduzido
'
'        If objNFiscalInfo.iMotivoNaoGerada = MOTIVO_NAOGERADA_POR_BLOQUEIO Then
'            GridPedido.TextMatrix(iLinha, iGrid_Motivo_Col) = MOTIVO_NAOGERADA_DESCRICAO_BLOQUEIO
'        ElseIf objNFiscalInfo.iMotivoNaoGerada = MOTIVO_NAOGERADA_POR_FALTAESTOQUE Then
'            GridPedido.TextMatrix(iLinha, iGrid_Motivo_Col) = MOTIVO_NAOGERADA_DESCRICAO_FALTA_ESTOQUE
'        ElseIf objNFiscalInfo.iMotivoNaoGerada = MOTIVO_NAOGERADA_OUTROS Then
'            GridPedido.TextMatrix(iLinha, iGrid_Motivo_Col) = MOTIVO_NAOGERADA_DESCRICAO_OUTROS
'        ElseIf objNFiscalInfo.iMotivoNaoGerada = MOTIVO_NAOGERADA_POR_BLOQUEIO_CREDITO Then
'            GridPedido.TextMatrix(iLinha, iGrid_Motivo_Col) = MOTIVO_NAOGERADA_DESCRICAO_BLOQUEIO_CREDITO
'        End If
'
'        ' Se a transportadora foi informada
'        If objNFiscalInfo.iCodTransp > 0 Then
'
'            objTransportadora.iCodigo = objNFiscalInfo.iCodTransp
'            'Lê a transportadora
'            lErro = CF("Transportadora_Le", objTransportadora)
'            If lErro <> SUCESSO And lErro <> 19250 Then Error 51360
'            If lErro <> SUCESSO Then Error 51361
'            'Coloca a transportadora no grid
'            objNFiscalInfo.sNomeRedTransp = objTransportadora.sNomeReduzido
'            GridPedido.TextMatrix(iLinha, iGrid_TransPortadora_Col) = objNFiscalInfo.iCodTransp & SEPARADOR & objNFiscalInfo.sNomeRedTransp
'
'        End If
'
'
'        If objNFiscalInfo.dtEmissaoPedido <> DATA_NULA And objNFiscalInfo.dtEmissaoPedido <> 0 Then GridPedido.TextMatrix(iLinha, iGrid_Emissao_Col) = Format(objNFiscalInfo.dtEmissaoPedido, "dd/mm/yyyy")
'        If objNFiscalInfo.dtEntregaPedido <> DATA_NULA And objNFiscalInfo.dtEntregaPedido Then GridPedido.TextMatrix(iLinha, iGrid_Entrega_Col) = Format(objNFiscalInfo.dtEntregaPedido, "dd/mm/yyyy")
'        GridPedido.TextMatrix(iLinha, iGrid_Valor_Col) = Format(objNFiscalInfo.dValorTotal, "Standard")
'        GridPedido.TextMatrix(iLinha, iGrid_Filial_Col) = objNFiscalInfo.iFilialCliente
'        GridPedido.TextMatrix(iLinha, iGrid_Estado_Col) = objNFiscalInfo.sSiglaEstadoEntrega
'        GridPedido.TextMatrix(iLinha, iGrid_Cidade_Col) = objNFiscalInfo.sCidadeEntrega
'        GridPedido.TextMatrix(iLinha, iGrid_Bairro_Col) = objNFiscalInfo.sBairroEntrega
'
'        bAchou = False
'        'Verifica se a FilialEmpresa do pedido já foi lida
'        For Each objCodigoNome In colCodigoNome
'            If objCodigoNome.iCodigo = objNFiscalInfo.iFilialEmpresa Then
'                objFilialEmpresa.iCodFilial = objNFiscalInfo.iFilialEmpresa
'                objFilialEmpresa.sNome = objCodigoNome.sNome
'                bAchou = True
'                Exit For
'            End If
'        Next
'        'Se ainda não foi lida
'        If Not bAchou Then
'
'            objFilialEmpresa.iCodFilial = objNFiscalInfo.iFilialEmpresa
'            objFilialEmpresa.lCodEmpresa = glEmpresa
'            'Lê a FilialEmpresa
'            lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
'            If lErro <> SUCESSO And lErro <> 27378 Then Error 51434
'            If lErro <> SUCESSO Then Error 51435
'            'Adiciona na coleção das filiais já lidas a filial lida agora
'            colCodigoNome.Add objFilialEmpresa.iCodFilial, objFilialEmpresa.sNome
'
'        End If
'        '???? O campo empresa do grid de pedidos é na verdade o campo filialempresa. Colocar na forma "codigo-nomered" e alterar o título.
'        'Coloca no grid o código-Nome da Filial do Pedido
'        GridPedido.TextMatrix(iLinha, iGrid_FilialEmpresa_Col) = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
'
'    Next
'
'    Call Grid_Refresh_Checkbox(objGrid)
'
'    'Passa para o Obj o número de NFiscais passados pela Coleção
'    objGrid.iLinhasExistentes = colNFiscalInfo.Count
'
'    Grid_Pedido_Preenche = SUCESSO
'
'    Exit Function
'
'Erro_Grid_Pedido_Preenche:
'
'    Grid_Pedido_Preenche = Err
'
'    Select Case Err
'
'        Case 19167
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_MAXIMO_PARCELAS_ULTRAPASSADO", Err, colNFiscalInfo.Count, NUM_MAXIMO_PARCELAS)
'
'        Case 51434
'
'        Case 51435
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err, objFilialEmpresa.iCodFilial)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160886)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function Saida_Celula(objGridInt As AdmGrid) As Long
''Faz a crítica da célula do grid que está deixando de ser a corrente
'
'Dim lErro As Long
'
'On Error GoTo Erro_Saida_Celula
'
'    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
'
'    If lErro = SUCESSO Then
'
'        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
'        If lErro <> SUCESSO Then Error 28481
'
'    End If
'
'    Saida_Celula = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula:
'
'    Saida_Celula = Err
'
'    Select Case Err
'
'        Case 28481
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 160887)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Trata_TabPedidos() As Long
'
'Dim lErro As Long
'
'On Error GoTo Erro_Trata_TabPedidos
'
'    If Ordenados.ListIndex = -1 Then
'
'        Ordenados.ListIndex = 0
'
'    Else
'
'        'Verifica se tem seleção e Preenche o Grid de acordo com a ordenação
'        lErro = Traz_Pedidos_Selecionados()
'        If lErro <> SUCESSO And lErro <> 51428 Then Error 58025
'        If lErro = 51428 Then Error 51429
'
'    End If
'
'    iTabPrincipalAlterado = 0
'
'    Exit Function
'
'Erro_Trata_TabPedidos:
'
'    Trata_TabPedidos = Err
'
'    Select Case Err
'
'        Case 58025, 51429
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160888)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Move_TabSelecao_Memoria() As Long
'
'Dim lErro As Long
'
'On Error GoTo Erro_Move_TabSelecao_Memoria
'
'    '???? Quando eu trago os clientes pelo browse antes de selecionar os pedidos ele não verifica se o cliente De é anterior ao Cliente Até
'
'    gobjGeracaoNFiscal.iTodosOsPedidos = ExibeTodos.Value
'
'    'Recolhe a data de emissão De
'    gobjGeracaoNFiscal.dtEmissaoDe = StrParaDate(DataEmissaoDe.Text)
'    'Recolhe a Data de emissão Até
'    gobjGeracaoNFiscal.dtEmissaoAte = StrParaDate(DataEmissaoAte.Text)
'
'    'Se a Data de emissão De e Até foram preenchidas
'    If gobjGeracaoNFiscal.dtEmissaoDe <> DATA_NULA And gobjGeracaoNFiscal.dtEmissaoAte <> DATA_NULA Then
'        'Verifica se Data Emissão De é anterior a Data de Emissão Até
'        If gobjGeracaoNFiscal.dtEmissaoAte < gobjGeracaoNFiscal.dtEmissaoDe Then Error 51444
'    End If
'
'    'Recolhe as data de entrega DE e ATÉ
'    gobjGeracaoNFiscal.dtEntregaDe = StrParaDate(DataEntregaDe.Text)
'    gobjGeracaoNFiscal.dtEntregaAte = StrParaDate(DataEntregaAte.Text)
'
'    'Se a Data de entrega De e Até foram preenchidas
'    If gobjGeracaoNFiscal.dtEntregaDe <> DATA_NULA And gobjGeracaoNFiscal.dtEntregaAte <> DATA_NULA Then
'        'Verifica se Data Entrega De é anterior a Data de Entrega Até
'        If gobjGeracaoNFiscal.dtEntregaAte < gobjGeracaoNFiscal.dtEntregaDe Then Error 51445
'    End If
'
'    'Recolhe Pedido De e Até
'    gobjGeracaoNFiscal.lPedidosDe = StrParaLong(PedidoInicial.Text)
'    gobjGeracaoNFiscal.lPedidosAte = StrParaLong(PedidoFinal.Text)
'
'    'Se PedidoFinal e PedidoInicial estão preenchidos
'    If gobjGeracaoNFiscal.lPedidosDe <> 0 And gobjGeracaoNFiscal.lPedidosAte <> 0 Then
'        'Verifica se Data Pedido De é menor que pedido Até
'        If gobjGeracaoNFiscal.lPedidosAte < gobjGeracaoNFiscal.lPedidosDe Then Error 51446
'    End If
'
'    'Recolhe o Cliente De e o ATé
'    gobjGeracaoNFiscal.lClientesDe = StrParaLong(ClienteDe.Text)
'    gobjGeracaoNFiscal.lClientesAte = StrParaLong(ClienteAte.Text)
'
'    'Se ClienteAté e ClienteDe estão preenchidos
'    If gobjGeracaoNFiscal.lClientesDe <> 0 And gobjGeracaoNFiscal.lClientesAte <> 0 Then
'        'Verifica se Cliente De é menor que Cliente Até
'        If gobjGeracaoNFiscal.lClientesAte < gobjGeracaoNFiscal.lClientesDe Then Error 51447
'    End If
'
'    'Sairam os campos Filiais
'    gobjGeracaoNFiscal.iFilialPedidoDe = 0
'    gobjGeracaoNFiscal.iFilialPedidoAte = 0
'
'    gobjGeracaoNFiscal.sOrdenacao = asOrdenacao(Ordenados.ListIndex)
'
'    Move_TabSelecao_Memoria = SUCESSO
'
'    Exit Function
'
'Erro_Move_TabSelecao_Memoria:
'
'    Move_TabSelecao_Memoria = Err
'
'    Select Case Err
'
'        Case 51444
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAODE_MAIOR_DATAEMISSAOATE", Err, DataEmissaoDe.Text, DataEmissaoAte.Text)
'
'        Case 51445
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAENTREGADE_MAIOR_DATAENTREGAATE", Err, DataEntregaDe.Text, DataEntregaAte.Text)
'
'        Case 51446
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOINICIAL_MAIOR_PEDIDOFINAL", Err)
'
'        Case 51447
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTEDE_MAIOR_CLIENTEATE", Err)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160889)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Sub DataSaida_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_DataSaida_Validate
'
'    'Verifica se a Data de Saida foi digitada
'    If Len(Trim(DataSaida.ClipText)) = 0 Then Exit Sub
'
'    'Critica a data digitada
'    lErro = Data_Critica(DataSaida.Text)
'    If lErro <> SUCESSO Then Error 31385
'
'    Exit Sub
'
'Erro_DataSaida_Validate:
'
'    Cancel = True
'
'
'    Select Case Err
'
'        Case 31385
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160890)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub UpDownSaida_DownClick()
'
'Dim lErro As Long
'Dim sData As String
'
'On Error GoTo Erro_UpDownSaida_DownClick
'
'    lErro = Data_Up_Down_Click(DataSaida, DIMINUI_DATA)
'    If lErro Then Error 31384
'
'    Exit Sub
'
'Erro_UpDownSaida_DownClick:
'
'    Select Case Err
'
'        Case 31384
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160891)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub UpDownSaida_UpClick()
'
'Dim lErro As Long
'Dim sData As String
'
'On Error GoTo Erro_UpDownSaida_UpClick
'
'    lErro = Data_Up_Down_Click(DataSaida, AUMENTA_DATA)
'    If lErro Then Error 31383
'
'    Exit Sub
'
'Erro_UpDownSaida_UpClick:
'
'    Select Case Err
'
'        Case 31383
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160892)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'
'Function Trata_Parametros() As Long
'
'    Trata_Parametros = SUCESSO
'
'End Function
'
''**** inicio do trecho a ser copiado *****
'Public Function Form_Load_Ocx() As Object
'
'    Parent.HelpContextID = IDH_GERACAO_NFISCAL_SELECAO
'    Set Form_Load_Ocx = Me
'    Caption = "Geração de Notas Fiscais"
'    Call Form_Load
'
'End Function
'
'Public Function Name() As String
'
'    Name = "GeracaoNFiscal"
'
'End Function
'
'Public Sub Show()
'    Parent.Show
'    Parent.SetFocus
'End Sub
'
'
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,Controls
'Public Property Get Controls() As Object
'    Set Controls = UserControl.Controls
'End Property
'
'Public Property Get hWnd() As Long
'    hWnd = UserControl.hWnd
'End Property
'
'Public Property Get Height() As Long
'    Height = UserControl.Height
'End Property
'
'Public Property Get Width() As Long
'    Width = UserControl.Width
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,ActiveControl
'Public Property Get ActiveControl() As Object
'    Set ActiveControl = UserControl.ActiveControl
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,Enabled
'Public Property Get Enabled() As Boolean
'    Enabled = UserControl.Enabled
'End Property
'
'Public Property Let Enabled(ByVal New_Enabled As Boolean)
'    UserControl.Enabled() = New_Enabled
'    PropertyChanged "Enabled"
'End Property
'
''Load property values from storage
'Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'
'    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
'End Sub
'
''Write property values to storage
'Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'
'    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
'End Sub
'
'Private Sub Unload(objme As Object)
'   ' Parent.UnloadDoFilho
'
'   RaiseEvent Unload
'
'End Sub
'
'Public Property Get Caption() As String
'    Caption = m_Caption
'End Property
'
'Public Property Let Caption(ByVal New_Caption As String)
'    Parent.Caption = New_Caption
'    m_Caption = New_Caption
'End Property
'
''***** fim do trecho a ser copiado ******
'
'Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = KEYCODE_BROWSER Then
'
'        If Me.ActiveControl Is Pedido Then
'            Call BotaoPedido_Click
'        ElseIf Me.ActiveControl Is PedidoInicial Then
'            Call LabelPedidoDe_Click
'        ElseIf Me.ActiveControl Is PedidoFinal Then
'            Call LabelPedidoAte_Click
'        ElseIf Me.ActiveControl Is ClienteDe Then
'            Call LabelClienteDe_Click
'        ElseIf Me.ActiveControl Is ClienteAte Then
'            Call LabelClienteAte_Click
'        End If
'
'    End If
'
'End Sub
'
'
'Private Sub NotaFiscal_Imprime()
'
'Dim lErro As Long
'Dim objRelatorio As New AdmRelatorio
'
'    If gobjGeracaoNFiscal.iImprime = IMPRIME_NOTA_FISCAL Then
'
'        'se estiver gerando uma nota fiscal fatura
'        If gobjGeracaoNFiscal.iTipoNFiscal = DOCINFO_NFISFVPV Then
'
'            Call objRelatorio.Rel_Menu_Executar("Emissão das Notas Fiscais Fatura", Serie.Text)
'
'        Else 'gerando uma nf que nao seja fatura
'            Call objRelatorio.Rel_Menu_Executar("Emissão das Notas Fiscais", Serie.Text)
'
'        End If
'
'    End If
'
'
'End Sub
'
'Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'    Call Controle_DragDrop(Label1(Index), Source, X, Y)
'End Sub
'
'Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
'End Sub
'
'
'Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label4, Source, X, Y)
'End Sub
'
'Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
'End Sub
'
'Private Sub label_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(label, Source, X, Y)
'End Sub
'
'Private Sub label_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(label, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label19, Source, X, Y)
'End Sub
'
'Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
'End Sub
'
'Private Sub LabelPedidoDe_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(LabelPedidoDe, Source, X, Y)
'End Sub
'
'Private Sub LabelPedidoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelPedidoDe, Button, Shift, X, Y)
'End Sub
'
'Private Sub LabelPedidoAte_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(LabelPedidoAte, Source, X, Y)
'End Sub
'
'Private Sub LabelPedidoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelPedidoAte, Button, Shift, X, Y)
'End Sub
'
'Private Sub LabelClienteAte_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(LabelClienteAte, Source, X, Y)
'End Sub
'
'Private Sub LabelClienteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelClienteAte, Button, Shift, X, Y)
'End Sub
'
'Private Sub LabelClienteDe_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(LabelClienteDe, Source, X, Y)
'End Sub
'
'Private Sub LabelClienteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelClienteDe, Button, Shift, X, Y)
'End Sub
'
'Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long
'
'Dim lErro As Long
'Dim iLinha As Integer
'Dim dQuantidadeConvertida As Double
'Dim sProdutoFormatado As String, sProdutoTela As String
'Dim iPreenchido As Integer
'Dim objEstoqueProduto As New ClassEstoqueProduto
'Dim bEncontrouProduto As Boolean
'Dim iLinha2 As Integer, objCodigoNome As New AdmlCodigoNome
'Dim objAlmoxarifado As New ClassAlmoxarifado
'Dim sContaMascarada As String, sAlmoxNomeRed As String
'Dim iAlmoxPadrao As Integer, objCliente As New ClassCliente
'Dim bEncontrouQuant As Boolean, objFilialCliente As New ClassFilialCliente
'Dim bEncontrouQuant2 As Boolean, objItem As ClassItemNF, objAlocacao As ClassItemNFAlocacao
'Dim objItemMovEstoque As New ClassItemMovEstoque
'Dim sDocInfo As String
'
'On Error GoTo Erro_Calcula_Mnemonico
'
'    Select Case objMnemonicoValor.sMnemonico
'
'
'        Case ESCANINHO_CUSTO
'
'            For Each objItem In gobjNFiscal.ColItensNF
'
'                For Each objAlocacao In objItem.colAlocacoes
'                    objMnemonicoValor.colValor.Add ESCANINHO_NOSSO
'                Next
'
'            Next
'
'        Case ESCANINHO_CUSTO_CONSIG
'
'            For Each objItem In gobjNFiscal.ColItensNF
'
'                For Each objAlocacao In objItem.colAlocacoes
'                    objMnemonicoValor.colValor.Add ESCANINHO_3_EM_CONSIGNACAO
'                Next
'
'            Next
'
'        Case QUANT_ALOCADA_CONSIG
'
'            For Each objItem In gobjNFiscal.ColItensNF
'
'                For Each objAlocacao In objItem.colAlocacoes
'
'                    'Define o produto que será passado como parâmetro para MovEstoque_Le_QuantVendConsig
'                    lErro = CF("Produto_Formata", objItem.sProduto, sProdutoFormatado, iPreenchido)
'                    If lErro <> SUCESSO Then gError 79012
'
'                    objItemMovEstoque.sProduto = sProdutoFormatado
'
'                    'Define o almoxarifado que será passado como parâmetro para MovEstoque_Le_QuantVendConsig
'                    objItemMovEstoque.iAlmoxarifado = objAlocacao.iAlmoxarifado
'
'                    'Define o tipo de movimento, o DocOrigem e o TipoNumIntoDocOrigem que serão passados como parâmetros para MovEstoque_Le_QuantVendConsig
'
'                    objItemMovEstoque.iTipoMov = MOV_EST_NF_VENDA_MAT_CONSIG
'
'                    'Define a sigla do DocInfo que será passado como parâmetro para a função MovEstoque_Le_QuantVendConsig
'                    Select Case gobjGeracaoNFiscal.iTipoNFiscal
'                        Case DOCINFO_NFISFVPV
'                            sDocInfo = "NFISFVPV"
'
'                        Case DOCINFO_NFISVPV
'                            sDocInfo = "NFISVPV"
'                    End Select
'
'                    objItemMovEstoque.sDocOrigem = sDocInfo & " " & Serie.Text & " " & gobjNFiscal.lNumNotaFiscal
'                    objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
'
'                    'Verifica se MovEstoque_Le_QuantVendConsig não encontrou erro
'                    lErro = CF("MovEstoque_Le_QuantVendConsig", objItemMovEstoque)
'                    If lErro <> SUCESSO And lErro <> 79003 Then gError 79010
'
'                    'se não conseguiu encontrar mov. estoque para os parametros em questao ==> quantidade é zerada
'                    If lErro = 79003 Then objItemMovEstoque.dQuantidade = 0
'
'                    'Passa para o mnemônico o valor encontrado por MovEstoque_Le_QuantVendConsig
'                    objMnemonicoValor.colValor.Add objItemMovEstoque.dQuantidade
'
'                    Next
'
'                Next
'
'        Case CTACONTABILEST1 'parametros: produto no formato da tela do grid de itens, produto no formato da tela do grid de alocacoes e nome reduzido do almoxarifado
'
'            For Each objItem In gobjNFiscal.ColItensNF
'
'                For Each objAlocacao In objItem.colAlocacoes
'
'                    If objMnemonicoValor.vParam(1) = objMnemonicoValor.vParam(2) Then
'
'                        objAlmoxarifado.sNomeReduzido = CStr(objMnemonicoValor.vParam(3))
'
'                        'Lê o Nome Reduzido do Almoxarifado
'                        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
'                        If lErro <> SUCESSO And lErro <> 25060 Then Error 59388
'
'                        'Se não encontrou ===> Erro
'                        If lErro = SUCESSO Then
'
'                            objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
'
'                            lErro = CF("Produto_Formata", objMnemonicoValor.vParam(1), sProdutoFormatado, iPreenchido)
'                            If iPreenchido <> PRODUTO_PREENCHIDO Then Error 59390
'
'                            objEstoqueProduto.sProduto = sProdutoFormatado
'
'                            'Lê a conta contábil do Produto no Almoxarifado
'                            lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
'                            If lErro <> SUCESSO And lErro <> 49991 Then Error 59391
'
'                            If Len(Trim(objEstoqueProduto.sContaContabil)) > 0 Then
'
'                                sContaMascarada = String(STRING_CONTA, 0)
'
'                                lErro = Mascara_RetornaContaTela(objEstoqueProduto.sContaContabil, sContaMascarada)
'                                If lErro <> SUCESSO Then Error 64225
'
'                                objMnemonicoValor.colValor.Add sContaMascarada
'                            Else
'                                objMnemonicoValor.colValor.Add ""
'                            End If
'                        Else
'                            objMnemonicoValor.colValor.Add ""
'                        End If
'                    Else
'                        objMnemonicoValor.colValor.Add ""
'                    End If
'
'                Next
'
'            Next
'
'        Case CODIGO1
'            objMnemonicoValor.colValor.Add gobjPedidoVenda.lCodigo
'
'        Case QUANT_ESTOQUE
'
'            For Each objItem In gobjNFiscal.ColItensNF
'
'                lErro = CF("UMEstoque_Conversao", objItem.sProduto, objItem.sUnidadeMed, objItem.dQuantidade, dQuantidadeConvertida)
'                If lErro <> SUCESSO Then Error 64214
'
'                objMnemonicoValor.colValor.Add dQuantidadeConvertida
'
'            Next
'
'        Case ALMOX1
'
'            If gcolAlmoxFilial.Count = 0 Then
'
'                lErro = CF("Almoxarifados_Le_FilialEmpresa", gobjNFiscal.iFilialEmpresa, gcolAlmoxFilial)
'                If lErro <> SUCESSO Then Error 59423
'
'            End If
'
'            For Each objItem In gobjNFiscal.ColItensNF
'
'                For Each objAlocacao In objItem.colAlocacoes
'
'                    sAlmoxNomeRed = ""
'
'                    For Each objAlmoxarifado In gcolAlmoxFilial
'
'                        If objAlocacao.iAlmoxarifado = objAlmoxarifado.iCodigo Then
'
'                            sAlmoxNomeRed = objAlmoxarifado.sNomeReduzido
'                            Exit For
'
'                        End If
'
'                    Next
'
'                    objMnemonicoValor.colValor.Add sAlmoxNomeRed
'
'                Next
'
'            Next
'
'        Case DATA_EMISSAO
'            objMnemonicoValor.colValor.Add gobjNFiscal.dtDataEmissao
'
'        Case DATA_SAIDA
'            objMnemonicoValor.colValor.Add gobjNFiscal.dtDataSaida
'
'        Case DESCONTO1
'            For Each objItem In gobjNFiscal.ColItensNF
'
'                objMnemonicoValor.colValor.Add objItem.dValorDesconto
'
'            Next
'
'        Case DESCRICAO_ITEM
'            For Each objItem In gobjNFiscal.ColItensNF
'
'                objMnemonicoValor.colValor.Add objItem.sDescricaoItem
'
'            Next
'
'        Case FILIAL1
'
'            objFilialCliente.lCodCliente = gobjNFiscal.lCliente
'            objFilialCliente.iCodFilial = gobjNFiscal.iFilialCli
'
'            lErro = CF("FilialCliente_Le", objFilialCliente)
'            If lErro <> SUCESSO And lErro <> 12567 Then Error 59392
'            If lErro <> SUCESSO Then Error 59393
'
'            objMnemonicoValor.colValor.Add objFilialCliente.sNome
'
'        Case CLIENTE1
'
'            objCodigoNome.lCodigo = gobjNFiscal.lCliente
'
'            lErro = CF("Cliente_Le_NomeRed", objCodigoNome)
'            If lErro <> SUCESSO And lErro <> 12553 Then Error 59394
'            If lErro <> SUCESSO Then Error 59395
'
'            objMnemonicoValor.colValor.Add objCodigoNome.sNome
'
'        Case NATUREZA_OP
'            objMnemonicoValor.colValor.Add gobjNFiscal.sNaturezaOp
'
'        Case NFISCAL1
'            objMnemonicoValor.colValor.Add gobjNFiscal.lNumNotaFiscal
'
'        Case PRODUTO1
'            For Each objItem In gobjNFiscal.ColItensNF
'
'                lErro = Mascara_RetornaProdutoTela(objItem.sProduto, sProdutoTela)
'                If lErro <> SUCESSO Then Error 59397
'
'                objMnemonicoValor.colValor.Add sProdutoTela
'
'            Next
'
'        Case QUANTIDADE1
'            For Each objItem In gobjNFiscal.ColItensNF
'
'                objMnemonicoValor.colValor.Add objItem.dQuantidade
'
'            Next
'
'        Case Serie1
'            objMnemonicoValor.colValor.Add gobjNFiscal.sSerie
'
'        Case UNIDADE_MED
'            For Each objItem In gobjNFiscal.ColItensNF
'
'                objMnemonicoValor.colValor.Add objItem.sUnidadeMed
'
'            Next
'
'        Case VALOR_TOTAL
'            objMnemonicoValor.colValor.Add gobjNFiscal.dValorTotal
'
'        Case PRECO_UNITARIO
'            For Each objItem In gobjNFiscal.ColItensNF
'
'                objMnemonicoValor.colValor.Add objItem.dPrecoUnitario
'
'            Next
'
'        Case PRECO_TOTAL
'            For Each objItem In gobjNFiscal.ColItensNF
'
'                objMnemonicoValor.colValor.Add objItem.dValorTotal
'
'            Next
'
'        Case ICMS
'            objMnemonicoValor.colValor.Add gobjNFiscal.objTributacaoNF.dICMSValor
'
'        Case ICMSSUBST
'            objMnemonicoValor.colValor.Add gobjNFiscal.objTributacaoNF.dICMSSubstValor
'
'        Case VALOR_FRETE
'            objMnemonicoValor.colValor.Add gobjNFiscal.dValorFrete
'
'        Case VALOR_SEGURO
'            objMnemonicoValor.colValor.Add gobjNFiscal.dValorSeguro
'
'        Case VALOR_DESPESAS
'            objMnemonicoValor.colValor.Add gobjNFiscal.dValorOutrasDespesas
'
'        Case IPI
'            objMnemonicoValor.colValor.Add gobjNFiscal.objTributacaoNF.dIPIValor
'
'        Case VALOR_DESCONTO
'            objMnemonicoValor.colValor.Add gobjNFiscal.dValorDesconto
'
'        Case ISS_VALOR
'            objMnemonicoValor.colValor.Add gobjNFiscal.objTributacaoNF.dISSValor
'
'        Case ISS_INCLUSO
'
'            objMnemonicoValor.colValor.Add gobjNFiscal.objTributacaoNF.iISSIncluso <> 0
'
'        Case VALOR_IRRF
'            objMnemonicoValor.colValor.Add gobjNFiscal.objTributacaoNF.dIRRFValor
'
'        Case VALOR_PRODUTOS
'            objMnemonicoValor.colValor.Add gobjNFiscal.dValorProdutos
'
'        Case PRODUTO_ALMOX
'            For Each objItem In gobjNFiscal.ColItensNF
'
'                For Each objAlocacao In objItem.colAlocacoes
'
'                    lErro = Mascara_RetornaProdutoTela(objItem.sProduto, sProdutoTela)
'                    If lErro <> SUCESSO Then Error 59409
'
'                    objMnemonicoValor.colValor.Add sProdutoTela
'
'                Next
'
'            Next
'
'        Case QUANT_ALOCADA
'
'            For Each objItem In gobjNFiscal.ColItensNF
'
'                If objMnemonicoValor.vParam(1) = objMnemonicoValor.vParam(2) Then
'
'                    For Each objAlocacao In objItem.colAlocacoes
'
'                        objMnemonicoValor.colValor.Add objAlocacao.dQuantidade
'
'                    Next
'
'                Else
'
'                    For Each objAlocacao In objItem.colAlocacoes
'
'                        objMnemonicoValor.colValor.Add 0
'
'                    Next
'
'                End If
'
'            Next
'
'        Case UNIDADE_MED_EST
'            For Each objItem In gobjNFiscal.ColItensNF
'
'                For Each objAlocacao In objItem.colAlocacoes
'
'                    If objMnemonicoValor.vParam(1) = objMnemonicoValor.vParam(2) Then
'                        objMnemonicoValor.colValor.Add objItem.sUMEstoque
'                    Else
'                        objMnemonicoValor.colValor.Add ""
'                    End If
'
'                Next
'
'            Next
'
'        Case Else
'            Error 59389
'
'    End Select
'
'    Calcula_Mnemonico = SUCESSO
'
'    Exit Function
'
'Erro_Calcula_Mnemonico:
'
'    Calcula_Mnemonico = Err
'
'    Select Case Err
'
'        Case 59389
'            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
'
'        Case 59388, 59390, 59391, 59392, 59394, 59397, 59409, 59410, 59423
'
'        Case 59393
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", Err, objFilialCliente.iCodFilial, objFilialCliente.lCodCliente)
'
'        Case 59395
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, gobjNFiscal.lCliente)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160893)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long
''esta funcao é chamada a cada atualizacao de nota fiscal e é responsavel por gerar a contabilizacao correspondente
'
'Dim lErro As Long, lDoc As Long, objItem As ClassItemNF, iNumAlocacoes As Integer
'
'On Error GoTo Erro_GeraContabilizacao
'
'    Set gobjContabAutomatica = objContabAutomatica
'    Set gobjNFiscal = vParams(0)
'    Set gobjPedidoVenda = vParams(1)
'
'    'percorre itens otendo qtde de alocacoes da nf como um todo
'    For Each objItem In gobjNFiscal.ColItensNF
'        iNumAlocacoes = iNumAlocacoes + objItem.colAlocacoes.Count
'    Next
'
'    GridAlocacao.Tag = iNumAlocacoes
'    GridItens.Tag = gobjNFiscal.ColItensNF.Count
'
'    'obter numero de Doc
'    lErro = CF("Voucher_Automatico_Trans", gobjNFiscal.iFilialEmpresa, giExercicio, giPeriodo, MODULO_FATURAMENTO, lDoc)
'    If lErro <> SUCESSO Then Error 59398
'
'    'grava a contabilizacao
'    lErro = objContabAutomatica.Gravar_Registro(Me, IIf(gobjNFiscal.iTipoNFiscal = DOCINFO_NFISFVPV, "NFiscalFaturaPedido", "NFiscalPedido"), gobjNFiscal.lNumIntDoc, gobjNFiscal.lCliente, gobjNFiscal.iFilialCli, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjNFiscal.iFilialEmpresa, gobjGeracaoNFiscal.iLoteContabil, gobjNFiscal.lNumNotaFiscal)
'    If lErro <> SUCESSO Then Error 59399
'
'    GeraContabilizacao = SUCESSO
'
'    Exit Function
'
'Erro_GeraContabilizacao:
'
'    GeraContabilizacao = Err
'
'    Select Case Err
'
'        Case 59398, 59399
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160894)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Sub CTBLabelLote_Click()
'
'    Call objContabil.Contabil_LabelLote_Click
'
'End Sub
'
'Public Sub CTBLote_Change()
'
'    Call objContabil.Contabil_Lote_Change
'
'End Sub
'
'Public Sub CTBLote_GotFocus()
'
'    Call objContabil.Contabil_Lote_GotFocus
'
'End Sub
'
'Public Sub CTBLote_Validate(Cancel As Boolean)
'
'    Call objContabil.Contabil_Lote_Validate(Cancel, Parent)
'
'End Sub
'
'Public Sub CTBDataContabil_Change()
'
'    Call objContabil.Contabil_DataContabil_Change
'
'End Sub
'
'Public Sub CTBDataContabil_GotFocus()
'
'    Call objContabil.Contabil_DataContabil_GotFocus
'
'End Sub
'
'Public Sub CTBDataContabil_Validate(Cancel As Boolean)
'
'    Call objContabil.Contabil_DataContabil_Validate(Cancel, Parent)
'
'End Sub
'
'Private Sub CTBLancAutomatico_Click()
'
'    Call objContabil.Contabil_LancAutomatico_Click
'
'End Sub
'
'Private Sub CTBAglutina_Click()
'
'    Call objContabil.Contabil_Aglutina_Click
'
'End Sub
'
'Private Sub objEventoLote_evSelecao(obj1 As Object)
''Traz o lote selecionado para a tela
'
'    Call objContabil.Contabil_objEventoLote_evSelecao(obj1)
'
'End Sub
'
'Private Function GeracaoNF_Prepara_CTB() As Long
''prepara informacoes necessarias para a contabilizacao
'
'Dim lErro As Long, objPeriodo As New ClassPeriodo
'Dim vbMsgRes As VbMsgBoxResult
'
'On Error GoTo Erro_GeracaoNF_Prepara_CTB
'
'    If (gcolModulo.ATIVO(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
'
'        'Verifica se a data contabil está preenchida
'        If Len(CTBDataContabil.ClipText) = 0 Then gError 59400
'
'        gobjGeracaoNFiscal.objTelaAtualizacao = Me
'        gobjGeracaoNFiscal.dtContabil = CDate(CTBDataContabil.Text)
'
'        If gobjGeracaoNFiscal.dtContabil <> gdtDataAtual Then
'
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DATA_CONTABIL_DIFERE_TRANSACAO", gobjGeracaoNFiscal.dtContabil, gdtDataAtual)
'
'            If vbMsgRes = vbNo Then gError 92041
'
'        End If
'
'        'se o lote estiver preenchido ==> não pode ser com o valor zero
'        If Len(CTBLote.ClipText) > 0 And giTipoVersao = VERSAO_FULL Then
'            gobjGeracaoNFiscal.iLoteContabil = CInt(CTBLote.ClipText)
'        Else
'            'se não estiver preenchido o lote ==> atualizacao imediata e o valor do lote será zero internamente
'            gobjGeracaoNFiscal.iLoteContabil = 0
'        End If
'
'        'Coloca o periodo relativo a data na tela
'        lErro = CF("Periodo_Le", gobjGeracaoNFiscal.dtContabil, objPeriodo)
'        If lErro <> SUCESSO Then gError 59401
'
'        giPeriodo = objPeriodo.iPeriodo
'        giExercicio = objPeriodo.iExercicio
'
'    End If
'
'    GeracaoNF_Prepara_CTB = SUCESSO
'
'    Exit Function
'
'Erro_GeracaoNF_Prepara_CTB:
'
'    GeracaoNF_Prepara_CTB = gErr
'
'    Select Case gErr
'
'        Case 59400
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_CONTABIL_NAO_PREENCHIDA", gErr)
'
'        Case 59401, 92041
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160895)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Sub LabelDataContabil_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelDataContabil, Button, Shift, X, Y)
'End Sub
'
'Private Sub LabelDataContabil_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(LabelDataContabil, Source, X, Y)
'End Sub
'
'Private Sub CTBLabelLote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(CTBLabelLote, Button, Shift, X, Y)
'End Sub
'
'Private Sub CTBLabelLote_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(CTBLabelLote, Source, X, Y)
'End Sub
'
'
'
'Private Sub TabStrip1_BeforeClick(Cancel As Integer)
'    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
'End Sub
'
'
'
'Private Sub CTBCclDescricao_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(CTBCclDescricao, Source, X, Y)
'End Sub
'
'Private Sub CTBCclDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(CTBCclDescricao, Button, Shift, X, Y)
'End Sub
'
'Private Sub CTBContaDescricao_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(CTBContaDescricao, Source, X, Y)
'End Sub
'
'Private Sub CTBContaDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(CTBContaDescricao, Button, Shift, X, Y)
'End Sub
'
'Private Sub CTBLabel7_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(CTBLabel7, Source, X, Y)
'End Sub
'
'Private Sub CTBLabel7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(CTBLabel7, Button, Shift, X, Y)
'End Sub
'
'Private Sub CTBCclLabel_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(CTBCclLabel, Source, X, Y)
'End Sub
'
'Private Sub CTBCclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(CTBCclLabel, Button, Shift, X, Y)
'End Sub
'
'Private Sub CTBLabelLote3_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(CTBLabelLote3, Source, X, Y)
'End Sub
'
'Private Sub CTBLabelLote3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(CTBLabelLote3, Button, Shift, X, Y)
'End Sub
'
'Private Sub CTBLabelDoc_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(CTBLabelDoc, Source, X, Y)
'End Sub
'
'Private Sub CTBLabelDoc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(CTBLabelDoc, Button, Shift, X, Y)
'End Sub
'
'Private Sub CTBLabel8_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(CTBLabel8, Source, X, Y)
'End Sub
'
'Private Sub CTBLabel8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(CTBLabel8, Button, Shift, X, Y)
'End Sub
'
'Private Sub CTBTotalCredito_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(CTBTotalCredito, Source, X, Y)
'End Sub
'
'Private Sub CTBTotalCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(CTBTotalCredito, Button, Shift, X, Y)
'End Sub
'
'Private Sub CTBTotalDebito_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(CTBTotalDebito, Source, X, Y)
'End Sub
'
'Private Sub CTBTotalDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(CTBTotalDebito, Button, Shift, X, Y)
'End Sub
'
'Private Sub CTBLabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(CTBLabelTotais, Source, X, Y)
'End Sub
'
'Private Sub CTBLabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(CTBLabelTotais, Button, Shift, X, Y)
'End Sub
'
'Private Sub CTBLabel1_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(CTBLabel1, Source, X, Y)
'End Sub
'
'Private Sub CTBLabel1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(CTBLabel1, Button, Shift, X, Y)
'End Sub
'
'Private Sub CTBLabelCcl_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(CTBLabelCcl, Source, X, Y)
'End Sub
'
'Private Sub CTBLabelCcl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(CTBLabelCcl, Button, Shift, X, Y)
'End Sub
'
'Private Sub CTBLabelContas_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(CTBLabelContas, Source, X, Y)
'End Sub
'
'Private Sub CTBLabelContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(CTBLabelContas, Button, Shift, X, Y)
'End Sub
'
'Private Sub CTBLabelHistoricos_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(CTBLabelHistoricos, Source, X, Y)
'End Sub
'
'Private Sub CTBLabelHistoricos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(CTBLabelHistoricos, Button, Shift, X, Y)
'End Sub
'
'Private Sub CTBLabel5_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(CTBLabel5, Source, X, Y)
'End Sub
'
'Private Sub CTBLabel5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(CTBLabel5, Button, Shift, X, Y)
'End Sub
'
'Private Sub CTBLabel13_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(CTBLabel13, Source, X, Y)
'End Sub
'
'Private Sub CTBLabel13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(CTBLabel13, Button, Shift, X, Y)
'End Sub
'
'Private Sub CTBExercicio_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(CTBExercicio, Source, X, Y)
'End Sub
'
'Private Sub CTBExercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(CTBExercicio, Button, Shift, X, Y)
'End Sub
'
'Private Sub CTBPeriodo_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(CTBPeriodo, Source, X, Y)
'End Sub
'
'Private Sub CTBPeriodo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(CTBPeriodo, Button, Shift, X, Y)
'End Sub
'
'Private Sub CTBLabel14_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(CTBLabel14, Source, X, Y)
'End Sub
'
'Private Sub CTBLabel14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(CTBLabel14, Button, Shift, X, Y)
'End Sub
'
'Private Sub CTBOrigem_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(CTBOrigem, Source, X, Y)
'End Sub
'
'Private Sub CTBOrigem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(CTBOrigem, Button, Shift, X, Y)
'End Sub
'
'Private Sub CTBLabel21_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(CTBLabel21, Source, X, Y)
'End Sub
'
'Private Sub CTBLabel21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(CTBLabel21, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label2, Source, X, Y)
'End Sub
'
'Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
'End Sub
'

Event Unload()

Private WithEvents objCT As CTGeracaoNFiscal
Attribute objCT.VB_VarHelpID = -1

Private Sub BotaoDesmarcar_Click()
    Call objCT.BotaoDesmarcar_Click
End Sub

Private Sub BotaoImprimirPI_Click()
     Call objCT.BotaoImprimirPI_Click
End Sub

Private Sub BotaoMarcar_Click()
    Call objCT.BotaoMarcar_Click
End Sub

Private Sub LabelVendedorAte_Click()
    Call objCT.LabelVendedorAte_Click
End Sub

Private Sub LabelVendedorDe_Click()
    Call objCT.LabelVendedorDe_Click
End Sub

Private Sub ListRegioes_Click()
    Call objCT.ListRegioes_Click
End Sub

Private Sub UserControl_Initialize()
    Set objCT = New CTGeracaoNFiscal
    Set objCT.objUserControl = Me
End Sub

Private Sub BotaoDesmarcarTodos_Click()
     Call objCT.BotaoDesmarcarTodos_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub BotaoMarcarTodos_Click()
     Call objCT.BotaoMarcarTodos_Click
End Sub

Private Sub BotaoNFiscal_Click()
     Call objCT.BotaoNFiscal_Click
End Sub

Private Sub BotaoNFiscalFatura_Click()
     Call objCT.BotaoNFiscalFatura_Click
End Sub

Private Sub BotaoNFiscalFaturaImprime_Click()
     Call objCT.BotaoNFiscalFaturaImprime_Click
End Sub

Private Sub BotaoNFiscalImprime_Click()
     Call objCT.BotaoNFiscalImprime_Click
End Sub

Private Sub BotaoPedido_Click()
     Call objCT.BotaoPedido_Click
End Sub

Private Sub ClienteAte_Change()
     Call objCT.ClienteAte_Change
End Sub

Private Sub ClienteAte_GotFocus()
     Call objCT.ClienteAte_GotFocus
End Sub

Private Sub ClienteAte_Validate(Cancel As Boolean)
     Call objCT.ClienteAte_Validate(Cancel)
End Sub

Private Sub ClienteDe_Change()
     Call objCT.ClienteDe_Change
End Sub

Private Sub ClienteDe_GotFocus()
     Call objCT.ClienteDe_GotFocus
End Sub

Private Sub ClienteDe_Validate(Cancel As Boolean)
     Call objCT.ClienteDe_Validate(Cancel)
End Sub

Private Sub CTBUpDown_DownClick()
     Call objCT.CTBUpDown_DownClick
End Sub

Private Sub CTBUpDown_UpClick()
     Call objCT.CTBUpDown_UpClick
End Sub

Private Sub DataEmissaoAte_Change()
     Call objCT.DataEmissaoAte_Change
End Sub

Private Sub DataEmissaoAte_GotFocus()
     Call objCT.DataEmissaoAte_GotFocus
End Sub

Private Sub DataEmissaoAte_Validate(Cancel As Boolean)
     Call objCT.DataEmissaoAte_Validate(Cancel)
End Sub

Private Sub DataEmissaoDe_Change()
     Call objCT.DataEmissaoDe_Change
End Sub

Private Sub DataEmissaoDe_GotFocus()
     Call objCT.DataEmissaoDe_GotFocus
End Sub

Private Sub DataEmissaoDe_Validate(Cancel As Boolean)
     Call objCT.DataEmissaoDe_Validate(Cancel)
End Sub

Private Sub DataEntregaAte_Change()
     Call objCT.DataEntregaAte_Change
End Sub

Private Sub DataEntregaAte_GotFocus()
     Call objCT.DataEntregaAte_GotFocus
End Sub

Private Sub DataEntregaAte_Validate(Cancel As Boolean)
     Call objCT.DataEntregaAte_Validate(Cancel)
End Sub

Private Sub DataEntregaDe_Change()
     Call objCT.DataEntregaDe_Change
End Sub

Private Sub DataEntregaDe_GotFocus()
     Call objCT.DataEntregaDe_GotFocus
End Sub

Private Sub DataEntregaDe_Validate(Cancel As Boolean)
     Call objCT.DataEntregaDe_Validate(Cancel)
End Sub

Private Sub DataSaida_Change()
     Call objCT.DataSaida_Change
End Sub

Private Sub DataSaida_GotFocus()
     Call objCT.DataSaida_GotFocus
End Sub

Private Sub ExibeTodos_Click()
     Call objCT.ExibeTodos_Click
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub GeraNFiscal_Click()
     Call objCT.GeraNFiscal_Click
End Sub

Private Sub GeraNFiscal_GotFocus()
     Call objCT.GeraNFiscal_GotFocus
End Sub

Private Sub GeraNFiscal_KeyPress(KeyAscii As Integer)
     Call objCT.GeraNFiscal_KeyPress(KeyAscii)
End Sub

Private Sub GeraNFiscal_Validate(Cancel As Boolean)
     Call objCT.GeraNFiscal_Validate(Cancel)
End Sub

Private Sub GridPedido_Click()
     Call objCT.GridPedido_Click
End Sub

Private Sub GridPedido_EnterCell()
     Call objCT.GridPedido_EnterCell
End Sub

Private Sub GridPedido_GotFocus()
     Call objCT.GridPedido_GotFocus
End Sub

Private Sub GridPedido_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridPedido_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridPedido_KeyPress(KeyAscii As Integer)
     Call objCT.GridPedido_KeyPress(KeyAscii)
End Sub

Private Sub GridPedido_LeaveCell()
     Call objCT.GridPedido_LeaveCell
End Sub

Private Sub GridPedido_Validate(Cancel As Boolean)
     Call objCT.GridPedido_Validate(Cancel)
End Sub

Private Sub GridPedido_RowColChange()
     Call objCT.GridPedido_RowColChange
End Sub

Private Sub GridPedido_Scroll()
     Call objCT.GridPedido_Scroll
End Sub

Private Sub LabelClienteAte_Click()
     Call objCT.LabelClienteAte_Click
End Sub

Private Sub LabelClienteDe_Click()
     Call objCT.LabelClienteDe_Click
End Sub

Private Sub LabelPedidoAte_Click()
     Call objCT.LabelPedidoAte_Click
End Sub

Private Sub LabelPedidoDe_Click()
     Call objCT.LabelPedidoDe_Click
End Sub

Private Sub Ordenados_Change()
     Call objCT.Ordenados_Change
End Sub

Private Sub Ordenados_Click()
     Call objCT.Ordenados_Click
End Sub

Private Sub PedidoFinal_Change()
     Call objCT.PedidoFinal_Change
End Sub

Private Sub PedidoFinal_GotFocus()
     Call objCT.PedidoFinal_GotFocus
End Sub

Private Sub PedidoFinal_Validate(Cancel As Boolean)
     Call objCT.PedidoFinal_Validate(Cancel)
End Sub

Private Sub PedidoInicial_Change()
     Call objCT.PedidoInicial_Change
End Sub

Private Sub PedidoInicial_GotFocus()
     Call objCT.PedidoInicial_GotFocus
End Sub

Private Sub PedidoInicial_Validate(Cancel As Boolean)
     Call objCT.PedidoInicial_Validate(Cancel)
End Sub

Private Sub Serie_Validate(Cancel As Boolean)
     Call objCT.Serie_Validate(Cancel)
End Sub

Private Sub TabStrip1_Click()
     Call objCT.TabStrip1_Click
End Sub

Private Sub UpDownEmissaoAte_DownClick()
     Call objCT.UpDownEmissaoAte_DownClick
End Sub

Private Sub UpDownEmissaoAte_UpClick()
     Call objCT.UpDownEmissaoAte_UpClick
End Sub

Private Sub UpDownEmissaoDe_DownClick()
     Call objCT.UpDownEmissaoDe_DownClick
End Sub

Private Sub UpDownEmissaoDe_UpClick()
     Call objCT.UpDownEmissaoDe_UpClick
End Sub

Private Sub UpDownEntregaAte_DownClick()
     Call objCT.UpDownEntregaAte_DownClick
End Sub

Private Sub UpDownEntregaAte_UpClick()
     Call objCT.UpDownEntregaAte_UpClick
End Sub

Private Sub UpDownEntregaDe_DownClick()
     Call objCT.UpDownEntregaDe_DownClick
End Sub

Private Sub UpDownEntregaDe_UpClick()
     Call objCT.UpDownEntregaDe_UpClick
End Sub

Private Sub DataSaida_Validate(Cancel As Boolean)
     Call objCT.DataSaida_Validate(Cancel)
End Sub

Private Sub UpDownSaida_DownClick()
     Call objCT.UpDownSaida_DownClick
End Sub

Private Sub UpDownSaida_UpClick()
     Call objCT.UpDownSaida_UpClick
End Sub

Function Trata_Parametros() As Long
     Trata_Parametros = objCT.Trata_Parametros()
End Function

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub
Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub
Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub
Private Sub label_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(label, Source, X, Y)
End Sub
Private Sub label_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(label, Button, Shift, X, Y)
End Sub
Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub
Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub
Private Sub LabelPedidoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPedidoDe, Source, X, Y)
End Sub
Private Sub LabelPedidoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPedidoDe, Button, Shift, X, Y)
End Sub
Private Sub LabelPedidoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPedidoAte, Source, X, Y)
End Sub
Private Sub LabelPedidoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPedidoAte, Button, Shift, X, Y)
End Sub
Private Sub LabelClienteAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteAte, Source, X, Y)
End Sub
Private Sub LabelClienteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteAte, Button, Shift, X, Y)
End Sub
Private Sub LabelClienteDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteDe, Source, X, Y)
End Sub
Private Sub LabelClienteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteDe, Button, Shift, X, Y)
End Sub
Private Sub CTBLancAutomatico_Click()
     Call objCT.CTBLancAutomatico_Click
End Sub

Private Sub CTBAglutina_Click()
     Call objCT.CTBAglutina_Click
End Sub

Private Sub LabelDataContabil_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataContabil, Button, Shift, X, Y)
End Sub
Private Sub LabelDataContabil_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataContabil, Source, X, Y)
End Sub
Private Sub CTBLabelLote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelLote, Button, Shift, X, Y)
End Sub
Private Sub CTBLabelLote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelLote, Source, X, Y)
End Sub
Private Sub TabStrip1_BeforeClick(Cancel As Integer)
     Call objCT.TabStrip1_BeforeClick(Cancel)
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
Private Sub CTBLabelLote3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelLote3, Source, X, Y)
End Sub
Private Sub CTBLabelLote3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelLote3, Button, Shift, X, Y)
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
'Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label2, Source, X, Y)
'End Sub
'Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
'End Sub
Public Function Form_Load_Ocx() As Object

    Call objCT.Form_Load_Ocx
    Set Form_Load_Ocx = Me

End Function

Public Sub Form_Unload(Cancel As Integer)
    If Not (objCT Is Nothing) Then
        Call objCT.Form_Unload(Cancel)
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

Private Sub VendedorFinal_Change()
    Call objCT.VendedorFinal_Change
End Sub

Private Sub VendedorFinal_Validate(Cancel As Boolean)
    Call objCT.VendedorFinal_Validate(Cancel)
End Sub

Private Sub VendedorInicial_Change()
    Call objCT.VendedorInicial_Change
End Sub

Private Sub VendedorInicial_Validate(Cancel As Boolean)
    Call objCT.VendedorInicial_Validate(Cancel)
End Sub

Private Sub BotaoExportar_Click()
    Call objCT.BotaoExportar_Click
End Sub
