VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.UserControl ConfiguraFATOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   ScaleHeight     =   6000
   ScaleMode       =   0  'User
   ScaleWidth      =   9510
   Begin VB.CommandButton BotaoConfigOutras 
      Caption         =   "Configurações Avançadas"
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
      Left            =   6450
      TabIndex        =   62
      Top             =   45
      Width           =   1680
   End
   Begin VB.Frame FrameC 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   5025
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   9240
      Begin VB.Frame Frame6 
         Caption         =   "Faturamento"
         Height          =   2160
         Left            =   75
         TabIndex        =   38
         Top             =   2700
         Width           =   8970
         Begin VB.CheckBox TestaDescontoMaxTabPreco 
            Caption         =   "Testar o desconto máximo da tabela de preço"
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
            Left            =   180
            TabIndex        =   61
            Top             =   255
            Width           =   4470
         End
         Begin VB.ComboBox TotTribTipo 
            Height          =   315
            ItemData        =   "ConfiguraFATOcx.ctx":0000
            Left            =   6450
            List            =   "ConfiguraFATOcx.ctx":000D
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   375
            Width           =   1905
         End
         Begin VB.Frame Frame8 
            Caption         =   "Faixa de faturamento"
            Height          =   1125
            Left            =   195
            TabIndex        =   43
            Top             =   900
            Width           =   3525
            Begin MSMask.MaskEdBox PercentMaisReceb 
               Height          =   315
               Left            =   2310
               TabIndex        =   44
               Top             =   300
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
               TabIndex        =   45
               Top             =   720
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
            Begin VB.Label Label5 
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
               TabIndex        =   47
               Top             =   780
               Width           =   1950
            End
            Begin VB.Label Label6 
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
               TabIndex        =   46
               Top             =   375
               Width           =   1785
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Faturamento fora da faixa"
            Height          =   1125
            Left            =   3900
            TabIndex        =   40
            Top             =   870
            Width           =   2100
            Begin VB.OptionButton RecebForaFaixa 
               Caption         =   "Não aceita"
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
               TabIndex        =   42
               Top             =   300
               Value           =   -1  'True
               Width           =   1635
            End
            Begin VB.OptionButton RecebForaFaixa 
               Caption         =   "Avisa e aceita"
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
               TabIndex        =   41
               Top             =   660
               Width           =   1605
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
            Left            =   180
            TabIndex        =   39
            Top             =   585
            Width           =   3585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total de Tributos:"
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
            Index           =   75
            Left            =   6615
            TabIndex        =   59
            Top             =   135
            Width           =   1575
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Registro de Notas Fiscais"
         Height          =   2520
         Left            =   5115
         TabIndex        =   14
         Top             =   60
         Width           =   3915
         Begin VB.Frame Frame4 
            Caption         =   "Alocação de Produtos"
            Height          =   645
            Left            =   195
            TabIndex        =   18
            Top             =   840
            Width           =   3645
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
               Left            =   2175
               TabIndex        =   20
               Top             =   300
               Width           =   1215
            End
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
               Left            =   450
               TabIndex        =   19
               Top             =   300
               Value           =   -1  'True
               Width           =   1425
            End
         End
         Begin VB.CheckBox CheckeditaComissoesNF 
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
            Left            =   195
            TabIndex        =   17
            Top             =   360
            Width           =   2535
         End
         Begin VB.Frame Frame5 
            Caption         =   "Fifo Notas Fiscais"
            Height          =   645
            Left            =   195
            TabIndex        =   15
            Top             =   1680
            Width           =   3645
            Begin VB.CheckBox CheckHabilitaFifoNF 
               Caption         =   "Habilitar Fifo Notas Fiscais"
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
               Left            =   360
               TabIndex        =   16
               Top             =   240
               Width           =   2655
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Registro de Pedidos de Venda"
         Height          =   2520
         Index           =   0
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   4650
         Begin VB.CheckBox CheckBloquearPrecoBaixoPV 
            Caption         =   "Bloquear Preço Baixo"
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
            Left            =   2340
            TabIndex        =   60
            ToolTipText     =   "Gerar bloqueio caso o preço esteja abaixo do mínimo para a tabela de preço ou por critério de análise de margem de contribuição."
            Top             =   2130
            Width           =   2205
         End
         Begin VB.ComboBox FilialFaturamento 
            Height          =   315
            ItemData        =   "ConfiguraFATOcx.ctx":0047
            Left            =   2055
            List            =   "ConfiguraFATOcx.ctx":0049
            TabIndex        =   12
            Text            =   "FilialFaturamento"
            Top             =   1620
            Width           =   2145
         End
         Begin VB.CheckBox CheckeditaComissoesPV 
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
            Left            =   240
            TabIndex        =   11
            Top             =   390
            Width           =   2460
         End
         Begin VB.Frame Frame2 
            Caption         =   "Reserva de Produtos"
            Height          =   645
            Left            =   195
            TabIndex        =   8
            Top             =   795
            Width           =   4320
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
               Left            =   450
               TabIndex        =   10
               Top             =   300
               Value           =   -1  'True
               Width           =   1905
            End
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
               TabIndex        =   9
               Top             =   300
               Width           =   1215
            End
         End
         Begin VB.CheckBox CheckValidaEmbalagemPV 
            Caption         =   "Valida Embalagem"
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
            TabIndex        =   7
            Top             =   2145
            Width           =   2460
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
            Height          =   195
            Index           =   17
            Left            =   300
            TabIndex        =   13
            Top             =   1680
            Width           =   1575
         End
      End
   End
   Begin VB.Frame FrameC 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   5025
      Index           =   2
      Left            =   90
      TabIndex        =   5
      Top             =   825
      Visible         =   0   'False
      Width           =   9240
      Begin VB.Frame Frame10 
         Caption         =   "Nota Fiscal Eletrônica"
         Height          =   2625
         Left            =   4965
         TabIndex        =   52
         Top             =   150
         Width           =   4275
         Begin VB.ComboBox versaoNFE 
            Height          =   315
            ItemData        =   "ConfiguraFATOcx.ctx":004B
            Left            =   900
            List            =   "ConfiguraFATOcx.ctx":004D
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   495
            Width           =   975
         End
         Begin VB.CheckBox GravaNFSE 
            Caption         =   "Gera NFSE ao gravar as Notas Fiscais"
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
            Left            =   135
            TabIndex        =   55
            Top             =   2190
            Width           =   3975
         End
         Begin VB.CheckBox UsaNFSE 
            Caption         =   "Usa Nota Fiscal de Serviços Eletrônica"
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
            Left            =   135
            TabIndex        =   54
            Top             =   1800
            Width           =   3930
         End
         Begin VB.CheckBox UsaNFe 
            Caption         =   "Usa Nota Fiscal Eletrônica"
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
            Left            =   165
            TabIndex        =   27
            Top             =   1050
            Width           =   3495
         End
         Begin VB.CheckBox GravaNFe 
            Caption         =   "Gera NFe ao gravar as Notas Fiscais"
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
            Left            =   150
            TabIndex        =   28
            Top             =   1425
            Width           =   3495
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "versão:"
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
            TabIndex        =   56
            Top             =   540
            Width           =   630
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Taxa Financeira"
         Height          =   720
         Left            =   270
         TabIndex        =   48
         Top             =   4065
         Width           =   4440
         Begin MSMask.MaskEdBox TaxaFinanceira 
            Height          =   300
            Left            =   1185
            TabIndex        =   26
            Top             =   285
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   529
            _Version        =   393216
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin VB.Label Label4 
            Caption         =   "Mensais:"
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
            TabIndex        =   51
            Top             =   315
            Width           =   795
         End
         Begin VB.Label Label10 
            Caption         =   "Diários:"
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
            Left            =   2250
            TabIndex        =   50
            Top             =   315
            Width           =   675
         End
         Begin VB.Label DespesasDiarias 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2970
            TabIndex        =   49
            Top             =   285
            Width           =   900
         End
      End
      Begin VB.Frame FrameOutros 
         Caption         =   "Outros"
         Height          =   4725
         Left            =   135
         TabIndex        =   29
         Top             =   150
         Width           =   4695
         Begin VB.Frame Frame11 
            Caption         =   "Crédito"
            Height          =   705
            Left            =   135
            TabIndex        =   53
            Top             =   3090
            Width           =   4440
            Begin VB.CheckBox VerificaLimCred 
               Caption         =   "Verificar limite de Crédito"
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
               TabIndex        =   32
               Top             =   315
               Width           =   3495
            End
         End
         Begin VB.Frame FrameBloqueios 
            Caption         =   "Bloqueios"
            Height          =   705
            Left            =   120
            TabIndex        =   35
            Top             =   1545
            Width           =   4455
            Begin MSMask.MaskEdBox DiasBloqueio 
               Height          =   315
               Left            =   3840
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   300
               Width           =   510
               _ExtentX        =   900
               _ExtentY        =   556
               _Version        =   393216
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
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Nº de Dias p/ Bloqueio por falta de Pagto:"
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
               Left            =   120
               TabIndex        =   37
               Top             =   330
               Width           =   3630
            End
         End
         Begin VB.Frame FrameContabilizacao 
            Caption         =   "Contabilização"
            Height          =   1200
            Left            =   120
            TabIndex        =   33
            Top             =   255
            Width           =   4455
            Begin VB.ListBox ListaConfigura 
               Height          =   735
               ItemData        =   "ConfiguraFATOcx.ctx":004F
               Left            =   120
               List            =   "ConfiguraFATOcx.ctx":005C
               Style           =   1  'Checkbox
               TabIndex        =   34
               Top             =   285
               Width           =   4215
            End
         End
         Begin VB.Frame FrameComissoes 
            Caption         =   "Comissões"
            Height          =   690
            Left            =   120
            TabIndex        =   30
            Top             =   2310
            Width           =   4455
            Begin VB.CheckBox UsaComissoesRegras 
               Caption         =   "Usar regras p/cálculos de comissões"
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
               TabIndex        =   31
               Top             =   300
               Width           =   3495
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Descontos por Adiantamento de Pagamento"
         Height          =   1980
         Index           =   1
         Left            =   4920
         TabIndex        =   21
         Top             =   2925
         Width           =   4290
         Begin VB.ComboBox TipoDesconto 
            Height          =   315
            Left            =   840
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   420
            Width           =   1890
         End
         Begin MSMask.MaskEdBox Dias 
            Height          =   225
            Left            =   2835
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   435
            Width           =   540
            _ExtentX        =   953
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
            Left            =   3540
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   435
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
            Height          =   1245
            Left            =   45
            TabIndex        =   25
            Top             =   270
            Width           =   4065
            _ExtentX        =   7170
            _ExtentY        =   2196
            _Version        =   393216
            Rows            =   4
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   8325
      ScaleHeight     =   495
      ScaleWidth      =   1035
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   1095
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   555
         Picture         =   "ConfiguraFATOcx.ctx":00EF
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "ConfiguraFATOcx.ctx":026D
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5415
      Left            =   45
      TabIndex        =   3
      Top             =   495
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9551
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Configuração Básica"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Outras"
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
Attribute VB_Name = "ConfiguraFATOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTConfiguraFAT
Attribute objCT.VB_VarHelpID = -1

Private Sub AlocacaoAutoNF_GotFocus()
     Call objCT.AlocacaoAutoNF_GotFocus
End Sub

Private Sub AlocacaoManNF_GotFocus()
     Call objCT.AlocacaoManNF_GotFocus
End Sub

Private Sub BotaoConfigOutras_Click()
     Call objCT.BotaoConfigOutras_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Function Trata_Parametros() As Long
     Trata_Parametros = objCT.Trata_Parametros()
End Function

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub Dias_Change()
     Call objCT.Dias_Change
End Sub

Private Sub Dias_GotFocus()
     Call objCT.Dias_GotFocus
End Sub

Private Sub Dias_KeyPress(KeyAscii As Integer)
     Call objCT.Dias_KeyPress(KeyAscii)
End Sub

Private Sub Dias_Validate(Cancel As Boolean)
     Call objCT.Dias_Validate(Cancel)
End Sub

Private Sub DiasBloqueio_Change()
    Call objCT.DiasBloqueio_Change 'por leo em 25/02/02
End Sub

Private Sub FilialFaturamento_Change()
     Call objCT.FilialFaturamento_Change
End Sub

Private Sub FilialFaturamento_Click()
     Call objCT.FilialFaturamento_Click
End Sub

Private Sub FilialFaturamento_Validate(Cancel As Boolean)
     Call objCT.FilialFaturamento_Validate(Cancel)
End Sub


Private Sub GravaNFe_Click()
    Call objCT.GravaNFe_Click
End Sub


Private Sub ListaConfigura_ItemCheck(Item As Integer)
     Call objCT.ListaConfigura_ItemCheck(Item)
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub PercentualDesc_Change()
     Call objCT.PercentualDesc_Change
End Sub

Private Sub PercentualDesc_GotFocus()
     Call objCT.PercentualDesc_GotFocus
End Sub

Private Sub PercentualDesc_KeyPress(KeyAscii As Integer)
     Call objCT.PercentualDesc_KeyPress(KeyAscii)
End Sub

Private Sub PercentualDesc_Validate(Cancel As Boolean)
     Call objCT.PercentualDesc_Validate(Cancel)
End Sub

Private Sub ReservaAutoPV_GotFocus()
     Call objCT.ReservaAutoPV_GotFocus
End Sub

Private Sub ReservaManPV_GotFocus()
     Call objCT.ReservaManPV_GotFocus
End Sub

Private Sub TipoDesconto_Change()
     Call objCT.TipoDesconto_Change
End Sub

Private Sub TipoDesconto_GotFocus()
     Call objCT.TipoDesconto_GotFocus
End Sub

Private Sub TipoDesconto_KeyPress(KeyAscii As Integer)
     Call objCT.TipoDesconto_KeyPress(KeyAscii)
End Sub

Private Sub TipoDesconto_Validate(Cancel As Boolean)
     Call objCT.TipoDesconto_Validate(Cancel)
End Sub

Private Sub GridDescontos_Click()
     Call objCT.GridDescontos_Click
End Sub

Private Sub GridDescontos_EnterCell()
     Call objCT.GridDescontos_EnterCell
End Sub

Private Sub GridDescontos_GotFocus()
     Call objCT.GridDescontos_GotFocus
End Sub

Private Sub GridDescontos_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridDescontos_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridDescontos_KeyPress(KeyAscii As Integer)
     Call objCT.GridDescontos_KeyPress(KeyAscii)
End Sub

Private Sub GridDescontos_LeaveCell()
     Call objCT.GridDescontos_LeaveCell
End Sub

Private Sub GridDescontos_Validate(Cancel As Boolean)
     Call objCT.GridDescontos_Validate(Cancel)
End Sub

Private Sub GridDescontos_RowColChange()
     Call objCT.GridDescontos_RowColChange
End Sub

Private Sub GridDescontos_Scroll()
     Call objCT.GridDescontos_Scroll
End Sub

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

Private Sub TotTribTipo_Click()
    Call objCT.TotTribTipo_Click
End Sub

Private Sub UsaComissoesRegras_Click() 'incluido por leo em 15/03/02
    Call objCT.UsaComissoesRegras_Click
End Sub

Private Sub UsaNFe_Click()
    objCT.UsaNFe_Click
End Sub

Private Sub UserControl_Initialize()
    Set objCT = New CTConfiguraFAT
    Set objCT.objUserControl = Me
End Sub

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

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Sub NaoTemFaixaReceb_Click()
     Call objCT.NaoTemFaixaReceb_Click
End Sub

Private Sub PercentMaisReceb_Change()
     Call objCT.PercentMaisReceb_Change
End Sub

Private Sub PercentMaisReceb_GotFocus()
     Call objCT.PercentMaisReceb_GotFocus
End Sub

Private Sub PercentMaisReceb_Validate(Cancel As Boolean)
     Call objCT.PercentMaisReceb_Validate(Cancel)
End Sub

Private Sub PercentMenosReceb_Change()
     Call objCT.PercentMenosReceb_Change
End Sub

Private Sub PercentMenosReceb_GotFocus()
     Call objCT.PercentMenosReceb_GotFocus
End Sub

Private Sub PercentMenosReceb_Validate(Cancel As Boolean)
     Call objCT.PercentMenosReceb_Validate(Cancel)
End Sub

Private Sub RecebForaFaixa_Click(Index As Integer)
     Call objCT.RecebForaFaixa_Click(Index)
End Sub

Private Sub TabStrip1_Click()
     Call objCT.TabStrip1_Click
End Sub

Private Sub TaxaFinanceira_Change()
    Call objCT.TaxaFinanceira_Change
End Sub

Private Sub TaxaFinanceira_Validate(Cancel As Boolean)
    Call objCT.TaxaFinanceira_Validate(Cancel)
End Sub

Private Sub VerificaLimCred_Click()
    Call objCT.VerificaLimCred_Click
End Sub

Private Sub UsaNFSE_Click()
    objCT.UsaNFSE_Click
End Sub

Private Sub GravaNFSE_Click()
    Call objCT.GravaNFSE_Click
End Sub


