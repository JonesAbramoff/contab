VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ConfiguraCOMOcx 
   ClientHeight    =   5310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9480
   ScaleHeight     =   5310
   ScaleWidth      =   9480
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4230
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   795
      Width           =   9060
      Begin VB.Frame Frame10 
         Caption         =   "Ponto de Pedido"
         Height          =   1905
         Left            =   180
         TabIndex        =   13
         Top             =   2010
         Width           =   8760
         Begin VB.Frame Frame8 
            Caption         =   "Estoque de Segurança"
            Height          =   1320
            Left            =   4695
            TabIndex        =   20
            Top             =   300
            Width           =   3870
            Begin MSMask.MaskEdBox ConsumoMedioMax 
               Height          =   315
               Left            =   2970
               TabIndex        =   22
               Top             =   375
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   3
               Format          =   "#0.#0\%"
               Mask            =   "###"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox TempoRessupMax 
               Height          =   315
               Left            =   2985
               TabIndex        =   24
               Top             =   855
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   3
               Format          =   "#0.#0\%"
               Mask            =   "###"
               PromptChar      =   " "
            End
            Begin VB.Label Label13 
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
               TabIndex        =   21
               Top             =   435
               Width           =   2490
            End
            Begin VB.Label Label14 
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
               Left            =   225
               TabIndex        =   23
               Top             =   885
               Width           =   2685
            End
         End
         Begin MSMask.MaskEdBox MesesMediaTempoRessup 
            Height          =   315
            Left            =   2955
            TabIndex        =   18
            Top             =   1035
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
            Left            =   2955
            TabIndex        =   15
            Top             =   660
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
         Begin VB.Label Label9 
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
            Left            =   300
            TabIndex        =   17
            Top             =   1095
            Width           =   2520
         End
         Begin VB.Label Label10 
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
            Left            =   225
            TabIndex        =   14
            Top             =   705
            Width           =   2595
         End
         Begin VB.Label Label11 
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
            Left            =   3525
            TabIndex        =   16
            Top             =   705
            Width           =   540
         End
         Begin VB.Label Label12 
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
            Left            =   3525
            TabIndex        =   19
            Top             =   1095
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Filial de Compras Default"
         Height          =   690
         Left            =   150
         TabIndex        =   2
         Top             =   15
         Width           =   8760
         Begin VB.ComboBox FilialCompra 
            Height          =   315
            Left            =   3495
            TabIndex        =   4
            Top             =   255
            Width           =   2700
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Filial Compra:"
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
            Left            =   2295
            TabIndex        =   3
            Top             =   315
            Width           =   1155
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Alçadas / Resíduo / Atraso / Ressuprimento"
         Height          =   1215
         Left            =   165
         TabIndex        =   5
         Top             =   780
         Width           =   8760
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
            Left            =   930
            TabIndex        =   6
            Top             =   360
            Width           =   2025
         End
         Begin MSMask.MaskEdBox Residuo 
            Height          =   315
            Left            =   2100
            TabIndex        =   10
            Top             =   765
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
            Left            =   7140
            TabIndex        =   8
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
            TabIndex        =   12
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
            Left            =   885
            TabIndex        =   9
            Top             =   825
            Width           =   1110
         End
         Begin VB.Label Label2 
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
            TabIndex        =   7
            Top             =   390
            Width           =   2940
         End
         Begin VB.Label Label15 
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
            TabIndex        =   11
            Top             =   825
            Width           =   2820
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4305
      Index           =   2
      Left            =   150
      TabIndex        =   25
      Top             =   825
      Visible         =   0   'False
      Width           =   9165
      Begin VB.Frame Frame9 
         Caption         =   "Cotações / Concorrências"
         Height          =   2340
         Left            =   135
         TabIndex        =   26
         Top             =   0
         Width           =   8955
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
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   690
            Width           =   4725
         End
         Begin VB.Frame Frame2 
            Caption         =   "Cotações Anteriores"
            Height          =   1965
            Left            =   5010
            TabIndex        =   30
            Top             =   210
            Width           =   3825
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
               TabIndex        =   31
               Top             =   270
               Width           =   3165
            End
            Begin VB.Frame Frame4 
               Caption         =   "Limites percentuais de quantidade para uso"
               Height          =   1245
               Left            =   210
               TabIndex        =   32
               Top             =   570
               Width           =   3435
               Begin MSMask.MaskEdBox PercentMaisQuantCotacaoAnterior 
                  Height          =   315
                  Left            =   2235
                  TabIndex        =   34
                  Top             =   360
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  AllowPrompt     =   -1  'True
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
                  Left            =   2250
                  TabIndex        =   36
                  Top             =   765
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
                  Left            =   360
                  TabIndex        =   33
                  Top             =   420
                  Width           =   1785
               End
               Begin VB.Label Label4 
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
                  Left            =   195
                  TabIndex        =   35
                  Top             =   840
                  Width           =   1950
               End
            End
         End
         Begin MSMask.MaskEdBox TaxaFinanceiraEmpresa 
            Height          =   315
            Left            =   3210
            TabIndex        =   29
            Top             =   1245
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
         Begin VB.Label Label7 
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
            Left            =   240
            TabIndex        =   28
            Top             =   1290
            Width           =   2880
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Recebimento"
         Height          =   1860
         Left            =   120
         TabIndex        =   37
         Top             =   2340
         Width           =   8970
         Begin VB.CheckBox AceitaDiferencaNFPC 
            Caption         =   "Aceita diferença no valor unitário e aliquotas ICMS/IPI entre Notas Fiscais e Pedidos de Compras"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   3960
            TabIndex        =   39
            Top             =   240
            Width           =   4875
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
            TabIndex        =   38
            Top             =   315
            Width           =   3585
         End
         Begin VB.Frame Frame7 
            Caption         =   "Recebimento fora da faixa"
            Height          =   1035
            Left            =   3900
            TabIndex        =   45
            Top             =   720
            Width           =   3540
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
               TabIndex        =   47
               Top             =   660
               Width           =   2655
            End
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
               TabIndex        =   46
               Top             =   300
               Value           =   -1  'True
               Width           =   2415
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Faixa de recebimento"
            Height          =   1125
            Left            =   210
            TabIndex        =   40
            Top             =   630
            Width           =   3525
            Begin MSMask.MaskEdBox PercentMaisReceb 
               Height          =   315
               Left            =   2310
               TabIndex        =   42
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
               TabIndex        =   44
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
               TabIndex        =   41
               Top             =   375
               Width           =   1785
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
               TabIndex        =   43
               Top             =   780
               Width           =   1950
            End
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   8160
      ScaleHeight     =   495
      ScaleWidth      =   1080
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   120
      Width           =   1140
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "ConfiguraCOMOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   585
         Picture         =   "ConfiguraCOMOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4770
      Left            =   135
      TabIndex        =   0
      Top             =   465
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   8414
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Configurações Básicas"
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
Attribute VB_Name = "ConfiguraCOMOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iFrameAtual As Integer
Public iAlterado As Integer

Function Traz_ConfiguraCOM_Tela(objConfiguraCOM As ClassConfiguraCOM) As Long
'Traz os dados de objConfiguraCOM para tela

Dim lErro As Long
Dim bCancel As Boolean

On Error GoTo Erro_Traz_ConfiguraCOM_Tela

    'Filial Compra
    If giTipoVersao = VERSAO_FULL Then
        If objConfiguraCOM.iFilialCompra > 0 Then
            FilialCompra.Text = objConfiguraCOM.iFilialCompra
            bCancel = False
            FilialCompra_Validate (bCancel)

        End If
    End If

    'Frame sem título
    CompradorAumentaQuant.Value = objConfiguraCOM.iCompradorAumentaQuant
    ControleAlcada.Value = objConfiguraCOM.iControleAlcada
    Residuo.Text = (objConfiguraCOM.dResiduo) * 100
    TaxaFinanceiraEmpresa.Text = objConfiguraCOM.dTaxaFinanceiraEmpresa * 100
    MesesConsumoMedio.Text = objConfiguraCOM.iMesesConsumoMedio
    MesesMediaTempoRessup.Text = objConfiguraCOM.iMesesMediaTempoRessup
    NumComprasMediaAtraso.Text = Format(objConfiguraCOM.iNumComprasMediaAtraso)
    NumComprasTempoRessup.Text = Format(objConfiguraCOM.iNumComprasTempoRessup)
    AceitaDiferencaNFPC.Value = objConfiguraCOM.iNFDiferentePC
    
    
    'Estoque de Segurança
    ConsumoMedioMax.Text = (objConfiguraCOM.dConsumoMedioMax) * 100
    TempoRessupMax.Text = objConfiguraCOM.dTempoRessupMax * 100
    
    'Cotações Anteriores
    If objConfiguraCOM.iConsideraQuantCotacaoAnterior = 0 Then
        NaoConsideraQuantCotacaoAnterior.Value = vbUnchecked
    Else
        NaoConsideraQuantCotacaoAnterior.Value = vbChecked
    End If
    
    PercentMenosQuantCotacaoAnterior.Text = (objConfiguraCOM.dPercentMenosQuantCotacaoAnterior) * 100
    PercentMaisQuantCotacaoAnterior.Text = (objConfiguraCOM.dPercentMaisQuantCotacaoAnterior) * 100
    
    'Faixa de Recebimento
    If objConfiguraCOM.iTemFaixaReceb = 0 Then
        NaoTemFaixaReceb.Value = vbUnchecked
    Else
        NaoTemFaixaReceb.Value = vbChecked
    End If
    
    PercentMenosReceb.Text = (objConfiguraCOM.dPercentMenosReceb) * 100
    PercentMaisReceb.Text = (objConfiguraCOM.dPercentMaisReceb) * 100

    RecebForaFaixa(objConfiguraCOM.iRecebForaFaixa).Value = True

    Exit Function

Erro_Traz_ConfiguraCOM_Tela:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154676)

    End Select

    Exit Function

End Function
Function Move_Tela_Memoria(objConfiguraCOM As ClassConfiguraCOM) As Long
'Move os dados da tela para objConfiguraCOM

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'FilialCompra
    If Len(Trim(FilialCompra.Text)) <> 0 Then
        objConfiguraCOM.iFilialCompra = Codigo_Extrai(FilialCompra.Text)
    Else
        objConfiguraCOM.iFilialCompra = 0
    End If
    
    'Frame sem titulo
    objConfiguraCOM.iCompradorAumentaQuant = CompradorAumentaQuant.Value
    objConfiguraCOM.iControleAlcada = ControleAlcada.Value
    objConfiguraCOM.iNFDiferentePC = AceitaDiferencaNFPC.Value
    
    If Len(Trim(Residuo.Text)) <> 0 Then
        objConfiguraCOM.dResiduo = StrParaDbl(Residuo.Text) / 100
    Else
        objConfiguraCOM.dResiduo = 0
    End If
    
    If Len(Trim(NumComprasMediaAtraso.Text)) <> 0 Then
        objConfiguraCOM.iNumComprasMediaAtraso = StrParaInt(NumComprasMediaAtraso.Text)
    Else
        objConfiguraCOM.iNumComprasMediaAtraso = 0
    End If
    
    If Len(Trim(NumComprasTempoRessup.Text)) <> 0 Then
        objConfiguraCOM.iNumComprasTempoRessup = StrParaInt(NumComprasTempoRessup.Text)
    Else
        objConfiguraCOM.iNumComprasTempoRessup = 0
    End If
    
    If Len(Trim(TaxaFinanceiraEmpresa.Text)) <> 0 Then
        objConfiguraCOM.dTaxaFinanceiraEmpresa = StrParaDbl(TaxaFinanceiraEmpresa.Text) / 100
    Else
        objConfiguraCOM.dTaxaFinanceiraEmpresa = 0
    End If
    objConfiguraCOM.iMesesConsumoMedio = StrParaInt(MesesConsumoMedio.Text)
    objConfiguraCOM.iMesesMediaTempoRessup = StrParaInt(MesesMediaTempoRessup.Text)

    'Estoque de Seguranca
    If Len(Trim(ConsumoMedioMax.Text)) <> 0 Then
        objConfiguraCOM.dConsumoMedioMax = StrParaDbl(ConsumoMedioMax.Text) / 100
    Else
        objConfiguraCOM.dConsumoMedioMax = 0
    End If
    If Len(Trim(TempoRessupMax.Text)) <> 0 Then
        objConfiguraCOM.dTempoRessupMax = StrParaDbl(TempoRessupMax.Text) / 100
    Else
        objConfiguraCOM.dTempoRessupMax = 0
    End If
    
    'Cotacoes Anteriores
    objConfiguraCOM.iConsideraQuantCotacaoAnterior = NaoConsideraQuantCotacaoAnterior.Value

    If Len(Trim(PercentMaisQuantCotacaoAnterior.Text)) <> 0 Then
        objConfiguraCOM.dPercentMaisQuantCotacaoAnterior = StrParaDbl(PercentMaisQuantCotacaoAnterior.Text) / 100
    Else
        objConfiguraCOM.dPercentMaisQuantCotacaoAnterior = 0
    End If

    If Len(Trim(PercentMenosQuantCotacaoAnterior.Text)) <> 0 Then
        objConfiguraCOM.dPercentMenosQuantCotacaoAnterior = StrParaDbl(PercentMenosQuantCotacaoAnterior.Text) / 100
    Else
        objConfiguraCOM.dPercentMenosQuantCotacaoAnterior = 0
    End If

    'Faixa de Recebimento
    If NaoTemFaixaReceb.Value = vbChecked Then
        objConfiguraCOM.iTemFaixaReceb = 1
    ElseIf NaoTemFaixaReceb.Value = vbUnchecked Then
        objConfiguraCOM.iTemFaixaReceb = 0
    End If
    
    If Len(Trim(PercentMenosReceb.Text)) <> 0 Then
        objConfiguraCOM.dPercentMenosReceb = StrParaDbl(PercentMenosReceb.Text) / 100
    Else
        objConfiguraCOM.dPercentMenosReceb = 0
    End If

    If Len(Trim(PercentMaisReceb.Text)) <> 0 Then
        objConfiguraCOM.dPercentMaisReceb = StrParaDbl(PercentMaisReceb.Text) / 100
    Else
        objConfiguraCOM.dPercentMaisReceb = 0
    End If

    If RecebForaFaixa(0).Value Then
        objConfiguraCOM.iRecebForaFaixa = REJEITA_RECEBIMENTO
    End If

    If RecebForaFaixa(1).Value Then
        objConfiguraCOM.iRecebForaFaixa = ACEITA_RECEBIMENTO
    End If
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154677)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objConfiguraCOM As New ClassConfiguraCOM

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se MesesConsumoMedio foi preenchido
    If Len(Trim(MesesConsumoMedio.Text)) = 0 Then gError 49367
    
    'Verifica se MesesMediaTempoRessup foi preenchido
    If Len(Trim(MesesMediaTempoRessup.Text)) = 0 Then gError 49368
    
    'Verifica se Residuo foi preenchido
    If Len(Trim(Residuo.Text)) = 0 Then gError 63314
    
    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objConfiguraCOM)
    If lErro <> SUCESSO Then gError 49344

    lErro = CF("ConfiguraCOM_Gravar", objConfiguraCOM)
    If lErro <> SUCESSO And lErro <> 72496 Then gError 49345
    
    'Marca novamente o Controle de Alcada
    If lErro = 72496 Then gError 72497
    
    Call gobjCOM.Inicializa

    Gravar_Registro = SUCESSO

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 49344, 49345

        Case 49366
            MesesMediaTempoRessup.SetFocus

        Case 49367
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MESESCONSUMOMEDIO_NAO_PREENCHIDO", gErr)

        Case 49368
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MESESMEDIATEMPORESSUP_NAO_PREENCHIDO", gErr)

        Case 63314
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RESIDUO_NAO_PREENCHIDO", gErr)
            
        Case 72497
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BLOQUEIO_ALCADA_EXISTENTE", gErr)
            ControleAlcada.Value = vbChecked
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154678)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

End Function

Private Sub AceitaDiferencaNFPC_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama a funcao Gravar_Registro
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 49325

    Call Rotina_Aviso(vbOKOnly, "AVISO_CONFIGURACAO_GRAVADA")
    
    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 49325

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154679)

    End Select

    Exit Sub

End Sub


Private Sub CompradorAumentaQuant_Click()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConsumoMedioMax_GotFocus()

    Call MaskEdBox_TrataGotFocus(ConsumoMedioMax, iAlterado)
    
End Sub

Private Sub MesesConsumoMedio_GotFocus()

    Call MaskEdBox_TrataGotFocus(MesesConsumoMedio, iAlterado)
    
End Sub

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154680)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub MesesMediaTempoRessup_GotFocus()

    Call MaskEdBox_TrataGotFocus(MesesMediaTempoRessup, iAlterado)
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154681)
            
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154682)

    End Select

    Exit Sub

End Sub

Private Sub ConsumoMedioMax_Change()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ControleAlcada_Click()

      iAlterado = REGISTRO_ALTERADO

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
        
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa, True)
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154683)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim objConfiguraCOM As New ClassConfiguraCOM
Dim colComprasConfig As New colComprasConfig

On Error GoTo Erro_Form_Load

    'Carrega a combo FilialCompra
    lErro = Carrega_FilialCompra()
    If lErro <> SUCESSO Then Error 49330

'    FilialCompra.ListIndex = 0
    
    'Visibilidade para versão LIGHT
    If giTipoVersao = VERSAO_LIGHT Then
        
        FilialCompra.Left = POSICAO_FORA_TELA
        FilialCompra.TabStop = False
        Label1.Left = POSICAO_FORA_TELA
        Label1.Visible = False
        Frame3.Left = POSICAO_FORA_TELA
        Frame3.Visible = False
        Frame11.Top = 80
        Frame10.Top = Frame11.Top + Frame11.Height + 40
        
        
    End If
    
    'Leitura da tabela de ComprasConfig
    lErro = CF("ComprasConfig_Le", objConfiguraCOM)
    If lErro <> SUCESSO Then Error 49360

    'Traz para tela os dados de objConfiguraCOM
    lErro = Traz_ConfiguraCOM_Tela(objConfiguraCOM)
    If lErro <> SUCESSO Then Error 49369

    iFrameAtual = 1
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case Err

        Case 49330, 49360, 49369 'Erros tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154684)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub
Private Function Carrega_FilialCompra() As Long
'Carrega a Combo FilialCompra com Codigo-Nome das filiais empresas

Dim lErro As Long
Dim colFiliais As New Collection
Dim objFiliais As AdmFiliais
Dim lCodEmpresa As Long

On Error GoTo Erro_Carrega_FilialCompra

    'Le todas as filiais e coloca todas as filiais em colFiliais
    lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
    If lErro <> SUCESSO Then Error 49353

    'Carrega a Filial na combo FilialCompra
    For Each objFiliais In colFiliais
        If objFiliais.iCodFilial <> EMPRESA_TODA Then
            FilialCompra.AddItem objFiliais.iCodFilial & SEPARADOR & objFiliais.sNome
        End If
    Next

    Carrega_FilialCompra = SUCESSO

    Exit Function

Erro_Carrega_FilialCompra:

    Carrega_FilialCompra = Err

    Select Case Err

        Case 49353

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154685)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Private Sub MesesConsumoMedio_Change()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MesesMediaTempoRessup_Change()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NaoTemFaixaReceb_Click()

Dim lErro As Long

On Error GoTo Erro_NaoTemFaixaReceb_Click

    'Verifica valor na checkbox
    If NaoTemFaixaReceb.Value = False Then

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
        
    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub

Erro_NaoTemFaixaReceb_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154686)

    End Select

    Exit Sub

End Sub

Private Sub NumComprasMediaAtraso_GotFocus()

    Call MaskEdBox_TrataGotFocus(NumComprasMediaAtraso, iAlterado)
    
End Sub

Private Sub NumComprasTempoRessup_GotFocus()

    Call MaskEdBox_TrataGotFocus(NumComprasTempoRessup, iAlterado)
    
End Sub

Private Sub PercentMaisQuantCotacaoAnterior_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PercentMaisQuantCotacaoAnterior_GotFocus()

    Call MaskEdBox_TrataGotFocus(PercentMaisQuantCotacaoAnterior, iAlterado)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154687)

    End Select

    Exit Sub

End Sub

Private Sub PercentMaisReceb_Change()

      iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PercentMaisReceb_GotFocus()

    Call MaskEdBox_TrataGotFocus(PercentMaisReceb, iAlterado)
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154688)

    End Select

    Exit Sub

End Sub

Private Sub PercentMenosQuantCotacaoAnterior_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PercentMenosQuantCotacaoAnterior_GotFocus()

    Call MaskEdBox_TrataGotFocus(PercentMenosQuantCotacaoAnterior, iAlterado)
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154689)

    End Select

    Exit Sub

End Sub

Private Sub PercentMenosReceb_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PercentMenosReceb_GotFocus()

    Call MaskEdBox_TrataGotFocus(PercentMenosReceb, iAlterado)
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154690)

    End Select

    Exit Sub

End Sub

Private Sub RecebForaFaixa_Click(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Residuo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Residuo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Residuo, iAlterado)
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154691)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se frame atual corresponde ao tab selecionado, sai da rotina
    If TabStrip1.SelectedItem.Index = iFrameAtual Then Exit Sub

    'Torna Frame correspondente ao Tab selecionado visivel
    Frame1(TabStrip1.SelectedItem.Index).Visible = True

    'Torna Frame atual invisivel
    Frame1(iFrameAtual).Visible = False

    'Armazena novo valor de iFrameAtual
    iFrameAtual = TabStrip1.SelectedItem.Index

    Exit Sub

Erro_TabStrip1_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154692)

    End Select

    Exit Sub

End Sub

Private Sub TaxaFinanceiraEmpresa_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TaxaFinanceiraEmpresa_GotFocus()

    Call MaskEdBox_TrataGotFocus(TaxaFinanceiraEmpresa, iAlterado)
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154693)

    End Select

    Exit Sub

End Sub

Private Sub TempoRessupMax_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Configuração do Módulo de Compras"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ConfiguraCOM"
    
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

Private Sub TempoRessupMax_GotFocus()

    Call MaskEdBox_TrataGotFocus(TempoRessupMax, iAlterado)
    
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

'**** fim do trecho a ser copiado *****


Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
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

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
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

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
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

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Public Function Trata_Parametros()
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
End Function

