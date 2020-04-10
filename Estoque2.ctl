VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl Estoque 
   ClientHeight    =   5745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9480
   KeyPreview      =   -1  'True
   ScaleHeight     =   5745
   ScaleWidth      =   9480
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   4770
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   810
      Width           =   9150
      Begin VB.Frame Frame6 
         Caption         =   "Local de Produção"
         Height          =   615
         Left            =   2595
         TabIndex        =   56
         Top             =   3975
         Width           =   3600
         Begin VB.CheckBox ProdNaFilial 
            Caption         =   "Produzido na Filial"
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
            Left            =   735
            TabIndex        =   57
            Top             =   225
            Width           =   2025
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Consumo"
         Height          =   2085
         Left            =   180
         TabIndex        =   35
         Top             =   1815
         Width           =   8610
         Begin VB.CheckBox CMCalculado 
            Caption         =   "Calculado"
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
            Left            =   4560
            TabIndex        =   5
            Top             =   270
            Value           =   1  'Checked
            Width           =   1275
         End
         Begin MSMask.MaskEdBox ConsumoMedio 
            Height          =   300
            Left            =   3420
            TabIndex        =   4
            Top             =   240
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MesesConsumoMedio 
            Height          =   300
            Left            =   3420
            TabIndex        =   7
            Top             =   1575
            Width           =   450
            _ExtentX        =   794
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
         Begin MSMask.MaskEdBox ConsumoMedioMax 
            Height          =   300
            Left            =   3420
            TabIndex        =   6
            Top             =   660
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Format          =   "0\%"
            Mask            =   "####"
            PromptChar      =   " "
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   735
            TabIndex        =   41
            Top             =   1635
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
            Left            =   3975
            TabIndex        =   40
            Top             =   1635
            Width           =   540
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Consumo Médio Máximo:"
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
            Left            =   1230
            TabIndex        =   39
            Top             =   1170
            Width           =   2100
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Consumo Médio:"
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
            Left            =   1920
            TabIndex        =   38
            Top             =   300
            Width           =   1410
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "% a Mais Máxima de Consumo Médio:"
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
            Left            =   150
            TabIndex        =   37
            Top             =   720
            Width           =   3180
         End
         Begin VB.Label ConsumoMedioMaxValor 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3420
            TabIndex        =   36
            Top             =   1125
            Width           =   915
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Classificação"
         Height          =   585
         Left            =   210
         TabIndex        =   33
         Top             =   3990
         Width           =   2145
         Begin VB.TextBox ClasseABC 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1500
            MaxLength       =   1
            TabIndex        =   8
            Top             =   210
            Width           =   255
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Classe ABC:"
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
            TabIndex        =   34
            Top             =   270
            Width           =   1050
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Produto"
         Height          =   1515
         Left            =   180
         TabIndex        =   26
         Top             =   210
         Width           =   8610
         Begin VB.ComboBox ControleEstoque 
            Height          =   315
            ItemData        =   "Estoque2.ctx":0000
            Left            =   3420
            List            =   "Estoque2.ctx":000D
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1050
            Width           =   2790
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   315
            Left            =   3405
            TabIndex        =   2
            Top             =   240
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label LblUMEstoque 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   0
            Left            =   7005
            TabIndex        =   32
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
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
            Left            =   2355
            TabIndex        =   31
            Top             =   690
            Width           =   930
         End
         Begin VB.Label Descricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3405
            TabIndex        =   30
            Top             =   660
            Width           =   4455
         End
         Begin VB.Label ProdutoLabel 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   2625
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   29
            Top             =   285
            Width           =   660
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Controle de Reserva/Estoque:"
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
            Left            =   690
            TabIndex        =   28
            Top             =   1095
            Width           =   2595
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "UM Estoque:"
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
            Left            =   5835
            TabIndex        =   27
            Top             =   300
            Width           =   1110
         End
      End
      Begin VB.CommandButton BotaoTipoProduto 
         Caption         =   "Traz dados default"
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
         Left            =   6525
         TabIndex        =   9
         Top             =   4065
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   4770
      Index           =   2
      Left            =   180
      TabIndex        =   10
      Top             =   810
      Visible         =   0   'False
      Width           =   9150
      Begin VB.Frame Frame2 
         Caption         =   "Quantidades"
         Height          =   1920
         Left            =   180
         TabIndex        =   48
         Top             =   2625
         Width           =   8460
         Begin VB.CheckBox ESCalculado 
            Caption         =   "Calculado"
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
            Left            =   3630
            TabIndex        =   18
            Top             =   915
            Value           =   1  'Checked
            Width           =   1275
         End
         Begin VB.CheckBox PPCalculado 
            Caption         =   "Calculado"
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
            Left            =   3630
            TabIndex        =   16
            Top             =   405
            Value           =   1  'Checked
            Width           =   1275
         End
         Begin MSMask.MaskEdBox EstoqueSeguranca 
            Height          =   285
            Left            =   2190
            TabIndex        =   17
            Top             =   885
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox EstoqueMaximo 
            Height          =   285
            Left            =   2190
            TabIndex        =   19
            Top             =   1365
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PontoPedido 
            Height          =   285
            Left            =   2175
            TabIndex        =   15
            Top             =   375
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox LoteEconomico 
            Height          =   285
            Left            =   6585
            TabIndex        =   21
            Top             =   1365
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox LoteMinimo 
            Height          =   285
            Left            =   6600
            TabIndex        =   20
            Top             =   885
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Lote Mínimo:"
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
            Left            =   5400
            TabIndex        =   55
            Top             =   930
            Width           =   1125
         End
         Begin VB.Label LblUMEstoque 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   1
            Left            =   6585
            TabIndex        =   54
            Top             =   360
            Width           =   705
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "UM:"
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
            Left            =   6165
            TabIndex        =   53
            Top             =   420
            Width           =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Lote Econômico:"
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
            Left            =   5085
            TabIndex        =   52
            Top             =   1410
            Width           =   1440
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Estoque Segurança:"
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
            Left            =   300
            TabIndex        =   51
            Top             =   930
            Width           =   1740
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Estoque Máximo:"
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
            Left            =   585
            TabIndex        =   50
            Top             =   1410
            Width           =   1455
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Ponto Pedido:"
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
            Left            =   825
            TabIndex        =   49
            Top             =   420
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ressuprimento (em dias)"
         Height          =   2370
         Left            =   180
         TabIndex        =   42
         Top             =   150
         Width           =   8460
         Begin VB.CheckBox TRCalculado 
            Caption         =   "Calculado"
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
            Left            =   3705
            TabIndex        =   12
            Top             =   360
            Value           =   1  'Checked
            Width           =   1275
         End
         Begin MSMask.MaskEdBox IntRessup 
            Height          =   315
            Left            =   2730
            TabIndex        =   14
            Top             =   1770
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TempoRessup 
            Height          =   315
            Left            =   2550
            TabIndex        =   11
            Top             =   345
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TempoRessupMax 
            Height          =   315
            Left            =   4110
            TabIndex        =   13
            Top             =   795
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Format          =   "0\%"
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "% a Mais Max de Tempo de Ressup:"
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
            Left            =   870
            TabIndex        =   47
            Top             =   840
            Width           =   3090
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Tempo de Ressup Máximo:"
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
            Left            =   825
            TabIndex        =   46
            Top             =   1260
            Width           =   2295
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Tempo de Ressup:"
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
            Left            =   855
            TabIndex        =   45
            Top             =   360
            Width           =   1605
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Intervalo de Ressup:"
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
            Left            =   870
            TabIndex        =   44
            Top             =   1830
            Width           =   1785
         End
         Begin VB.Label TempoRessupMaxValor 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3225
            TabIndex        =   43
            Top             =   1230
            Width           =   705
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7665
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   90
      Width           =   1695
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "Estoque2.ctx":003E
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   615
         Picture         =   "Estoque2.ctx":01BC
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "Estoque2.ctx":06EE
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Gravar"
         Top             =   105
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5160
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   9102
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ponto de Pedido"
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
Attribute VB_Name = "Estoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'todas as variáveis devem ser declaradas

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iProdutoAlterado As Integer
Dim iFrameAtual As Integer
Dim iValorAlterado As Integer
Dim iPontoPedidoAlterado As Integer
Dim iTempoRessupAlterado As Integer
Dim iEstoqueSegurancaAlterado As Integer

Dim WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_DadosPrincipais = 1
Private Const TAB_PontoDePedido = 2

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 53578

    'Limpa a Tela
    Call Limpa_Tela_Estoque

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 53578 'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 159545)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

   'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 53577

    'Limpa a Tela
    Call Limpa_Tela_Estoque

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 53577 'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159546)

    End Select

    Exit Sub

End Sub

Private Sub BotaoTipoProduto_Click()
'Traz dados default de Tipo Produto

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoComprasConfig_Click

    'Verifica preenchimento de Produto
    If Len(Trim(Produto.ClipText)) = 0 Then Error 25627

    sCodProduto = Produto.Text

    'Critica o formato do Produto e se existe no BD
    lErro = CF("Produto_Critica_Compra", sCodProduto, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 25605 Then Error 25628

    'Verifica se produto foi preenchido
    If iProdutoPreenchido = PRODUTO_VAZIO Then Error 25630

    'Não encontrou o Produto no BD
    If lErro = 25605 Then Error 25629
    
    'Limpa os campos que serão preenchidos com os valores default do Tipo de Produto
    
        'Frame Consumo
        ConsumoMedioMax.Text = "" '% a Mais Máxima de Consumo Médio:
        MesesConsumoMedio.Text = "" 'Faixa Cálculo Consumo Médio:____ meses
        ConsumoMedioMaxValor.Caption = "" 'Consumo Médio Máximo:
        '*********************************
        
        'Frame Ressuprimento (em dias)
        IntRessup.Text = "" 'Intervalo de Ressup:
        TempoRessupMax.Text = "" '% a Mais Max de Tempo de Ressup:
        TempoRessupMaxValor.Caption = "" 'Tempo de Ressup Máximo:
        '*********************************
        
    '**********************************************************************************
    
    'Traz dados do TipoProduto para tela
    lErro = Traz_TipoProduto_Tela(objProduto)
    If lErro <> SUCESSO Then Error 25631

    Exit Sub

Erro_BotaoComprasConfig_Click:

    Select Case Err

        Case 25628, 25631 'Tratado na rotina chamada
        
        Case 25627, 25630
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)
        
        Case 25629
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, sCodProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159547)

    End Select

    Exit Sub

End Sub

Private Sub ClasseABC_Change()

    'Transforma em maiúscula o que for digitado em ClasseABC.text
    ClasseABC.Text = UCase(ClasseABC.Text)
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ClasseABC_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sClasse As String

On Error GoTo Erro_ClasseABC_Validate

    'Verifica se a ClasseABC foi preenchida
    If Len(Trim(ClasseABC.Text)) = 0 Then Exit Sub

    sClasse = ClasseABC.Text

    'Se for diferente de A, B e C -> erro
    If (sClasse <> "A") And (sClasse <> "B") And (sClasse <> "C") Then Error 53588

    Exit Sub

Erro_ClasseABC_Validate:

    Cancel = True
    
    Select Case Err

        Case 53588
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLASSE_PRODUTO_INEXISTENTE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159548)

    End Select

    Exit Sub
End Sub

Private Sub CMCalculado_Click()

    iAlterado = REGISTRO_ALTERADO
    
    'Se calculado foi marcado
    If CMCalculado.Value = vbChecked Then
        ConsumoMedio.Text = ""
        ConsumoMedioMaxValor.Caption = ""
        iValorAlterado = 0
    End If
    
End Sub

Private Sub ConsumoMedio_Change()

    iValorAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ConsumoMedio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ConsumoMedio_Validate
    
    If iValorAlterado = 0 Then Exit Sub
        
    'verifica se o campo ConsumoMedio está preenchido
    If Len(Trim(ConsumoMedio.Text)) > 0 Then
            
        CMCalculado.Value = vbUnchecked

        'faz a crítica do valor informado
        lErro = Valor_Positivo_Critica(ConsumoMedio.Text)
        If lErro <> SUCESSO Then Error 53589

        'Se ConsumoMedioMax estiver preenchido,
        'Calcula o valor de Consumo Médio Máximo e coloca na tela
        If Len(Trim(ConsumoMedioMax.Text)) <> 0 Then
            lErro = Calcula_ConsumoMedioMaxValor(ConsumoMedio.Text, ConsumoMedioMax.Text)
            If lErro <> SUCESSO Then Error 53608
        End If

    End If

    Exit Sub

Erro_ConsumoMedio_Validate:

    Cancel = True
    
    Select Case Err

        Case 53589, 53608 'Erros tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159549)

    End Select

    Exit Sub

End Sub

Private Sub ConsumoMedioMax_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Function Trata_Parametros(Optional objProduto As ClassProduto) As Long

Dim lErro As Long, sProduto As String
Dim objProdutoFilial As New ClassProdutoFilial
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Trata_Parametros

    'Verifica se foi passado algum Produto
    If Not (objProduto Is Nothing) Then

        objProdutoFilial.sProduto = objProduto.sCodigo
        
        'Traz os dados para a Tela
        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 53602
    
        'Coloca na tela o Produto selecionado
        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True
        
        sProduto = Produto.Text

        'Critica o formato do Produto e se existe no BD
        lErro = CF("Produto_Critica2", sProduto, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 69115

        'Não encontrou o Produto
        If lErro = 25041 Then gError 69116
    
        lErro = Traz_Produto_Tela(objProduto)
        If lErro <> SUCESSO Then gError 83026
    
    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 53602, 51333, 69115, 83026

        Case 69116
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159550)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Sub ConsumoMedioMax_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ConsumoMedioMax, iAlterado)

End Sub

Private Sub ConsumoMedioMax_Validate(Cancel As Boolean)

Dim lErro As Long


On Error GoTo Erro_ConsumoMedioMax_Validate

    'Verifica se o campo ConsumoMedioMax está preenchido
    If Len(Trim(ConsumoMedioMax.Text)) > 0 Then

        'Faz a crítica do valor informado
        lErro = Valor_Critica(ConsumoMedioMax.Text)
        If lErro <> SUCESSO Then Error 53595
        
        'Se ConsumoMedio e ConsumoMedioMax estiverem preenchidos,
        'Calcula o valor de Consumo Médio Máximo e coloca na tela
        If Len(Trim(ConsumoMedio.Text)) > 0 Then
            lErro = Calcula_ConsumoMedioMaxValor(ConsumoMedio.Text, ConsumoMedioMax.Text)
            If lErro <> SUCESSO Then Error 53609
        End If


    End If

    Exit Sub

Erro_ConsumoMedioMax_Validate:

    Select Case Err

        Case 53595, 53609 'Erros tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159551)

    End Select

    Exit Sub

End Sub


Private Sub ControleEstoque_Click()

Dim lErro As Long

On Error GoTo Error_ControleEstoque_Click

    iAlterado = REGISTRO_ALTERADO

    'Verifica se a ComboBox ControleEstoque foi preenchida
    If ControleEstoque.ListIndex = -1 Then Exit Sub

    'Se o produto não tem controle de estoque
    If ControleEstoque.ItemData(ControleEstoque.ListIndex) = PRODUTO_CONTROLE_SEM_ESTOQUE Then

        'Limpa os controles da tela, com exceção do frame Produto
        Call Limpa_Tela_Estoque1

        'Desabilita os controles da tela (com exceção do frame Produto), pois os mesmos
        'não podem ser preenchidos para produtos sem controle de estoque
        
        Call Habilita_Controles(False)
    
    'Se o produto tem controle de estoque
    Else
        
        'Habilita os controles da tela
        Call Habilita_Controles(True)
    End If

    Exit Sub

Error_ControleEstoque_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159552)

    End Select

    Exit Sub

End Sub

Private Sub EstoqueMaximo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EstoqueMaximo_Validate

    'Verifica se EstoqueMáximo foi digitado
    If Len(Trim(EstoqueMaximo.Text)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_Positivo_Critica(EstoqueMaximo.Text)
    If lErro <> SUCESSO Then Error 53584

    Exit Sub

Erro_EstoqueMaximo_Validate:

    Cancel = True
    
    Select Case Err

        Case 53584
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159553)

    End Select

    Exit Sub

End Sub


Private Sub EstoqueSeguranca_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EstoqueSeguranca_Validate

    'Verifica se EstoqueSeguranca foi preenchido
    If Len(Trim(EstoqueSeguranca.Text)) = 0 Then Exit Sub
    
    ESCalculado.Value = vbUnchecked
    
    If iEstoqueSegurancaAlterado = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(EstoqueSeguranca.Text)
    If lErro <> SUCESSO Then Error 53586

    Exit Sub

Erro_EstoqueSeguranca_Validate:

    Cancel = True
    
    Select Case Err

        Case 53586

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159554)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iFrameAtual = 1

    Set objEventoProduto = New AdmEvento

'    'Carrega a árvore de Produtos com os Produtos do BD
'    lErro = CF("Carga_Arvore_Produto",TvwProduto.Nodes)
'    If lErro <> SUCESSO Then Error 53579

    'Inicializa as máscaras de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then Error 53580

    'Verifica se módulo de Estoque não faz parte do pacote
    If gcolModulo.Ativo(MODULO_ESTOQUE) <> MODULO_ATIVO Then

        'Se o estoque não estiver ativo não permitir selecionar controle de estoque na combo: nem reserva+estoque nem estoque
        'Selecionar a outra opção e desabilitar a combo
        ControleEstoque.ListIndex = 2
        ControleEstoque.Enabled = False

    End If
    
    ConsumoMedio.Format = FORMATO_ESTOQUE

    CMCalculado.Value = vbChecked
    
    'Verifica se módulo de Compras não está ativo
    If gcolModulo.Ativo(MODULO_COMPRAS) <> MODULO_ATIVO Then

        'Se Compras não estiver ativo desabilita as CheckBoxs de Calculado
        TRCalculado.Value = vbUnchecked
        TRCalculado.Enabled = False
        PPCalculado.Value = vbUnchecked
        PPCalculado.Enabled = False
        ESCalculado.Value = vbUnchecked
        ESCalculado.Enabled = False
        CMCalculado.Value = vbUnchecked
        CMCalculado.Enabled = False
        
    End If

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 53579, 53580 'Erros tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159555)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Private Sub IntRessup_GotFocus()

    Call MaskEdBox_TrataGotFocus(IntRessup, iAlterado)

End Sub

Private Sub LoteMinimo_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub LoteMinimo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LoteMinimo_Validate

    'Verifica se o Lote Minino foi preenchido
    If Len(Trim(LoteMinimo.Text)) = 0 Then Exit Sub

    'Critica o valor do lote mínimo
    lErro = Valor_NaoNegativo_Critica(LoteMinimo.Text)
    If lErro <> SUCESSO Then gError 111427

    Exit Sub

Erro_LoteMinimo_Validate:

    Cancel = True
    
    Select Case gErr

        Case 111427
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159556) '??? gerr

    End Select

    Exit Sub

End Sub

Private Sub MesesConsumoMedio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MesesConsumoMedio_Validate

    If Len(Trim(MesesConsumoMedio.Text)) = 0 Then Exit Sub

    'Se o usuário digitar alguma coisa neste campo, tem que ser maior que zero.
    lErro = Valor_Positivo_Critica(MesesConsumoMedio.Text)
    If lErro <> SUCESSO Then Error 53507

    Exit Sub

Erro_MesesConsumoMedio_Validate:

    Cancel = True
    
    Select Case Err
        
        Case 53507
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159557)

    End Select

    Exit Sub

End Sub

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(Opcao.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index

        Select Case iFrameAtual
        
            Case TAB_DadosPrincipais
                Parent.HelpContextID = IDH_ESTOQUE_DADOS_PRINCIPAIS
            
            Case TAB_PontoDePedido
                Parent.HelpContextID = IDH_ESTOQUE_PONTO_PEDIDO
        
        End Select
        
    End If

End Sub

Private Sub ProdNaFilial_Click()
    
   iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim vbMsgRes As VbMsgBoxResult
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer
Dim iCodigo As Integer
Dim objProdutoFilial As New ClassProdutoFilial
Dim iIndice As Integer
Dim objTipoDeProduto As New ClassTipoDeProduto

On Error GoTo Erro_Produto_Validate

    'Se Produto não foi alterado, sai
    If iProdutoAlterado <> REGISTRO_ALTERADO Then Exit Sub

    'Limpa descrição e UM
    Descricao.Caption = ""
    LblUMEstoque(0).Caption = ""
    LblUMEstoque(1).Caption = ""

    'Verifica preenchimento de Produto
    If Len(Trim(Produto.ClipText)) > 0 Then

        sCodProduto = Produto.Text
        
        'Critica o formato do Produto e se existe no BD
        lErro = CF("Produto_Critica2", sCodProduto, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then Error 53596

        'Não encontrou o Produto
        If lErro = 25041 Then Error 53597

        'Preenche ProdutoDescricao com Descrição do Produto
        Descricao.Caption = objProduto.sDescricao
        
        'Preenche a Unidade de Medida
        LblUMEstoque(0).Caption = objProduto.sSiglaUMEstoque
        LblUMEstoque(1).Caption = objProduto.sSiglaUMEstoque

    End If

    iProdutoAlterado = 0

    Exit Sub

Erro_Produto_Validate:

    Cancel = True
    
    Select Case Err

        Case 53596, 51332 'Tratados na rotina chamada
    
        Case 53597
            'Não encontrou Produto no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de Produtos
                Call Chama_Tela("Produto", objProduto)
            
            End If
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159558)

    End Select

    Exit Sub

End Sub

Private Sub TempoRessup_GotFocus()

    Call MaskEdBox_TrataGotFocus(TempoRessup, iAlterado)
    
End Sub

Private Sub TempoRessup_Validate(Cancel As Boolean)
'Se TempoRessup estiver preenchido,
'Calcula o valor do Tempo de Ressup Maximo e coloca na tela

Dim lErro As Long

On Error GoTo Erro_TempoRessup_Validate
    
    If iTempoRessupAlterado = REGISTRO_ALTERADO Then
        
        TRCalculado.Value = vbUnchecked
        
        If Len(Trim(TempoRessup.Text)) > 0 Then
            If Len(Trim(TempoRessupMax.Text)) <> 0 Then
                lErro = Calcula_TempoRessupMaxValor(TempoRessup.Text, TempoRessupMax.Text)
                If lErro <> SUCESSO Then Error 53613
            End If
        End If
    
    End If
    
    Exit Sub

Erro_TempoRessup_Validate:

    Cancel = True
    
    Select Case Err

        Case 53613 'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159559)

    End Select

    Exit Sub

End Sub

Private Sub TempoRessupMax_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TempoRessupMax_GotFocus()

    Call MaskEdBox_TrataGotFocus(TempoRessupMax, iAlterado)
    
End Sub

Private Sub TempoRessupMax_Validate(Cancel As Boolean)
'Se TempoRessupMax estiver preenchido,
'Calcula o valor do Tempo de Ressup Maximo e coloca na tela

Dim lErro As Long

On Error GoTo Erro_TempoRessupMax_Validate

    If Len(Trim(TempoRessupMax.Text)) > 0 Then
        If Len(Trim(TempoRessup.Text)) <> 0 Then
            lErro = Calcula_TempoRessupMaxValor(TempoRessup.Text, TempoRessupMax.Text)
            If lErro <> SUCESSO Then Error 53614
        End If
    End If

    Exit Sub

Erro_TempoRessupMax_Validate:

    Cancel = True
    
    Select Case Err

        Case 53614 'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159560)

    End Select

    Exit Sub


End Sub

'Private Sub TvwProduto_Expand(ByVal objNode As MSComctlLib.Node)
'
'Dim lErro As Long
'
'On Error GoTo Erro_TvwProduto_Expand
'
'    If (objNode.Tag <> NETOS_NA_ARVORE) Then
'
'        'Move os dados do plano de contas do Banco de Dados para a árvore
'        lErro = CF("Carga_Arvore_Produto_Netos",objNode, TvwProduto.Nodes)
'        If lErro <> SUCESSO Then Error 53590
'
'    End If
'
'    Exit Sub
'
'Erro_TvwProduto_Expand:
'
'    Select Case Err
'
'        Case 53590 'Erro tratado na rotina chamada
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159561)
'
'    End Select
'
'    Exit Sub
'
'End Sub

'Private Sub TvwProduto_NodeClick(ByVal Node As MSComctlLib.Node)
'
'Dim lErro As Long
'Dim sCodigo As String
'Dim objProduto As New ClassProduto
'Dim objProdutoFilial As New ClassProdutoFilial
'Dim bCancel As Boolean
'
'On Error GoTo Erro_TvwProduto_NodeClick
'
'    'Verifica se produto tem filhos
'    If Node.Children > 0 Then Exit Sub
'
'    'Armazena key do nó clicado sem caracter inicial
'    sCodigo = Right(Node.Key, Len(Node.Key) - 1)
'
'    objProduto.sCodigo = sCodigo
'
'    'Lê Produto
'    lErro = CF("Produto_Le",objProduto)
'    If lErro <> SUCESSO And lErro <> 28030 Then Error 53591
'
'    'Verifica se Produto é gerencial
'    If objProduto.iGerencial = GERENCIAL Then Exit Sub
'
'    'Mostra Unidade de Medida na tela
'    LblUMEstoque(0).Caption = objProduto.sSiglaUMEstoque
'    LblUMEstoque(1).Caption = objProduto.sSiglaUMEstoque
'
'    lErro = CF("Traz_Produto_MaskEd",sCodigo, Produto, Descricao)
'    If lErro <> SUCESSO Then Error 53592
'
'    'Chama Validate de Produto
'    Call Produto_Validate(bCancel)
'
'    'Fecha comando de setas se estiver aberto
'    lErro = ComandoSeta_Fechar(Me.Name)
'
'    iAlterado = 0
'
'    Exit Sub
'
'Erro_TvwProduto_NodeClick:
'
'    Select Case Err
'
'        Case 53591, 53592
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159562)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Function Gravar_Registro() As Long
'Verifica se dados de Estoque necessários foram preenchidos

Dim lErro As Long
Dim objProdutoFilial As New ClassProdutoFilial
Dim iControleEstoque As Integer
Dim sClasse As String

On Error GoTo Erro_Gravar_Registro
    
    'Verifica se o Código do Produto foi preenchido
    If Len(Trim(Produto.ClipText)) = 0 Then gError 53603

    'Verifica se o Controle Estoque foi informado
    If Len(Trim(ControleEstoque.Text)) = 0 Then gError 53604

    'Verifica se a classeABC foi preenchida
    If Len(Trim(ClasseABC.Text)) <> 0 Then

        sClasse = ClasseABC.Text

        'Se for diferente de A, B ou C -> erro
        If (sClasse <> "A") Then
            If (sClasse <> "B") Then
                    If (sClasse <> "C") Then gError 51334
            End If
        End If
    End If

    
    'Lê os dados da Tela relacionados ao Estoque
    lErro = Move_Tela_Memoria(objProdutoFilial, iControleEstoque)
    If lErro <> SUCESSO Then gError 53606

    'Verifica se Lote mínimo e Lote econômico estão preenchidos
    If objProdutoFilial.dLoteEconomico <> 0 And objProdutoFilial.dLoteMinimo <> 0 Then
        
        'Verifica se o Lote mínimo é maior do que o Lote econômico*** se for Erro
        If objProdutoFilial.dLoteMinimo > objProdutoFilial.dLoteEconomico Then gError 111428
        
    End If
    
    lErro = Trata_Alteracao(objProdutoFilial, objProdutoFilial.sProduto)
    If lErro <> SUCESSO Then gError 32311
    
    'Grava os dados do Controle de Estoque do Produto no BD
    lErro = CF("Estoque_Grava", objProdutoFilial, iControleEstoque)
    If lErro <> SUCESSO Then gError 53607

    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 32311

        Case 53603
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 53604
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTROLE_ESTOQUE_NAO_PREENCHIDO", gErr)

        Case 53606, 53607

        Case 51334
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLASSE_PRODUTO_INEXISTENTE", gErr, ClasseABC.Text)

        Case 111428
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTEMINIMO_MAIOR", gErr, objProdutoFilial.dLoteEconomico, objProdutoFilial.dLoteMinimo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159563)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objProdutoFilial As ClassProdutoFilial, iControleEstoque As Integer) As Long

'Move os dados da tela para memória
Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_Move_Tela_Memoria

    objProdutoFilial.iFilialEmpresa = giFilialEmpresa

    'Verifica se o Produto foi preenchido
    If Len(Trim(Produto.ClipText)) > 0 Then

        'Passa para o formato do BD
        lErro = CF("Produto_Formata", Produto.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then Error 53601

        'Testa se o codigo está preenchido
        If iPreenchido = PRODUTO_PREENCHIDO Then objProdutoFilial.sProduto = sProduto
        
        lErro = CF("ProdutoFilial_Le", objProdutoFilial)
        If lErro <> SUCESSO And lErro <> 28261 Then gError 53601

    End If
    
    'Move o que está selecionado na Combobox ControleEstoque para iControleEstoque
    If ControleEstoque.ListIndex <> -1 Then iControleEstoque = ControleEstoque.ItemData(ControleEstoque.ListIndex)

    'Frame Consumo
    objProdutoFilial.dConsumoMedio = StrParaDbl(ConsumoMedio.Text)
    objProdutoFilial.dConsumoMedioMax = StrParaDbl(ConsumoMedioMax.Text) / 100
    objProdutoFilial.iCMCalculado = StrParaInt(CMCalculado.Value)
    objProdutoFilial.iMesesConsumoMedio = StrParaInt(MesesConsumoMedio.Text)

    'Frame Ressuprimento
    objProdutoFilial.iTempoRessup = StrParaInt(TempoRessup.Text)
    objProdutoFilial.dTempoRessupMax = StrParaDbl(TempoRessupMax.Text) / 100
    objProdutoFilial.iTRCalculado = StrParaInt(TRCalculado.Value)
    objProdutoFilial.iIntRessup = StrParaInt(IntRessup.Text)

    'Frame Quantidades
    objProdutoFilial.dEstoqueSeguranca = StrParaDbl(EstoqueSeguranca.Text)
    objProdutoFilial.dEstoqueMaximo = StrParaDbl(EstoqueMaximo.Text)
    objProdutoFilial.dPontoPedido = StrParaDbl(PontoPedido.Text)
    objProdutoFilial.dLoteEconomico = StrParaDbl(LoteEconomico.Text)
    objProdutoFilial.dLoteMinimo = StrParaDbl(LoteMinimo.Text)
    objProdutoFilial.iPPCalculado = StrParaInt(PPCalculado.Value)
    objProdutoFilial.iESCalculado = StrParaInt(ESCalculado.Value)

    'Frame Classificação
    objProdutoFilial.sClasseABC = ClasseABC.Text
    
    'CheckBox ProduzidoNaFilial
    objProdutoFilial.iProdNaFilial = ProdNaFilial.Value
 
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 53601

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159564)

    End Select

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

'Extrai os campos da tela que correspondem aos campos no BD
Dim lErro As Long
Dim objProdutoFilial As New ClassProdutoFilial
Dim iControleEstoque As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "ProdutosFilial"

    'Lê os dados da Tela
    lErro = Move_Tela_Memoria(objProdutoFilial, iControleEstoque)
    If lErro <> SUCESSO Then Error 53593

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Produto", objProdutoFilial.sProduto, STRING_PRODUTO, "Produto"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 53593 'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159565)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objProdutoFilial As New ClassProdutoFilial
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim bCancel As Boolean

On Error GoTo Erro_Tela_Preenche

    'Passa o produto da coleção de campos-valores para objprodutofilial.sproduto
    objProdutoFilial.sProduto = colCampoValor.Item("Produto").vValor
    objProdutoFilial.iFilialEmpresa = giFilialEmpresa

    'Se o produto existir
    If objProdutoFilial.sProduto <> "" Then

        lErro = Mascara_RetornaProdutoEnxuto(objProdutoFilial.sProduto, sProduto)
        If lErro <> SUCESSO Then gError 53594

        'Coloca na tela o Produto selecionado
        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True

        objProduto.sCodigo = objProdutoFilial.sProduto

        'Lê Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 83022
    
        If lErro = 28030 Then gError 83023
  
        lErro = Traz_Produto_Tela(objProduto)
        If lErro <> SUCESSO Then gError 83024

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

         Case 53594, 83022, 83024

         Case 83023
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159566)

    End Select

    Exit Sub

End Sub

Private Function Traz_Estoque_Tela(objProdutoFilial As ClassProdutoFilial, iControleEstoque As Integer, objProduto As ClassProduto) As Long
'Traz os dados do Estoque para tela

Dim lErro As Long
Dim sProduto As String

On Error GoTo Erro_Traz_Estoque_Tela

    'Função que limpa os controles da tela dos frames de Consumo, Ressuprimento, Quantidades e Classificação
    Call Limpa_Tela_Estoque1

    'Coloca os demais dados do Estoque na tela
    If iControleEstoque > 0 Then
        ControleEstoque.ListIndex = iControleEstoque - 1
    Else
        ControleEstoque.ListIndex = -1
    End If

    objProdutoFilial.iFilialEmpresa = giFilialEmpresa

    'Lê o ProdutoFilial
    lErro = CF("ProdutoFilial_Le", objProdutoFilial)
    If lErro <> SUCESSO And lErro <> 28261 Then gError 53599

    'Não encontrou o ProdutoFilial ==> erro
    If lErro = 28261 Then gError 53600

    lErro = Mascara_RetornaProdutoEnxuto(objProdutoFilial.sProduto, sProduto)
    If lErro <> SUCESSO Then gError 53605

    'Frame Consumo
    If objProdutoFilial.dConsumoMedio <> 0 Then
        ConsumoMedio.Text = objProdutoFilial.dConsumoMedio
    Else
        ConsumoMedio.Text = ""
    End If
    
    ConsumoMedioMax.Text = objProdutoFilial.dConsumoMedioMax * 100
    
    CMCalculado.Value = objProdutoFilial.iCMCalculado

    If objProdutoFilial.iMesesConsumoMedio <> 0 Then
        MesesConsumoMedio.Text = (objProdutoFilial.iMesesConsumoMedio)
    Else
        MesesConsumoMedio.Text = ""
    End If

    'Frame Ressuprimento
    TempoRessup.Text = objProdutoFilial.iTempoRessup

    TempoRessupMax.Text = objProdutoFilial.dTempoRessupMax * 100

    IntRessup.Text = objProdutoFilial.iIntRessup

    TRCalculado.Value = objProdutoFilial.iTRCalculado

    'Frame Quantidades
    EstoqueSeguranca.Text = objProdutoFilial.dEstoqueSeguranca

    If objProdutoFilial.dEstoqueMaximo <> 0 Then
        EstoqueMaximo.Text = objProdutoFilial.dEstoqueMaximo
    Else
        EstoqueMaximo.Text = ""
    End If

    If objProdutoFilial.dPontoPedido <> 0 Then
        PontoPedido.Text = objProdutoFilial.dPontoPedido
    Else
        PontoPedido.Text = ""
    End If

    If objProdutoFilial.dLoteEconomico <> 0 Then
        LoteEconomico.Text = objProdutoFilial.dLoteEconomico
    Else
        LoteEconomico.Text = ""
    End If

    '***** Incluido P/Sergio para Trazer para a Tela o Lote Minimo de um Produto dia 30/10/2002
    If objProdutoFilial.dLoteMinimo <> 0 Then
        LoteMinimo.Text = objProdutoFilial.dLoteMinimo
    Else
        LoteMinimo.Text = ""
    End If

    PPCalculado.Value = objProdutoFilial.iPPCalculado
    ESCalculado.Value = objProdutoFilial.iESCalculado

    ClasseABC.Text = objProdutoFilial.sClasseABC

    'Se ConsumoMedio e ConsumoMedioMax estiverem preenchidos,
    'Calcula o valor do Consumo Medio Maximo e coloca na tela
    If Len(Trim(ConsumoMedio.Text)) <> 0 And Len(Trim(ConsumoMedioMax.Text)) <> 0 Then
        lErro = Calcula_ConsumoMedioMaxValor(ConsumoMedio.Text, ConsumoMedioMax.Text)
        If lErro <> SUCESSO Then gError 53610
    End If

    'Se TempoRessup e TempoRessupMax estiverem preenchidos,
    'Calcula o valor do Tempo de Ressup Maximo e coloca na tela
    If Len(Trim(TempoRessup.Text)) <> 0 And Len(Trim(TempoRessupMax.Text)) <> 0 Then
        lErro = Calcula_TempoRessupMaxValor(TempoRessup.Text, TempoRessupMax.Text)
        If lErro <> SUCESSO Then gError 53615
    End If
    
    'preenchimento da checkbox PreoduzidoNaFilial
    If objProdutoFilial.iProdNaFilial = PRODUZIDO_NA_FILIAL Then
        ProdNaFilial.Value = vbChecked
    Else
        ProdNaFilial.Value = vbUnchecked
    End If
    
    LblUMEstoque(0).Caption = objProduto.sSiglaUMEstoque
    LblUMEstoque(1).Caption = objProduto.sSiglaUMEstoque
    
    iAlterado = 0
    
    Traz_Estoque_Tela = SUCESSO

    Exit Function

Erro_Traz_Estoque_Tela:

    Traz_Estoque_Tela = gErr

    Select Case gErr

        Case 53605, 53599, 53610, 53615, 75469 'Erros tratados nas rotinas chamadas

        Case 53600 'ProdutoFilial não existe no BD

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159567)

    End Select

    Exit Function

End Function

Private Sub ControleEstoque_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Descricao_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ESCalculado_Click()

    iAlterado = REGISTRO_ALTERADO
    
    If ESCalculado.Value = vbChecked Then
        EstoqueSeguranca.Text = ""
    End If
    
End Sub

Private Sub EstoqueMaximo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub EstoqueSeguranca_Change()

    iAlterado = REGISTRO_ALTERADO
    iEstoqueSegurancaAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoProduto = Nothing

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub IntRessup_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LblUMEstoque_Click(Index As Integer)

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LoteEconomico_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LoteEconomico_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LoteEconomico_Validate

    'Verifica se o Lote Economico foi preenchido
    If Len(Trim(LoteEconomico.Text)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(LoteEconomico.Text)
    If lErro <> SUCESSO Then Error 53587

    Exit Sub

Erro_LoteEconomico_Validate:

    Cancel = True
    
    Select Case Err

        Case 53587
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159568)

    End Select

    Exit Sub

End Sub

Private Sub MesesConsumoMedio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MesesConsumoMedio_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(MesesConsumoMedio, iAlterado)
    
End Sub

Private Sub PontoPedido_Change()

    iAlterado = REGISTRO_ALTERADO
    iPontoPedidoAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub PontoPedido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PontoPedido_Validate

    'Verifica se o Ponto Pedido foi preenchido
    If Len(Trim(PontoPedido.Text)) = 0 Then Exit Sub

    If iPontoPedidoAlterado = 0 Then Exit Sub

    PPCalculado.Value = vbUnchecked
    
    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(PontoPedido.Text)
    If lErro <> SUCESSO Then Error 53585

    Exit Sub

Erro_PontoPedido_Validate:

    Cancel = True
    
    Select Case Err

        Case 53585
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159569)

    End Select

    Exit Sub

End Sub

Private Sub PPCalculado_Click()

    iAlterado = REGISTRO_ALTERADO
    
    If PPCalculado.Value = vbChecked Then
        PontoPedido.Text = ""
    End If
    
End Sub

Private Sub Produto_Change()

    iProdutoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoLabel_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_ProdutoLabel_Click

    'Verifica se o produto foi preenchido
    If Len(Trim(Produto.ClipText)) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 71934

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_ProdutoLabel_Click:

    Select Case gErr

        Case 71934

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159570)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim bCancel As Boolean

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 71935

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 71936

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, Produto, Descricao)
    If lErro <> SUCESSO Then gError 53592

    lErro = Traz_Produto_Tela(objProduto)
    If lErro <> SUCESSO Then gError 83025

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 71935, 83025

        Case 71936
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159571)

    End Select

    Exit Sub

End Sub

Private Sub TempoRessup_Change()

    iAlterado = REGISTRO_ALTERADO
    iTempoRessupAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TRCalculado_Click()

    iAlterado = REGISTRO_ALTERADO
    
    'Se Calculado foi marcado, limpa o tempo de Ressuprimento
    If TRCalculado.Value = vbChecked Then
        TempoRessup.Text = ""
        TempoRessupMaxValor.Caption = ""
    End If
    
End Sub


Private Function Limpa_Tela_Estoque() As Long
'Limpa os campos do Frame Produto
'Chama a função Limpa_Tela_Estoque1 para limpar os controles dos frames restantes
'Limpa iAlterado

    'Frame Produto
    
        'Código:
        Produto.PromptInclude = False
        Produto.Text = "" 'Código:
        Produto.PromptInclude = True
        '*****
        
        LblUMEstoque(0).Caption = "" ' UM Estoque:
        Descricao.Caption = "" 'Descrição:
        ControleEstoque.ListIndex = -1 'Controle de Reserva/Estoque:
    '*******************************

    'Limpa os controles restantes
    Call Limpa_Tela_Estoque1
    
    'Habilita os controles que, eventualmente, poderão estar desabilitadas em função do último item selecionado na ComboBox ControleEstoque
    Call Habilita_Controles(True)
    '********************************************************************
    
    'Seta iAlterado como não alterado
    iAlterado = 0

End Function

Public Function Limpa_Tela_Estoque1()
'Limpa os controles dos frames Consumo, Classificação, Ressuprimento (em dias) e Quantidades
'Marca as CheckBox da tela
'Pode ser chamada a partir da função Limpa_Tela_Estoque (essa função limpa apenas o frame Produto)
'E pode ser chamada a partir de outros pontos da tela onde é necessário limpar apenas os frames acima citados
    
    'Marca as CheckBox da tela
    If gcolModulo.Ativo(MODULO_COMPRAS) = MODULO_ATIVO Then
        CMCalculado.Value = vbChecked
        ESCalculado.Value = vbChecked
        PPCalculado.Value = vbChecked
        TRCalculado.Value = vbChecked
    End If
    
    'Produto
    LblUMEstoque(0).Caption = "" 'UM Estoque:
    '*******************************
    
    'Frame Consumo
    ConsumoMedio.Text = "" 'Consumo Médio:
    ConsumoMedioMax.Text = "" '% a Mais Máxima de Consumo Médio:
    ConsumoMedioMaxValor.Caption = "" 'Consumo Médio Máximo:
    MesesConsumoMedio.Text = "" 'Faixa Cálculo Consumo Médio:____ meses
    '*************************************
    
    'Classificação
    ClasseABC.Text = "" 'Classe ABC:
    '************************
    
    ' Ressuprimento (em dias)
    TempoRessup.Text = "" 'Tempo de Ressup:
    TempoRessupMax.Text = "" '% a Mais Max de Tempo de Ressup:
    TempoRessupMaxValor.Caption = "" 'Tempo de Ressup Máximo:
    IntRessup.Text = "" 'Intervalo de Ressup:
    '*************************************
    
    'Quantidades
    PontoPedido.Text = "" 'Ponto Pedido:
    LblUMEstoque(1).Caption = "" 'UM:
    EstoqueSeguranca.Text = "" 'Estoque Segurança:
    EstoqueMaximo.Text = "" 'Estoque Máximo:
    LoteEconomico.Text = "" 'Lote Econômico:
    LoteMinimo.Text = "" 'Lote Mímimo
    '*********************************
    
    'Checkbox ProdNaFilial
    ProdNaFilial.Value = vbUnchecked 'Produzido na Filial
    '*********************************

End Function

Function Traz_TipoProduto_Tela(objProduto As ClassProduto) As Long
'Traz os dados de TipoProduto para tela
'O tipo deve estar passado dentro de objProduto

Dim lErro As Long
Dim objTipoDeProduto As New ClassTipoDeProduto

On Error GoTo Erro_Traz_TipoProduto_Tela
  
    objTipoDeProduto.iTipo = objProduto.iTipo

    'Lê o TipoDeProduto
    lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
    If lErro <> SUCESSO And lErro <> 22531 Then Error 53582

    'Não encontrou o TipoDeProduto
    If lErro = 22531 Then Error 53583
    
    'Dados de Consumo
    ConsumoMedioMax.Text = (objTipoDeProduto.dConsumoMedioMax) * 100

    If objTipoDeProduto.iMesesConsumoMedio <> 0 Then
        MesesConsumoMedio.Text = CStr(objTipoDeProduto.iMesesConsumoMedio)
    Else
        MesesConsumoMedio.Text = ""
    End If

    'Se ConsumoMedio e ConsumoMedioMax estiverem preenchidos,
    'calcula o valor do Consumo Medio Maximo e coloca na tela
    If Len(Trim(ConsumoMedio.Text)) <> 0 And Len(Trim(ConsumoMedioMax.Text)) <> 0 Then
        lErro = Calcula_ConsumoMedioMaxValor(ConsumoMedio.Text, ConsumoMedioMax.Text)
        If lErro <> SUCESSO Then Error 53611
    End If
    
    'Dados de ressuprimento
    IntRessup.Text = CStr(objTipoDeProduto.iIntRessup)
    TempoRessupMax.Text = (objTipoDeProduto.dTempoRessupMax) * 100

    'Se TempoRessup e TempoRessupMax estiverem preenchidos,
    'calcula o valor do Tempo de Ressup Maximo e coloca na tela
    If Len(Trim(TempoRessup.Text)) <> 0 And Len(Trim(TempoRessupMax.Text)) <> 0 Then
        lErro = Calcula_TempoRessupMaxValor(TempoRessup.Text, TempoRessupMax.Text)
        If lErro <> SUCESSO Then Error 53612
    End If

    Traz_TipoProduto_Tela = SUCESSO

    Exit Function

Erro_Traz_TipoProduto_Tela:

    Traz_TipoProduto_Tela = Err

    Select Case Err

        Case 53582, 53611, 53612  'Tratados nas rotinas chamadas

        Case 53583
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_TIPO", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159572)

    End Select

    Exit Function

End Function


Function Calcula_ConsumoMedioMaxValor(sConsumoMedio As String, sConsumoMedioMax As String) As Long

Dim dConsumoMedio As Double
Dim dConsumoMedioMax As Double
Dim dResultado As Double
Dim sConsumoMedioMaxValor As String

    dConsumoMedioMax = StrParaDbl(sConsumoMedioMax)
    dConsumoMedioMax = dConsumoMedioMax / 100
    dConsumoMedio = StrParaDbl(sConsumoMedio)

    dResultado = dConsumoMedio * (1 + dConsumoMedioMax)

    sConsumoMedioMaxValor = CStr(dResultado)
    ConsumoMedioMaxValor.Caption = sConsumoMedioMaxValor

End Function

Function Calcula_TempoRessupMaxValor(sTempoRessup As String, sTempoRessupMax As String) As Long
'A partir de sTempoRessup e de sTempoRessupMax (%) calcula
'sTempoRessupMaxValor e coloca na tela

Dim dTempoRessup As Double
Dim dTempoRessupMax As Double
Dim dResultado As Double
Dim sTempoRessupMaxValor As String

    'Conversão de tipos
    dTempoRessupMax = StrParaDbl(sTempoRessupMax)
    dTempoRessupMax = dTempoRessupMax / 100
    dTempoRessup = CDbl(sTempoRessup)

    'Cálculo
    dResultado = dTempoRessup * (1 + dTempoRessupMax)

    'Coloca resultado na tela
    sTempoRessupMaxValor = CStr(dResultado)
    TempoRessupMaxValor.Caption = sTempoRessupMaxValor

End Function

Private Function Traz_Produto_Tela(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objProdutoFilial As New ClassProdutoFilial

On Error GoTo Erro_Traz_Produto_Tela

    'Preenche ProdutoDescricao com Descrição do Produto
    Descricao.Caption = objProduto.sDescricao

    'Preenche a Unidade de Medida
    LblUMEstoque(0).Caption = objProduto.sSiglaUMEstoque
    LblUMEstoque(1).Caption = objProduto.sSiglaUMEstoque
    
    'Preenche a combo Controle de Estoque
    For iIndice = 0 To (ControleEstoque.ListCount - 1)
        If (ControleEstoque.ItemData(iIndice)) = objProduto.iControleEstoque Then
            ControleEstoque.ListIndex = iIndice
            Exit For
        End If
    Next

    'Passa o código do produto para o objProdutoFilial
    objProdutoFilial.sProduto = objProduto.sCodigo

    'Tenta trazer os dados de ProdutoFilial
    lErro = Traz_Estoque_Tela(objProdutoFilial, objProduto.iControleEstoque, objProduto)
    If lErro <> SUCESSO And lErro <> 53600 Then Error 51335
            
    Traz_Produto_Tela = SUCESSO
            
    Exit Function

Erro_Traz_Produto_Tela:

    Traz_Produto_Tela = Err

    Select Case Err
    
        Case 51335

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159573)
            
    End Select
    
    Exit Function
    
End Function

Private Function Habilita_Controles(bHabilita As Boolean)
'Habilita / Desabilita os controles dos frames Consumo, Classificação, Ressuprimento (em dias) e Quantidades

        'Frame Consumo
        ConsumoMedio.Enabled = bHabilita 'Consumo Médio:
        CMCalculado.Enabled = bHabilita 'Calculado
        ConsumoMedioMax.Enabled = bHabilita '% a Mais Máxima de Consumo Médio:
        MesesConsumoMedio.Enabled = bHabilita 'Faixa cálculo Consumo Médio: ____meses
        '*******************************
        
        'Frame Classificação
        ClasseABC.Enabled = bHabilita 'Classe ABC:
        '************************
        
        'Frame Ressuprimento (em dias)
        TempoRessup.Enabled = bHabilita 'Tempo Ressup:
        TRCalculado.Enabled = bHabilita 'Calculado
        TempoRessupMax.Enabled = bHabilita '% a mais Max de Tempo de Ressup:
        IntRessup.Enabled = bHabilita 'Intervalo de Ressup:
        '******************************
        
        'Frame Quantidades
        PontoPedido.Enabled = bHabilita 'Ponto Pedido:
        PPCalculado.Enabled = bHabilita 'Calculado
        EstoqueSeguranca.Enabled = bHabilita 'Estoque Segurança:
        ESCalculado.Enabled = bHabilita 'Calculado
        EstoqueMaximo.Enabled = bHabilita 'Estoque Máximo:
        LoteEconomico.Enabled = bHabilita 'Lote Econômico:
        LoteMinimo.Enabled = bHabilita 'Lote Mínimo:
        '********************************************
        
        'Frame Produção
        'ProdNaFilial.Enabled = bHabilita 'Produzido na Filial
        '********************************************

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ESTOQUE_DADOS_PRINCIPAIS
    Set Form_Load_Ocx = Me
    Caption = "Controle de Estoque"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Estoque"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Produto Then
            Call ProdutoLabel_Click
        End If
    End If

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

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub ConsumoMedioMaxValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ConsumoMedioMaxValor, Source, X, Y)
End Sub

Private Sub ConsumoMedioMaxValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ConsumoMedioMaxValor, Button, Shift, X, Y)
End Sub

Private Sub Label39_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label39, Source, X, Y)
End Sub

Private Sub Label39_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label39, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Descricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Descricao, Source, X, Y)
End Sub

Private Sub Descricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Descricao, Button, Shift, X, Y)
End Sub

Private Sub ProdutoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ProdutoLabel, Source, X, Y)
End Sub

Private Sub ProdutoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label29_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label29, Source, X, Y)
End Sub

Private Sub Label29_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label29, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub TempoRessupMaxValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TempoRessupMaxValor, Source, X, Y)
End Sub

Private Sub TempoRessupMaxValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TempoRessupMaxValor, Button, Shift, X, Y)
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

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

