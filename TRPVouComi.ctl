VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TRPVouComi 
   ClientHeight    =   5895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   9510
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   8325
      ScaleHeight     =   450
      ScaleWidth      =   1005
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   45
      Width           =   1065
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "TRPVouComi.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   540
         Picture         =   "TRPVouComi.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoHist 
      Caption         =   "Histórico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4560
      TabIndex        =   4
      Top             =   135
      Width           =   1740
   End
   Begin VB.Frame Frame2 
      Caption         =   "Voucher"
      Height          =   975
      Left            =   105
      TabIndex        =   23
      Top             =   660
      Width           =   9300
      Begin VB.CheckBox Antc 
         Caption         =   "Antc"
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
         Left            =   7815
         TabIndex        =   63
         Top             =   585
         Width           =   945
      End
      Begin VB.CheckBox Cartao 
         Caption         =   "Cartão"
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
         Left            =   5835
         TabIndex        =   62
         Top             =   600
         Width           =   945
      End
      Begin VB.Label Destino 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1650
         TabIndex        =   60
         Top             =   570
         Width           =   2970
      End
      Begin VB.Label Produto 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   5850
         TabIndex        =   58
         Top             =   195
         Width           =   3345
      End
      Begin VB.Label Label1 
         Caption         =   "Destino:"
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
         Index           =   44
         Left            =   900
         TabIndex        =   61
         Top             =   615
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "Produto:"
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
         Index           =   1
         Left            =   5100
         TabIndex        =   59
         Top             =   225
         Width           =   1020
      End
      Begin VB.Label DataEmissaoVou 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1680
         TabIndex        =   26
         Top             =   195
         Width           =   1290
      End
      Begin VB.Label ValorBrutoVou 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   3675
         TabIndex        =   24
         Top             =   195
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Data de Emissão:"
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
         Index           =   5
         Left            =   150
         TabIndex        =   27
         Top             =   255
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "Bruto:"
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
         Index           =   3
         Left            =   3135
         TabIndex        =   25
         Top             =   255
         Width           =   615
      End
   End
   Begin VB.CommandButton BotaoTrazerVou 
      Height          =   330
      Left            =   3930
      Picture         =   "TRPVouComi.ctx":02D8
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Trazer Dados"
      Top             =   150
      Width           =   360
   End
   Begin VB.Frame FrameSuporte 
      Caption         =   "Simular Comissão"
      Height          =   1260
      Left            =   105
      TabIndex        =   34
      Top             =   5880
      Width           =   9270
      Begin VB.CheckBox Import 
         Caption         =   "Simular importação"
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
         Left            =   2925
         TabIndex        =   17
         Top             =   585
         Width           =   2085
      End
      Begin VB.CommandButton BotaoPrimeiraComissao 
         Caption         =   "Primeira Comissão"
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
         Left            =   165
         TabIndex        =   77
         Top             =   960
         Width           =   1800
      End
      Begin VB.CheckBox CartaoNovo 
         Caption         =   "Cartão"
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
         Left            =   5865
         TabIndex        =   14
         Top             =   240
         Width           =   945
      End
      Begin VB.CheckBox AntcNovo 
         Caption         =   "Antc"
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
         Left            =   7845
         TabIndex        =   15
         Top             =   225
         Width           =   945
      End
      Begin VB.CommandButton BotaoExcluirComissao 
         Caption         =   "Excluir Comissão"
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
         Left            =   2040
         TabIndex        =   19
         Top             =   960
         Width           =   1740
      End
      Begin MSMask.MaskEdBox BrutoNovo 
         Height          =   315
         Left            =   1680
         TabIndex        =   16
         Top             =   555
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSMask.MaskEdBox ProdutoNovo 
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Top             =   180
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox DestinoNovo 
         Height          =   315
         Left            =   5865
         TabIndex        =   18
         Top             =   555
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Cliente 
         Height          =   315
         Left            =   5865
         TabIndex        =   76
         Top             =   900
         Visible         =   0   'False
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Destino:"
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
         Index           =   22
         Left            =   5055
         TabIndex        =   75
         Top             =   600
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "Produto:"
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
         Index           =   21
         Left            =   870
         TabIndex        =   74
         Top             =   225
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Bruto:"
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
         Index           =   20
         Left            =   1065
         TabIndex        =   73
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Comissão Nova"
      Height          =   1800
      Left            =   105
      TabIndex        =   33
      Top             =   3855
      Width           =   9285
      Begin MSMask.MaskEdBox RepresentanteNovo 
         Height          =   315
         Left            =   1665
         TabIndex        =   6
         Top             =   570
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox PercComiRepNovo 
         Height          =   315
         Left            =   5850
         TabIndex        =   7
         Top             =   570
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CorrentistaNovo 
         Height          =   315
         Left            =   1665
         TabIndex        =   8
         Top             =   930
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox PercComiCorNovo 
         Height          =   315
         Left            =   5850
         TabIndex        =   9
         Top             =   930
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox EmissorNovo 
         Height          =   315
         Left            =   1665
         TabIndex        =   10
         Top             =   1275
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox PercComiEmiNovo 
         Height          =   315
         Left            =   5850
         TabIndex        =   11
         Top             =   1275
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PercComiAgeNovo 
         Height          =   315
         Left            =   5850
         TabIndex        =   5
         Top             =   210
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PromotorNovo 
         Height          =   315
         Left            =   1665
         TabIndex        =   12
         Top             =   1635
         Visible         =   0   'False
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin VB.Label AgenciaNovo 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1665
         TabIndex        =   80
         Top             =   195
         Width           =   3000
      End
      Begin VB.Label LabelPromotor 
         Caption         =   "Promotor:"
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
         Left            =   765
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   72
         Top             =   1665
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "% Comissão:"
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
         Index           =   18
         Left            =   4755
         TabIndex        =   71
         Top             =   225
         Width           =   1140
      End
      Begin VB.Label LabelAgencia 
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
         Height          =   330
         Left            =   930
         TabIndex        =   70
         Top             =   225
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "% Comissão:"
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
         Index           =   16
         Left            =   4755
         TabIndex        =   69
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label LabelEmissor 
         Caption         =   "Emissor:"
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
         Left            =   870
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   68
         Top             =   1290
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "% Comissão:"
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
         Index           =   14
         Left            =   4755
         TabIndex        =   67
         Top             =   945
         Width           =   1140
      End
      Begin VB.Label LabelCorrentista 
         Caption         =   "Correntista:"
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
         Left            =   585
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   66
         Top             =   975
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "% Comissão:"
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
         Index           =   6
         Left            =   4755
         TabIndex        =   65
         Top             =   1305
         Width           =   1140
      End
      Begin VB.Label LabelRepresentante 
         Caption         =   "Representante:"
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
         Left            =   255
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   64
         Top             =   615
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Comissão Atual"
      Height          =   2115
      Left            =   105
      TabIndex        =   28
      Top             =   1680
      Width           =   9285
      Begin VB.Label Label1 
         Caption         =   "Promotor:"
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
         Index           =   4
         Left            =   810
         TabIndex        =   79
         Top             =   1725
         Width           =   840
      End
      Begin VB.Label Promotor 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1710
         TabIndex        =   78
         Top             =   1680
         Width           =   2925
      End
      Begin VB.Label VlrComiAge 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   7770
         TabIndex        =   56
         Top             =   195
         Width           =   1425
      End
      Begin VB.Label PercComiAge 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   5850
         TabIndex        =   54
         Top             =   195
         Width           =   1065
      End
      Begin VB.Label Agencia 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1710
         TabIndex        =   53
         Top             =   195
         Width           =   2925
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
         Height          =   330
         Index           =   12
         Left            =   7125
         TabIndex        =   57
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "% Comissão:"
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
         Index           =   11
         Left            =   4785
         TabIndex        =   55
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label1 
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
         Height          =   330
         Index           =   10
         Left            =   960
         TabIndex        =   52
         Top             =   240
         Width           =   750
      End
      Begin VB.Label VlrComiEmi 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   7770
         TabIndex        =   47
         Top             =   1305
         Width           =   1425
      End
      Begin VB.Label VlrComiCor 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   7770
         TabIndex        =   45
         Top             =   945
         Width           =   1425
      End
      Begin VB.Label VlrComiRep 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   7770
         TabIndex        =   43
         Top             =   570
         Width           =   1425
      End
      Begin VB.Label PercComiEmi 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   5850
         TabIndex        =   41
         Top             =   1305
         Width           =   1065
      End
      Begin VB.Label PercComiCor 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   5850
         TabIndex        =   37
         Top             =   945
         Width           =   1065
      End
      Begin VB.Label PercComiRep 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   5850
         TabIndex        =   31
         Top             =   570
         Width           =   1065
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
         Height          =   330
         Index           =   9
         Left            =   7125
         TabIndex        =   48
         Top             =   1350
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
         Height          =   330
         Index           =   8
         Left            =   7125
         TabIndex        =   46
         Top             =   975
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
         Height          =   330
         Index           =   7
         Left            =   7125
         TabIndex        =   44
         Top             =   600
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "% Comissão:"
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
         Index           =   61
         Left            =   4785
         TabIndex        =   42
         Top             =   1350
         Width           =   1140
      End
      Begin VB.Label Emissor 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1710
         TabIndex        =   40
         Top             =   1305
         Width           =   2925
      End
      Begin VB.Label Label1 
         Caption         =   "Emissor:"
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
         Index           =   50
         Left            =   900
         TabIndex        =   39
         Top             =   1350
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "% Comissão:"
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
         Index           =   59
         Left            =   4785
         TabIndex        =   38
         Top             =   975
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Correntista:"
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
         Index           =   60
         Left            =   615
         TabIndex        =   36
         Top             =   1005
         Width           =   1080
      End
      Begin VB.Label Correntista 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1710
         TabIndex        =   35
         Top             =   945
         Width           =   2925
      End
      Begin VB.Label Label1 
         Caption         =   "% Comissão:"
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
         Index           =   58
         Left            =   4785
         TabIndex        =   32
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Representante:"
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
         Index           =   57
         Left            =   285
         TabIndex        =   30
         Top             =   615
         Width           =   1410
      End
      Begin VB.Label Representante 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1710
         TabIndex        =   29
         Top             =   570
         Width           =   2925
      End
   End
   Begin MSMask.MaskEdBox TipoVou 
      Height          =   315
      Left            =   735
      TabIndex        =   0
      Top             =   150
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   1
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox SerieVou 
      Height          =   315
      Left            =   1770
      TabIndex        =   1
      Top             =   150
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   1
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NumeroVou 
      Height          =   315
      Left            =   3045
      TabIndex        =   2
      Top             =   150
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   6
      Mask            =   "######"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
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
      Height          =   330
      Index           =   0
      Left            =   240
      TabIndex        =   51
      Top             =   195
      Width           =   435
   End
   Begin VB.Label Label1 
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
      Height          =   330
      Index           =   2
      Left            =   1200
      TabIndex        =   50
      Top             =   195
      Width           =   480
   End
   Begin VB.Label LabelNumVou 
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
      Height          =   330
      Left            =   2265
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   49
      Top             =   195
      Width           =   750
   End
End
Attribute VB_Name = "TRPVouComi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim lNumvouAnt As Long
Dim sSerieAnt As String
Dim sTipoAnt As String

Private WithEvents objEventoEmissor As AdmEvento
Attribute objEventoEmissor.VB_VarHelpID = -1
Private WithEvents objEventoRepresentante As AdmEvento
Attribute objEventoRepresentante.VB_VarHelpID = -1
Private WithEvents objEventoCorrentista As AdmEvento
Attribute objEventoCorrentista.VB_VarHelpID = -1
Private WithEvents objEventoAgencia As AdmEvento
Attribute objEventoAgencia.VB_VarHelpID = -1
Private WithEvents objEventoPromotor As AdmEvento
Attribute objEventoPromotor.VB_VarHelpID = -1
Private WithEvents objEventoVoucher As AdmEvento
Attribute objEventoVoucher.VB_VarHelpID = -1

Dim iAlterado As Integer

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Alterações no comissionamento do voucher"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TRPVouComi"

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

Private Sub LabelNumVou_Click()
    Call BotaoVou_Click
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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoEmissor = Nothing
    Set objEventoRepresentante = Nothing
    Set objEventoCorrentista = Nothing
    Set objEventoAgencia = Nothing
    Set objEventoPromotor = Nothing
    Set objEventoVoucher = Nothing

    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198028)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iAlterado = 0
    
    Set objEventoEmissor = New AdmEvento
    Set objEventoRepresentante = New AdmEvento
    Set objEventoCorrentista = New AdmEvento
    Set objEventoAgencia = New AdmEvento
    Set objEventoPromotor = New AdmEvento
    Set objEventoVoucher = New AdmEvento

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198029)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objVou As ClassTRPVouchers) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objVou Is Nothing) Then

        lErro = Traz_TRPVouchers_Tela(objVou)
        If lErro <> SUCESSO Then gError 198030

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 198030

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198031)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(ByVal objVou As ClassTRPVouchers) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objVou.lNumVou = StrParaLong(NumeroVou.Text)
    objVou.sSerie = SerieVou.Text
    objVou.sTipVou = TipoVou.Text
    
    lErro = CF("TRPVouchers_Le", objVou)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 198200
    
    If lErro <> SUCESSO Then gError 198201
    
    If Len(Trim(PercComiAgeNovo.Text)) <> 0 Then
        objVou.lClienteComissao = LCodigo_Extrai(AgenciaNovo.Caption)
        objVou.dComissaoAg = StrParaDbl(PercComiAgeNovo.Text) / 100
    End If
    
    If LCodigo_Extrai(RepresentanteNovo.Text) <> 0 Then
        objVou.lRepresentante = LCodigo_Extrai(RepresentanteNovo.Text)
        objVou.dComissaoRep = StrParaDbl(PercComiRepNovo.Text) / 100
    End If
    
    If LCodigo_Extrai(CorrentistaNovo.Text) <> 0 Then
        objVou.lCorrentista = LCodigo_Extrai(CorrentistaNovo.Text)
        objVou.dComissaoCorr = StrParaDbl(PercComiCorNovo.Text) / 100
    End If
    
    If LCodigo_Extrai(EmissorNovo.Text) <> 0 Then
        objVou.lEmissor = LCodigo_Extrai(EmissorNovo.Text)
        objVou.dComissaoEmissor = StrParaDbl(PercComiEmiNovo.Text) / 100
    End If

    objVou.iPromotor = Codigo_Extrai(PromotorNovo.Text)
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 198200
        
        Case 198201
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_NAO_CADASTRADO", gErr, objVou.lNumVou, objVou.sSerie, objVou.sTipVou)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198032)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objVou As New ClassTRPVouchers

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TRPVouchers"

    objVou.lNumVou = StrParaLong(NumeroVou.Text)
    objVou.sSerie = SerieVou.Text
    objVou.sTipVou = TipoVou.Text

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumVou", objVou.lNumVou, 0, "NumVou"
    colCampoValor.Add "Serie", objVou.sSerie, STRING_TRP_OCR_SERIE, "Serie"
    colCampoValor.Add "TipVou", objVou.sTipVou, STRING_TRP_OCR_TIPOVOU, "TipVou"

    colSelecao.Add "Status", OP_DIFERENTE, STATUS_TRP_VOU_CANCELADO
    colSelecao.Add "GeraComissao", OP_IGUAL, DESMARCADO

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 198033

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198034)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objVou As New ClassTRPVouchers

On Error GoTo Erro_Tela_Preenche

    objVou.lNumVou = colCampoValor.Item("NumVou").vValor
    objVou.sSerie = colCampoValor.Item("Serie").vValor
    objVou.sTipVou = colCampoValor.Item("TipVou").vValor
    
    If objVou.lNumVou <> 0 Then

        lErro = Traz_TRPVouchers_Tela(objVou)
        If lErro <> SUCESSO Then gError 198035

    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 198035

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198036)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objVou As New ClassTRPVouchers
Dim bSimulaImport As Boolean

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(NumeroVou.Text)) = 0 Then gError 198037
    '#####################

    'Preenche o objTRPTiposOcorrencia
    lErro = Move_Tela_Memoria(objVou)
    If lErro <> SUCESSO Then gError 198038

    lErro = Trata_Alteracao(objVou, objVou.sTipVou, objVou.sSerie, objVou.lNumVou)
    If lErro <> SUCESSO Then gError 198039
    
    If Import.Value = vbChecked Then
        bSimulaImport = True
    Else
        bSimulaImport = False
    End If

    'Grava o/a TRPTiposOcorrencia no Banco de Dados
    lErro = CF("TRPVouComi_Grava", objVou)
    If lErro <> SUCESSO Then gError 198040
    
    'Limpa Tela
    Call Limpa_Tela_TRPVouchers
    
    Call Traz_TRPVouchers_Tela(objVou)

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 198037
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TRPTIPOSOCORRENCIA_NAO_PREENCHIDO", gErr)

        Case 198038, 198039, 198040

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198041)

    End Select

    Exit Function

End Function

Function Limpa_Tela_TRPVouchers() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_TRPVouchers

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Cartao.Value = vbUnchecked
    Antc.Value = vbUnchecked
    CartaoNovo.Value = Cartao.Value
    AntcNovo.Value = Antc.Value
    Import.Value = vbUnchecked
    
    DataEmissaoVou.Caption = ""
    ValorBrutoVou.Caption = ""
    Produto.Caption = ""
    Destino.Caption = ""
    Representante.Caption = ""
    Correntista.Caption = ""
    Emissor.Caption = ""
    Agencia.Caption = ""
    PercComiRep.Caption = ""
    PercComiCor.Caption = ""
    PercComiEmi.Caption = ""
    PercComiAge.Caption = ""
    VlrComiRep.Caption = ""
    VlrComiCor.Caption = ""
    VlrComiEmi.Caption = ""
    VlrComiAge.Caption = ""
    Promotor.Caption = ""
    AgenciaNovo.Caption = ""

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    iAlterado = 0

    Limpa_Tela_TRPVouchers = SUCESSO

    Exit Function

Erro_Limpa_Tela_TRPVouchers:

    Limpa_Tela_TRPVouchers = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198042)

    End Select

    Exit Function

End Function

Function Traz_TRPVouchers_Tela(ByVal objVou As ClassTRPVouchers) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_TRPVouchers_Tela

    Call Limpa_Tela_TRPVouchers
    
    NumeroVou.PromptInclude = False
    NumeroVou.Text = CStr(objVou.lNumVou)
    NumeroVou.PromptInclude = True
    
    SerieVou.Text = objVou.sSerie
    TipoVou.Text = objVou.sTipVou
    
    Call TrazerVou_Click

    iAlterado = 0

    Traz_TRPVouchers_Tela = SUCESSO

    Exit Function

Erro_Traz_TRPVouchers_Tela:

    Traz_TRPVouchers_Tela = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198044)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 198045
'
'    'Limpa Tela
'    Call Limpa_Tela_TRPVouchers

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 198045

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198046)

    End Select

    Exit Sub

End Sub

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198047)

    End Select

    Exit Sub

End Sub


Private Sub AgenciaNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub AntcNovo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub BrutoNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CartaoNovo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CorrentistaNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DestinoNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub EmissorNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PercComiAgeNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PercComiCorNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PercComiEmiNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PercComiRepNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ProdutoNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PromotorNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub RepresentanteNovo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub SerieVou_Validate(Cancel As Boolean)
    If SerieVou.Text <> sSerieAnt Then
        Call Limpa_Vou
    End If
End Sub

Private Sub TipoVou_Validate(Cancel As Boolean)
    If TipoVou.Text <> sTipoAnt Then
        Call Limpa_Vou
    End If
End Sub

Private Sub TipoVou_Change()
    iAlterado = REGISTRO_ALTERADO
    If Len(Trim(TipoVou.ClipText)) > 0 Then
        If SerieVou.Visible Then SerieVou.SetFocus
    End If
End Sub

Private Sub SerieVou_Change()
    iAlterado = REGISTRO_ALTERADO
    If Len(Trim(SerieVou.ClipText)) > 0 Then
        If NumeroVou.Visible Then NumeroVou.SetFocus
    End If
End Sub

Private Sub TipoVou_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub SerieVou_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub NumeroVou_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NumeroVou_Validate(Cancel As Boolean)
    If StrParaLong(NumeroVou.Text) <> lNumvouAnt Then
        Call Limpa_Vou
    End If
End Sub

Private Sub BotaoTrazerVou_Click()
    Call TrazerVou_Click
End Sub

Private Sub TrazerVou_Click()

Dim lErro As Long
Dim objVoucher As New ClassTRPVouchers
Dim objVouInfo As New ClassTRPVoucherInfo
Dim objCliente As New ClassCliente
Dim objForn As New ClassFornecedor
Dim objVendedor As New ClassVendedor
Dim sSiglaDest As String
Dim sDescricaoDest As String

On Error GoTo Erro_TrazerVou_Click

    objVoucher.lNumVou = StrParaLong(NumeroVou.Text)
    objVoucher.sSerie = SerieVou.Text
    objVoucher.sTipVou = TipoVou.Text
    
    lErro = CF("TRPVouchers_Le", objVoucher)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194421
    
    If lErro <> SUCESSO Then gError 194425
    
    If objVoucher.iStatus = STATUS_TRP_VOU_CANCELADO Then gError 194422
    
    'Busca o Cliente no BD
    If objVoucher.lRepresentante <> 0 Then
        Cliente.Text = objVoucher.lRepresentante
        lErro = TP_Cliente_Le2(Cliente, objCliente)
        If lErro <> SUCESSO Then gError 190204
        Representante.Caption = Cliente.Text
    End If
    
    'Busca o Cliente no BD
    If objVoucher.lCorrentista <> 0 Then
        Cliente.Text = objVoucher.lCorrentista
        lErro = TP_Cliente_Le2(Cliente, objCliente)
        If lErro <> SUCESSO Then gError 190204
        Correntista.Caption = Cliente.Text
    End If
    
    'Busca o Cliente no BD
    If objVoucher.lClienteVou <> 0 Then
        Cliente.Text = objVoucher.lClienteVou
        lErro = TP_Cliente_Le2(Cliente, objCliente)
        If lErro <> SUCESSO Then gError 190204
        Agencia.Caption = Cliente.Text
        AgenciaNovo.Caption = Cliente.Text
    End If
    
    'Busca o Cliente no BD
    If objVoucher.lEmissor <> 0 Then
        Cliente.Text = objVoucher.lEmissor
        lErro = TP_Fornecedor_Le2(Cliente, objForn)
        If lErro <> SUCESSO Then gError 190204
        Emissor.Caption = Cliente.Text
    End If
    
    'Busca o Cliente no BD
    If objVoucher.iPromotor <> 0 Then
        Cliente.Text = objVoucher.iPromotor
        lErro = TP_Vendedor_Le2(Cliente, objVendedor)
        If lErro <> SUCESSO Then gError 190204
        Promotor.Caption = Cliente.Text
    End If
    
    DataEmissaoVou.Caption = Format(objVoucher.dtData, "dd/mm/yyyy")
    ValorBrutoVou.Caption = Format(objVoucher.dValorBruto, "STANDARD")
    
    Produto.Caption = objVoucher.sProduto
    
    PercComiCor.Caption = Format(objVoucher.dComissaoCorr, "PERCENT")
    PercComiAge.Caption = Format(objVoucher.dComissaoAg, "PERCENT")
    PercComiRep.Caption = Format(objVoucher.dComissaoRep, "PERCENT")
    PercComiEmi.Caption = Format(objVoucher.dComissaoEmissor, "PERCENT")
    
    VlrComiCor.Caption = Format(objVoucher.dValorBruto * objVoucher.dComissaoCorr, "STANDARD")
    VlrComiAge.Caption = Format(objVoucher.dValorBruto * objVoucher.dComissaoAg, "STANDARD")
    VlrComiRep.Caption = Format(objVoucher.dValorBruto * objVoucher.dComissaoRep, "STANDARD")
    VlrComiEmi.Caption = Format(objVoucher.dValorBruto * objVoucher.dComissaoEmissor, "STANDARD")
    
    If objVoucher.iCartao = MARCADO Then
        Cartao.Value = vbChecked
    Else
        Cartao.Value = vbUnchecked
    End If
    
    If objVoucher.iDiasAntc = MARCADO Then
        Antc.Value = vbChecked
    Else
        Antc.Value = vbUnchecked
    End If
    
    CartaoNovo.Value = Cartao.Value
    AntcNovo.Value = Antc.Value
    
    lErro = CF("TRPDestino_Le", objVoucher.iDestino, sSiglaDest, sDescricaoDest)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192646
    
    Destino.Caption = objVoucher.iDestino & SEPARADOR & sDescricaoDest
    
    lNumvouAnt = objVoucher.lNumVou
    sSerieAnt = objVoucher.sSerie
    sTipoAnt = objVoucher.sTipVou

    Exit Sub

Erro_TrazerVou_Click:

    Select Case gErr
    
        Case 194421, 194423, 192646
        
        Case 194422
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_JA_CANCELADO", gErr)
            
        Case 194424
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_SEM_DADOS_SIGAV", gErr)
            
        Case 194425
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_NAO_CADASTRADO", gErr, objVoucher.lNumVou, objVoucher.sSerie, objVoucher.sTipVou)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194426)

    End Select

    Exit Sub
    
End Sub

Private Sub Limpa_Vou()

Dim lNumVou As Long
Dim sSerie As String
Dim sTipVou As String

    lNumVou = StrParaLong(NumeroVou.Text)
    sSerie = SerieVou.Text
    sTipVou = TipoVou.Text

    If lNumVou <> lNumvouAnt Or sSerie <> sSerieAnt Or sTipVou <> sTipoAnt Then
    
        Call Limpa_Tela_TRPVouchers
        
        NumeroVou.PromptInclude = False
        NumeroVou.Text = CStr(lNumVou)
        NumeroVou.PromptInclude = True
        
        SerieVou.Text = sSerie
        TipoVou.Text = sTipVou
        
    End If

End Sub

Private Sub RepresentanteNovo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_RepresentanteNovo_Validate

    If Len(Trim(RepresentanteNovo.Text)) > 0 Then
    
        RepresentanteNovo.Text = LCodigo_Extrai(RepresentanteNovo.Text)

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(RepresentanteNovo, objCliente)
        If lErro <> SUCESSO Then gError 195843
        
    End If
    
    Exit Sub

Erro_RepresentanteNovo_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 195843

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195845)
    
    End Select

End Sub

Private Sub CorrentistaNovo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_CorrentistaNovo_Validate

    If Len(Trim(CorrentistaNovo.Text)) > 0 Then
    
        CorrentistaNovo.Text = LCodigo_Extrai(CorrentistaNovo.Text)

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(CorrentistaNovo, objCliente)
        If lErro <> SUCESSO Then gError 195843
        
    End If
    
    Exit Sub

Erro_CorrentistaNovo_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 195843

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195845)
    
    End Select

End Sub
'
'Private Sub AgenciaNovo_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objCliente As New ClassCliente
'Dim objVoucher As New ClassTRPVouchers
'
'On Error GoTo Erro_AgenciaNovo_Validate
'
'    If Len(Trim(AgenciaNovo.Text)) > 0 Then
'
'        AgenciaNovo.Text = LCodigo_Extrai(AgenciaNovo.Text)
'
'        'Tenta ler o Vendedor (NomeReduzido ou Código)
'        lErro = TP_Cliente_Le2(AgenciaNovo, objCliente)
'        If lErro <> SUCESSO Then gError 195843
'
'        objVoucher.lNumVou = StrParaLong(NumeroVou.Text)
'        objVoucher.sSerie = SerieVou.Text
'        objVoucher.sTipVou = TipoVou.Text
'
'        lErro = CF("TRPVouchers_Le", objVoucher)
'        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 195843
'
'        If lErro = SUCESSO Then
'
'            lErro = CF("TRPVou_Le_Dados_Comis_Cliente", objVoucher)
'            If lErro <> SUCESSO Then gError 195843
'
'            If LCodigo_Extrai(Representante.Caption) <> objVoucher.lRepresentante And objVoucher.lRepresentante <> 0 Then
'                RepresentanteNovo.Text = objVoucher.lRepresentante
'                Call RepresentanteNovo_Validate(bSGECancelDummy)
'                PercComiRepNovo.Text = objVoucher.dComissaoRep * 100
'                Call PercComiRepNovo_Validate(bSGECancelDummy)
'            End If
'
'            If LCodigo_Extrai(Correntista.Caption) <> objVoucher.lCorrentista And objVoucher.lCorrentista <> 0 Then
'                CorrentistaNovo.Text = objVoucher.lCorrentista
'                Call CorrentistaNovo_Validate(bSGECancelDummy)
'                PercComiCorNovo.Text = objVoucher.dComissaoCorr * 100
'                Call PercComiCorNovo_Validate(bSGECancelDummy)
'            End If
'
'            If LCodigo_Extrai(Emissor.Caption) <> objVoucher.lEmissor And objVoucher.lEmissor <> 0 Then
'                EmissorNovo.Text = objVoucher.lEmissor
'                Call EmissorNovo_Validate(bSGECancelDummy)
'                PercComiEmiNovo.Text = objVoucher.dComissaoEmissor * 100
'                Call PercComiEmiNovo_Validate(bSGECancelDummy)
'            End If
'
'            PercComiAgeNovo.Text = objVoucher.dComissaoAg * 100
'            Call PercComiAgeNovo_Validate(bSGECancelDummy)
'
'        End If
'
'    End If
'
'    Exit Sub
'
'Erro_AgenciaNovo_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 195843
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195845)
'
'    End Select
'
'End Sub

Private Sub EmissorNovo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objForn As New ClassFornecedor

On Error GoTo Erro_EmissorNovo_Validate

    If Len(Trim(EmissorNovo.Text)) > 0 Then
    
        EmissorNovo.Text = LCodigo_Extrai(EmissorNovo.Text)

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Fornecedor_Le2(EmissorNovo, objForn)
        If lErro <> SUCESSO Then gError 195843
        
    End If
    
    Exit Sub

Erro_EmissorNovo_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 195843

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195845)
    
    End Select

End Sub

Private Sub PromotorNovo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_PromotorNovo_Validate

    If Len(Trim(PromotorNovo.Text)) > 0 Then
    
        PromotorNovo.Text = LCodigo_Extrai(PromotorNovo.Text)

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(PromotorNovo, objVendedor)
        If lErro <> SUCESSO Then gError 195843
        
    End If
    
    Exit Sub

Erro_PromotorNovo_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 195843

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195845)
    
    End Select

End Sub

Private Sub PercComiAgeNovo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercComiAgeNovo_Validate

    'Verifica se foi preenchido a Comissao de Venda
    If Len(Trim(PercComiAgeNovo.Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Porcentagem_Critica(PercComiAgeNovo.Text)
    If lErro <> SUCESSO Then Error 195853

    'Formata
    PercComiAgeNovo.Text = Format(PercComiAgeNovo.Text, "Fixed")

    Exit Sub

Erro_PercComiAgeNovo_Validate:

    Cancel = True

    Select Case gErr

        Case 195853
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195854)
    
    End Select

    Exit Sub

End Sub

Private Sub PercComiCorNovo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercComiCorNovo_Validate

    'Verifica se foi preenchido a Comissao de Venda
    If Len(Trim(PercComiCorNovo.Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Porcentagem_Critica(PercComiCorNovo.Text)
    If lErro <> SUCESSO Then Error 195853

    'Formata
    PercComiCorNovo.Text = Format(PercComiCorNovo.Text, "Fixed")

    Exit Sub

Erro_PercComiCorNovo_Validate:

    Cancel = True

    Select Case gErr

        Case 195853
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195854)
    
    End Select

    Exit Sub

End Sub

Private Sub PercComiRepNovo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercComiRepNovo_Validate

    'Verifica se foi preenchido a Comissao de Venda
    If Len(Trim(PercComiRepNovo.Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Porcentagem_Critica(PercComiRepNovo.Text)
    If lErro <> SUCESSO Then Error 195853

    'Formata
    PercComiRepNovo.Text = Format(PercComiRepNovo.Text, "Fixed")

    Exit Sub

Erro_PercComiRepNovo_Validate:

    Cancel = True

    Select Case gErr

        Case 195853
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195854)
    
    End Select

    Exit Sub

End Sub

Private Sub PercComiEmiNovo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercComiEmiNovo_Validate

    'Verifica se foi preenchido a Comissao de Venda
    If Len(Trim(PercComiEmiNovo.Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Porcentagem_Critica(PercComiEmiNovo.Text)
    If lErro <> SUCESSO Then Error 195853

    'Formata
    PercComiEmiNovo.Text = Format(PercComiEmiNovo.Text, "Fixed")

    Exit Sub

Erro_PercComiEmiNovo_Validate:

    Cancel = True

    Select Case gErr

        Case 195853
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195854)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoHist_Click()

Dim lErro As Long
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoHist_Click

    colSelecao.Add StrParaLong(NumeroVou.Text)
    colSelecao.Add TipoVou.Text
    colSelecao.Add SerieVou.Text

    Call Chama_Tela("TRPVoucherInfoLista", colSelecao, Nothing, Nothing, "NumVou= ? AND TipVou = ? AND Serie = ?")

    Exit Sub

Erro_BotaoHist_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190226)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluirComissao_Click()

Dim lErro As Long
Dim objVou As New ClassTRPVouchers

On Error GoTo Erro_BotaoExcluirComissao_Click

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(NumeroVou.Text)) = 0 Then gError 198216
    '#####################

    'Preenche o objTRPTiposOcorrencia
    lErro = Move_Tela_Memoria(objVou)
    If lErro <> SUCESSO Then gError 198217

    'Grava o/a TRPTiposOcorrencia no Banco de Dados
    lErro = CF("TRPVoucher_Exclui_Comissao", objVou)
    If lErro <> SUCESSO Then gError 198218

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluirComissao_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 198216
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TRPTIPOSOCORRENCIA_NAO_PREENCHIDO", gErr)
            
        Case 198217, 198218

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198219)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoPrimeiraComissao_Click()

Dim lErro As Long
Dim objVou As New ClassTRPVouchers

On Error GoTo Erro_BotaoPrimeiraComissao_Click

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(NumeroVou.Text)) = 0 Then gError 198216
    '#####################

    'Preenche o objTRPTiposOcorrencia
    lErro = Move_Tela_Memoria(objVou)
    If lErro <> SUCESSO Then gError 198217

    'Grava o/a TRPTiposOcorrencia no Banco de Dados
    lErro = CF("TRPVouComi_Grava", objVou, True)
    If lErro <> SUCESSO Then gError 198218

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoPrimeiraComissao_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 198216
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TRPTIPOSOCORRENCIA_NAO_PREENCHIDO", gErr)
            
        Case 198217, 198218

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198219)

    End Select

    Exit Sub
    
End Sub

Public Sub LabelRepresentante_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
Dim sNomeBrowse As String

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(RepresentanteNovo.Text)) > 0 Then
        objCliente.lCodigo = LCodigo_Extrai(RepresentanteNovo.Text)
        objCliente.sNomeReduzido = RepresentanteNovo.Text
    End If

    sNomeBrowse = "ClientesLista"

    Call CF("Cliente_Obtem_NomeBrowse", sNomeBrowse)

    'Chama Tela ClienteLista
    Call Chama_Tela(sNomeBrowse, colSelecao, objCliente, objEventoRepresentante)

End Sub

Public Sub objEventoRepresentante_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    RepresentanteNovo.Text = objCliente.lCodigo
    Call RepresentanteNovo_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Public Sub LabelCorrentista_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
Dim sNomeBrowse As String

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(CorrentistaNovo.Text)) > 0 Then
        objCliente.lCodigo = LCodigo_Extrai(CorrentistaNovo.Text)
        objCliente.sNomeReduzido = CorrentistaNovo.Text
    End If

    sNomeBrowse = "ClientesLista"

    Call CF("Cliente_Obtem_NomeBrowse", sNomeBrowse)

    'Chama Tela ClienteLista
    Call Chama_Tela(sNomeBrowse, colSelecao, objCliente, objEventoCorrentista)

End Sub

Public Sub objEventoCorrentista_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    CorrentistaNovo.Text = objCliente.lCodigo
    Call CorrentistaNovo_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Public Sub LabelAgencia_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
Dim sNomeBrowse As String

'    'Preenche NomeReduzido com o cliente da tela
'    If Len(Trim(AgenciaNovo.Caption)) > 0 Then
'        objCliente.lCodigo = LCodigo_Extrai(AgenciaNovo.Text)
'        objCliente.sNomeReduzido = AgenciaNovo.Text
'    End If
'
'    sNomeBrowse = "ClientesLista"
'
'    Call CF("Cliente_Obtem_NomeBrowse", sNomeBrowse)
'
'    'Chama Tela ClienteLista
'    Call Chama_Tela(sNomeBrowse, colSelecao, objCliente, objEventoAgencia)

End Sub

Public Sub objEventoAgencia_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

'    Set objCliente = obj1
'
'    AgenciaNovo.Caption = objCliente.lCodigo
'    Call AgenciaNovo_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Public Sub LabelEmissor_Click()

Dim objForn As New ClassFornecedor
Dim colSelecao As Collection

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(EmissorNovo.Text)) > 0 Then
        objForn.lCodigo = LCodigo_Extrai(EmissorNovo.Text)
        objForn.sNomeReduzido = EmissorNovo.Text
    End If


    'Chama Tela ClienteLista
    Call Chama_Tela("FornecedorLista", colSelecao, objForn, objEventoEmissor)

End Sub

Public Sub objEventoEmissor_evSelecao(obj1 As Object)

Dim objForn As ClassFornecedor
Dim bCancel As Boolean

    Set objForn = obj1

    EmissorNovo.Text = objForn.lCodigo
    Call EmissorNovo_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Public Sub LabelPromotor_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(PromotorNovo.Text)) > 0 Then
        objVendedor.iCodigo = Codigo_Extrai(PromotorNovo.Text)
        objVendedor.sNomeReduzido = PromotorNovo.Text
    End If

    'Chama Tela ClienteLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoPromotor)

End Sub

Public Sub objEventoPromotor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor
Dim bCancel As Boolean

    Set objVendedor = obj1

    PromotorNovo.Text = objVendedor.iCodigo
    Call PromotorNovo_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is AgenciaNovo Then Call LabelAgencia_Click
        If Me.ActiveControl Is RepresentanteNovo Then Call LabelRepresentante_Click
        If Me.ActiveControl Is CorrentistaNovo Then Call LabelCorrentista_Click
        If Me.ActiveControl Is EmissorNovo Then Call LabelEmissor_Click
        If Me.ActiveControl Is PromotorNovo Then Call LabelPromotor_Click
    
    End If
    
End Sub

Private Sub BotaoVou_Click()

Dim lErro As Long
Dim objVoucher As New ClassTRPVouchers
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoVou_Click

    objVoucher.lNumVou = StrParaLong(NumeroVou.Text)
    objVoucher.sSerie = SerieVou.Text
    objVoucher.sTipVou = TipoVou.Text

    Call Chama_Tela("TRPVoucherRapidoLista", colSelecao, objVoucher, objEventoVoucher)

    Exit Sub

Erro_BotaoVou_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190160)

    End Select

    Exit Sub

End Sub

Private Sub objEventoVoucher_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objVou As ClassTRPVouchers

On Error GoTo Erro_objEventoVoucher_evSelecao

    Set objVou = obj1

    'Mostra os dados do TRPVouchers na tela
    lErro = Traz_TRPVouchers_Tela(objVou)
    If lErro <> SUCESSO Then gError 190909

    Me.Show

    Exit Sub

Erro_objEventoVoucher_evSelecao:

    Select Case gErr

        Case 190909

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143951)

    End Select

    Exit Sub

End Sub
