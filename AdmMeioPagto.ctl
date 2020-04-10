VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl AdmMeioPagto 
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9255
   KeyPreview      =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   9255
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3390
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   975
      Width           =   8970
      Begin VB.CheckBox Ativo 
         Caption         =   "Ativo"
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
         Left            =   2430
         TabIndex        =   5
         Top             =   120
         Value           =   1  'Checked
         Width           =   810
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   300
         Left            =   1845
         Picture         =   "AdmMeioPagto.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Numeração Automática"
         Top             =   90
         Width           =   315
      End
      Begin VB.ComboBox TipoMeioPagto 
         Height          =   315
         ItemData        =   "AdmMeioPagto.ctx":00EA
         Left            =   1425
         List            =   "AdmMeioPagto.ctx":00EC
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Tipo do meio de pagamento"
         Top             =   1455
         Width           =   2880
      End
      Begin VB.CheckBox CheckNaoTitRec 
         Caption         =   "Não gera título a receber"
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
         Left            =   405
         TabIndex        =   21
         ToolTipText     =   "Indica se esse meio de pagamento deve gerar um título a receber"
         Top             =   3000
         Width           =   2580
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1455
         TabIndex        =   3
         ToolTipText     =   "Código do meio de pagamento"
         Top             =   75
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin VB.ComboBox ContaCorrenteInterna 
         Height          =   315
         Left            =   2835
         TabIndex        =   20
         ToolTipText     =   "Conta corrente onde a administradora deve fazer o depósito"
         Top             =   2490
         Width           =   2850
      End
      Begin VB.ListBox Administradoras 
         Height          =   2985
         ItemData        =   "AdmMeioPagto.ctx":00EE
         Left            =   6255
         List            =   "AdmMeioPagto.ctx":00F0
         Sorted          =   -1  'True
         TabIndex        =   23
         ToolTipText     =   "Meios de pagamento cadastrados"
         Top             =   315
         Width           =   2670
      End
      Begin VB.ComboBox Rede 
         Height          =   315
         ItemData        =   "AdmMeioPagto.ctx":00F2
         Left            =   3990
         List            =   "AdmMeioPagto.ctx":00F4
         TabIndex        =   7
         ToolTipText     =   "Rede a qual a administradora está vinculada"
         Top             =   75
         Width           =   1725
      End
      Begin MSMask.MaskEdBox TaxaVista 
         Height          =   285
         Left            =   1425
         TabIndex        =   11
         ToolTipText     =   "Taxa cobrada por pagamentos à vista"
         Top             =   1020
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Nome 
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         ToolTipText     =   "Nome do meio de pagamento"
         Top             =   540
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DefasagemPagtoVista 
         Height          =   285
         Left            =   2415
         TabIndex        =   17
         ToolTipText     =   "Em quantos dias a administradora faz o pagamento"
         Top             =   1980
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TaxaPrazo 
         Height          =   285
         Left            =   4950
         TabIndex        =   13
         ToolTipText     =   "Taxa cobrada por pagamentos à prazo"
         Top             =   990
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   6
         Left            =   855
         TabIndex        =   14
         ToolTipText     =   "Tipo do meio de pagamento"
         Top             =   1500
         Width           =   450
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   660
         TabIndex        =   2
         ToolTipText     =   "Código do meio de pagamento"
         Top             =   135
         Width           =   660
      End
      Begin VB.Label LblConta 
         AutoSize        =   -1  'True
         Caption         =   "Conta Corrente p/ Depósito:"
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
         Left            =   285
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   19
         ToolTipText     =   "Conta corrente onde a administradora deve fazer o depósito"
         Top             =   2535
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "dias"
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
         Index           =   8
         Left            =   2925
         TabIndex        =   18
         ToolTipText     =   "Em quantos dias a administradora faz o pagamento"
         Top             =   2010
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Demora Pagto a Vista:"
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
         Left            =   405
         TabIndex        =   16
         ToolTipText     =   "Em quantos dias a administradora faz o pagamento"
         Top             =   2010
         Width           =   1920
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Meios de Pagamento"
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
         Left            =   6225
         TabIndex        =   22
         ToolTipText     =   "Meios de Pagamento"
         Top             =   75
         Width           =   1785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
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
         Index           =   3
         Left            =   765
         TabIndex        =   8
         ToolTipText     =   "Nome do meio de pagamento"
         Top             =   600
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Taxa a Vista:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   10
         ToolTipText     =   "Taxa cobrada por pagamentos à vista"
         Top             =   1050
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Taxa a Prazo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   5
         Left            =   3570
         TabIndex        =   12
         ToolTipText     =   "Taxa cobrada por pagamentos à prazo"
         Top             =   1050
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Rede:"
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
         Left            =   3360
         TabIndex        =   6
         ToolTipText     =   "Rede a qual o meio de pagamento está vinculado"
         Top             =   135
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3450
      Index           =   2
      Left            =   270
      TabIndex        =   24
      Top             =   915
      Visible         =   0   'False
      Width           =   8880
      Begin VB.CheckBox PreDatado 
         Caption         =   "Pré Datado"
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
         TabIndex        =   86
         Top             =   225
         Value           =   1  'Checked
         Width           =   1410
      End
      Begin VB.PictureBox Picture2 
         Height          =   540
         Left            =   7440
         ScaleHeight     =   480
         ScaleWidth      =   1095
         TabIndex        =   83
         Top             =   75
         Width           =   1155
         Begin VB.CommandButton BotaoGravarParc 
            Height          =   360
            Left            =   75
            Picture         =   "AdmMeioPagto.ctx":00F6
            Style           =   1  'Graphical
            TabIndex        =   85
            ToolTipText     =   "Gravar"
            Top             =   75
            Width           =   420
         End
         Begin VB.CommandButton BotaoExcluirParc 
            Height          =   360
            Left            =   600
            Picture         =   "AdmMeioPagto.ctx":0250
            Style           =   1  'Graphical
            TabIndex        =   84
            ToolTipText     =   "Excluir"
            Top             =   90
            Width           =   420
         End
      End
      Begin VB.CheckBox AtivoParc 
         Caption         =   "Ativo"
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
         Left            =   4605
         TabIndex        =   82
         Top             =   210
         Value           =   1  'Checked
         Width           =   900
      End
      Begin VB.Frame Frame2 
         Caption         =   "Parcelamento"
         Height          =   675
         Left            =   330
         TabIndex        =   47
         ToolTipText     =   "Juros por conta da ..."
         Top             =   2715
         Width           =   3240
         Begin VB.OptionButton ParcelamentoLoja 
            Caption         =   "Loja"
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
            Left            =   420
            TabIndex        =   40
            Top             =   240
            Value           =   -1  'True
            Width           =   870
         End
         Begin VB.OptionButton ParcelamentoAdm 
            Caption         =   "Administradora"
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
            Left            =   1440
            TabIndex        =   41
            Top             =   315
            Width           =   1575
         End
      End
      Begin MSMask.MaskEdBox IntervaloParcela 
         Height          =   270
         Left            =   6750
         TabIndex        =   43
         Top             =   1290
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox PercRecebimento 
         Height          =   270
         Left            =   5445
         TabIndex        =   42
         Top             =   1275
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox ValorMinimo 
         Height          =   300
         Left            =   1740
         TabIndex        =   37
         ToolTipText     =   "Qual o mínimo para aceitar esse parcelamento"
         Top             =   1905
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.ComboBox Parcelamentos 
         Height          =   315
         Left            =   1740
         TabIndex        =   26
         ToolTipText     =   "Formas de Parcelamento"
         Top             =   150
         Width           =   2550
      End
      Begin MSMask.MaskEdBox TaxaParcelamento 
         Height          =   300
         Left            =   1725
         TabIndex        =   36
         ToolTipText     =   "Taxa que a administradora cobra por esse parcelamento"
         Top             =   1500
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         _Version        =   393216
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ParcelasRecebto 
         Height          =   300
         Left            =   1740
         TabIndex        =   30
         ToolTipText     =   "Número de parcelas que a administradora pagará ao Estabalecimento"
         Top             =   1065
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ParcelasPagto 
         Height          =   300
         Left            =   1740
         TabIndex        =   28
         ToolTipText     =   "Número de parcelas que o cliente vai pagar"
         Top             =   615
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desconto 
         Height          =   300
         Left            =   1740
         TabIndex        =   38
         ToolTipText     =   "Desconto oferecido neste parcelamento"
         Top             =   -2430
         Visible         =   0   'False
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         _Version        =   393216
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox JurosParcelas 
         Height          =   300
         Left            =   1740
         TabIndex        =   39
         ToolTipText     =   "% de acréscimo"
         Top             =   2355
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         _Version        =   393216
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridParcelas 
         Height          =   1395
         Left            =   4305
         TabIndex        =   48
         ToolTipText     =   "Reflete a forma que o estabelecimento receberá da administradora"
         Top             =   690
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   2461
         _Version        =   393216
         Rows            =   7
         Cols            =   3
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label PercTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5265
         TabIndex        =   81
         Top             =   3030
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total :"
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
         Index           =   29
         Left            =   4650
         TabIndex        =   80
         Top             =   3060
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desconto :"
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
         Index           =   27
         Left            =   -20000
         TabIndex        =   79
         ToolTipText     =   "Desconto oferecido neste parcelamento"
         Top             =   2600
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Juros Parcelas:"
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
         Index           =   28
         Left            =   375
         TabIndex        =   46
         ToolTipText     =   "% de acréscimo"
         Top             =   2370
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Recebe em "
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
         Index           =   11
         Left            =   690
         TabIndex        =   29
         ToolTipText     =   "Número de parcelas que a administradora pagará ao Estabalecimento"
         Top             =   1095
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Taxa:"
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
         Index           =   13
         Left            =   1200
         TabIndex        =   44
         ToolTipText     =   "Taxa que a administradora cobra por esse parcelamento"
         Top             =   1545
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Parcelas:"
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
         Index           =   10
         Left            =   885
         TabIndex        =   27
         ToolTipText     =   "Número de parcelas que o cliente vai pagar"
         Top             =   675
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor Mínimo:"
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
         Left            =   525
         TabIndex        =   45
         ToolTipText     =   "Qual o mínimo para aceitar esse parcelamento"
         Top             =   1980
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Parcelamentos:"
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
         Index           =   9
         Left            =   390
         TabIndex        =   25
         ToolTipText     =   "Formas de Parcelamento"
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "parcelas"
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
         Left            =   2370
         TabIndex        =   31
         ToolTipText     =   "Número de parcelas que a administradora pagará ao Estabalecimento"
         Top             =   1095
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3450
      Index           =   3
      Left            =   165
      TabIndex        =   49
      Top             =   870
      Visible         =   0   'False
      Width           =   8880
      Begin VB.Frame FrameBanco 
         Caption         =   "Banco para envio dos meios pagamento"
         Height          =   1140
         Left            =   270
         TabIndex        =   73
         Top             =   120
         Visible         =   0   'False
         Width           =   8595
         Begin VB.TextBox Agencia 
            Height          =   300
            Left            =   5655
            MaxLength       =   6
            TabIndex        =   76
            ToolTipText     =   "Agência para a qual os meios de pagamentos devem ser enviados"
            Top             =   510
            Width           =   735
         End
         Begin MSMask.MaskEdBox CodBanco 
            Height          =   300
            Left            =   2100
            TabIndex        =   75
            ToolTipText     =   "Banco para o qual os meios de pagamentos devem ser enviados"
            Top             =   540
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            Mask            =   "999"
            PromptChar      =   " "
         End
         Begin VB.Label LabelBanco 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
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
            Left            =   1380
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   74
            ToolTipText     =   "Banco para o qual os meios de pagamentos devem ser enviados"
            Top             =   570
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Agência:"
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
            Index           =   15
            Left            =   4785
            TabIndex        =   77
            ToolTipText     =   "Agência para a qual os meios de pagamentos devem ser enviados"
            Top             =   570
            Width           =   765
         End
      End
      Begin VB.Frame FrameEndereco 
         Caption         =   "Endereço para envio dos meios de pagamento"
         Height          =   2970
         Left            =   225
         TabIndex        =   50
         Top             =   135
         Width           =   8595
         Begin VB.TextBox Endereco 
            Height          =   315
            Left            =   1305
            MaxLength       =   40
            MultiLine       =   -1  'True
            TabIndex        =   52
            ToolTipText     =   "Endereço para qual os meios de pagamentos devem ser enviados"
            Top             =   450
            Width           =   6345
         End
         Begin VB.ComboBox Pais 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4020
            TabIndex        =   57
            ToolTipText     =   "Endereço para qual os meios de pagamentos devem ser enviados"
            Top             =   1455
            Width           =   1995
         End
         Begin VB.ComboBox Estado 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1305
            TabIndex        =   56
            ToolTipText     =   "Endereço para qual os meios de pagamentos devem ser enviados"
            Top             =   1440
            Width           =   630
         End
         Begin MSMask.MaskEdBox Cidade 
            Height          =   315
            Left            =   4050
            TabIndex        =   54
            ToolTipText     =   "Endereço para qual os meios de pagamentos devem ser enviados"
            Top             =   960
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Bairro 
            Height          =   315
            Left            =   1335
            TabIndex        =   53
            ToolTipText     =   "Endereço para qual os meios de pagamentos devem ser enviados"
            Top             =   990
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CEP 
            Height          =   315
            Left            =   6705
            TabIndex        =   55
            ToolTipText     =   "Endereço para qual os meios de pagamentos devem ser enviados"
            Top             =   945
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#####-###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Telefone1 
            Height          =   315
            Left            =   1290
            TabIndex        =   58
            ToolTipText     =   "Telefone de contato com a administradora"
            Top             =   1920
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Telefone2 
            Height          =   315
            Left            =   1290
            TabIndex        =   60
            ToolTipText     =   "Telefone de contato com a administradora"
            Top             =   2415
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Email 
            Height          =   315
            Left            =   4020
            TabIndex        =   61
            ToolTipText     =   "Email de contato com a administradora"
            Top             =   2400
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Contato 
            Height          =   315
            Left            =   6660
            TabIndex        =   62
            ToolTipText     =   "Nome do contato na administradora"
            Top             =   2355
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Fax 
            Height          =   315
            Left            =   4020
            TabIndex        =   59
            ToolTipText     =   "Fax de contato com a administradora"
            Top             =   1905
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Contato:"
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
            Index           =   26
            Left            =   5865
            TabIndex        =   72
            ToolTipText     =   "Nome do contato na administradora"
            Top             =   2430
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "CEP:"
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
            Index           =   19
            Left            =   6165
            TabIndex        =   65
            ToolTipText     =   "Endereço para qual os meios de pagamentos devem ser enviados"
            Top             =   1005
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fax:"
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
            Index           =   23
            Left            =   3570
            TabIndex        =   69
            ToolTipText     =   "Faz de contato com a administradora"
            Top             =   1950
            Width           =   405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "e-mail:"
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
            Left            =   3375
            TabIndex        =   71
            ToolTipText     =   "Email de contato com a administradora"
            Top             =   2430
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Telefone 2:"
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
            Index           =   24
            Left            =   195
            TabIndex        =   70
            ToolTipText     =   "Telefone de contato com a administradora"
            Top             =   2430
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Telefone 1:"
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
            Left            =   195
            TabIndex        =   68
            ToolTipText     =   "Telefone de contato com a administradora"
            Top             =   1950
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
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
            Left            =   615
            TabIndex        =   63
            ToolTipText     =   "Endereço para qual os meios de pagamentos devem ser enviados"
            Top             =   975
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
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
            Index           =   20
            Left            =   525
            TabIndex        =   66
            ToolTipText     =   "Endereço para qual os meios de pagamentos devem ser enviados"
            Top             =   1470
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
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
            Left            =   3300
            TabIndex        =   64
            ToolTipText     =   "Endereço para qual os meios de pagamentos devem ser enviados"
            Top             =   1005
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
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
            Index           =   16
            Left            =   315
            TabIndex        =   51
            ToolTipText     =   "Endereço para qual os meios de pagamentos devem ser enviados"
            Top             =   495
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "País:"
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
            Left            =   3495
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   67
            ToolTipText     =   "Endereço para qual os meios de pagamentos devem ser enviados"
            Top             =   1485
            Width           =   495
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7005
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "AdmMeioPagto.ctx":03DA
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "AdmMeioPagto.ctx":0558
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "AdmMeioPagto.ctx":0A8A
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "AdmMeioPagto.ctx":0C14
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3900
      Left            =   90
      TabIndex        =   0
      Top             =   525
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   6879
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Formas de Parcelamento"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Envio de Pagamento"
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
Attribute VB_Name = "AdmMeioPagto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Constante Limpa Combo
Const LIMPA_COMBO = -1
 
'Constantes Relacionadas as Colunas do Grid
Dim iGrid_Parcela_Col As Integer
Dim iGrid_Recebimento_Col As Integer
Dim iGrid_IntervalosPagamentos_Col As Integer

'Variáveis Globais
Dim gbCarregando As Boolean
Dim giTipoMeioPagto As Integer

'Flag de alteração dos campos da tela
Dim iAlterado As Integer
Dim iNumParcelas As Integer

'Indica qual dos frames do tab está visível no momento
Dim iFrameAtual As Integer

'Coleções da AdmMeioPagto
Public gcolAdmMeioPagtoCondPagto As New Collection
Public gcolTipoMeioPagto As New Collection

'Variável que guarda as características do grid da tela
Dim objGridParcelas As AdmGrid

Private WithEvents objEventoBanco As AdmEvento
Attribute objEventoBanco.VB_VarHelpID = -1

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Meios de Pagamento"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "AdmMeioPagto"

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



Private Sub PreDatado_Click()
    iAlterado = REGISTRO_ALTERADO
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

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Implementado pois agora é possível ter constantes cutomizadas em função de tamanhos de campos do BD. AdmLib.ClassConsCust
    Endereco.MaxLength = STRING_ENDERECO
    Bairro.MaxLength = STRING_BAIRRO
    Cidade.MaxLength = STRING_CIDADE
    Telefone1.MaxLength = STRING_TELEFONE
    Telefone2.MaxLength = STRING_TELEFONE
    Fax.MaxLength = STRING_FAX
    Email.MaxLength = STRING_EMAIL
    Contato.MaxLength = STRING_CONTATO
    
    'Variável Global Booleana é Setada no Form Load com  False
    'Varial que no Meio da Função Traz AdmMeioPagto para tela, Recebe True
    'para indicar que na tela estão sendo carregados dados do BD
    'eliminando assim a necessidade de automatismos de preenchimento
    gbCarregando = False

    'A tela abre com o primeiro frame visível
    iFrameAtual = 1
    
    'Inicializa o Evento de Browser
    Set objEventoBanco = New AdmEvento

    'Inicializa o código do próximo parcelamento
    'Todos os meios de pagamento tem o Parcelamento a vista "1"
    iCodigo = 1

    'Carrega a listbox de administradoras
    lErro = Carrega_Administradoras()
    If lErro <> SUCESSO Then gError 104000

    'Carrega a combo de Tipos de Meios de Pagamento
    lErro = Carrega_TipoMeioPagto()
    If lErro <> SUCESSO Then gError 104001

    'Carrega a combo de contas correntes
    lErro = Carrega_ContaCorrenteInterna()
    If lErro <> SUCESSO Then gError 104002
    
    'Carrega a Combo de Parcelamentos
    Call Carrega_Parcelamentos
    
    'Carrega as combos de Paises e a de estados
    lErro = Carrega_Pais_Estados()
    If lErro <> SUCESSO Then gError 104004

    'Carrega a combo de redes
    lErro = Carrega_Redes()
    If lErro <> SUCESSO Then gError 104005

    'Inicializa o obj relacionado ao Grid de parcelas
    Set objGridParcelas = New AdmGrid
    
    'Inicialização de Grid de Parcelamento
    lErro = Inicializa_Grid_Parcelas(objGridParcelas)
    If lErro <> SUCESSO Then gError 104006

    'Zera o flag de alterações indicando que não houve nenhuma ainda
    iAlterado = 0

    'Indica que o carregamento da tela aconteceu com sucesso
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        'Erros tratados na rotina chamadora
        Case 104000 To 104006
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142437)

    End Select

    Exit Sub

End Sub

Private Function Carrega_Administradoras() As Long
'Carrega a Lista administradora com código e o nome administradora em questão

Dim lErro As Long
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim colAdmMeioPagto As New Collection

On Error GoTo Erro_Carrega_Administradoras

    'Le os meios de pagamento
    lErro = CF("AdmMeioPagto_Le_Todas1", giFilialEmpresa, colAdmMeioPagto)
    If lErro <> SUCESSO Then gError 104033
    
    'Adcionar todos os Meios de Pagto na ListBox
    For Each objAdmMeioPagto In colAdmMeioPagto
               
        Administradoras.AddItem objAdmMeioPagto.sNome
        Administradoras.ItemData(Administradoras.NewIndex) = objAdmMeioPagto.iCodigo
        
    Next

    Carrega_Administradoras = SUCESSO

    Exit Function

Erro_Carrega_Administradoras:

    Carrega_Administradoras = gErr

    Select Case gErr

        Case 104033
         'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142438)

    End Select

    Exit Function
    
End Function

Function Carrega_TipoMeioPagto() As Long
'Carrega a Combo de Tipo meio Pagto com  as informações lida no BD
Dim lErro As Long
Dim objTipoMeioPagto As New ClassTMPLoja

On Error GoTo Erro_Carrega_TipoMeioPagto

    'Le os Tipos de Pagamentos e Carrega a Coleção Global
    lErro = CF("TipoMeioPagto_Le_Todas", gcolTipoMeioPagto)
    If lErro <> SUCESSO Then gError 104034
    
    'Adcionar na Combo TipoMeioPagto
    For Each objTipoMeioPagto In gcolTipoMeioPagto
        
        If objTipoMeioPagto.iPossuiAdm = POSSUI_ADM And objTipoMeioPagto.iTipo <> TIPOMEIOPAGTOLOJA_TEF Then
        
            TipoMeioPagto.AddItem objTipoMeioPagto.iTipo & SEPARADOR & objTipoMeioPagto.sDescricao
            TipoMeioPagto.ItemData(TipoMeioPagto.NewIndex) = objTipoMeioPagto.iTipo
      
        End If
      
    Next

    Carrega_TipoMeioPagto = SUCESSO

    Exit Function

Erro_Carrega_TipoMeioPagto:

    Carrega_TipoMeioPagto = gErr

    Select Case gErr

        'Erro Tradado dentro da Função
        Case 104034
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142439)

    End Select

    Exit Function

End Function

Function Carrega_ContaCorrenteInterna() As Long
'Função que carrega a combo de conta corrente Interna, com informações lida no BD

Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome
Dim lErro As Long

On Error GoTo Erro_Carrega_ContaCorrenteInterna
    
    'Lê cada código e descrição da tabela Paises
    lErro = CF("Cod_Nomes_Le", "ContasCorrentesInternas", "Codigo", "NomeReduzido", STRING_CONTA_CORRENTE_NOME_REDUZIDO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 104038


    'Preenche cada ComboBox ContaCorrente com os objetos da coleção colCodigoDescrição
    For Each objCodigoDescricao In colCodigoDescricao

       ContaCorrenteInterna.AddItem objCodigoDescricao.iCodigo & SEPARADOR & objCodigoDescricao.sNome
       ContaCorrenteInterna.ItemData(ContaCorrenteInterna.NewIndex) = objCodigoDescricao.iCodigo
    
    Next

    Carrega_ContaCorrenteInterna = SUCESSO

    Exit Function

Erro_Carrega_ContaCorrenteInterna:

   Carrega_ContaCorrenteInterna = gErr

    Select Case gErr

        Case 104038
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142440)

    End Select

    Exit Function

End Function
Sub Carrega_Parcelamentos()
'Carrega  Coleção Global com Informações do Parcelamento à Vista

Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim objAdmMeioPagtoParcelas As ClassAdmMeioPagtoParcelas

On Error GoTo Erro_Carrega_Parcelamentos
    
    'Limpa a Combo de Parcelamentos
    Parcelamentos.Clear
    
    'limpar A coleção
    Set gcolAdmMeioPagtoCondPagto = New Collection
    
    'Estanciar o objadmMeioPagtoCondPagto para ser Adcionado na Coleção
    Set objAdmMeioPagtoCondPagto = New ClassAdmMeioPagtoCondPagto
    
    objAdmMeioPagtoCondPagto.sNomeParcelamento = NOME_A_VISTA
    objAdmMeioPagtoCondPagto.iNumParcelas = 1
    objAdmMeioPagtoCondPagto.iParcelamento = COD_A_VISTA
    objAdmMeioPagtoCondPagto.iParcelasRecebto = 1
    objAdmMeioPagtoCondPagto.dTaxa = PercentParaDbl(TaxaVista.Text)
    objAdmMeioPagtoCondPagto.dValorMinimo = 0
    objAdmMeioPagtoCondPagto.dDesconto = 0
    objAdmMeioPagtoCondPagto.iJurosParcelamento = JUROS_LOJA
    objAdmMeioPagtoCondPagto.dJuros = 0
        
        
    'Estanciar o objAdmMeioPagtoParcelas para ser Adcionado na Coleção
    Set objAdmMeioPagtoParcelas = New ClassAdmMeioPagtoParcelas
    
    objAdmMeioPagtoParcelas.iParcela = 1
    objAdmMeioPagtoParcelas.dPercRecebimento = 1
    objAdmMeioPagtoParcelas.iIntervaloRecebimento = StrParaInt(DefasagemPagtoVista.Text)
    objAdmMeioPagtoParcelas.iParcelamento = 1
    
    'Adicionar informações de Parcelamntos na Coleção Parcelamento
    objAdmMeioPagtoCondPagto.colParcelas.Add objAdmMeioPagtoParcelas
    
    'Adicionar informações de AdmMeiopagtoCondPagto na Coleção Global
    gcolAdmMeioPagtoCondPagto.Add objAdmMeioPagtoCondPagto
    
    'Faz a Chamada da Função que Carrega a Combo de Parcelamento
    Call Carrega_ComboParcelas
    
    Exit Sub
    
Erro_Carrega_Parcelamentos:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142441)

    End Select

    Exit Sub
    
End Sub

Function Carrega_Pais_Estados() As Long
'Carrega a Combo de Pais e Estados com informações carregadas do BD

Dim colCodigo As New Collection
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Carrega_Pais_Estados

    'Lê cada código da tabela Estados e poe na coleção colCodigo
    lErro = CF("Codigos_Le", "Estados", "Sigla", TIPO_STR, colCodigo, STRING_ESTADOS_SIGLA)
    If lErro <> SUCESSO Then gError 104045

    For iIndice = 1 To colCodigo.Count
        Estado.AddItem colCodigo(iIndice)
    Next
    
    'Inicializa a coleção de código descrição
    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê cada código e descrição da tabela Paises
    lErro = CF("Cod_Nomes_Le", "Paises", "Codigo", "Nome", STRING_PAISES_NOME, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 104046

    'Preenche cada ComboBox País com os objetos da coleção colCodigoDescrição
    For Each objCodigoDescricao In colCodigoDescricao

        Pais.AddItem objCodigoDescricao.iCodigo & SEPARADOR & objCodigoDescricao.sNome
        Pais.ItemData(Pais.NewIndex) = objCodigoDescricao.iCodigo

    Next

    'Seleciona Brasil se existir
    For iIndice = 0 To Pais.ListCount - 1
        If Right(Pais.List(iIndice), 6) = "Brasil" Then
            Pais.ListIndex = iIndice
            Exit For
        End If
    Next

    Carrega_Pais_Estados = SUCESSO

    Exit Function

Erro_Carrega_Pais_Estados:

   Carrega_Pais_Estados = gErr

    Select Case gErr

        Case 104045, 104046
            'Erro tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142442)

    End Select

    Exit Function

End Function

Private Function Carrega_Redes() As Long
'Função que Carrega a Combo de Redes com Informações lidas no BD
Dim lErro As Long
'Dim objCodigoNome As AdmCodigoNome
'Dim colCodigoNome As New AdmColCodigoNome
Dim objRede As ClassRede
Dim colRedes As New Collection

On Error GoTo Erro_Carrega_Redes

    'Lê o Código e o Nome de Todas as Redes do BD
'    lErro = CF("Cod_Nomes_Le", "Redes", "Codigo", "Nome", STRING_REDE_NOME, colCodigoNome)
'    If lErro <> SUCESSO Then gError 104047

    lErro = CF("Redes_Le_Todas", colRedes)
    If lErro <> SUCESSO Then gError 104047

    'Carrega a combo de Redes
'    For Each objCodigoNome In colCodigoNome
'        Rede.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
'        Rede.ItemData(Rede.NewIndex) = objCodigoNome.iCodigo
'    Next
    For Each objRede In colRedes
        Rede.AddItem CStr(objRede.iCodigo) & SEPARADOR & objRede.sNome
        Rede.ItemData(Rede.NewIndex) = objRede.iCodigo
    Next

    Carrega_Redes = SUCESSO

    Exit Function

Erro_Carrega_Redes:

    Carrega_Redes = gErr

    Select Case gErr

        Case 104047
         'Erro tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142443)

    End Select

    Exit Function

End Function

Function Inicializa_Grid_Parcelas(objGridInt As AdmGrid) As Long

   'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("% Recebimento")
    objGridInt.colColuna.Add ("Intervalos de Pagamento")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (PercRecebimento.Name)
    objGridInt.colCampo.Add (IntervaloParcela.Name)

    'Colunas do Grid
    iGrid_Parcela_Col = 0
    iGrid_Recebimento_Col = 1
    iGrid_IntervalosPagamentos_Col = 2

    'Grid do GridInterno
    objGridInt.objGrid = GridParcelas

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_PARCELAS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridParcelas.ColWidth(0) = 900

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridParcelas)
    
    Inicializa_Grid_Parcelas = SUCESSO

    Exit Function

End Function

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Function Trata_Parametros(Optional objAdmMeioPagto As ClassAdmMeioPagto) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se houver POS passado como parâmetro, exibe seus dados
    If Not (objAdmMeioPagto Is Nothing) Then

        objAdmMeioPagto.iFilialEmpresa = giFilialEmpresa

        If objAdmMeioPagto.iCodigo > 0 Then

            'Lê POS no BD a partir do código
            lErro = CF("AdmMeioPagto_Le", objAdmMeioPagto)
            If lErro <> SUCESSO And lErro <> 104017 Then gError 104018
            If lErro = SUCESSO Then

                'Exibe os dados de AdmMeioPagto
                lErro = Traz_AdmMeioPagto_Tela(objAdmMeioPagto)
                If lErro <> SUCESSO Then gError 104019
            Else
    
                Codigo.Text = objAdmMeioPagto.iCodigo
                Nome.Text = objAdmMeioPagto.sNome
                    
            End If

        End If
    
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 104018, 104019
            'Erro tratado dentro da Função Chamadora
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142444)

    End Select

    Exit Function

End Function

Function Traz_AdmMeioPagto_Tela(objAdmMeioPagto As ClassAdmMeioPagto) As Long
'Função que Traz as Informações da Admnistradoras contida no objAdmMeioPagto para Tela

Dim lErro As Long
Dim iIndice As Integer
Dim objEndereco As New ClassEndereco
Dim objContasCorrentesInternas As New ClassContasCorrentesInternas
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_Traz_AdmMeioPagto_Tela

    'Variável gbCarregnado recebe True
    gbCarregando = True
    
    Call Limpa_Tela_AdmMeioPagto
    
    If objAdmMeioPagto.iAtivo = ADMMEIOPAGTO_ATIVO Then
        Ativo.Value = vbChecked
    Else
        Ativo.Value = vbUnchecked
    End If
    
    'Lê para cada admnistradoras os Pacelamentos Vinculados
    lErro = CF("AdmMeioPagtoCondPagto_Le", objAdmMeioPagto)
    If lErro <> SUCESSO Then gError 104008
    
   'Preenche o endereço
    If objAdmMeioPagto.lEndereco > 0 Then

        objAdmMeioPagto.objEndereco.lCodigo = objAdmMeioPagto.lEndereco

        'Lê o endereço
        lErro = CF("Endereco_Le", objAdmMeioPagto.objEndereco)
        If lErro <> SUCESSO And lErro <> 12513 Then gError 104012
        'Se não encontrou então Erro
        If lErro = 12513 Then gError 104013

    End If
    
    'Traz os dados para tela
    Codigo.Text = objAdmMeioPagto.iCodigo
    
    'Procedimento quando a Combo é Editável
    If objAdmMeioPagto.iRede > 0 Then
        Rede.Text = objAdmMeioPagto.iRede
        Call Rede_Validate(False)
    End If
    
    Nome.Text = objAdmMeioPagto.sNome
    
    'verifica o valor do taxavista
    If objAdmMeioPagto.dTaxaVista > 0 Then TaxaVista.Text = Format(objAdmMeioPagto.dTaxaVista * 100, "Fixed")
    

    'verifica o valor do taxaparcelado
    If objAdmMeioPagto.dTaxaParcelado > 0 Then TaxaPrazo.Text = Format(objAdmMeioPagto.dTaxaParcelado * 100, "Fixed")
    
   'Procedimento usado quando a combox não é editavél.
    For iIndice = 0 To TipoMeioPagto.ListCount - 1
       If TipoMeioPagto.ItemData(iIndice) = objAdmMeioPagto.iTipoMeioPagto Then
            TipoMeioPagto.ListIndex = iIndice
            Exit For
        End If
    Next
    
    If objAdmMeioPagto.iDefasagemPagtoVista > 0 Then DefasagemPagtoVista.Text = objAdmMeioPagto.iDefasagemPagtoVista
    
    If objAdmMeioPagto.iContaCorrenteInterna > 0 Then
        'Procedimento quando a Combo é Editável
        ContaCorrenteInterna.Text = objAdmMeioPagto.iContaCorrenteInterna
        Call ContaCorrenteInterna_Validate(False)
    
    End If
    
    CheckNaoTitRec.Value = objAdmMeioPagto.iGeraTituloRec
    
    
    'Faz com que a coleção global receba as Informações contidas na coleção de colCondPagtoLoja
    Set gcolAdmMeioPagtoCondPagto = objAdmMeioPagto.colCondPagtoLoja
    
    'Função para carregar as Parcelas do Grid relacionadas ao Parcelamento
    Call Carrega_ComboParcelas
    
    'Verifica e o Envio é Banco
    If objAdmMeioPagto.iCodBanco > 0 Then
        
        CodBanco.Text = objAdmMeioPagto.iCodBanco
        Agencia.Text = objAdmMeioPagto.sAgencia
    
    End If
    'Verifica e o Envio é Endereço
    If objAdmMeioPagto.lEndereco > 0 Then
        Email.Text = objAdmMeioPagto.objEndereco.sEmail
        Estado.Text = objAdmMeioPagto.objEndereco.sSiglaEstado
        Endereco.Text = objAdmMeioPagto.objEndereco.sEndereco
        Bairro.Text = objAdmMeioPagto.objEndereco.sBairro
        Telefone1.Text = objAdmMeioPagto.objEndereco.sTelefone1
        Telefone2.Text = objAdmMeioPagto.objEndereco.sTelefone2
        Cidade.Text = objAdmMeioPagto.objEndereco.sCidade
        Fax.Text = objAdmMeioPagto.objEndereco.sFax
        CEP.PromptInclude = False
        CEP.Text = objAdmMeioPagto.objEndereco.sCEP
        CEP.PromptInclude = True
        Contato.Text = objAdmMeioPagto.objEndereco.sContato
        
        'se for Endereço então Preenecher a Combo de Pais
        If objAdmMeioPagto.objEndereco.iCodigoPais > 0 Then

            'Procedimento quando a Combo é Editável
            Pais.Text = objAdmMeioPagto.objEndereco.iCodigoPais
            Call Pais_Validate(False)
       
       End If
       'se for Endereço então Preenecher a Combo de Estado
        If objAdmMeioPagto.objEndereco.sSiglaEstado <> "" Then
            
            Estado.Text = objAdmMeioPagto.objEndereco.sSiglaEstado
            Call Estado_Validate(False)
       
        End If
       
    End If
    
    iAlterado = 0

    'A gbCarregando Recebe False
    gbCarregando = False
    
    Exit Function

Erro_Traz_AdmMeioPagto_Tela:

    Traz_AdmMeioPagto_Tela = gErr

    gbCarregando = False

    Select Case gErr

        'Erro Tradado na Funçào Chamadora
        Case 104008, 104009, 104012
        
        Case 104013
             Call Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142445)

    End Select

    Exit Function

End Function

Sub Limpa_Tela_AdmMeioPagto()
'Limpa a Tela AdmMeioPagto
 Dim iIndice As Integer
 
    'Limpa a tela
    Call Limpa_Tela(Me)
    
    'Limpando combos Editáveis
    Rede.Text = ""
    Pais.Text = ""
    Estado.Text = ""
    ContaCorrenteInterna.Text = ""

    'Limpando Combo não editáveis
    TipoMeioPagto.ListIndex = LIMPA_COMBO
    
    'Desmarcando as Checks
    CheckNaoTitRec.Value = CHECK_DESMARC
    
    Ativo.Value = vbChecked
    
    'Limpa Parcelamentos
    Call Limpar_Parcelamentos
    
    Call Carrega_Parcelamentos
    
    iAlterado = 0

End Sub

Sub Limpar_Parcelamentos()
    
   'Limpar os Parcelamento
    Parcelamentos.Text = ""
    ParcelasPagto.Text = ""
    ParcelasRecebto.Text = ""
    TaxaParcelamento.Text = ""
    ValorMinimo.Text = ""
    Desconto.Text = ""
    JurosParcelas.Text = ""
    AtivoParc.Value = False
    PreDatado.Value = False
    
    'Limpar a Soma do Percentual das Parcelas
     PercTotal.Caption = ""
    
    'Limpar o GridParcelas
    Call Grid_Limpa(objGridParcelas)

End Sub

Sub Carrega_ComboParcelas()
'Carrega a Combo com Parcelamento

Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
    
    'Limpar a Combo Parcelamentos
    Parcelamentos.Clear
    
    For Each objAdmMeioPagtoCondPagto In gcolAdmMeioPagtoCondPagto
            
        Parcelamentos.AddItem objAdmMeioPagtoCondPagto.sNomeParcelamento
        Parcelamentos.ItemData(Parcelamentos.NewIndex) = objAdmMeioPagtoCondPagto.iParcelamento
                    
    Next
    
    Exit Sub
    
End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD
Dim lErro As Long
Dim objAdmMeioPagto As New ClassAdmMeioPagto

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "AdmMeioPagtoTipo"

    'Le os dados da Tela AdmMeioPagto
    lErro = Move_Tela_Memoria(objAdmMeioPagto)
    If lErro <> SUCESSO Then gError 104007

    'Preenche a coleção colCampoValor, com nome do campo,
    colCampoValor.Add "Codigo", objAdmMeioPagto.iCodigo, 0, "Codigo"
    colCampoValor.Add "Nome", objAdmMeioPagto.sNome, STRING_ADMMEIOPAGTO_NOME, "Nome"
    colCampoValor.Add "TaxaVista", objAdmMeioPagto.dTaxaVista, 0, "TaxaVista"
    colCampoValor.Add "TaxaParcelado", objAdmMeioPagto.dTaxaParcelado, 0, "TaxaParcelado"
    colCampoValor.Add "TipoMeioPagto", objAdmMeioPagto.iTipoMeioPagto, 0, "TipoMeioPagto"
    colCampoValor.Add "DefasagemPagtoVista", objAdmMeioPagto.iDefasagemPagtoVista, 0, "DefasagemPagtoVista"
    colCampoValor.Add "Endereco", objAdmMeioPagto.lEndereco, 0, "Endereco"
    colCampoValor.Add "CodBanco", objAdmMeioPagto.iCodBanco, 0, "CodBanco"
    colCampoValor.Add "Agencia", objAdmMeioPagto.sAgencia, STRING_ADMMEIOPAGTO_AGENCIA, "Agencia"
    colCampoValor.Add "ContaCorrenteInterna", objAdmMeioPagto.iContaCorrenteInterna, 0, "ContaCorrenteInterna"
    colCampoValor.Add "FilialEmpresa", objAdmMeioPagto.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "GeraTituloRec", objAdmMeioPagto.iGeraTituloRec, 0, "GeraTituloRec"
    colCampoValor.Add "Rede", objAdmMeioPagto.iRede, 0, "Rede"
    colCampoValor.Add "Ativo", objAdmMeioPagto.iAtivo, 0, "Ativo"
    
   'Utilizado na hora de passar o parâmetro FilialEmpresa
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "PossuiAdm", OP_IGUAL, MARCADO
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr
        'Erro tratado na rotina chamadora
        Case 104007
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142446)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objAdmMeioPagto As New ClassAdmMeioPagto

On Error GoTo Erro_Tela_Preenche

    objAdmMeioPagto.iCodigo = colCampoValor.Item("Codigo").vValor
            
    If objAdmMeioPagto.iCodigo > 0 Then
        
        'Carrega objAdmMeioPagto com os dados passados em colCampoValor
        objAdmMeioPagto.sNome = colCampoValor.Item("Nome").vValor
        objAdmMeioPagto.dTaxaVista = colCampoValor.Item("TaxaVista").vValor
        objAdmMeioPagto.dTaxaParcelado = colCampoValor.Item("TaxaParcelado").vValor
        objAdmMeioPagto.iTipoMeioPagto = colCampoValor.Item("TipoMeioPagto").vValor
        objAdmMeioPagto.iDefasagemPagtoVista = colCampoValor.Item("DefasagemPagtoVista").vValor
        objAdmMeioPagto.lEndereco = colCampoValor.Item("Endereco").vValor
        objAdmMeioPagto.iCodBanco = colCampoValor.Item("CodBanco").vValor
        objAdmMeioPagto.sAgencia = colCampoValor.Item("Agencia").vValor
        objAdmMeioPagto.iContaCorrenteInterna = colCampoValor.Item("ContaCorrenteInterna").vValor
        objAdmMeioPagto.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
        objAdmMeioPagto.iGeraTituloRec = colCampoValor.Item("GeraTituloRec").vValor
        objAdmMeioPagto.iRede = colCampoValor.Item("Rede").vValor
        objAdmMeioPagto.iAtivo = colCampoValor.Item("Ativo").vValor
        
        'Traz dados da Administradora para a Tela
        lErro = Traz_AdmMeioPagto_Tela(objAdmMeioPagto)
        If lErro <> SUCESSO Then gError 104108
        
    End If
        
    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 104108
        'Erro tratado na rotina chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142447)

    End Select
    
    Exit Sub

End Sub
Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Private Sub Codigo_GotFocus()
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
End Sub

Private Sub BotaoProxNum_Click()
'Botão que Gera um Proximo Código para o Meio de Pagto

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    lErro = AdmMeioPagto_Codigo_Automatico(lCodigo)
    If lErro <> SUCESSO Then gError 107667

    Codigo.Text = lCodigo

    Exit Sub
    
Erro_BotaoProxNum_Click:

    Select Case gErr
        
        Case 107667
            'Erro Tratado Dentro da Função Chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142448)

        End Select

    Exit Sub

End Sub

Function AdmMeioPagto_Codigo_Automatico(lCodigo As Long) As Long

Dim lErro As Long

On Error GoTo Erro_AdmMeioPagto_Codigo_Automatico

    'Chama a rotina que gera o sequencial Para Back
    lErro = CF("Config_ObterAutomatico", "LojaConfig", "COD_PROX_ADMMEIOPAGTO", "AdmMeioPagto", "Codigo", lCodigo)
    If lErro <> SUCESSO Then gError 107668

    AdmMeioPagto_Codigo_Automatico = SUCESSO
    
    Exit Function

Erro_AdmMeioPagto_Codigo_Automatico:

    Select Case gErr

        Case 107668
            'Erro Tratado Dentro da Função Chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142449)

        End Select

    Exit Function

End Function
Private Sub Ativo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Rede_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Rede_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Rede_Validate(Cancel As Boolean)
'função que Valida os Dados na Comdo de Redes

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objRede As New ClassRede
Dim iCodigo As Integer

On Error GoTo Erro_Rede_Validate

    If Len(Trim(Rede.Text)) = 0 Then Exit Sub

    'Verifica se há Algo Selecionado  Selecionado
    If Rede.ListIndex <> -1 Then Exit Sub

    'Verifica se existe o item na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Rede, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 104024

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        'Lê o Código ea Filial Empresa
        objRede.iCodigo = iCodigo
        objRede.iFilialEmpresa = giFilialEmpresa
        
        'Tenta ler Rede com esse código no BD
        lErro = CF("Rede_Le", objRede)
        If lErro <> SUCESSO And lErro <> 104244 Then gError 104025

        If lErro = 104244 Then gError 104026 'Pergunta se Deseja Cadastrar Rede

        'Encontrou Rede no BD, Adciona na Combo
        Rede.AddItem objRede.iCodigo & SEPARADOR & objRede.sNome
        Rede.ItemData(Rede.NewIndex) = objRede.iCodigo

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 104027 'Pergunta se Deseja Cadastrar a Rede

    Exit Sub

Erro_Rede_Validate:

    Cancel = True

    Select Case gErr
        
        Case 104024, 104025
        'Erro tratado na rotina chamadora

        Case 104026
            'Pergunta ao usuário se ele deseja cadastrar A Rede
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_CADASTRAR_REDE", objRede.iCodigo)
                        
            'Se sim
            If vbMsgRes = vbYes Then
                
                'Chama a tela Rede
                Call Chama_Tela("Rede", objRede)
            
            End If
            
        Case 104027
            'Pergunta ao usuário se ele deseja cadastrar A Rede
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_CADASTRAR_REDE", objRede.sNome)
                        
            'Se sim
            If vbMsgRes = vbYes Then
                
                'Chama a tela Rede
                Call Chama_Tela("Rede", objRede)
            
            End If
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142450)

    End Select

    Exit Sub

End Sub
Private Sub Nome_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Private Sub TaxaVista_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Private Sub TaxaVista_Validate(Cancel As Boolean)
'Função que valida as informações contidas em Taxa a Vista
Dim lErro As Long

On Error GoTo Erro_TaxaVista_Validate
    
    'Verifica se Taxa à Vista foi preenchida
    If Len(Trim(TaxaVista.Text)) > 0 Then
       'Retorna Erro ao Usuário
       lErro = Porcentagem_Critica2(TaxaVista.Text)
       If lErro <> SUCESSO Then gError 104027
    
    End If

    Exit Sub

Erro_TaxaVista_Validate:

    Cancel = True

    Select Case gErr

        Case 104027
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142451)

    End Select

End Sub

Private Sub TaxaPrazo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TaxaPrazo_Validate(Cancel As Boolean)
'Função que valida as Informações contidas em taxa a Prazo
Dim lErro As Long

On Error GoTo Erro_TaxaPrazo_Validate

    'Verifica se Taxa à Prazo foi preenchida
    If Len(Trim(TaxaPrazo.Text)) > 0 Then
    
        'Função que verifica se a porcentagem é válida
        lErro = Porcentagem_Critica2(TaxaPrazo.Text)
        If lErro <> SUCESSO Then gError 104028

    End If
    
    Exit Sub

Erro_TaxaPrazo_Validate:

    Cancel = True

    Select Case gErr

        Case 104028
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142452)

    End Select
    
    Exit Sub

End Sub
Private Sub TipoMeioPagto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoMeioPagto_Click()
'Define se vai ser Habilitado o Frame  Endereço ou Banco, Verifica se o Tipo de pagto é Carnê se for traz para tela, permite alteração

Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim objTipoMeioPagtoLoja As New ClassTMPLoja
Dim lErro As Long
Dim vbResultado As VbMsgBoxResult
Dim bDesvio As Boolean
Dim iIndice As Integer

On Error GoTo Erro_TipoMeioPagto_Click

    'Guarda no objAdmMeioPagto a FilialEmpresa que se Esta Trabalhando
    objAdmMeioPagto.iFilialEmpresa = giFilialEmpresa
    
    If TipoMeioPagto.ListIndex = -1 Then Exit Sub
    
    'Se o meio de pagamento foi carregado com sucesso, não Permite que seja Alterado _
    o Tipo de Pagto
    If gbCarregando = False And _
       giTipoMeioPagto <> Codigo_Extrai(TipoMeioPagto.Text) And _
       Codigo_Extrai(TipoMeioPagto.Text) = TIPOMEIOPAGTOLOJA_CARNE And _
       StrParaInt(Codigo.Text) <> MEIO_PAGAMENTO_CARNE Then
        
        'Ao tentar Ser Feito alteração no Tipo de Pagto, pergunta se deseja carrregar o tipo de pagto de Carnê
        vbResultado = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_CARREGAR_TIPOMEIOPAGTO_CARNE")
    
        'Se sim
        If vbResultado = vbYes Then
                   
            objAdmMeioPagto.iCodigo = MEIO_PAGAMENTO_CARNE
     
            lErro = CF("AdmMeioPagto_Le", objAdmMeioPagto)
            If lErro <> SUCESSO Then gError 107643
            
            lErro = Traz_AdmMeioPagto_Tela(objAdmMeioPagto)
            If lErro <> SUCESSO Then gError 107641
        
            bDesvio = True
        Else
            
            For iIndice = 0 To TipoMeioPagto.ListCount - 1
                
                If TipoMeioPagto.ItemData(iIndice) = giTipoMeioPagto Then
                
                    'Carrega o Tipo de Pagto que estana variável global giTipoMeioPagto
                    TipoMeioPagto.ListIndex = iIndice
            
                End If
                
            Next
        
        End If
        
    'Se o Tipo de Meio de Pagamento For Carnê
    ElseIf TipoMeioPagto.ItemData(TipoMeioPagto.ListIndex) = TIPOMEIOPAGTOLOJA_CARNE And gbCarregando = False Then
        
        'Passa por parâmetro o Codigo do Meio de Pagamento que Corresponde ao Carnê
        objAdmMeioPagto.iCodigo = MEIO_PAGAMENTO_CARNE
        
        lErro = CF("AdmMeioPagto_Le", objAdmMeioPagto)
        If lErro <> SUCESSO Then gError 107640
        
        'Traz o para a Tela o Meio de Pagmento Carnê
        lErro = Traz_AdmMeioPagto_Tela(objAdmMeioPagto)
        If lErro <> SUCESSO Then gError 107636
        
        bDesvio = True
        
    End If
    
    If bDesvio = False Then
        For Each objTipoMeioPagtoLoja In gcolTipoMeioPagto
        
            'Verificar se é Cobrança é Endereço ou Banco
            If TipoMeioPagto.ItemData(TipoMeioPagto.ListIndex) = objTipoMeioPagtoLoja.iTipo Then
            'Visualização de um Frame ou de Outro , se o Tipo for Envio_Endereço
                If objTipoMeioPagtoLoja.iEnvioPagamento = ENVIO_ENDERECO Then
            
                    FrameEndereco.Visible = True
                    FrameBanco.Visible = False
                'Se não
                Else
        
                    FrameEndereco.Visible = False
                    FrameBanco.Visible = True
            
                End If
        
            End If
                
        Next
        
    End If
    
    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub
    
Erro_TipoMeioPagto_Click:

    Select Case gErr

        Case 107636, 107640, 107641, 107643
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142453)

    End Select

    Exit Sub

End Sub
Private Sub ContaCorrenteInterna_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Private Sub ContaCorrenteInterna_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ContaCorrenteInterna_Validate(Cancel As Boolean)
'valida as Informações contidas na combo de Conta Corrente Interna

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objContaCorrenteInterna As New ClassContasCorrentesInternas
Dim iCodigo As Integer

On Error GoTo Erro_ContaCorrenteInterna_Validate

    If Len(Trim(ContaCorrenteInterna.Text)) = 0 Then Exit Sub
    
    'Verifica se está preenchida com o item selecionado na ComboBox ContaCorrenteInterna
    If ContaCorrenteInterna.ListIndex <> -1 Then Exit Sub

    'Verifica se existe o item na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(ContaCorrenteInterna, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 104021
    
    'Nao Encontrou nas Lista o Texto
    If lErro = 6731 Then gError 104020
    
    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        'Tenta ler ContaCorrenteInterna com esse código no BD
        lErro = CF("ContaCorrenteInt_Le", iCodigo, objContaCorrenteInterna)
        If lErro <> SUCESSO And lErro <> 11807 Then gError 104022

        'Não encontrou ContaCorrenteInterna no BD
        If lErro = 11807 Then gError 104023

        'Encontrou ContaCorrenteInterna no BD, coloca no Text da Combo
        ContaCorrenteInterna.Text = objContaCorrenteInterna.iCodigo & SEPARADOR & objContaCorrenteInterna.sNomeReduzido

    End If

    Exit Sub

Erro_ContaCorrenteInterna_Validate:

    Cancel = True

    Select Case gErr

        Case 104020
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_CORRENTE_NAO_ENCONTRADA", gErr)
        
        Case 104021, 104022
        'Erro tratado na rotina chamadora

        Case 104023
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CORRENTE_NAO_ENCONTRADA", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142454)

    End Select

    Exit Sub

End Sub

Private Sub CheckNaoTitRec_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub
Private Sub Administradoras_Click()
'Preencher a tela com os dados da Admnistradoras atraves do Click
Dim lErro As Long
Dim objAdmMeioPagto As New ClassAdmMeioPagto

On Error GoTo Erro_Administradoras
    
    If Administradoras.ListIndex = -1 Then Exit Sub
    
    objAdmMeioPagto.iCodigo = Administradoras.ItemData(Administradoras.ListIndex)
    
    'Preencher Filial Empreasa
    objAdmMeioPagto.iFilialEmpresa = giFilialEmpresa
    
    'Lê no banco de Dados, Os Dados da Admnistradoras Selecionada
    lErro = CF("AdmMeioPagto_Le", objAdmMeioPagto)
    If lErro <> SUCESSO And lErro <> 104017 Then gError 104048
   
    If lErro = 104017 Then gError 104081
    
    'Preenche a tela com os Dados da Administradora Carregada
    lErro = Traz_AdmMeioPagto_Tela(objAdmMeioPagto)
    If lErro <> SUCESSO Then gError 104049

    Exit Sub

Erro_Administradoras:

    Select Case gErr

        Case 104048, 104049
            'Erro tratado na rotina chamadora
    
        Case 104081
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_CADASTRADO", gErr, objAdmMeioPagto.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142455)

    End Select

    Exit Sub

End Sub
Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.index

    End If

End Sub
Private Sub Parcelamentos_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Parcelamentos_Click()
' Traz as Informações de AdmMeioPagtoCondPagto para tela atraves do Parcelamento
Dim lErro As Long
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_Parcelamentos
    'Se nenhum parcelamento estiver Selecionado então sai
    If Parcelamentos.ListIndex = -1 Then Exit Sub
    
    'Percorrer A coleção global a Procura do parcelamento Selecionado na Combo
    For Each objAdmMeioPagtoCondPagto In gcolAdmMeioPagtoCondPagto
        
        'Se encontrou
        If objAdmMeioPagtoCondPagto.sNomeParcelamento = Parcelamentos.List(Parcelamentos.ListIndex) Then
            'Traz AdmMeioPagtoCondPagto para Tela
            lErro = Traz_AdmMeioPagtoCondPagto_Tela(objAdmMeioPagtoCondPagto)
            If lErro <> SUCESSO Then gError 104082
            Exit For
        
        End If
    
    Next
    
    'Calcula e verifica se o Percentual de Parcelas é menor que 100%
    Call Calcular_Total_Parcelas
    
    Exit Sub
    
Erro_Parcelamentos:
    
    Select Case gErr
        
        Case 104082
            'Erro tratado Dentro da Função Chamadora
            
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142456)
             
    End Select
    
    Exit Sub

End Sub

Function Traz_AdmMeioPagtoCondPagto_Tela(objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_AdmMeioPagtoCondPagto_Tela

    'verifica o valor de Parcelamento é Maior que Zero
    If objAdmMeioPagtoCondPagto.dTaxa > 0 Then
        
        TaxaParcelamento.Text = Format(objAdmMeioPagtoCondPagto.dTaxa * 100, "Fixed")
    Else
        
        TaxaParcelamento.Text = ""
    End If
    
    'verifica o valor do Desconto
    If objAdmMeioPagtoCondPagto.dDesconto > 0 Then
        
        Desconto.Text = Format(objAdmMeioPagtoCondPagto.dDesconto * 100, "Fixed")
    
    Else
        
        Desconto.Text = ""
    
    End If
    
    'Verfica o Valor de Juros
    If objAdmMeioPagtoCondPagto.dJuros > 0 Then
        
        JurosParcelas.Text = Format(objAdmMeioPagtoCondPagto.dJuros * 100, "Fixed")
    
    Else
        
        JurosParcelas.Text = ""
        
    End If
    
    'Verifica se o Valor minimo é zero
    If objAdmMeioPagtoCondPagto.dValorMinimo = 0 Then
        
        ValorMinimo.Text = ""
    Else
    
        ValorMinimo.Text = objAdmMeioPagtoCondPagto.dValorMinimo
    
    End If
    
    If objAdmMeioPagtoCondPagto.iJurosParcelamento = JUROS_LOJA Then
        
        ParcelamentoLoja.Value = True
    
    Else
        
        ParcelamentoAdm.Value = True
        
    End If
    
    If objAdmMeioPagtoCondPagto.iAtivo = ADMMEIOPAGTOCONDPAGTO_ATIVO Then
        AtivoParc.Value = MARCADO
    Else
        AtivoParc.Value = DESMARCADO
    End If
    
    If objAdmMeioPagtoCondPagto.iPreDatado = ADMMEIOPAGTOCONDPAGTO_PREDATADO Then
        PreDatado.Value = MARCADO
    Else
        PreDatado.Value = DESMARCADO
    End If
    
    
    'Preenchimento da Tela AdmMeioPagtoCondPagto
    ParcelasPagto.Text = objAdmMeioPagtoCondPagto.iNumParcelas
    ParcelasRecebto.Text = objAdmMeioPagtoCondPagto.iParcelasRecebto
    
    
    'Preenche o Grid com dados do Parcelamento
    lErro = Traz_Parcelas_Tela(objAdmMeioPagtoCondPagto)
    If lErro <> SUCESSO Then gError 104083
    
    Traz_AdmMeioPagtoCondPagto_Tela = SUCESSO
    
    Exit Function

Erro_Traz_AdmMeioPagtoCondPagto_Tela:
    
    Traz_AdmMeioPagtoCondPagto_Tela = gErr
    
    Select Case gErr
         
        Case 104083
            'Erro Tratado dentro da Função que a Chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142457)
            
    End Select
    
    Exit Function
    
End Function


Function Traz_Parcelas_Tela(objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto) As Long

Dim lErro As Long
Dim objAdmMeioPagtoParcelas As ClassAdmMeioPagtoParcelas
Dim iIndice As Integer
Dim dResto As Double
Dim dSomaPercentual As Double

On Error GoTo Erro_Traz_Parcelas_Tela
    
    'Limpa o Grid Parcelas
    Call Grid_Limpa(objGridParcelas)

    'Iniciliza Linhas Existentes
    objGridParcelas.iLinhasExistentes = objAdmMeioPagtoCondPagto.iParcelasRecebto
    
    'Verificando a Coleção para encontrar as parcelas vinculadas
    For Each objAdmMeioPagtoParcelas In objAdmMeioPagtoCondPagto.colParcelas
        'preencher o Grid
         iIndice = iIndice + 1
                
        GridParcelas.TextMatrix(iIndice, iGrid_Parcela_Col) = iIndice
        GridParcelas.TextMatrix(iIndice, iGrid_Recebimento_Col) = Format(objAdmMeioPagtoParcelas.dPercRecebimento, "PERCENT")
        GridParcelas.TextMatrix(iIndice, iGrid_IntervalosPagamentos_Col) = objAdmMeioPagtoParcelas.iIntervaloRecebimento
       
    Next
        
    Traz_Parcelas_Tela = SUCESSO
            
    Exit Function

Erro_Traz_Parcelas_Tela:
    
    Traz_Parcelas_Tela = gErr
    
    Select Case gErr
        'Neste caso só Retorna Erro fornecido pelo Vb
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142458)
    
    End Select
       
    Exit Function

End Function

Private Sub ParcelasPagto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Private Sub ParcelasPagto_GotFocus()

    iNumParcelas = StrParaInt(ParcelasPagto.Text)
    Call MaskEdBox_TrataGotFocus(ParcelasPagto, iAlterado)

End Sub
Private Sub ParcelasRecebto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ParcelasRecebto_GotFocus()

    iNumParcelas = StrParaInt(ParcelasRecebto.Text)
    Call MaskEdBox_TrataGotFocus(ParcelasRecebto, iAlterado)

End Sub

Private Sub ParcelasRecebto_Validate(Cancel As Boolean)
'Valida as informações contidas em parcelas Recebto
Dim lErro As Long

On Error GoTo Erro_ParcelasRecebto_Validate
    
    If iNumParcelas = StrParaInt(ParcelasRecebto.Text) Then Exit Sub
      
    If Len(Trim(ParcelasRecebto.Text)) = 0 Then
        
        Call Grid_Limpa(objGridParcelas)
    
    Else
        If StrParaInt(ParcelasRecebto.Text) = 0 Then gError 104050
        
        If StrParaInt(ParcelasRecebto.Text) >= (NUM_MAXIMO_PARC - 1) Then
            
            GridParcelas.Rows = StrParaInt(ParcelasRecebto.Text) + 2
            
            Call Grid_Inicializa(objGridParcelas)
        
        End If
        
        'Carrega os Parcelamentos do Grid
        Call Carrega_Parcelas(StrParaInt(ParcelasRecebto.Text))
        
        'Calcula e verifica se o Percentual de Parcelas é nemor que 100%
        Call Calcular_Total_Parcelas
           
    End If
                
    Exit Sub

Erro_ParcelasRecebto_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 104050
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_ZERO_EM_RECEBIMENTO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142459)

        End Select
        
   Exit Sub
    
End Sub

Sub Carrega_Parcelas(iNumParcelasRecebto As Integer)
'Carrega as Parcelas Relacionada a um parcelamento
Dim iIndice As Integer
Dim dSomaPercentual As Double
Dim dValor As Double
Dim dTesteValor As Double
Dim dResto As Double
Dim dSoma As Double

    Call Grid_Limpa(objGridParcelas)
    'Carrega no Grid o Numero de Pacelas de Rcebimento de parcelas
    'Preenche até o numero de Parcelas para Recebimento
    For iIndice = 1 To iNumParcelasRecebto
        
        GridParcelas.TextMatrix(iIndice, iGrid_Parcela_Col) = iIndice
        GridParcelas.TextMatrix(iIndice, iGrid_Recebimento_Col) = Format((1 / iNumParcelasRecebto), "Percent")
        dResto = PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Recebimento_Col))
        dSomaPercentual = dSomaPercentual + PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Recebimento_Col))
        
        'o Primeiro recebimento tem que ser = a defasagem para a primeira
        If iIndice = 1 Then
            If StrParaInt(DefasagemPagtoVista.Text) > 0 Then
                GridParcelas.TextMatrix(iIndice, iGrid_IntervalosPagamentos_Col) = DefasagemPagtoVista.Text
            Else
                GridParcelas.TextMatrix(iIndice, iGrid_IntervalosPagamentos_Col) = 0
            End If
        
        Else
            'se não for a primeira parcela então receber em intervalos de 30 dias
            GridParcelas.TextMatrix(iIndice, iGrid_IntervalosPagamentos_Col) = 30
            
        End If
            
    Next
    
    'Acerto da última parcela para ficar com 100% no somatório das percentagem
    dValor = dSomaPercentual - PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Recebimento_Col))
    GridParcelas.TextMatrix(iNumParcelasRecebto, iGrid_Recebimento_Col) = Format((1 - dValor) + dResto, "Percent")
   
   'Atualizar Linhas existentes no Grid Para o Numero de Parcelas
    objGridParcelas.iLinhasExistentes = iNumParcelasRecebto
    
    Exit Sub
    
End Sub

Sub Calcular_Total_Parcelas()
'Calcula se o somatório das porcentagem das parcelas do grid é maior que 100%
Dim dValorPercentual As Double
Dim iIndice As Integer
    
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        dValorPercentual = dValorPercentual + PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Recebimento_Col))
    Next
    
    'Imprime na Tela a Porcentagem contida no grid
    PercTotal.Caption = Format(dValorPercentual, "Percent")
   
    Exit Sub

End Sub

Private Sub TaxaParcelamento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TaxaParcelamento_Validate(Cancel As Boolean)
'Valida a Taxa de Parcelamnto
Dim lErro As Long
 
On Error GoTo Erro_TaxaParcelamento_Validate

    'se Taxa de Pracelamnto for zero sai do Validate
    If Len(Trim(TaxaParcelamento.Text)) = 0 Then Exit Sub
    
    'Verefica se a Porcentagem é Valida
    lErro = Porcentagem_Critica2(TaxaParcelamento.Text)
    If lErro <> SUCESSO Then gError 104060
    
    Exit Sub
    
Erro_TaxaParcelamento_Validate:

    Cancel = True
    
    Select Case gErr
   
        Case 104060
        'Erro Tratado Dentro da Função
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142460)

        End Select
        
   Exit Sub
    
End Sub
Private Sub ValorMinimo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Private Sub ValorMinimo_Validate(Cancel As Boolean)
'Valida as Informações contidas em Valor Minimo, verificando se o valor á positivo
Dim lErro As Long

On Error GoTo Erro_ValorMinimo_Validate

    'Verifica se Valor Minimo foi preenchidO
    If Len(Trim(ValorMinimo.Text)) = 0 Then Exit Sub

    lErro = Valor_Positivo_Critica(ValorMinimo.Text)
    If lErro <> SUCESSO Then gError 104062

    Exit Sub

Erro_ValorMinimo_Validate:

    Cancel = True

    Select Case gErr

        Case 104062
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142461)

    End Select
    
    Exit Sub

End Sub
Private Sub JurosParcelas_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub JurosParcelas_Validate(Cancel As Boolean)
'Valida  as Informações contidas em juros de Parcelas verificando se a porcentagem é Valida
Dim lErro As Long
 
On Error GoTo Erro_JurosParcelas_Validate
    
    If Len(Trim(JurosParcelas.Text)) = 0 Then Exit Sub
     
    lErro = Porcentagem_Critica2(JurosParcelas.Text)
    If lErro <> SUCESSO Then gError 104069
    
    Exit Sub

Erro_JurosParcelas_Validate:

    Cancel = True

    Select Case gErr
    
        Case 104069
        'Erro tratado dentro da Função Chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142462)
         
    End Select
    
    Exit Sub

End Sub

Private Sub ParcelamentoLoja_Click()
'Se o parcelamnto for pela Loja então parcela de Recebimento =1

'    If StrParaInt(ParcelasRecebto.Text) <> 1 Then
'
'        ParcelasRecebto.Text = 1
'
'        Call ParcelasRecebto_Validate(bSGECancelDummy)
'
'    End If
'    'Numero de Parcelas domGrid Fica igual a um
'    iNumParcelas = 1
    
    iAlterado = REGISTRO_ALTERADO
        
End Sub
Private Sub ParcelamentoAdm_Click()
'Verifica se om Parcelamento vai ser pela Admnistradoras então parcelas recbto ficará com o Mesmo Valor de parcelas de Pagamento

    If Len(Trim(ParcelasRecebto.Text)) = 0 Then Exit Sub
    
    If StrParaInt(ParcelasRecebto.Text) <> StrParaInt(ParcelasPagto.Text) Then
        
        ParcelasRecebto.Text = ParcelasPagto.Text
        
        Call ParcelasRecebto_Validate(bSGECancelDummy)
    
    End If
    'Atualiza a Varialvel com o Numero de Parcelas do Grid
    iNumParcelas = StrParaInt(ParcelasRecebto.Text)

End Sub

Private Sub AtivoParc_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub GridParcelas_Click()

    Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        'Variavel não definida
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Private Sub GridParcelas_EnterCell()
    'Parametro não opcional
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

Private Sub GridParcelas_LostFocus()

    Call Grid_Libera_Foco(objGridParcelas)

End Sub
Private Sub GridParcelas_RowColChange()

    Call Grid_RowColChange(objGridParcelas)

End Sub

Private Sub GridParcelas_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridParcelas)

End Sub
Private Sub GridParcelas_Scroll()

    Call Grid_Scroll(objGridParcelas)

End Sub
Private Sub PercRecebimento_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PercRecebimento_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)


End Sub

Private Sub PercRecebimento_KeyPress(KeyAscii As Integer)


    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub PercRecebimento_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = PercRecebimento
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub IntervaloParcela_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub IntervaloParcela_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub IntervaloParcela_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub IntervaloParcela_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = IntervaloParcela
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Agencia_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Bairro_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CEP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Cidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodBanco_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub


Private Sub Contato_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DefasagemPagtoVista_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Email_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Endereco_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Fax_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub



Private Sub Telefone1_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Telefone2_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub Parcelamentos_Validate(Cancel As Boolean)
'Valida as Informações contidas na Combo de Parcelamento
Dim lErro As Long

On Error GoTo Erro_Parcelamentos_Validate
    
    'Verifica se Nenhum parcelamento foi Informado
    If Len(Trim(Parcelamentos.Text)) = 0 Then Exit Sub
    'Verfica se não existe ninguem Marcado
    If Parcelamentos.ListIndex <> -1 Then Exit Sub
    
    'Verifica se existe o item na Combo Parcelamnetos, se existir seleciona o item
    lErro = Combo_Item_Igual(Parcelamentos)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 104044
    
    Exit Sub

Erro_Parcelamentos_Validate:

    Cancel = True
   
    Select Case gErr
        'Erro Tratado Dentro da Função Chamdora
        Case 104044
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142463)

        End Select
        
   Exit Sub
    
End Sub

Private Sub ParcelasPagto_Validate(Cancel As Boolean)
'Valida as Informações em Parcela Pagto

On Error GoTo Erro_ParcelasPagto_Validate

    'Verfica se ParcelasPagto for igual a Zero ou paracelemanto loja Marcado sai do Validate
    If Len(Trim(ParcelasPagto.Text)) = 0 Or ParcelamentoLoja.Value = True Then Exit Sub
    
    'Verifica se iNumParcelas for diferente de Parcelas pagto então sai do Validate
    If iNumParcelas <> StrParaInt(ParcelasPagto.Text) Then Exit Sub
    
    'Verifica se Parcelas Recebto é Diferente de Parcelas Pagto
    If StrParaInt(ParcelasRecebto.Text) <> StrParaInt(ParcelasPagto.Text) Then
        
        ParcelasRecebto.Text = ParcelasPagto.Text
        'Chamar o Validate de Parcela Recebto
        Call ParcelasRecebto_Validate(bSGECancelDummy)
    
    End If
    
    Exit Sub

Erro_ParcelasPagto_Validate:

    Cancel = True
   
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142464)

        End Select
        
   Exit Sub

End Sub

Private Sub Desconto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Desconto_Validate(Cancel As Boolean)
'Valida as Informações em Desconto
Dim lErro As Long

On Error GoTo Erro_Desconto_Validate
    
    If Len(Trim(Desconto.Text)) = 0 Then Exit Sub
    'Valida a Procentagem
    lErro = Porcentagem_Critica2(Desconto.Text)
    If lErro <> SUCESSO Then gError 104061
    
    Exit Sub

Erro_Desconto_Validate:

    Cancel = True

    Select Case gErr
    
        Case 104061
            'Erro tratado dentro da Função
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142465)
                   
    End Select
    
    Exit Sub
 
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Verifica qual a coluna atual do Grid
        Select Case objGridInt.objGrid.Col

            'PercRecebimento
            Case iGrid_Recebimento_Col
                lErro = Saida_Celula_PercRecebimento(objGridInt)
                If lErro <> SUCESSO Then gError 104064

            'IntervaloParcela
            Case iGrid_IntervalosPagamentos_Col
                lErro = Saida_Celula_IntervaloParcela(objGridInt)
                If lErro <> SUCESSO Then gError 104065


        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 104066

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 104064 To 104066
            ' Erros Tratados Dentro da Função
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142466)

    End Select

    Exit Function

End Function


Private Sub Estado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Estado_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Estado_Validate(Cancel As Boolean)
'Valida as Informações na Combo de Estados
Dim lErro As Long

On Error GoTo Erro_Estado_Validate

    'Verifica se foi preenchido o Estado
    If Len(Trim(Estado.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o item selecionado na ComboBox Estado
    If Estado.ListIndex <> -1 Then Exit Sub

    'Verifica se existe o item no Estado, se existir seleciona o item
    lErro = Combo_Item_Igual_CI(Estado)
    If lErro <> SUCESSO And lErro <> 58583 Then gError 104042

    'Nao existe o item na ComboBox Estado
    If lErro <> SUCESSO Then gError 104043

    Exit Sub

Erro_Estado_Validate:

    Cancel = True

    Select Case gErr

    Case 104042
    'Erro tratado na rotina chamadora

    Case 104043
        Call Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", gErr, Estado.Text)

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142467)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_PercRecebimento(objGridInt As AdmGrid) As Long
'Função de Saida de Célula de grid de Parcelamentos
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_PercRecebimento

    Set objGridParcelas.objControle = PercRecebimento
    
    'Se necessário cria uma nova linha no Grid
    If Len(Trim(PercRecebimento.Text)) > 0 Then
    
        lErro = Porcentagem_Critica(PercRecebimento.Text)
        If lErro <> SUCESSO Then gError 104067
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 104068
    
    Call Calcular_Total_Parcelas
    
    Saida_Celula_PercRecebimento = SUCESSO

    Exit Function

Erro_Saida_Celula_PercRecebimento:

    Saida_Celula_PercRecebimento = gErr

    Select Case gErr

        Case 104067 To 104068
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142468)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_IntervaloParcela(objGridInt As AdmGrid) As Long
'Saida de Célula do Contole IntervaloParcela
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_IntervaloParcela

    Set objGridParcelas.objControle = IntervaloParcela
    
    If Len(Trim(IntervaloParcela.Text)) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(IntervaloParcela.Text)
        If lErro <> SUCESSO Then gError 104054
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 104055
    
    Saida_Celula_IntervaloParcela = SUCESSO

    Exit Function

Erro_Saida_Celula_IntervaloParcela:

    Saida_Celula_IntervaloParcela = gErr

    Select Case gErr
        Case 104053
            Call Rotina_Erro(vbOKOnly, "ERRO_INTERVALO_PARCELA_NAO_PREENCHIDO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridParcelas)
      
        Case 104054
            'Erro Tratado Dentro da Função
        
        Case 104055
            Call Grid_Trata_Erro_Saida_Celula(objGridParcelas)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142469)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoBanco = Nothing
    Set objGridParcelas = Nothing
    Set gcolAdmMeioPagtoCondPagto = Nothing
    
    'Libera a referência da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Private Sub Pais_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Private Sub Pais_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Pais_Validate(Cancel As Boolean)
'Valida as Informações contidas na Combo Pais
Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Pais_Validate

    'Verifica se foi preenchida a Combo Pais
    If Len(Trim(Pais.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o item selecionado na ComboBox Pais
    If Pais.ListIndex <> -1 Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Pais, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 104041

    'Não existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then gError 104039
    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 104040

    Exit Sub

Erro_Pais_Validate:

    Cancel = True

    Select Case gErr

        Case 104039
            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO1", gErr, Trim(Pais.Text))
    
        Case 104041
        'Erro tratado na rotina chamadora
    
        Case 104040
            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO", gErr, iCodigo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142470)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
'Botão que limpa a Tela
Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    'Limpa todo os Contoles menos Combo e Label's
    Call Teste_Salva(Me, iAlterado)
    'Limpa Combo e Label's
    Call Limpa_Tela_AdmMeioPagto

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Exit Sub

Erro_Botaolimpar_Click:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 104088
        'Erro tratado na rotina chamadora

        Case Else
           lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142471)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravarParc_Click()
'Botão que Salva as Informações do parcelamento em quetão na Tela

Dim lErro As Long
Dim bAchou As Boolean
Dim iIndice As Integer
Dim iCont As Integer
Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim vbResultado As VbMsgBoxResult
Dim sNomeParcelamento As String

On Error GoTo Erro_BotaoGravarParc_Click
    
    'Valida o Grid de Parcelas
    lErro = Valida_Grid_Parcelas()
    If lErro <> SUCESSO Then Exit Sub

    'Procedimento Combo Editável, Procedimnto p/ gravar um novo Parcelamento
    If Len(Trim(Parcelamentos.Text)) = 0 Then gError 104056
    
    If Len(Trim(ParcelasPagto.Text)) = 0 Then gError 104057
    
    If TipoMeioPagto.ListIndex > -1 Then
    
        'se se tratar dos meios de pagamento outros ou vale ticket que só admitem um parcelamento ativo
        If TipoMeioPagto.ItemData(TipoMeioPagto.ListIndex) = TIPOMEIOPAGTOLOJA_OUTROS Or TipoMeioPagto.ItemData(TipoMeioPagto.ListIndex) = TIPOMEIOPAGTOLOJA_VALE_TICKET Then
        
            'Buscar na coleção global um parcelamento com o mesmo Nome
            For Each objAdmMeioPagtoCondPagto In gcolAdmMeioPagtoCondPagto
        
            
                    'se tem um outro parcelamento ativo ==> avisa que vai desativa-lo
                    If objAdmMeioPagtoCondPagto.iAtivo = ADMMEIOPAGTOCONDPAGTO_ATIVO And AtivoParc.Value = MARCADO And objAdmMeioPagtoCondPagto.sNomeParcelamento <> Parcelamentos.Text Then
        
                        Call Rotina_Aviso(vbOKOnly, "AVISO_ALTERACAO_STATUS_ADMMEIOPAGTOCONDPAGTO", objAdmMeioPagtoCondPagto.sNomeParcelamento)
                        objAdmMeioPagtoCondPagto.iAtivo = ADMMEIOPAGTOCONDPAGTO_INATIVO
                
                    End If
            
            Next
            
        End If
        
        'so cartao de debito pode ser pre-datado
        If TipoMeioPagto.ItemData(TipoMeioPagto.ListIndex) <> TIPOMEIOPAGTOLOJA_CARTAO_DEBITO And PreDatado.Value = MARCADO Then gError 214530


    End If

    'Buscar na coleção global um parcelamento com o mesmo Nome
    For Each objAdmMeioPagtoCondPagto In gcolAdmMeioPagtoCondPagto

        If objAdmMeioPagtoCondPagto.sNomeParcelamento = Parcelamentos.Text Then

            bAchou = True
            Exit For

        End If
        
        If UCase(Parcelamentos.Text) = "AVISTA" Or UCase(Parcelamentos.Text) = "A VISTA" Or _
           UCase(Parcelamentos.Text) = "ÀVISTA" Or UCase(Parcelamentos.Text) = "À VISTA" Then
        
            If UCase(objAdmMeioPagtoCondPagto.sNomeParcelamento) = "AVISTA" Or _
               UCase(objAdmMeioPagtoCondPagto.sNomeParcelamento) = "A VISTA" Or _
               UCase(objAdmMeioPagtoCondPagto.sNomeParcelamento) = "ÀVISTA" Or _
               UCase(objAdmMeioPagtoCondPagto.sNomeParcelamento) = "À VISTA" Then
                gError 126050
                
            End If
            
        End If
         
    Next

    If Not bAchou Then
        'Criando novo Obj
        Set objAdmMeioPagtoCondPagto = New ClassAdmMeioPagtoCondPagto
        objAdmMeioPagtoCondPagto.sNomeParcelamento = Parcelamentos.Text
        
    End If
    
    'o objAdmMeioPagtoCondPagto recebe as Informações contidas na tela
    objAdmMeioPagtoCondPagto.iFilialEmpresa = giFilialEmpresa
    objAdmMeioPagtoCondPagto.dDesconto = PercentParaDbl(Desconto.FormattedText)
    objAdmMeioPagtoCondPagto.dJuros = PercentParaDbl(JurosParcelas.FormattedText)
    objAdmMeioPagtoCondPagto.dTaxa = PercentParaDbl(TaxaParcelamento.FormattedText)
    objAdmMeioPagtoCondPagto.dValorMinimo = StrParaDbl(ValorMinimo.Text)
    objAdmMeioPagtoCondPagto.iAdmMeioPagto = StrParaInt(Codigo.Text)
    objAdmMeioPagtoCondPagto.dJuros = PercentParaDbl(JurosParcelas.FormattedText)
    objAdmMeioPagtoCondPagto.iNumParcelas = StrParaInt(ParcelasPagto.Text)
    objAdmMeioPagtoCondPagto.iParcelasRecebto = StrParaInt(ParcelasRecebto.Text)
    objAdmMeioPagtoCondPagto.iAtivo = AtivoParc.Value
    objAdmMeioPagtoCondPagto.iPreDatado = PreDatado.Value
        
    'Verificar se o juro são por conta da Administradora ou Loja
    If ParcelamentoAdm.Value = True Then
        
        objAdmMeioPagtoCondPagto.iJurosParcelamento = JUROS_ADM
    
    End If
        
    If ParcelamentoLoja.Value = True Then
        
        objAdmMeioPagtoCondPagto.iJurosParcelamento = JUROS_LOJA
    
    End If
    
    lErro = Move_GridParcelas_Memoria(objAdmMeioPagtoCondPagto)
    If lErro <> SUCESSO Then gError 104089
     
    If Not bAchou Then
   
        'Adcionar um novo Obj na Coleção
        gcolAdmMeioPagtoCondPagto.Add objAdmMeioPagtoCondPagto
    
    End If
    
    'Preenche a Combo Parcelamentos
    Call Carrega_ComboParcelas
    
    'Função Limpar Psarcelamentos
    Call Limpar_Parcelamentos
        
    'Zera o Numero de Parcelas
    iNumParcelas = 0
    
    Exit Sub

Erro_BotaoGravarParc_Click:

    Select Case gErr
    
        Case 104056
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_PARCELAMENTO_SELECIONADO", gErr)
        
        Case 104057
            Call Rotina_Erro(vbOKOnly, "ERRO_NAO_HA_PARCELAS_PGTO", gErr)
        
        Case 104089
            
        Case 115021
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAMENTO_EXISTENTE", gErr)
            
        Case 115030
            
        Case 126050
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAMENTO_AVISTA_CADASTRADO", gErr)
            
        Case 214530
            Call Rotina_Erro(vbOKOnly, "ERRO_SO_CARTAO_DEBITO_PREDATADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142472)
            
        End Select
       
    Exit Sub
    
End Sub

Function Move_GridParcelas_Memoria(objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto) As Long

Dim iIndice As Integer
Dim objAdmMeioPagtoParcelas As ClassAdmMeioPagtoParcelas
    
On Error GoTo Erro_Move_GridParcelas_Memoria
    
    'Inicializar a Coleção
    Set objAdmMeioPagtoCondPagto.colParcelas = New Collection
    
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        
        Set objAdmMeioPagtoParcelas = New ClassAdmMeioPagtoParcelas
        
        objAdmMeioPagtoParcelas.dPercRecebimento = PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Recebimento_Col))
        objAdmMeioPagtoParcelas.iIntervaloRecebimento = StrParaInt(GridParcelas.TextMatrix(iIndice, iGrid_IntervalosPagamentos_Col))
         
        'Adcionar na Coleção de Parcelas
        objAdmMeioPagtoCondPagto.colParcelas.Add objAdmMeioPagtoParcelas
    
    Next
                        
    'Verifica a Soma dos Percentuais das Parcelas Se for Maior
    If PercentParaDbl(PercTotal.Caption) > 1 Then
        gError 104179
                        
    ElseIf PercentParaDbl(PercTotal.Caption) < 1 Then gError 104269
    
    End If
    
    Move_GridParcelas_Memoria = SUCESSO

    Exit Function
    
Erro_Move_GridParcelas_Memoria:
    
    Move_GridParcelas_Memoria = gErr
    
    Select Case gErr
    
        Case 104179
            Call Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_PARCELAS_INVALIDO_MAIOR", gErr)
            
        Case 104269
            Call Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_PARCELAS_INVALIDO_MENOR", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142473)
            
    End Select
       
    Exit Function
    
        
End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim objAdmMeioPagto As New ClassAdmMeioPagto

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 104090

    Call Limpa_Tela_AdmMeioPagto

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 104090
            'Erro tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142474)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim objAdmMeioPagtoBD As New ClassAdmMeioPagto
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim iAtivo As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'se estiver operando no backoffice ==> nao pode gravar
    If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then gError 110085

    'Verifica se o Código do meio de pagto é Carnê se for,
    If StrParaInt(Codigo.Text) = MEIO_PAGAMENTO_CARNE Then
      
        If TipoMeioPagto.ListIndex = -1 Then
            objAdmMeioPagto.iCodigo = MEIO_PAGAMENTO_CARNE
            gError 107639
        End If
        
        'Verificar se o Tipo de pagto é igual ao tipo Carnê se não erro
        If TipoMeioPagto.ItemData(TipoMeioPagto.ListIndex) <> TIPOMEIOPAGTOLOJA_CARNE Then
            'passa para o objAdmMeioPagto o Código do meio que Não pode ser Alterado
            objAdmMeioPagto.iCodigo = MEIO_PAGAMENTO_CARNE
            gError 107639
        End If
    
    End If
    
    'verifica preenchimento do codigo
    If Len(Trim(Codigo.Text)) = 0 Then gError 104091

    'verifica preenchimento do nome
    If Len(Trim(Nome.Text)) = 0 Then gError 104092

    'Preenche objAdmMeioPagto com Dados Obtidos na Tela
    lErro = Move_Tela_Memoria(objAdmMeioPagto)
    If lErro <> SUCESSO Then gError 104095

    'Verifica se o Códigos Está relacionado aos Códigos que vem Previamnto Cadastrados por Default, não podem ser  Gravados.
    If StrParaInt(Codigo.Text) = MEIO_PAGAMENTO_DINHEIRO Or StrParaInt(Codigo.Text) = MEIO_PAGAMENTO_CHEQUE Or StrParaInt(Codigo.Text) = MEIO_PAGAMENTO_TROCA Or StrParaInt(Codigo.Text) = MEIO_PAGAMENTO_CONTRAVALE Then
    
        'passa para o objAdmMeioPagto o Código do meio que Não pode ser Alterado
        objAdmMeioPagto.iCodigo = StrParaInt(Codigo.Text)
        objAdmMeioPagto.iFilialEmpresa = giFilialEmpresa
        
        Call Rotina_Aviso(vbOKOnly, "AVISO_GRAVACAO_STATUS_ADMMEIOPAGTO")
        
        lErro = CF("AdmMeioPagto_Grava_Ativo", objAdmMeioPagto)
        If lErro <> SUCESSO Then gError 107638
        
    Else
      
        'verifica se nada foi Selecionado na Combo
        If TipoMeioPagto.ListIndex = -1 Then gError 104093
       
        'Verifica se o Tipo Meio de Pagamento é Outros
        If TipoMeioPagto.ItemData(TipoMeioPagto.ListIndex) = TIPOMEIOPAGTOLOJA_OUTROS Then
        
            For Each objAdmMeioPagtoCondPagto In gcolAdmMeioPagtoCondPagto
                iAtivo = iAtivo + objAdmMeioPagtoCondPagto.iAtivo
            Next
            
            If iAtivo > 1 Then gError 107716
            
        End If
    
        'Verifica se o Tipo Meio de Pagamento é Vale Ticket
        If TipoMeioPagto.ItemData(TipoMeioPagto.ListIndex) = TIPOMEIOPAGTOLOJA_VALE_TICKET Then
        
            For Each objAdmMeioPagtoCondPagto In gcolAdmMeioPagtoCondPagto
                iAtivo = iAtivo + objAdmMeioPagtoCondPagto.iAtivo
            Next
            
            If iAtivo > 1 Then gError 110087
        
        End If
       
        'Alterado por cyntia para incluir FilialEmpresa como parâmetro
        lErro = Trata_Alteracao(objAdmMeioPagto, objAdmMeioPagto.iCodigo, objAdmMeioPagto.iFilialEmpresa)
        If lErro <> SUCESSO Then gError 104193
    
        objAdmMeioPagtoBD.iCodigo = objAdmMeioPagto.iCodigo
        objAdmMeioPagtoBD.iFilialEmpresa = objAdmMeioPagto.iFilialEmpresa
        
    
        lErro = CF("AdmMeioPagto_Le", objAdmMeioPagtoBD)
        If lErro <> SUCESSO And lErro <> 104017 Then gError 126054

        If lErro = SUCESSO And objAdmMeioPagtoBD.dtDataLog <> DATA_NULA Then Call Rotina_Aviso(vbOKOnly, "AVISO_ALTERACAO_ADMMEIOPAGTO")

        lErro = CF("AdmMeioPagto_Grava", objAdmMeioPagto)
        If lErro <> SUCESSO Then gError 104097
    
        'Exclui da ListBox
        Call Administradoras_Exclui(objAdmMeioPagto)
    
        'Inclui na ListBox
        Call Administradoras_Inclui(objAdmMeioPagto)

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
        
        Case 104095, 104097, 104193, 107638, 126054
        'Erro tratado dentro da Função chamadora
        
        Case 104094
            Call Rotina_Erro(vbOKOnly, "ERRO_TAXA_NAO_PREENCHIDA", gErr)

        Case 104091
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 104092
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_NAO_PREENCHIDO", gErr)

        Case 104093
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_NAO_PREENCHIDO", gErr)
               
        Case 107639
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_ALTERACAO_TIPO_DIFERENTE", gErr, objAdmMeioPagto.iCodigo)

        Case 107716
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_OUTROS_PARCELAMENTO_UNICO", gErr)

        Case 110087
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_TICKET_PARCELAMENTO_UNICO", gErr)

        Case 110085
            Call Rotina_Erro(vbOKOnly, "ERRO_GRAVACAO_ADMMEIOPAGTO_BACKOFFICE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142475)

    End Select

    Exit Function

End Function

Function Move_Tela_Memoria(objAdmMeioPagto As ClassAdmMeioPagto) As Long
'Move os dados da tela para o objAdmMeioPagto

Dim lErro As Long
Dim objTipoMeioPagto As New ClassTMPLoja
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_Move_Tela_Memoria
      
    If Ativo.Value = vbUnchecked Then
        objAdmMeioPagto.iAtivo = ADMMEIOPAGTO_INATIVO
        For Each objAdmMeioPagtoCondPagto In gcolAdmMeioPagtoCondPagto
            objAdmMeioPagtoCondPagto.iAtivo = ADMMEIOPAGTOCONDPAGTO_INATIVO
        Next
        
    Else
        objAdmMeioPagto.iAtivo = ADMMEIOPAGTO_ATIVO
    End If
      
    objAdmMeioPagto.iFilialEmpresa = giFilialEmpresa
   
    'Move o Codigo Para Memoria
    objAdmMeioPagto.iCodigo = StrParaInt(Codigo.Text)
    
    'Move o Nome Para Memoria
    objAdmMeioPagto.sNome = Nome.Text
    
    'Move Taxa a Vista para Memória
    objAdmMeioPagto.dTaxaVista = PercentParaDbl(TaxaVista.FormattedText)
    
    'Move taxa parcelado para Memória
    objAdmMeioPagto.dTaxaParcelado = PercentParaDbl(TaxaPrazo.FormattedText)
    
    objAdmMeioPagto.iRede = Codigo_Extrai(Rede.Text)
    objAdmMeioPagto.iTipoMeioPagto = Codigo_Extrai(TipoMeioPagto.Text)
    objAdmMeioPagto.iDefasagemPagtoVista = StrParaInt(DefasagemPagtoVista.Text)
    objAdmMeioPagto.iContaCorrenteInterna = Codigo_Extrai(ContaCorrenteInterna.Text)
    objAdmMeioPagto.iGeraTituloRec = CheckNaoTitRec.Value
    
    'Procurar o Tipo de Pagamento na Coleção Geral
    For Each objTipoMeioPagto In gcolTipoMeioPagto
        'Se Encontrou
        If objTipoMeioPagto.iTipo = objAdmMeioPagto.iTipoMeioPagto Then
            
            Exit For
        
        End If
    
    Next
    
     If objTipoMeioPagto.iEnvioPagamento = ENVIO_ENDERECO Then
        
        objAdmMeioPagto.lEndereco = ENVIO_ENDERECO
        objAdmMeioPagto.objEndereco.sEndereco = Endereco.Text
        objAdmMeioPagto.objEndereco.sBairro = Bairro.Text
        objAdmMeioPagto.objEndereco.sCidade = Cidade.Text
        objAdmMeioPagto.objEndereco.sCEP = CEP.ClipText
        objAdmMeioPagto.objEndereco.sSiglaEstado = Estado.Text
        objAdmMeioPagto.objEndereco.iCodigoPais = Codigo_Extrai(Pais.Text)
        objAdmMeioPagto.objEndereco.sTelefone1 = Telefone1.Text
        objAdmMeioPagto.objEndereco.sTelefone2 = Telefone2.Text
        objAdmMeioPagto.objEndereco.sFax = Fax.Text
        objAdmMeioPagto.objEndereco.sContato = Contato.Text
        objAdmMeioPagto.objEndereco.sEmail = Email.Text
    
    Else
         
        FrameBanco.Visible = True
        objAdmMeioPagto.sAgencia = Agencia.Text
        objAdmMeioPagto.iCodBanco = StrParaInt(CodBanco.ClipText)
      
    End If
    
    'Preenche a Coleção do Obj com a Coleção Global
    Set objAdmMeioPagto.colCondPagtoLoja = gcolAdmMeioPagtoCondPagto
        
    Move_Tela_Memoria = SUCESSO
    
    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142476)

    End Select

    Exit Function

End Function

Private Sub Administradoras_Inclui(objAdmMeioPagto As ClassAdmMeioPagto)
'Adiciona na ListBox informações da Administradora
Dim iIndice As Integer

        Administradoras.AddItem objAdmMeioPagto.sNome
        Administradoras.ItemData(Administradoras.NewIndex) = objAdmMeioPagto.iCodigo

    Exit Sub

End Sub

Private Sub Administradoras_Exclui(objAdmMeioPagto As ClassAdmMeioPagto)
'Percorre a ListBox de Administradoras para remover a informação em questão

Dim iIndice As Integer
    'Percorre a listBox
    For iIndice = 0 To Administradoras.ListCount - 1
        'se o Codigo For Igual então é Excluida da List
        If Administradoras.ItemData(iIndice) = objAdmMeioPagto.iCodigo Then
            Administradoras.RemoveItem (iIndice)
            Exit For
        End If
     Next

End Sub

Private Sub BotaoExcluirParc_Click()

Dim lErro As Long
Dim vbResultado As VbMsgBoxResult
Dim iIndice As Integer
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_BotaoExcluirParc_Click

    'Se não existir nada Selecionado na Combo erro
    If Parcelamentos.ListIndex = -1 Then gError 104059
    
    Set objAdmMeioPagtoCondPagto = gcolAdmMeioPagtoCondPagto.Item(Parcelamentos.ListIndex + 1)
    
    lErro = CF("AdmMeioPagtoCondPagto_Le_Parcelamento", objAdmMeioPagtoCondPagto)
    If lErro <> SUCESSO And lErro <> 107297 Then gError 117570
    
    'se o parcelamento está cadastrado e já foi transferido ==> nao pode excluir
    If lErro = SUCESSO And objAdmMeioPagtoCondPagto.dtDataLog <> DATA_NULA Then gError 117571
    
    'Pergunta se Deseja Realmente Excluir o Parcelamento
    vbResultado = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_EXCLUIR_PARCELAMENTO", Parcelamentos.Text)
    
    'Se sim
    If vbResultado = vbNo Then Exit Sub
       
    'Remove na Coleção Global o Parcelamento Selecionado
    gcolAdmMeioPagtoCondPagto.Remove (Parcelamentos.ListIndex + 1)
    
    'Remover da Combo o Parcelamento Selecionado
    Parcelamentos.RemoveItem (Parcelamentos.ListIndex)
    
    'Carrega a Combo Parcelamentos
    Call Carrega_ComboParcelas
    
    'Limpar os Parcelamentos
    Call Limpar_Parcelamentos
    
    Exit Sub
 
Erro_BotaoExcluirParc_Click:

    Select Case gErr
    
        Case 104059
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAMENTO_NAO_SELECIONADO", gErr)
      
        Case 117570
    
        Case 117571
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTOCONDPAGTO_DATALOG", gErr, objAdmMeioPagtoCondPagto.iAdmMeioPagto, objAdmMeioPagtoCondPagto.iParcelamento)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142477)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'se estiver operando no backoffice ==> nao pode gravar
    If giLocalOperacao = LOCALOPERACAO_BACKOFFICE Then gError 110086

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o codigo foi preenchido
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 104144
    
    'Verifica se o Códigos Está relacionado aos Códigos que vem Previamnto Cadastrados por Default, não podem ser  Excluidos.
    If Codigo.Text = MEIO_PAGAMENTO_DINHEIRO Or Codigo.Text = MEIO_PAGAMENTO_CHEQUE Or Codigo.Text = MEIO_PAGAMENTO_CARNE Or Codigo.Text = MEIO_PAGAMENTO_TROCA Or Codigo.Text = MEIO_PAGAMENTO_CONTRAVALE Then gError 107637
    
    'Lê o Codigo e Filial Empresa e Carrega no objAdmMeioPagto
    objAdmMeioPagto.iCodigo = StrParaInt(Codigo.Text)
    objAdmMeioPagto.iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("AdmMeioPagto_Le", objAdmMeioPagto)
    If lErro <> SUCESSO And lErro <> 104017 Then gError 104145
    
    'Verifica se administradora não está cadastrada
    If lErro = 104017 Then gError 104146
    
    'Envia aviso perguntando se realmente deseja excluir administradora
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_MEIO_PAGAMENTO", objAdmMeioPagto.iCodigo)

    If vbMsgRes = vbYes Then

        'Carega a Coleção de Parcelamento Relacionados A Administradoras
        lErro = CF("AdmMeioPagtoCondPagto_Le", objAdmMeioPagto)
        If lErro <> SUCESSO And lErro <> 104086 Then gError 104194
    
        'Exclui Administradora
        lErro = CF("AdmMeioPagto_Exclui", objAdmMeioPagto)
        If lErro <> SUCESSO Then gError 104203

        'Exclui da List
        Call Administradoras_Exclui(objAdmMeioPagto)

        Call Limpa_Tela_AdmMeioPagto

        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)

        iAlterado = 0
        
    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 104144
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 104145, 104203, 104194
            'Erro tratado na rotina chamadora

        Case 104146
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_CADASTRADO", gErr, objAdmMeioPagto.iCodigo)

        Case 107637
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMMEIOPAGTO_NAO_PERMITE_EXCLUSAO", gErr, Codigo.Text)

        Case 110086
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ADMMEIOPAGTO_BACKOFFICE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142478)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

End Sub

'*************************************************************
'Função temporária para testar o Log gerado pela Tela
'*************************************************************

Private Sub TipoMeioPagto_GotFocus()
'Função que Serve para setar a Variavel Globál

Dim lErro As Long

On Error GoTo Erro_TipoMeioPagto_GotFocus

    'Quando For Carregado a Combo de Tipo meio de Pagamentos a Variável Recebe o Código do Tipo de Pagto
    giTipoMeioPagto = Codigo_Extrai(TipoMeioPagto.Text)
    
    Exit Sub
    
Erro_TipoMeioPagto_GotFocus:
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142479)

        End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is CodBanco Then
            Call LabelBanco_Click
    
        End If
        
    End If

End Sub

Private Function Valida_Grid_Parcelas() As Long
'Valida o conteúdo do grid de Parcelas

Dim iIndice As Integer
Dim lErro As Long
Dim dValorTotal As Double

On Error GoTo Erro_Valida_Grid_Parcelas
    
    'Verificar se o Grid está Vazio
    If objGridParcelas.iLinhasExistentes = 0 Then gError 104058
    
    'Para cada item do grid...
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
    
        'Verifica se a % de recebimento está preenchida
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Recebimento_Col))) = 0 Then gError 115031
        
        'Verifica se o Intervalo de Pagamento está preenchido
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_IntervalosPagamentos_Col))) = 0 Then gError 115032
        
    Next
    
    Valida_Grid_Parcelas = SUCESSO
    
    Exit Function
    
Erro_Valida_Grid_Parcelas:

    Select Case gErr
    
        Case 104058
            Call Rotina_Erro(vbOKOnly, "ERRO_NAO_HA_PARCELAS_RECEPTO", gErr)

        Case 115031
            Call Rotina_Erro(vbOKOnly, "ERRO_NAO_HA_PORCENTAGEM_RECEBTO", gErr)

        Case 115032
            Call Rotina_Erro(vbOKOnly, "ERRO_INTERVALO_ENTRE_PARCELAS_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142480)

        End Select

    Exit Function

End Function



Private Sub DefasagemPagtoVista_GotFocus()
    Call MaskEdBox_TrataGotFocus(DefasagemPagtoVista, iAlterado)
End Sub

Private Sub DefasagemPagtoVista_Validate(Cancel As Boolean)

Dim objAdmMeioPagtoCondPagto As New ClassAdmMeioPagtoCondPagto
Dim iIndice As Integer

    For Each objAdmMeioPagtoCondPagto In gcolAdmMeioPagtoCondPagto
        
        iIndice = iIndice + 1
        objAdmMeioPagtoCondPagto.colParcelas(iIndice).iIntervaloRecebimento = StrParaInt(DefasagemPagtoVista.Text)
        
    Next
    
    Exit Sub

End Sub


Private Sub LabelBanco_Click()

Dim objBanco As New ClassBanco
Dim colSelecao As New Collection

    objBanco.iCodBanco = StrParaInt(CodBanco.Text)
    
    Call Chama_Tela("BancoLista", colSelecao, objBanco, objEventoBanco)

    Exit Sub

End Sub

Private Sub objEventoBanco_evSelecao(obj1 As Object)

Dim objBanco As ClassBanco

    Set objBanco = obj1
    
    If objBanco.iCodBanco > 0 Then CodBanco.Text = objBanco.iCodBanco
        
    Me.Show
    
    Exit Sub

End Sub


