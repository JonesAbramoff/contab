VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TipoDeTributacaoOcx 
   ClientHeight    =   6675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6675
   ScaleWidth      =   9510
   Begin VB.Frame Frame9 
      Caption         =   "Código de Base de Cálculo do Crédito"
      Height          =   570
      Left            =   2280
      TabIndex        =   61
      Top             =   6030
      Width           =   7140
      Begin VB.ComboBox NatBCCred 
         Height          =   315
         Left            =   855
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   195
         Width           =   6210
      End
      Begin VB.Label Label9 
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
         Height          =   225
         Left            =   135
         TabIndex        =   62
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Identificação"
      Height          =   1050
      Left            =   135
      TabIndex        =   49
      Top             =   15
      Width           =   7005
      Begin VB.Frame Frame11 
         Height          =   420
         Left            =   4680
         TabIndex        =   55
         Top             =   180
         Width           =   2250
         Begin VB.OptionButton OptISS 
            Caption         =   "ISS"
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
            Left            =   1230
            TabIndex        =   4
            Top             =   135
            Width           =   900
         End
         Begin VB.OptionButton OptICMS 
            Caption         =   "ICMS"
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
            Left            =   165
            TabIndex        =   3
            Top             =   105
            Value           =   -1  'True
            Width           =   1350
         End
      End
      Begin VB.TextBox Descricao 
         Height          =   312
         Left            =   1770
         MaxLength       =   100
         TabIndex        =   5
         Top             =   645
         Width           =   5160
      End
      Begin VB.Frame Frame3 
         Height          =   420
         Left            =   2415
         TabIndex        =   50
         Top             =   180
         Width           =   2250
         Begin VB.OptionButton Entrada 
            Caption         =   "Entrada"
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
            Left            =   165
            TabIndex        =   1
            Top             =   105
            Value           =   -1  'True
            Width           =   990
         End
         Begin VB.OptionButton Saida 
            Caption         =   "Saída"
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
            Left            =   1230
            TabIndex        =   2
            Top             =   135
            Width           =   900
         End
      End
      Begin MSMask.MaskEdBox Tipo 
         Height          =   315
         Left            =   1770
         TabIndex        =   0
         Top             =   285
         Width           =   615
         _ExtentX        =   1085
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
      Begin VB.Label Label3 
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
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   825
         TabIndex        =   52
         Top             =   675
         Width           =   1065
      End
      Begin VB.Label TipoLabel 
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
         Left            =   30
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   51
         Top             =   300
         Width           =   1755
      End
   End
   Begin VB.Frame FrameISS 
      Caption         =   "ISSQN"
      Height          =   570
      Left            =   120
      TabIndex        =   47
      Top             =   4905
      Width           =   9300
      Begin VB.ComboBox ISSIndExigibilidade 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "TipoDeTributacaoOcx.ctx":0000
         Left            =   4395
         List            =   "TipoDeTributacaoOcx.ctx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   210
         Width           =   4845
      End
      Begin VB.ComboBox TipoTributacaoISS 
         Height          =   315
         ItemData        =   "TipoDeTributacaoOcx.ctx":00CE
         Left            =   555
         List            =   "TipoDeTributacaoOcx.ctx":00D0
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   180
         Width           =   2580
      End
      Begin VB.CheckBox IncideISS 
         Caption         =   "Incide"
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
         Left            =   7635
         TabIndex        =   48
         Top             =   480
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.CheckBox ReterISS 
         Caption         =   "Com retenção"
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
         Left            =   9000
         TabIndex        =   21
         Top             =   255
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Exigibilidade:"
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
         Index           =   44
         Left            =   3225
         TabIndex        =   64
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label8 
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
         Height          =   225
         Left            =   120
         TabIndex        =   56
         Top             =   225
         Width           =   510
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "CSLL"
      Height          =   570
      Left            =   120
      TabIndex        =   46
      Top             =   6030
      Width           =   2115
      Begin VB.CheckBox ReterCSLL 
         Caption         =   "Com retenção"
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
         Left            =   135
         TabIndex        =   27
         Top             =   225
         Width           =   1515
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "COFINS"
      Height          =   570
      Left            =   120
      TabIndex        =   45
      Top             =   4335
      Width           =   9315
      Begin VB.ComboBox TipoTributacaoCOFINS 
         Height          =   315
         Left            =   570
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   180
         Width           =   4830
      End
      Begin VB.CheckBox CreditaCOFINS 
         Caption         =   "Credita"
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
         Left            =   8280
         TabIndex        =   19
         Top             =   255
         Width           =   960
      End
      Begin VB.CheckBox ReterCOFINS 
         Caption         =   "Com retenção"
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
         Left            =   5445
         TabIndex        =   18
         Top             =   255
         Width           =   1515
      End
      Begin VB.Label Label7 
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
         Height          =   225
         Left            =   120
         TabIndex        =   54
         Top             =   225
         Width           =   510
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "PIS"
      Height          =   570
      Left            =   120
      TabIndex        =   44
      Top             =   3765
      Width           =   9315
      Begin VB.ComboBox TipoTributacaoPIS 
         Height          =   315
         Left            =   570
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   165
         Width           =   4830
      End
      Begin VB.CheckBox ReterPIS 
         Caption         =   "Com retenção"
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
         Left            =   5445
         TabIndex        =   15
         Top             =   225
         Width           =   1515
      End
      Begin VB.CheckBox CreditaPIS 
         Caption         =   "Credita"
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
         Left            =   8280
         TabIndex        =   16
         Top             =   255
         Width           =   960
      End
      Begin VB.Label Label6 
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
         Height          =   225
         Left            =   120
         TabIndex        =   53
         Top             =   210
         Width           =   510
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "IR"
      Height          =   570
      Left            =   5850
      TabIndex        =   42
      Top             =   5385
      Width           =   3570
      Begin VB.CheckBox IncideIR 
         Caption         =   "Com retenção"
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
         Left            =   1995
         TabIndex        =   26
         Top             =   240
         Width           =   1515
      End
      Begin MSMask.MaskEdBox IRAliquota 
         Height          =   315
         Left            =   990
         TabIndex        =   25
         Top             =   165
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   15
         Format          =   "#0.#0\%"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Alíquota:"
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
         Left            =   90
         TabIndex        =   43
         Top             =   195
         Width           =   795
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "INSS"
      Height          =   570
      Left            =   120
      TabIndex        =   39
      Top             =   5475
      Width           =   5730
      Begin MSMask.MaskEdBox INSSRetencaoMinima 
         Height          =   315
         Left            =   1260
         TabIndex        =   22
         Top             =   180
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   15
         Format          =   "standard"
         PromptChar      =   "_"
      End
      Begin VB.CheckBox INSSRetencao 
         Caption         =   "Com retenção"
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
         Left            =   4185
         TabIndex        =   24
         Top             =   225
         Width           =   1500
      End
      Begin MSMask.MaskEdBox INSSAliquota 
         Height          =   315
         Left            =   3195
         TabIndex        =   23
         Top             =   165
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   15
         Format          =   "#0.#0\%"
         PromptChar      =   "_"
      End
      Begin VB.Label LabelINSSAliquota 
         AutoSize        =   -1  'True
         Caption         =   "Alíquota:"
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
         Left            =   2310
         TabIndex        =   41
         Top             =   225
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mínima (R$):"
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
         Left            =   75
         TabIndex        =   40
         Top             =   225
         Width           =   1110
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7290
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TipoDeTributacaoOcx.ctx":00D2
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "TipoDeTributacaoOcx.ctx":022C
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "TipoDeTributacaoOcx.ctx":03B6
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "TipoDeTributacaoOcx.ctx":08E8
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FrameICMS 
      Caption         =   "ICMS"
      Height          =   1485
      Left            =   120
      TabIndex        =   38
      Top             =   1080
      Width           =   9300
      Begin VB.Frame Frame2 
         Caption         =   "Para o Simples Nacional"
         Height          =   585
         Left            =   30
         TabIndex        =   59
         Top             =   795
         Width           =   9240
         Begin VB.ComboBox TipoTributacaoICMSSimples 
            Height          =   315
            Left            =   525
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   195
            Width           =   4800
         End
         Begin VB.Label Label2 
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
            Height          =   180
            Left            =   45
            TabIndex        =   60
            Top             =   225
            Width           =   450
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Para Tributação Normal"
         Height          =   585
         Left            =   30
         TabIndex        =   57
         Top             =   210
         Width           =   9240
         Begin VB.CheckBox CreditaICMS 
            Caption         =   "Credita"
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
            Left            =   8220
            TabIndex        =   8
            Top             =   255
            Width           =   924
         End
         Begin VB.ComboBox TipoTributacaoICMS 
            Height          =   315
            Left            =   525
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   195
            Width           =   4800
         End
         Begin VB.CheckBox IncluiIPI 
            Caption         =   "Inclui IPI na base "
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
            Left            =   5385
            TabIndex        =   7
            Top             =   165
            Width           =   1950
         End
         Begin VB.Label Label4 
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
            Height          =   180
            Left            =   45
            TabIndex        =   58
            Top             =   225
            Width           =   450
         End
      End
      Begin VB.CheckBox IncideICMS 
         Caption         =   "Incide"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   7575
         TabIndex        =   33
         Top             =   15
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   852
      End
   End
   Begin VB.Frame FrameIPI 
      Caption         =   "IPI"
      Height          =   1185
      Left            =   120
      TabIndex        =   36
      Top             =   2565
      Width           =   9300
      Begin VB.CheckBox IncideIPI 
         Caption         =   "Incide"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   8595
         TabIndex        =   34
         Top             =   1095
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.CheckBox Destaca 
         Caption         =   "Destaca"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   5445
         TabIndex        =   11
         Top             =   210
         Width           =   1104
      End
      Begin VB.CheckBox SobreFrete 
         Caption         =   "Sobre o Frete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   6660
         TabIndex        =   12
         Top             =   225
         Width           =   1500
      End
      Begin VB.ComboBox TipoTributacaoIPI 
         Height          =   315
         Left            =   570
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   165
         Width           =   4830
      End
      Begin VB.CheckBox CreditaIPI 
         Caption         =   "Credita"
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
         Left            =   8280
         TabIndex        =   13
         Top             =   225
         Width           =   990
      End
      Begin MSMask.MaskEdBox IPICodEnq 
         Height          =   315
         Left            =   1905
         TabIndex        =   65
         Top             =   525
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   3
         Format          =   "000"
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin VB.Label IPICodEnqDesc 
         BorderStyle     =   1  'Fixed Single
         Height          =   585
         Left            =   2610
         TabIndex        =   67
         Top             =   525
         Width           =   6630
      End
      Begin VB.Label IPICodEnqLabel 
         AutoSize        =   -1  'True
         Caption         =   "Cód.Enquadramento:"
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
         Left            =   105
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   66
         Top             =   570
         Width           =   1770
      End
      Begin VB.Label TipoIPI 
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
         Height          =   225
         Left            =   120
         TabIndex        =   37
         Top             =   210
         Width           =   510
      End
   End
End
Attribute VB_Name = "TipoDeTributacaoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gcolTiposTribICMS As New Collection, gcolTiposTribICMSSimples As New Collection
Dim gcolTiposTribIPI As New Collection, gcolTiposTribPISCOFINS As New Collection
Dim gcolTiposTribISS As New Collection

Dim iEntradaAnt As Integer, iISSAnt As Integer

Dim iAlterado As Integer
Private WithEvents objEventoTipo As AdmEvento
Attribute objEventoTipo.VB_VarHelpID = -1
Private WithEvents objEventoIPICodEnq As AdmEvento
Attribute objEventoIPICodEnq.VB_VarHelpID = -1

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 33279

    'Limpa a Tela
    Call Limpa_Tela_TipoTributacao

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 33279

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174784)

    End Select

    Exit Sub

End Sub

Private Sub CreditaCOFINS_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CreditaICMS_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CreditaIPI_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CreditaPIS_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Destaca_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim sCodigo As String
'Dim colCodigoDescricao As New AdmColCodigoNome
'Dim objCodigoDescricao As New AdmCodigoNome
'Dim colTiposTribICMS As New AdmColCodigoNome
'Dim objTiposTribICMS As New AdmCodigoNome
'Dim colTiposTribIPI As New AdmColCodigoNome
'Dim objTiposTribIPI As New AdmCodigoNome
Dim colTiposTribICMS As New Collection, colTiposTribICMSSimples As New Collection
Dim colTiposTribIPI As New Collection, colTiposTribPISCOFINS As New Collection
Dim colTiposTribISS As New Collection
Dim objTipoTribISS As ClassTipoTribISS
Dim objTipoTribICMS As ClassTipoTribICMS
Dim objTipoTribICMSSimples As ClassTipoTribICMSSimples

On Error GoTo Erro_Form_Load

    Set objEventoTipo = New AdmEvento
    Set objEventoIPICodEnq = New AdmEvento

'    'Lê Tipos e Descrição da tabela TiposDeTributacaoMovto e devolve na coleção
'    lErro = CF("Cod_Nomes_Le", "TiposDeTributacaoMovto", "Tipo", "Descricao", TIPO_TRIBUTACAO_DESCRICAO, colCodigoDescricao)
'    If lErro <> SUCESSO Then Error 33276


'    'Lê cada Código e Descrição da tabela TiposTribICMS e põe na coleção
'    lErro = CF("Cod_Nomes_Le", "TiposTribICMS", "Tipo", "Descricao", STRING_TIPO_ICMS_DESCRICAO, colTiposTribICMS)
'    If lErro <> SUCESSO Then Error 33277
'
'    'Preenche TipoTributacaoICMS
'    For Each objTiposTribICMS In colTiposTribICMS
'        sCodigo = CStr(objTiposTribICMS.iCodigo) & SEPARADOR & objTiposTribICMS.sNome
'        TipoTributacaoICMS.AddItem (sCodigo)
'        TipoTributacaoICMS.ItemData(TipoTributacaoICMS.NewIndex) = objTiposTribICMS.iCodigo
'    Next
'
'    'Le cada Codigo e Descrição da tabela TiposTribIPI e poe na colecao
'    lErro = CF("Cod_Nomes_Le", "TiposTribIPI", "Tipo", "Descricao", STRING_TIPO_IPI_DESCRICAO, colTiposTribIPI)
'    If lErro <> SUCESSO Then Error 33278
'
'    'Preenche TipoTributacaoIPI
'    For Each objTiposTribIPI In colTiposTribIPI
'        sCodigo = CStr(objTiposTribIPI.iCodigo) & SEPARADOR & objTiposTribIPI.sNome
'        TipoTributacaoIPI.AddItem (sCodigo)
'        TipoTributacaoIPI.ItemData(TipoTributacaoIPI.NewIndex) = objTiposTribIPI.iCodigo
'    Next
'
'    'Marca Incide ICMS e IPI
'    IncideICMS.Value = TRIBUTO_INCIDE
'    IncideIPI.Value = TRIBUTO_INCIDE
'
'    TipoTributacaoICMS.ListIndex = 1
'    TipoTributacaoIPI.ListIndex = 1

    lErro = CF("TiposTribICMS_Le_Todos", colTiposTribICMS)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Set gcolTiposTribICMS = colTiposTribICMS

    lErro = CF("TiposTribICMSSimples_Le_Todos", colTiposTribICMSSimples)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Set gcolTiposTribICMSSimples = colTiposTribICMSSimples

    lErro = CF("TiposTribIPI_Le_Todos", colTiposTribIPI)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Set gcolTiposTribIPI = colTiposTribIPI

    lErro = CF("TiposTribPISCOFINS_Le_Todos", colTiposTribPISCOFINS)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Set gcolTiposTribPISCOFINS = colTiposTribPISCOFINS

    lErro = CF("TiposTribISS_Le_Todos", colTiposTribISS)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Set gcolTiposTribISS = colTiposTribISS
    
    For Each objTipoTribISS In colTiposTribISS
        TipoTributacaoISS.AddItem objTipoTribISS.sDescricao
        TipoTributacaoISS.ItemData(TipoTributacaoISS.NewIndex) = objTipoTribISS.iTipo
    Next
    
'    RegimeTributario.AddItem REGIME_TRIBUTARIO_NORMAL_TEXTO
'    RegimeTributario.ItemData(RegimeTributario.NewIndex) = REGIME_TRIBUTARIO_NORMAL
'
'    RegimeTributario.AddItem REGIME_TRIBUTARIO_SIMPLES_TEXTO
'    RegimeTributario.ItemData(RegimeTributario.NewIndex) = REGIME_TRIBUTARIO_SIMPLES

    For Each objTipoTribICMSSimples In gcolTiposTribICMSSimples
        TipoTributacaoICMSSimples.AddItem Format(objTipoTribICMSSimples.iCSOSN, "000") & SEPARADOR & objTipoTribICMSSimples.sDescricao
        TipoTributacaoICMSSimples.ItemData(TipoTributacaoICMSSimples.NewIndex) = objTipoTribICMSSimples.iTipo
    Next

    For Each objTipoTribICMS In gcolTiposTribICMS
        TipoTributacaoICMS.AddItem Format(objTipoTribICMS.iTipoTribCST, "00") & SEPARADOR & objTipoTribICMS.sDescricao
        TipoTributacaoICMS.ItemData(TipoTributacaoICMS.NewIndex) = objTipoTribICMS.iTipo
    Next
    
    lErro = CF("Carrega_Combo", NatBCCred, "NatBCCred", "Codigo", TIPO_STR, "Descricao", TIPO_STR)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    NatBCCred.AddItem " "
    NatBCCred.ItemData(NatBCCred.NewIndex) = 0
    
    Call Default_Tela
        
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174785)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Sub Default_Tela()

    TipoTributacaoISS.ListIndex = -1
    TipoTributacaoICMS.ListIndex = 0
    TipoTributacaoICMSSimples.ListIndex = 0

    iEntradaAnt = -1
    'iRegimeAnt = -1
    iISSAnt = -1

    OptICMS.Value = True
    Call Trata_ICMS_ISS
    
    Entrada.Value = False
    Call Trata_Entrada_Saida
    
    'Call Combo_Seleciona_ItemData(RegimeTributario, REGIME_TRIBUTARIO_NORMAL)
    'Call Trata_Regime_Tributario
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 33280

    'Limpa a Tela
    Call Limpa_Tela_TipoTributacao

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 33280

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 174786)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTipoTributacao As New ClassTipoDeTributacaoMovto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se o Tipo foi preenchido
    If Len(Trim(Tipo.Text)) = 0 Then Error 33309

    'Verifica se o Tipo de Tributação existe
    objTipoTributacao.iTipo = CInt(Tipo.Text)
    lErro = CF("TipoTributacao_Le", objTipoTributacao)
    If lErro <> SUCESSO And lErro <> 27259 Then Error 33310

    'Não encontrou o Tipo de Tributação ==> Erro
    If lErro = 27259 Then Error 33311

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_TIPO_TRIBUTACAO", objTipoTributacao.iTipo)

    If vbMsgRes = vbNo Then Exit Sub

    'Exclui o Tipo de Tributação
    lErro = CF("TipoTributacao_Exclui", objTipoTributacao)
    If lErro <> SUCESSO Then Error 33312

   
    Call Limpa_Tela_TipoTributacao

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case Err

        Case 33309
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_NAO_PREENCHIDO", Err)

        Case 33310, 33312

        Case 33311
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_NAO_CADASTRADO", Err, objTipoTributacao.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174787)

    End Select

    Exit Sub

End Sub



Private Function Limpa_Tela_TipoTributacao() As Long
'Limpa os campos tela TipoDeTributacao

Dim lErro As Long

   'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    Tipo.Text = ""
    Descricao.Text = ""
    
    'IncideICMS.Value = 1
    'TipoTributacaoICMS.ListIndex = 1
    'IncluiIPI.Value = 0
    'CreditaIPI.Value = 0
    'CreditaICMS.Value = 0
    
    'IncideIPI.Value = 1
    'TipoTributacaoIPI.ListIndex = 1
    'Destaca.Value = 0
    'SobreFrete.Value = 0
    
    'IncideISS.Value = 0
    IncideIR.Value = 0
    INSSRetencao.Value = 0
    
    CreditaPIS.Value = 0
    ReterPIS.Value = 0
    CreditaCOFINS.Value = 0
    ReterCOFINS.Value = 0
    ReterCSLL.Value = 0
    
    ReterISS.Value = 0
    
    NatBCCred.ListIndex = -1
    ISSIndExigibilidade.ListIndex = -1
    
    IPICodEnqDesc.Caption = ""
    
    Call Default_Tela
    
    iAlterado = 0

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTipoTributacao As New ClassTipoDeTributacaoMovto

On Error GoTo Erro_Gravar_Registro

    'Verifica se foi Tipo foi preenchido
    If Len(Trim(Tipo.Text)) = 0 Then gError 33281
    
    'Verifica se o tipo de tributação é menor que 500
    If StrParaInt(Tipo.Text) < NUMERO_PRIMEIRO_TIPOTRIB_USUARIO Then gError 81144
    
    'Verifica se foi preenchida Descrição
    If Len(Trim(Descricao.Text)) = 0 Then gError 33282
    
    lErro = IPICodEnq_Valida()
    If lErro <> SUCESSO Then gError 33283

    'Verifica se o Tipo de Tributação ICMS foi preenchido
'    If Len(Trim(TipoTributacaoICMS.Text)) = 0 Then gError 33400

    'Verifica se o Tipo de Tributação IPI foi preenchido
'    If Len(Trim(TipoTributacaoIPI.Text)) = 0 Then gError 33401
    
    'Lê os dados da Tela relacionados ao Tipo de Tributação
    lErro = Move_Tela_Memoria(objTipoTributacao)
    If lErro <> SUCESSO Then gError 33283

    lErro = Trata_Alteracao(objTipoTributacao, objTipoTributacao.iTipo)
    If lErro <> SUCESSO Then Error 32323

    'Grava o Tipo de Tributação no BD
    lErro = CF("TipoTributacao_Grava", objTipoTributacao)
    If lErro <> SUCESSO Then gError 33284

    'Atualiza ListBox de Tipos

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 32323

        Case 33281
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_NAO_PREENCHIDO", gErr)

        Case 33282
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_TIPO_TRIBUTACAO_NAO_PREENCHIDO", gErr)

        Case 33283, 33284

        Case 33400
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_ICMS_NAO_PREENCHIDO", gErr)

        Case 33401
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_IPI_NAO_PREENCHIDO", gErr)
            
        Case 81144
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PRIMEIRO_TIPOTRIB_USUARIO", gErr, NUMERO_PRIMEIRO_TIPOTRIB_USUARIO)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174788)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objTipoTributacao As ClassTipoDeTributacaoMovto) As Long
'Move os dados da tela para memória

Dim lErro As Long
Dim objTipoTribICMS As New ClassTipoTribICMS
Dim objTipoTribIPI As New ClassTipoTribIPI

On Error GoTo Erro_Move_Tela_Memoria

    If Len(Trim(Tipo.Text)) > 0 Then objTipoTributacao.iTipo = CInt(Tipo.Text)

    objTipoTributacao.sDescricao = Descricao.Text

    objTipoTributacao.iEntrada = IIf(Entrada.Value, 1, 0)
    
    'obtem atributos correspondentes ao ICMS
    objTipoTributacao.iICMSIncide = IncideICMS.Value
    'objTipoTributacao.iICMSTipo = TipoTributacaoICMS.ItemData(TipoTributacaoICMS.ListIndex)
    objTipoTributacao.iICMSBaseComIPI = IncluiIPI.Value
    objTipoTributacao.iICMSCredita = IIf(objTipoTributacao.iEntrada = 1, IIf(CreditaICMS.Value = vbChecked, TIPOTRIB_CREDITA, TIPOTRIB_SEMCREDDEB), IIf(CreditaICMS.Value = vbChecked, TIPOTRIB_DEBITA, TIPOTRIB_SEMCREDDEB))
    
    'obtem atributos correspondentes ao IPI
    objTipoTributacao.iIPIIncide = IncideIPI.Value
    'objTipoTributacao.iIPITipo = TipoTributacaoIPI.ItemData(TipoTributacaoIPI.ListIndex)
    objTipoTributacao.iIPIDestaca = Destaca.Value
    objTipoTributacao.iIPIFrete = SobreFrete.Value
    objTipoTributacao.iIPICredita = IIf(objTipoTributacao.iEntrada = 1, IIf(CreditaIPI.Value = vbChecked, TIPOTRIB_CREDITA, TIPOTRIB_SEMCREDDEB), IIf(CreditaIPI.Value = vbChecked, TIPOTRIB_DEBITA, TIPOTRIB_SEMCREDDEB))
    
    'obtem atributos correspondentes ao IR
    objTipoTributacao.dIRAliquota = PercentParaDbl(IRAliquota.FormattedText)
    objTipoTributacao.iIRIncide = IncideIR.Value
   
    'obtem atributos correspondentes ao ISS
    objTipoTributacao.dINSSAliquota = PercentParaDbl(INSSAliquota.FormattedText)
    objTipoTributacao.dINSSRetencaoMinima = StrParaDbl(INSSRetencaoMinima)
    objTipoTributacao.iINSSIncide = INSSRetencao.Value
    objTipoTributacao.iISSRetencao = ReterISS.Value
    
    'Verifica os demais impostos
    objTipoTributacao.iISSIncide = IncideISS.Value
    objTipoTributacao.iIRIncide = IncideIR.Value
    objTipoTributacao.iPISCredita = IIf(objTipoTributacao.iEntrada = 1, IIf(CreditaPIS.Value = vbChecked, TIPOTRIB_CREDITA, TIPOTRIB_SEMCREDDEB), IIf(CreditaPIS.Value = vbChecked, TIPOTRIB_DEBITA, TIPOTRIB_SEMCREDDEB))
    objTipoTributacao.iPISRetencao = ReterPIS.Value
    objTipoTributacao.iCOFINSCredita = IIf(objTipoTributacao.iEntrada = 1, IIf(CreditaCOFINS.Value = vbChecked, TIPOTRIB_CREDITA, TIPOTRIB_SEMCREDDEB), IIf(CreditaCOFINS.Value = vbChecked, TIPOTRIB_DEBITA, TIPOTRIB_SEMCREDDEB))
    objTipoTributacao.iCOFINSRetencao = ReterCOFINS.Value
    objTipoTributacao.iCSLLRetencao = ReterCSLL.Value
    
    objTipoTributacao.iPISTipo = TipoTributacaoPIS.ItemData(TipoTributacaoPIS.ListIndex)
    objTipoTributacao.iCOFINSTipo = TipoTributacaoCOFINS.ItemData(TipoTributacaoCOFINS.ListIndex)
    'objTipoTributacao.iRegimeTributario = RegimeTributario.ItemData(RegimeTributario.ListIndex)
    
    If OptICMS.Value Then
'        If objTipoTributacao.iRegimeTributario = REGIME_TRIBUTARIO_SIMPLES Then
         If TipoTributacaoICMSSimples.ListIndex <> -1 Then objTipoTributacao.iICMSSimplesTipo = TipoTributacaoICMSSimples.ItemData(TipoTributacaoICMSSimples.ListIndex)
'        Else
            objTipoTributacao.iICMSTipo = TipoTributacaoICMS.ItemData(TipoTributacaoICMS.ListIndex)
'        End If
        objTipoTributacao.iIPITipo = TipoTributacaoIPI.ItemData(TipoTributacaoIPI.ListIndex)
    Else
        objTipoTributacao.iISSTipo = TipoTributacaoISS.ItemData(TipoTributacaoISS.ListIndex)
    End If
   
    objTipoTributacao.sNatBCCred = SCodigo_Extrai(NatBCCred.Text)
   
    objTipoTributacao.iISSIndExigibilidade = Codigo_Extrai(ISSIndExigibilidade.Text)
    
    If Len(Trim(IPICodEnq.Text)) > 0 Then objTipoTributacao.sIPICodEnq = Format(Trim(IPICodEnq.Text), "000")
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174789)

    End Select

    Exit Function

End Function


'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objTipoTributacao As New ClassTipoDeTributacaoMovto

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TiposDeTributacaoMovto"

    'Lê os dados da Tela
    lErro = Move_Tela_Memoria(objTipoTributacao)
    If lErro <> SUCESSO Then Error 33289

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Tipo", objTipoTributacao.iTipo, 0, "Tipo"
    colCampoValor.Add "Descricao", objTipoTributacao.sDescricao, TIPO_TRIBUTACAO_DESCRICAO, "Descricao"
    colCampoValor.Add "ICMSIncide", objTipoTributacao.iICMSIncide, 0, "ICMSIncide"
    colCampoValor.Add "ICMSTipo", objTipoTributacao.iICMSTipo, 0, "ICMSTipo"
    colCampoValor.Add "ICMSBaseComIPI", objTipoTributacao.iICMSBaseComIPI, 0, "ICMSBaseComIPI"
    colCampoValor.Add "ICMSCredita", objTipoTributacao.iICMSCredita, 0, "ICMSCredita"
    colCampoValor.Add "IPIIncide", objTipoTributacao.iIPIIncide, 0, "IPIIncide"
    colCampoValor.Add "IPITipo", objTipoTributacao.iIPITipo, 0, "IPITipo"
    colCampoValor.Add "IPIFrete", objTipoTributacao.iIPIFrete, 0, "IPIFrete"
    colCampoValor.Add "IPIDestaca", objTipoTributacao.iIPIDestaca, 0, "IPIDestaca"
    colCampoValor.Add "IPICredita", objTipoTributacao.iIPICredita, 0, "IPICredita"
    colCampoValor.Add "ISSIncide", objTipoTributacao.iISSIncide, 0, "ISSIncide"
    colCampoValor.Add "IRIncide", objTipoTributacao.iIRIncide, 0, "IRIncide"
    colCampoValor.Add "IRAliquota", objTipoTributacao.dIRAliquota, 0, "IRAliquota"
    colCampoValor.Add "INSSAliquota", objTipoTributacao.dINSSAliquota, 0, "INSSAliquota"
    colCampoValor.Add "INSSRetencaoMinima", objTipoTributacao.dINSSRetencaoMinima, 0, "INSSRetencaoMinima"
    colCampoValor.Add "INSSIncide", objTipoTributacao.iINSSIncide, 0, "INSSIncide"
    colCampoValor.Add "PISCredita", objTipoTributacao.iPISCredita, 0, "PISCredita"
    colCampoValor.Add "PISRetencao", objTipoTributacao.iPISRetencao, 0, "PISRetencao"
    colCampoValor.Add "COFINSCredita", objTipoTributacao.iCOFINSCredita, 0, "COFINSCredita"
    colCampoValor.Add "COFINSRetencao", objTipoTributacao.iCOFINSRetencao, 0, "COFINSRetencao"
    colCampoValor.Add "CSLLRetencao", objTipoTributacao.iCSLLRetencao, 0, "CSLLRetencao"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 33289

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174790)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objTipoTributacao As New ClassTipoDeTributacaoMovto

On Error GoTo Erro_Tela_Preenche

    objTipoTributacao.iTipo = colCampoValor.Item("Tipo").vValor

    If objTipoTributacao.iTipo <> 0 Then

        'Chama Traz_TipoTributacao_Tela
        lErro = Traz_TipoTributacao_Tela(objTipoTributacao)
        If lErro <> SUCESSO Then Error 33290

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 33290

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174791)

    End Select

    Exit Sub

End Sub

Private Function Traz_TipoTributacao_Tela(objTipoTributacao As ClassTipoDeTributacaoMovto) As Long
'Traz os dados do Tipo de Tributação para tela

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer

On Error GoTo Erro_Traz_TipoTributacao_Tela

    'Lê o Tipo de Tributação
    lErro = CF("TipoTributacao_Le", objTipoTributacao)
    If lErro <> SUCESSO And lErro <> 27259 Then Error 33291

    'Não encontrou o Tipo de Tributação ==> erro
    If lErro = 27259 Then Error 33292

    If objTipoTributacao.iTipo <> 0 Then
        Tipo.Text = CStr(objTipoTributacao.iTipo)
    Else
        Tipo.Text = ""
    End If

    Descricao.Text = objTipoTributacao.sDescricao
    
    iEntradaAnt = -1
    'iRegimeAnt = -1
    iISSAnt = -1
    
    If objTipoTributacao.iISSIncide = MARCADO Then
        OptISS.Value = True
    Else
        OptICMS.Value = True
    End If
    Call Trata_ICMS_ISS
    
    Entrada.Value = (objTipoTributacao.iEntrada = 1)
    Saida.Value = (objTipoTributacao.iEntrada = 0)
    Call Trata_Entrada_Saida
            
'    Call Combo_Seleciona_ItemData(RegimeTributario, objTipoTributacao.iRegimeTributario)
'    Call Trata_Regime_Tributario
    
    Call Combo_Seleciona_ItemData(TipoTributacaoPIS, objTipoTributacao.iPISTipo)
    Call Combo_Seleciona_ItemData(TipoTributacaoCOFINS, objTipoTributacao.iCOFINSTipo)
    
    ''Verifica os demais impostos
    'IncideICMS.Value = objTipoTributacao.iICMSIncide
'
'    'Percorre todos os elementos da ComboBox
'    For iIndice = 0 To TipoTributacaoICMS.ListCount - 1
'
'        'Compara se código já existe na ComboBox
'        If TipoTributacaoICMS.ItemData(iIndice) = objTipoTributacao.iICMSTipo Then
'
'            'Seleciona o item na ComboBox
'            TipoTributacaoICMS.ListIndex = iIndice
'            Exit For
'
'        End If
'
'    Next

    If OptICMS.Value Then
        IncluiIPI.Value = objTipoTributacao.iICMSBaseComIPI
        CreditaICMS.Value = IIf(objTipoTributacao.iICMSCredita <> TIPOTRIB_SEMCREDDEB, vbChecked, vbUnchecked)
    
        Destaca.Value = objTipoTributacao.iIPIDestaca
        SobreFrete.Value = objTipoTributacao.iIPIFrete
        CreditaIPI.Value = IIf(objTipoTributacao.iIPICredita <> TIPOTRIB_SEMCREDDEB, vbChecked, vbUnchecked)
    
        Call Combo_Seleciona_ItemData(TipoTributacaoIPI, objTipoTributacao.iIPITipo)
        
'        If objTipoTributacao.iRegimeTributario = REGIME_TRIBUTARIO_SIMPLES Then
         If objTipoTributacao.iICMSSimplesTipo <> 0 Then Call Combo_Seleciona_ItemData(TipoTributacaoICMSSimples, objTipoTributacao.iICMSSimplesTipo)
'        Else
            Call Combo_Seleciona_ItemData(TipoTributacaoICMS, objTipoTributacao.iICMSTipo)
'        End If
    Else
        ReterISS.Value = objTipoTributacao.iISSRetencao
        Call Combo_Seleciona_ItemData(TipoTributacaoISS, objTipoTributacao.iISSTipo)
    End If
    
    
'    IncideIPI.Value = objTipoTributacao.iIPIIncide
'
'    'Percorre todos os elementos da ComboBox
'    For iIndice1 = 0 To TipoTributacaoIPI.ListCount - 1
'
'        'Compara se código já existe na ComboBox
'        If TipoTributacaoIPI.ItemData(iIndice1) = objTipoTributacao.iIPITipo Then
'
'            'Seleciona o item na ComboBox
'            TipoTributacaoIPI.ListIndex = iIndice1
'            Exit For
'
'        End If
'
'    Next

'    Destaca.Value = objTipoTributacao.iIPIDestaca
'    SobreFrete.Value = objTipoTributacao.iIPIFrete
'    CreditaIPI.Value = IIf(objTipoTributacao.iIPICredita <> TIPOTRIB_SEMCREDDEB, vbChecked, vbUnchecked)
    
'    IncideISS.Value = objTipoTributacao.iISSIncide
'    ReterISS.Value = objTipoTributacao.iISSRetencao
    
    IncideIR.Value = objTipoTributacao.iIRIncide
    INSSRetencao.Value = objTipoTributacao.iINSSIncide
    
    IRAliquota.Text = (objTipoTributacao.dIRAliquota) * 100
    INSSRetencaoMinima.Text = objTipoTributacao.dINSSRetencaoMinima
    INSSAliquota.Text = (objTipoTributacao.dINSSAliquota) * 100
    
    CreditaPIS.Value = IIf(objTipoTributacao.iPISCredita <> TIPOTRIB_SEMCREDDEB, vbChecked, vbUnchecked)
    
    ReterPIS.Value = objTipoTributacao.iPISRetencao
    CreditaCOFINS.Value = IIf(objTipoTributacao.iCOFINSCredita <> TIPOTRIB_SEMCREDDEB, vbChecked, vbUnchecked)
    ReterCOFINS.Value = objTipoTributacao.iCOFINSRetencao
    ReterCSLL.Value = objTipoTributacao.iCSLLRetencao
    
    Call CF("SCombo_Seleciona2", NatBCCred, objTipoTributacao.sNatBCCred)
    
    If objTipoTributacao.iISSIndExigibilidade <> 0 Then
        Call Combo_Seleciona_ItemData(ISSIndExigibilidade, objTipoTributacao.iISSIndExigibilidade)
    Else
        ISSIndExigibilidade.ListIndex = -1
    End If
    
    IPICodEnq.Text = objTipoTributacao.sIPICodEnq
    Call IPICodEnq_Validate(bSGECancelDummy)
    
    iAlterado = 0
    
    Traz_TipoTributacao_Tela = SUCESSO

    Exit Function

Erro_Traz_TipoTributacao_Tela:

    Traz_TipoTributacao_Tela = Err

    Select Case Err

        Case 33291

        Case 33292
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_NAO_CADASTRADO", Err, objTipoTributacao.iTipo)

        Case 33396, 33398

        Case 33397
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_CADASTRADO_TIPOSTRIBICMS1", Err, TipoTributacaoICMS.Text)
            TipoTributacaoICMS.SetFocus

        Case 33399
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_CADASTRADO_TIPOSTRIBIPI1", Err, TipoTributacaoIPI.Text)
            TipoTributacaoIPI.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174792)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoTipo = Nothing
    Set objEventoIPICodEnq = Nothing

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
     
End Sub

Private Sub Entrada_Click()

    CreditaICMS.Caption = IIf(Entrada.Value, "Credita", "Debita")
    CreditaIPI.Caption = IIf(Entrada.Value, "Credita", "Debita")
    CreditaPIS.Caption = IIf(Entrada.Value, "Credita", "Debita")
    CreditaCOFINS.Caption = IIf(Entrada.Value, "Credita", "Debita")
    
    iAlterado = REGISTRO_ALTERADO
    Call Trata_Entrada_Saida
    
End Sub

Private Sub IncideICMS_Click()
'
'    iAlterado = REGISTRO_ALTERADO
'
'    If IncideICMS.Value = TRIBUTO_NAO_INCIDE Then
'        'Limpa e desabilita os campos correspondentes ao ICMS
'        TipoTributacaoICMS.ListIndex = 0
'        TipoTributacaoICMS.Enabled = False
'        IncluiIPI.Value = 0
'        IncluiIPI.Enabled = False
'        CreditaICMS.Value = 0
'        CreditaICMS.Enabled = False
'    Else
'        TipoTributacaoICMS.ListIndex = 1
'        TipoTributacaoICMS.Enabled = True
'        IncluiIPI.Enabled = True
'        CreditaICMS.Enabled = True
'    End If

End Sub

Private Sub IncideIPI_Click()
'
'    iAlterado = REGISTRO_ALTERADO
'
'    If IncideIPI.Value = TRIBUTO_NAO_INCIDE Then
'        'Limpa e desabilita os campos correspondentes ao IPI
'        TipoTributacaoIPI.ListIndex = 0
'        TipoTributacaoIPI.Enabled = False
'        Destaca.Value = 0
'        Destaca.Enabled = False
'        SobreFrete.Value = 0
'        SobreFrete.Enabled = False
'        CreditaIPI.Value = 0
'        CreditaIPI.Enabled = False
'    Else
'        TipoTributacaoIPI.ListIndex = 1
'        TipoTributacaoIPI.Enabled = True
'        Destaca.Enabled = True
'        SobreFrete.Enabled = True
'        CreditaIPI.Enabled = True
'    End If

End Sub

Private Sub IncideIR_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub IncideISS_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub IncluiIPI_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub



Private Sub INSSAliquota_Change()

    iAlterado = REGISTRO_ALTERADO


End Sub

Private Sub INSSAliquota_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_IINSSAliquota_Validate
    
    If INSSAliquota.Text = "" Then Exit Sub

    'Verifica se o valor percentual informado está entre 0 e 100
    lErro = Porcentagem_Critica(INSSAliquota.Text)
    If lErro <> SUCESSO Then gError 70417
    
    Exit Sub

Erro_IINSSAliquota_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 70417
               
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174793)
    
    End Select
    
    Exit Sub


End Sub

Private Sub INSSRetencao_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub IRAliquota_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub IRAliquota_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_IRAliquota_Validate
    
    If IRAliquota.Text = "" Then Exit Sub

    'Verifica se o valor percentual informado está entre 0 e 100
    lErro = Porcentagem_Critica(IRAliquota.Text)
    If lErro <> SUCESSO Then gError 70414
    
    Exit Sub

Erro_IRAliquota_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 70414
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174794)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub INSSRetencaoMinima_Change()

    iAlterado = REGISTRO_ALTERADO


End Sub

Private Sub ISSIndExigibilidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoTipo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoTributacao As ClassTipoDeTributacaoMovto

On Error GoTo Erro_objEventoTipo_evSelecao

    Set objTipoTributacao = obj1

    'Mostra o tipo de tributação na tela
    lErro = Traz_TipoTributacao_Tela(objTipoTributacao)
    If lErro <> SUCESSO Then Error 33325

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    Me.Show

    Exit Sub

Erro_objEventoTipo_evSelecao:

    Select Case Err

        Case 33325

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 174795)

    End Select

    Exit Sub

End Sub

Private Sub OptICMS_Click()
    iAlterado = REGISTRO_ALTERADO
    Call Trata_ICMS_ISS
End Sub

Private Sub OptISS_Click()
    iAlterado = REGISTRO_ALTERADO
    Call Trata_ICMS_ISS
End Sub

'Private Sub RegimeTributario_Change()
'    iAlterado = REGISTRO_ALTERADO
'    Call Trata_Regime_Tributario
'End Sub
'
'Private Sub RegimeTributario_Click()
'    iAlterado = REGISTRO_ALTERADO
'    Call Trata_Regime_Tributario
'End Sub

Private Sub ReterCOFINS_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ReterCSLL_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ReterPIS_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ReterISS_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Saida_Click()

    CreditaICMS.Caption = IIf(Entrada.Value, "Credita", "Debita")
    CreditaIPI.Caption = IIf(Entrada.Value, "Credita", "Debita")
    CreditaPIS.Caption = IIf(Entrada.Value, "Credita", "Debita")
    CreditaCOFINS.Caption = IIf(Entrada.Value, "Credita", "Debita")
    
    iAlterado = REGISTRO_ALTERADO
    Call Trata_Entrada_Saida

End Sub

Private Sub SobreFrete_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tipo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tipo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Tipo, iAlterado)

End Sub

Private Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Tipo_Validate

    'Verifica se Tipo foi digitado
    If Len(Trim(Tipo.Text)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(Tipo.Text)
    If lErro <> SUCESSO Then Error 33293

    'Verifica se é um número inteiro
    lErro = Inteiro_Critica(Tipo.Text)
    If lErro <> SUCESSO Then Error 33294

    Exit Sub

Erro_Tipo_Validate:

    Cancel = True

    Select Case Err

        Case 33293, 33294

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174796)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objTipoTributacao As ClassTipoDeTributacaoMovto) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se foi passado algum Tipo de Tributação
    If Not (objTipoTributacao Is Nothing) Then

        'Traz os dados para a Tela
        lErro = Traz_TipoTributacao_Tela(objTipoTributacao)
        If lErro <> SUCESSO And lErro <> 33292 Then Error 33328

        If lErro = 33292 Then Tipo.Text = CStr(objTipoTributacao.iTipo)
    
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 33328

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 174797)

    End Select
    
    iAlterado = 0

    Exit Function

    Trata_Parametros = SUCESSO

End Function

Private Sub TipoLabel_Click()

Dim colSelecao As New Collection
Dim objTipoTributacao As New ClassTipoDeTributacaoMovto

    colSelecao.Add "1"
    colSelecao.Add "0"
    
    'Chama a tela de browse
    Call Chama_Tela("TiposDeTribMovtoLista", colSelecao, objTipoTributacao, objEventoTipo)

End Sub

Private Sub TipoTributacaoCOFINS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TipoTributacaoCOFINS_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TipoTributacaoICMS_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoTributacaoICMS_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoTributacaoIPI_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoTributacaoIPI_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object
    
    Parent.HelpContextID = IDH_TIPO_TRIBUTACAO
    Set Form_Load_Ocx = Me
    Caption = "Tipos de Tributação"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TipoDeTributacao"
    
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

Private Sub TipoTributacaoISS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TipoTributacaoISS_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TipoTributacaoPIS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TipoTributacaoPIS_Click()
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
        
        If Me.ActiveControl Is Tipo Then
            Call TipoLabel_Click
        End If
    
    End If

End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub TipoIPI_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoIPI, Source, X, Y)
End Sub

Private Sub TipoIPI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoIPI, Button, Shift, X, Y)
End Sub

Private Sub TipoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoLabel, Source, X, Y)
End Sub

Private Sub TipoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Trata_ICMS_ISS()

    If OptICMS.Value Then
        TipoTributacaoISS.ListIndex = -1
        If TipoTributacaoICMS.ListCount > 0 Then TipoTributacaoICMS.ListIndex = 0
        If TipoTributacaoIPI.ListCount > 0 Then TipoTributacaoIPI.ListIndex = 0
        FrameICMS.Enabled = True
        FrameIPI.Enabled = True
        FrameISS.Enabled = False
        ReterISS.Value = vbUnchecked
        IncideIPI.Value = vbChecked
        IncideICMS.Value = vbChecked
        IncideISS.Value = vbUnchecked
    Else
        TipoTributacaoICMS.ListIndex = -1
        TipoTributacaoICMSSimples.ListIndex = -1
        TipoTributacaoIPI.ListIndex = -1
        If TipoTributacaoISS.ListCount > 0 Then TipoTributacaoISS.ListIndex = 0
        FrameICMS.Enabled = False
        FrameIPI.Enabled = False
        FrameISS.Enabled = True
        Destaca.Value = vbUnchecked
        SobreFrete.Value = vbUnchecked
        CreditaIPI.Value = vbUnchecked
        IncluiIPI.Value = vbUnchecked
        CreditaICMS.Value = vbUnchecked
        IncideIPI.Value = vbUnchecked
        IncideICMS.Value = vbUnchecked
        IncideISS.Value = vbChecked
    End If

End Sub

Private Sub Trata_Entrada_Saida()

Dim lErro As Long, iEntrada As Integer
Dim objTipoTribPISCOFINS As ClassTipoTribPISCOFINS
Dim objTipoTribIPI As ClassTipoTribIPI

On Error GoTo Erro_Trata_Entrada_Saida

    If Entrada.Value Then
        iEntrada = MARCADO
    Else
        iEntrada = DESMARCADO
    End If
    
    If iEntradaAnt <> iEntrada Then

        TipoTributacaoPIS.Clear
        TipoTributacaoCOFINS.Clear
        For Each objTipoTribPISCOFINS In gcolTiposTribPISCOFINS
            If (iEntrada = MARCADO And objTipoTribPISCOFINS.iEntrada = MARCADO) Or (iEntrada = DESMARCADO And objTipoTribPISCOFINS.iSaida = MARCADO) Then
                TipoTributacaoPIS.AddItem Format(objTipoTribPISCOFINS.iTipo, "00") & SEPARADOR & objTipoTribPISCOFINS.sDescricao
                TipoTributacaoPIS.ItemData(TipoTributacaoPIS.NewIndex) = objTipoTribPISCOFINS.iTipo
            
                TipoTributacaoCOFINS.AddItem Format(objTipoTribPISCOFINS.iTipo, "00") & SEPARADOR & objTipoTribPISCOFINS.sDescricao
                TipoTributacaoCOFINS.ItemData(TipoTributacaoCOFINS.NewIndex) = objTipoTribPISCOFINS.iTipo
            End If
        Next
        
        TipoTributacaoIPI.Clear
        For Each objTipoTribIPI In gcolTiposTribIPI
            If iEntrada = MARCADO Then
                TipoTributacaoIPI.AddItem Format(objTipoTribIPI.iCSTEntrada, "00") & SEPARADOR & objTipoTribIPI.sDescricao
                TipoTributacaoIPI.ItemData(TipoTributacaoIPI.NewIndex) = objTipoTribIPI.iTipo
            Else
                TipoTributacaoIPI.AddItem Format(objTipoTribIPI.iCSTSaida, "00") & SEPARADOR & objTipoTribIPI.sDescricao
                TipoTributacaoIPI.ItemData(TipoTributacaoIPI.NewIndex) = objTipoTribIPI.iTipo
            End If
        Next
        
        TipoTributacaoPIS.ListIndex = 0
        TipoTributacaoCOFINS.ListIndex = 0
        
        If OptICMS.Value Then TipoTributacaoIPI.ListIndex = 0
        
        iEntradaAnt = iEntrada
        
    End If

    Exit Sub

Erro_Trata_Entrada_Saida:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174796)

    End Select

    Exit Sub

End Sub

'Private Sub Trata_Regime_Tributario()
'
'Dim lErro As Long
'Dim objTipoTribICMS As ClassTipoTribICMS
'Dim objTipoTribICMSSimples As ClassTipoTribICMSSimples
'
'On Error GoTo Erro_Trata_Regime_Tributario
'
'    If RegimeTributario.ListIndex >= 0 Then
'
'        If iRegimeAnt <> RegimeTributario.ItemData(RegimeTributario.ListIndex) Then
'
'            TipoTributacaoICMS.Clear
'            If RegimeTributario.ItemData(RegimeTributario.ListIndex) = REGIME_TRIBUTARIO_SIMPLES Then
'
'                For Each objTipoTribICMSSimples In gcolTiposTribICMSSimples
'                    TipoTributacaoICMS.AddItem Format(objTipoTribICMSSimples.iCSOSN, "000") & SEPARADOR & objTipoTribICMSSimples.sDescricao
'                    TipoTributacaoICMS.ItemData(TipoTributacaoICMS.NewIndex) = objTipoTribICMSSimples.iTipo
'                Next
'
'            Else
'
'                For Each objTipoTribICMS In gcolTiposTribICMS
'                    TipoTributacaoICMS.AddItem Format(objTipoTribICMS.iTipoTribCST, "00") & SEPARADOR & objTipoTribICMS.sDescricao
'                    TipoTributacaoICMS.ItemData(TipoTributacaoICMS.NewIndex) = objTipoTribICMS.iTipo
'                Next
'
'            End If
'
'            If OptICMS.Value Then TipoTributacaoICMS.ListIndex = 0
'
'            iRegimeAnt = RegimeTributario.ItemData(RegimeTributario.ListIndex)
'
'        End If
'
'    End If
'
'    Exit Sub
'
'Erro_Trata_Regime_Tributario:
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174796)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Private Sub TipoTributacaoICMSSimples_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TipoTributacaoICMSSimples_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NatBCCred_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub IPICodEnqLabel_Click()

Dim colSelecao As New Collection
Dim objIPICodEnq As New ClassIPICodEnquadramento

    If TipoTributacaoIPI.ListIndex <> -1 Then

        colSelecao.Add TipoTributacaoIPI.ItemData(TipoTributacaoIPI.ListIndex)
        
        objIPICodEnq.sCodigo = Trim(IPICodEnq.Text)
        
        'Chama a tela de browse
        Call Chama_Tela("IPICodEnquadramentoLista", colSelecao, objIPICodEnq, objEventoIPICodEnq, "TipoIPI = ?")

    End If
    
End Sub

Private Sub objEventoIPICodEnq_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objIPICodEnq As ClassIPICodEnquadramento

On Error GoTo Erro_objEventoIPICodEnq_evSelecao

    Set objIPICodEnq = obj1

    IPICodEnq.Text = objIPICodEnq.sCodigo
    Call IPICodEnq_Validate(bSGECancelDummy)

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    Me.Show

    Exit Sub

Erro_objEventoIPICodEnq_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174795)

    End Select

    Exit Sub

End Sub

Private Sub IPICodEnq_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub IPICodEnq_GotFocus()
    Call MaskEdBox_TrataGotFocus(IPICodEnq, iAlterado)
End Sub

Private Sub IPICodEnq_Validate(Cancel As Boolean)
Dim lErro As Long
On Error GoTo Erro_IPICodEnq_Validate
    lErro = IPICodEnq_Valida
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    Exit Sub
Erro_IPICodEnq_Validate:
    Cancel = True
End Sub

Private Function IPICodEnq_Valida() As Long

Dim lErro As Long
Dim objIPICodEnq As New ClassIPICodEnquadramento

On Error GoTo Erro_IPICodEnq_Valida

    If Len(Trim(IPICodEnq.Text)) <> 0 Then

        objIPICodEnq.sCodigo = Format(Trim(IPICodEnq.Text), "000")
        
        lErro = CF("IPICodEnquadramento_Le", objIPICodEnq)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 216110 'Não cadastrado
        
        If TipoTributacaoIPI.ListIndex = -1 Then gError 216111 'Tem que preencher o CST do IPI
        
        If TipoTributacaoIPI.ItemData(TipoTributacaoIPI.ListIndex) <> objIPICodEnq.iTipoIPI Then gError 216112 'Tipo incompatível
        
        IPICodEnqDesc.Caption = objIPICodEnq.sDescCompleta
        
    Else
    
        IPICodEnqDesc.Caption = ""

    End If

    IPICodEnq_Valida = SUCESSO

    Exit Function

Erro_IPICodEnq_Valida:

    IPICodEnq_Valida = gErr

    Select Case gErr
    
        Case 216110
            Call Rotina_Erro(vbOKOnly, "ERRO_IPICODENQUADRAMENTO_NAO_CADASTRADO", gErr, objIPICodEnq.sCodigo)
                
        Case 216111
            Call Rotina_Erro(vbOKOnly, "ERRO_IPICODENQ_TIPOIPI_NAO_CADASTRADO", gErr)
        
        Case 216112
            Call Rotina_Erro(vbOKOnly, "ERRO_IPICODENQ_TIPOIPI_INCOMPATIVEL", gErr)
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216109)

    End Select

    Exit Function
    
End Function
