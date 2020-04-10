VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl OVAcompanhamentoOcx 
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   KeyPreview      =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   16995
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8295
      Index           =   2
      Left            =   75
      TabIndex        =   43
      Top             =   645
      Visible         =   0   'False
      Width           =   16560
      Begin VB.TextBox OBSGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   8355
         TabIndex        =   149
         Top             =   1410
         Visible         =   0   'False
         Width           =   30
      End
      Begin VB.TextBox FilialGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   3540
         TabIndex        =   103
         Top             =   1605
         Width           =   900
      End
      Begin VB.TextBox StatusGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   6825
         TabIndex        =   86
         Top             =   2715
         Width           =   3915
      End
      Begin VB.TextBox ClienteGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   3225
         TabIndex        =   85
         Top             =   2355
         Width           =   3255
      End
      Begin VB.Frame Frame3 
         Caption         =   "Detalhe"
         Height          =   4170
         Left            =   45
         TabIndex        =   75
         Top             =   4110
         Width           =   16470
         Begin VB.CommandButton BotaoImportarCRM 
            Caption         =   "Importar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   12135
            TabIndex        =   148
            Top             =   1530
            Width           =   1650
         End
         Begin VB.CommandButton BtnCRMPrimeiro 
            Height          =   1155
            Left            =   105
            Picture         =   "OVAcompanhamento.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   136
            ToolTipText     =   "Trazer para a tela o 1o registro"
            Top             =   2925
            Width           =   360
         End
         Begin VB.CommandButton BtnCRMAnterior 
            Height          =   1155
            Left            =   525
            Picture         =   "OVAcompanhamento.ctx":0312
            Style           =   1  'Graphical
            TabIndex        =   135
            ToolTipText     =   "Trazer para a tela o registro anterior"
            Top             =   2925
            Width           =   360
         End
         Begin VB.CommandButton BtnCRMProximo 
            Height          =   1155
            Left            =   12135
            Picture         =   "OVAcompanhamento.ctx":04BC
            Style           =   1  'Graphical
            TabIndex        =   134
            ToolTipText     =   "Trazer para a tela o registro seguinte"
            Top             =   2925
            Width           =   360
         End
         Begin VB.CommandButton BtnCRMUltimo 
            Height          =   1155
            Left            =   12540
            Picture         =   "OVAcompanhamento.ctx":0666
            Style           =   1  'Graphical
            TabIndex        =   133
            ToolTipText     =   "Trazer para a tela o último registro"
            Top             =   2925
            Width           =   360
         End
         Begin VB.CommandButton BotaoDocOriginal 
            Height          =   570
            Left            =   12135
            Picture         =   "OVAcompanhamento.ctx":0978
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   855
            Width           =   1650
         End
         Begin VB.TextBox Assunto 
            Height          =   1215
            Left            =   900
            MaxLength       =   510
            MultiLine       =   -1  'True
            TabIndex        =   38
            Top             =   1560
            Width           =   11175
         End
         Begin VB.CommandButton BotaoRelac 
            Caption         =   "Relacionamentos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   12135
            TabIndex        =   40
            Top             =   1950
            Width           =   1650
         End
         Begin VB.CommandButton BotaoGravar 
            Caption         =   "Gravar "
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
            Left            =   12150
            TabIndex        =   39
            Top             =   2370
            Width           =   1635
         End
         Begin VB.ComboBox Status 
            Height          =   315
            Left            =   915
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   1200
            Width           =   3375
         End
         Begin MSComCtl2.UpDown UpDownDataPrev 
            Height          =   300
            Left            =   6675
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   1185
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataPrev 
            Height          =   300
            Left            =   5685
            TabIndex        =   33
            ToolTipText     =   "Informe a data prevista para o recebimento."
            Top             =   1200
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataProx 
            Height          =   300
            Left            =   9555
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   1200
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataProx 
            Height          =   300
            Left            =   8565
            TabIndex        =   35
            ToolTipText     =   "Informe a data prevista para o recebimento."
            Top             =   1200
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LblDataCCRM 
            Height          =   195
            Left            =   8190
            TabIndex        =   147
            Top             =   2895
            Width           =   1140
         End
         Begin VB.Label LblDataFCRM 
            Height          =   195
            Left            =   10815
            TabIndex        =   146
            Top             =   2895
            Width           =   1140
         End
         Begin VB.Label Label1 
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
            Index           =   35
            Left            =   5250
            TabIndex        =   145
            Top             =   2895
            Width           =   480
         End
         Begin VB.Label LblDataCRM 
            Height          =   195
            Left            =   5790
            TabIndex        =   144
            Top             =   2895
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Próx.Contato:"
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
            Index           =   34
            Left            =   6930
            TabIndex        =   143
            Top             =   2895
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fechamento:"
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
            Index           =   33
            Left            =   9660
            TabIndex        =   142
            Top             =   2895
            Width           =   1110
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
            Height          =   195
            Index           =   32
            Left            =   3225
            TabIndex        =   141
            Top             =   2895
            Width           =   645
         End
         Begin VB.Label LblCodCRM 
            Height          =   195
            Left            =   3915
            TabIndex        =   140
            Top             =   2895
            Width           =   1140
         End
         Begin VB.Label LabelNumCRM 
            AutoSize        =   -1  'True
            Caption         =   "0 de 0"
            Height          =   195
            Left            =   2355
            TabIndex        =   139
            Top             =   2895
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Histórico CRM:"
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
            Index           =   31
            Left            =   960
            TabIndex        =   138
            Top             =   2895
            Width           =   1275
         End
         Begin VB.Label LblHistoricoCRM 
            BorderStyle     =   1  'Fixed Single
            Height          =   990
            Left            =   945
            TabIndex        =   137
            Top             =   3090
            Width           =   11115
         End
         Begin VB.Label ProjetoDesc 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   9780
            TabIndex        =   106
            Top             =   510
            Width           =   4005
         End
         Begin VB.Label ClienteNome 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1785
            TabIndex        =   105
            Top             =   510
            Width           =   4215
         End
         Begin VB.Label Filial 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   165
            TabIndex        =   104
            Top             =   1950
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label Responsavel 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   9735
            TabIndex        =   102
            Top             =   165
            Width           =   4035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Resp:"
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
            Left            =   9210
            TabIndex        =   101
            Top             =   195
            Width           =   510
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
            Height          =   195
            Index           =   10
            Left            =   6195
            TabIndex        =   100
            Top             =   195
            Width           =   510
         End
         Begin VB.Label Valor 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6705
            TabIndex        =   99
            Top             =   165
            Width           =   1200
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
            Height          =   195
            Index           =   9
            Left            =   4110
            TabIndex        =   98
            Top             =   225
            Width           =   750
         End
         Begin VB.Label Emissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4890
            TabIndex        =   97
            Top             =   180
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Projeto:"
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
            Left            =   7890
            TabIndex        =   96
            Top             =   555
            Width           =   675
         End
         Begin VB.Label Projeto 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   8565
            TabIndex        =   95
            Top             =   510
            Width           =   1200
         End
         Begin VB.Label Label1 
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
            Height          =   195
            Index           =   7
            Left            =   225
            TabIndex        =   94
            Top             =   555
            Width           =   660
         End
         Begin VB.Label Cliente 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   915
            TabIndex        =   93
            Top             =   510
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tel2:"
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
            Index           =   6
            Left            =   3405
            TabIndex        =   92
            Top             =   915
            Width           =   450
         End
         Begin VB.Label Telefone2 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3900
            TabIndex        =   91
            Top             =   855
            Width           =   2115
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tel1:"
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
            Left            =   435
            TabIndex        =   90
            Top             =   915
            Width           =   450
         End
         Begin VB.Label Telefone1 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   915
            TabIndex        =   89
            Top             =   855
            Width           =   2115
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
            Index           =   0
            Left            =   7815
            TabIndex        =   88
            Top             =   900
            Width           =   735
         End
         Begin VB.Label Contato 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   8565
            TabIndex        =   87
            Top             =   855
            Width           =   3540
         End
         Begin VB.Label LabelAssunto 
            AutoSize        =   -1  'True
            Caption         =   "Assunto:"
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
            Left            =   120
            TabIndex        =   84
            Top             =   1545
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fechamento:"
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
            Index           =   1
            Left            =   4575
            TabIndex        =   83
            Top             =   1245
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Próximo Contato:"
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
            Index           =   2
            Left            =   7095
            TabIndex        =   82
            Top             =   1245
            Width           =   1440
         End
         Begin VB.Label Label1 
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
            Height          =   195
            Index           =   4
            Left            =   180
            TabIndex        =   81
            Top             =   225
            Width           =   705
         End
         Begin VB.Label NumOV 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   915
            TabIndex        =   80
            Top             =   180
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Versão:"
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
            Index           =   5
            Left            =   2070
            TabIndex        =   79
            Top             =   225
            Width           =   645
         End
         Begin VB.Label Versao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2730
            TabIndex        =   78
            Top             =   180
            Width           =   375
         End
         Begin VB.Label NumIntOV 
            BorderStyle     =   1  'Fixed Single
            Height          =   105
            Left            =   6105
            TabIndex        =   77
            Top             =   240
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Status:"
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
            Left            =   255
            TabIndex        =   76
            Top             =   1230
            Width           =   615
         End
      End
      Begin MSMask.MaskEdBox DataPrevGrid 
         Height          =   225
         Left            =   8055
         TabIndex        =   49
         Top             =   2025
         Width           =   1365
         _ExtentX        =   2408
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
      Begin MSMask.MaskEdBox DataProxGrid 
         Height          =   225
         Left            =   6135
         TabIndex        =   48
         Top             =   675
         Width           =   1365
         _ExtentX        =   2408
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
      Begin MSMask.MaskEdBox EmissaoGrid 
         Height          =   225
         Left            =   1980
         TabIndex        =   47
         Top             =   2025
         Width           =   1095
         _ExtentX        =   1931
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
      Begin MSMask.MaskEdBox ValorGrid 
         Height          =   225
         Left            =   1065
         TabIndex        =   45
         Top             =   2475
         Width           =   1140
         _ExtentX        =   2011
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
      Begin VB.TextBox ProjetoGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   5175
         TabIndex        =   46
         Top             =   2025
         Width           =   1425
      End
      Begin VB.TextBox NumOVGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   315
         TabIndex        =   44
         Top             =   2025
         Width           =   735
      End
      Begin MSFlexGridLib.MSFlexGrid GridOV 
         Height          =   3705
         Left            =   60
         TabIndex        =   32
         Top             =   0
         Width           =   16455
         _ExtentX        =   29025
         _ExtentY        =   6535
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
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
         Index           =   24
         Left            =   13965
         TabIndex        =   53
         Top             =   3765
         Width           =   1005
      End
      Begin VB.Label ValorTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   15030
         TabIndex        =   52
         Top             =   3750
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8355
      Index           =   1
      Left            =   90
      TabIndex        =   42
      Top             =   585
      Width           =   16440
      Begin VB.Frame Frame2 
         Caption         =   "Filtros"
         Height          =   8130
         Left            =   300
         TabIndex        =   54
         Top             =   75
         Width           =   12615
         Begin VB.Frame Frame1 
            Caption         =   "Data de Emissão"
            Height          =   1620
            Index           =   0
            Left            =   555
            TabIndex        =   110
            Top             =   1395
            Width           =   8925
            Begin VB.OptionButton EntreEmi 
               Caption         =   "Entre"
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
               TabIndex        =   132
               Top             =   1230
               Width           =   780
            End
            Begin VB.OptionButton ApenasEmi 
               Caption         =   "Apenas"
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
               TabIndex        =   131
               Top             =   750
               Width           =   1050
            End
            Begin VB.OptionButton FaixaDataEmi 
               Caption         =   "Faixa de Datas"
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
               TabIndex        =   130
               Top             =   270
               Value           =   -1  'True
               Width           =   1620
            End
            Begin VB.Frame FrameD4 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   480
               Index           =   0
               Left            =   2295
               TabIndex        =   123
               Top             =   105
               Width           =   5130
               Begin MSMask.MaskEdBox DataEmiAte 
                  Height          =   300
                  Left            =   2325
                  TabIndex        =   124
                  Top             =   105
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   529
                  _Version        =   393216
                  MaxLength       =   8
                  Format          =   "dd/mm/yyyy"
                  Mask            =   "##/##/##"
                  PromptChar      =   " "
               End
               Begin MSComCtl2.UpDown UpDownDataEmiDe 
                  Height          =   300
                  Left            =   1440
                  TabIndex        =   125
                  TabStop         =   0   'False
                  Top             =   105
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   529
                  _Version        =   393216
                  Enabled         =   -1  'True
               End
               Begin MSMask.MaskEdBox DataEmiDe 
                  Height          =   300
                  Left            =   450
                  TabIndex        =   126
                  Top             =   105
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   529
                  _Version        =   393216
                  MaxLength       =   8
                  Format          =   "dd/mm/yyyy"
                  Mask            =   "##/##/##"
                  PromptChar      =   " "
               End
               Begin MSComCtl2.UpDown UpDownDataEmiAte 
                  Height          =   300
                  Left            =   3300
                  TabIndex        =   127
                  TabStop         =   0   'False
                  Top             =   105
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
                  Index           =   30
                  Left            =   1920
                  TabIndex        =   129
                  Top             =   165
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
                  Index           =   29
                  Left            =   90
                  TabIndex        =   128
                  Top             =   165
                  Width           =   315
               End
            End
            Begin VB.Frame FrameD4 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   405
               Index           =   1
               Left            =   1185
               TabIndex        =   119
               Top             =   645
               Width           =   6195
               Begin VB.ComboBox ApenasQualifEmi 
                  Height          =   315
                  ItemData        =   "OVAcompanhamento.ctx":388E
                  Left            =   90
                  List            =   "OVAcompanhamento.ctx":3898
                  Style           =   2  'Dropdown List
                  TabIndex        =   120
                  Top             =   45
                  Width           =   1365
               End
               Begin MSMask.MaskEdBox ApenasDiasEmi 
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   121
                  Top             =   45
                  Width           =   555
                  _ExtentX        =   979
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   4
                  Mask            =   "####"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "dia(s)"
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
                  Index           =   28
                  Left            =   2265
                  TabIndex        =   122
                  Top             =   90
                  Width           =   480
               End
            End
            Begin VB.Frame FrameD4 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   405
               Index           =   2
               Left            =   1230
               TabIndex        =   111
               Top             =   1185
               Width           =   6195
               Begin VB.ComboBox EntreQualifEmiDe 
                  Height          =   315
                  ItemData        =   "OVAcompanhamento.ctx":38B3
                  Left            =   1515
                  List            =   "OVAcompanhamento.ctx":38BD
                  Style           =   2  'Dropdown List
                  TabIndex        =   113
                  Top             =   0
                  Width           =   1365
               End
               Begin VB.ComboBox EntreQualifEmiAte 
                  Height          =   315
                  ItemData        =   "OVAcompanhamento.ctx":38D2
                  Left            =   4755
                  List            =   "OVAcompanhamento.ctx":38DC
                  Style           =   2  'Dropdown List
                  TabIndex        =   112
                  Top             =   15
                  Width           =   1365
               End
               Begin MSMask.MaskEdBox EntreDiasEmiDe 
                  Height          =   315
                  Left            =   45
                  TabIndex        =   114
                  Top             =   0
                  Width           =   555
                  _ExtentX        =   979
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   4
                  Mask            =   "####"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox EntreDiasEmiAte 
                  Height          =   315
                  Left            =   3390
                  TabIndex        =   115
                  Top             =   0
                  Width           =   555
                  _ExtentX        =   979
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   4
                  Mask            =   "####"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "dia(s)"
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
                  Index           =   27
                  Left            =   720
                  TabIndex        =   118
                  Top             =   45
                  Width           =   480
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "dia(s)"
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
                  Index           =   26
                  Left            =   4095
                  TabIndex        =   117
                  Top             =   45
                  Width           =   480
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "e"
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
                  Index           =   25
                  Left            =   3090
                  TabIndex        =   116
                  Top             =   45
                  Width           =   120
               End
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Ignorar Orçamento"
            Height          =   675
            Left            =   555
            TabIndex        =   109
            Top             =   435
            Width           =   8925
            Begin VB.CheckBox SoOVNaoPerdido 
               Caption         =   "Perdido"
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
               Left            =   6105
               TabIndex        =   2
               Top             =   315
               Value           =   1  'Checked
               Width           =   1200
            End
            Begin VB.CheckBox SoOVNaoFaturado 
               Caption         =   "Faturado"
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
               Left            =   3135
               TabIndex        =   1
               ToolTipText     =   "Vinculado à Pedido de Venda Baixado ou à Nota Fiscal"
               Top             =   315
               Value           =   1  'Checked
               Width           =   1215
            End
            Begin VB.CheckBox SoOVEmPV 
               Caption         =   "Sem Pedido"
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
               TabIndex        =   0
               ToolTipText     =   "Não vinculado à pedido de vendas"
               Top             =   315
               Width           =   1455
            End
         End
         Begin VB.Frame FrameVendedor 
            Caption         =   "Vendedor"
            Height          =   795
            Left            =   555
            TabIndex        =   107
            Top             =   7005
            Width           =   8925
            Begin MSMask.MaskEdBox Vendedor 
               Height          =   300
               Left            =   1275
               TabIndex        =   29
               Top             =   300
               Width           =   2895
               _ExtentX        =   5106
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               PromptChar      =   " "
            End
            Begin VB.Label LabelVendedor 
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
               Left            =   330
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   108
               Top             =   330
               Width           =   885
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Data Previsão Recebimento"
            Height          =   1725
            Index           =   3
            Left            =   555
            TabIndex        =   65
            Top             =   5010
            Width           =   8925
            Begin VB.Frame FrameD3 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   405
               Index           =   2
               Left            =   1275
               TabIndex        =   71
               Top             =   1155
               Width           =   6105
               Begin VB.ComboBox EntreQualifPrevDe 
                  Height          =   315
                  ItemData        =   "OVAcompanhamento.ctx":38F1
                  Left            =   1530
                  List            =   "OVAcompanhamento.ctx":38FB
                  Style           =   2  'Dropdown List
                  TabIndex        =   26
                  Top             =   15
                  Width           =   1365
               End
               Begin VB.ComboBox EntreQualifPrevAte 
                  Height          =   315
                  ItemData        =   "OVAcompanhamento.ctx":3910
                  Left            =   4740
                  List            =   "OVAcompanhamento.ctx":391A
                  Style           =   2  'Dropdown List
                  TabIndex        =   28
                  Top             =   30
                  Width           =   1365
               End
               Begin MSMask.MaskEdBox EntreDiasPrevDe 
                  Height          =   315
                  Left            =   0
                  TabIndex        =   25
                  Top             =   15
                  Width           =   555
                  _ExtentX        =   979
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   4
                  Mask            =   "####"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox EntreDiasPrevAte 
                  Height          =   315
                  Left            =   3405
                  TabIndex        =   27
                  Top             =   0
                  Width           =   555
                  _ExtentX        =   979
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   4
                  Mask            =   "####"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "dia(s)"
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
                  Index           =   15
                  Left            =   705
                  TabIndex        =   74
                  Top             =   60
                  Width           =   480
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "dia(s)"
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
                  Index           =   14
                  Left            =   4110
                  TabIndex        =   73
                  Top             =   60
                  Width           =   480
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "e"
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
                  Index           =   13
                  Left            =   3105
                  TabIndex        =   72
                  Top             =   60
                  Width           =   120
               End
            End
            Begin VB.Frame FrameD3 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   405
               Index           =   1
               Left            =   1275
               TabIndex        =   69
               Top             =   690
               Width           =   4965
               Begin VB.ComboBox ApenasQualifPrev 
                  Height          =   315
                  ItemData        =   "OVAcompanhamento.ctx":392F
                  Left            =   0
                  List            =   "OVAcompanhamento.ctx":3939
                  Style           =   2  'Dropdown List
                  TabIndex        =   23
                  Top             =   0
                  Width           =   1365
               End
               Begin MSMask.MaskEdBox ApenasDiasPrev 
                  Height          =   315
                  Left            =   1530
                  TabIndex        =   24
                  Top             =   0
                  Width           =   555
                  _ExtentX        =   979
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   4
                  Mask            =   "####"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "dia(s)"
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
                  Index           =   16
                  Left            =   2175
                  TabIndex        =   70
                  Top             =   45
                  Width           =   480
               End
            End
            Begin VB.Frame FrameD3 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   405
               Index           =   0
               Left            =   2445
               TabIndex        =   66
               Top             =   225
               Width           =   4965
               Begin MSMask.MaskEdBox DataPrevAte 
                  Height          =   300
                  Left            =   2235
                  TabIndex        =   21
                  Top             =   0
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   529
                  _Version        =   393216
                  MaxLength       =   8
                  Format          =   "dd/mm/yyyy"
                  Mask            =   "##/##/##"
                  PromptChar      =   " "
               End
               Begin MSComCtl2.UpDown UpDownDataPrevDe 
                  Height          =   300
                  Left            =   1350
                  TabIndex        =   20
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   529
                  _Version        =   393216
                  Enabled         =   -1  'True
               End
               Begin MSMask.MaskEdBox DataPrevDe 
                  Height          =   300
                  Left            =   360
                  TabIndex        =   19
                  Top             =   0
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   529
                  _Version        =   393216
                  MaxLength       =   8
                  Format          =   "dd/mm/yyyy"
                  Mask            =   "##/##/##"
                  PromptChar      =   " "
               End
               Begin MSComCtl2.UpDown UpDownDataPrevAte 
                  Height          =   300
                  Left            =   3210
                  TabIndex        =   22
                  TabStop         =   0   'False
                  Top             =   0
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
                  Index           =   18
                  Left            =   1830
                  TabIndex        =   68
                  Top             =   60
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
                  Index           =   17
                  Left            =   0
                  TabIndex        =   67
                  Top             =   60
                  Width           =   315
               End
            End
            Begin VB.OptionButton EntrePrev 
               Caption         =   "Entre"
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
               TabIndex        =   18
               Top             =   1230
               Width           =   780
            End
            Begin VB.OptionButton ApenasPrev 
               Caption         =   "Apenas"
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
               Left            =   240
               TabIndex        =   17
               Top             =   750
               Width           =   1050
            End
            Begin VB.OptionButton FaixaDataPrev 
               Caption         =   "Faixa de Datas"
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
               TabIndex        =   16
               Top             =   270
               Value           =   -1  'True
               Width           =   1620
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Data Próximo Contato"
            Height          =   1620
            Index           =   4
            Left            =   555
            TabIndex        =   55
            Top             =   3195
            Width           =   8925
            Begin VB.Frame FrameD1 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   405
               Index           =   2
               Left            =   1230
               TabIndex        =   61
               Top             =   1185
               Width           =   6195
               Begin VB.ComboBox EntreQualifProxAte 
                  Height          =   315
                  ItemData        =   "OVAcompanhamento.ctx":3954
                  Left            =   4755
                  List            =   "OVAcompanhamento.ctx":395E
                  Style           =   2  'Dropdown List
                  TabIndex        =   15
                  Top             =   15
                  Width           =   1365
               End
               Begin VB.ComboBox EntreQualifProxDe 
                  Height          =   315
                  ItemData        =   "OVAcompanhamento.ctx":3973
                  Left            =   1515
                  List            =   "OVAcompanhamento.ctx":397D
                  Style           =   2  'Dropdown List
                  TabIndex        =   13
                  Top             =   0
                  Width           =   1365
               End
               Begin MSMask.MaskEdBox EntreDiasProxDe 
                  Height          =   315
                  Left            =   45
                  TabIndex        =   12
                  Top             =   0
                  Width           =   555
                  _ExtentX        =   979
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   4
                  Mask            =   "####"
                  PromptChar      =   " "
               End
               Begin MSMask.MaskEdBox EntreDiasProxAte 
                  Height          =   315
                  Left            =   3390
                  TabIndex        =   14
                  Top             =   0
                  Width           =   555
                  _ExtentX        =   979
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   4
                  Mask            =   "####"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "e"
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
                  Index           =   12
                  Left            =   3090
                  TabIndex        =   64
                  Top             =   45
                  Width           =   120
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "dia(s)"
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
                  Index           =   19
                  Left            =   4095
                  TabIndex        =   63
                  Top             =   45
                  Width           =   480
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "dia(s)"
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
                  Index           =   20
                  Left            =   720
                  TabIndex        =   62
                  Top             =   45
                  Width           =   480
               End
            End
            Begin VB.Frame FrameD1 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   405
               Index           =   1
               Left            =   1185
               TabIndex        =   59
               Top             =   645
               Width           =   6195
               Begin VB.ComboBox ApenasQualifProx 
                  Height          =   315
                  ItemData        =   "OVAcompanhamento.ctx":3992
                  Left            =   90
                  List            =   "OVAcompanhamento.ctx":399C
                  Style           =   2  'Dropdown List
                  TabIndex        =   10
                  Top             =   45
                  Width           =   1365
               End
               Begin MSMask.MaskEdBox ApenasDiasProx 
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   11
                  Top             =   45
                  Width           =   555
                  _ExtentX        =   979
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   4
                  Mask            =   "####"
                  PromptChar      =   " "
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "dia(s)"
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
                  Index           =   21
                  Left            =   2265
                  TabIndex        =   60
                  Top             =   90
                  Width           =   480
               End
            End
            Begin VB.Frame FrameD1 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   480
               Index           =   0
               Left            =   2295
               TabIndex        =   56
               Top             =   105
               Width           =   5130
               Begin MSMask.MaskEdBox DataProxAte 
                  Height          =   300
                  Left            =   2325
                  TabIndex        =   8
                  Top             =   105
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   529
                  _Version        =   393216
                  MaxLength       =   8
                  Format          =   "dd/mm/yyyy"
                  Mask            =   "##/##/##"
                  PromptChar      =   " "
               End
               Begin MSComCtl2.UpDown UpDownDataProxDe 
                  Height          =   300
                  Left            =   1440
                  TabIndex        =   7
                  TabStop         =   0   'False
                  Top             =   105
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   529
                  _Version        =   393216
                  Enabled         =   -1  'True
               End
               Begin MSMask.MaskEdBox DataProxDe 
                  Height          =   300
                  Left            =   450
                  TabIndex        =   6
                  Top             =   105
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   529
                  _Version        =   393216
                  MaxLength       =   8
                  Format          =   "dd/mm/yyyy"
                  Mask            =   "##/##/##"
                  PromptChar      =   " "
               End
               Begin MSComCtl2.UpDown UpDownDataProxAte 
                  Height          =   300
                  Left            =   3300
                  TabIndex        =   9
                  TabStop         =   0   'False
                  Top             =   105
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   529
                  _Version        =   393216
                  Enabled         =   -1  'True
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
                  Index           =   22
                  Left            =   90
                  TabIndex        =   58
                  Top             =   165
                  Width           =   315
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
                  Index           =   23
                  Left            =   1920
                  TabIndex        =   57
                  Top             =   165
                  Width           =   360
               End
            End
            Begin VB.OptionButton FaixaDataProx 
               Caption         =   "Faixa de Datas"
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
               TabIndex        =   3
               Top             =   270
               Value           =   -1  'True
               Width           =   1620
            End
            Begin VB.OptionButton ApenasProx 
               Caption         =   "Apenas"
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
               TabIndex        =   4
               Top             =   750
               Width           =   1050
            End
            Begin VB.OptionButton EntreProx 
               Caption         =   "Entre"
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
               TabIndex        =   5
               Top             =   1230
               Width           =   780
            End
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   15720
      ScaleHeight     =   450
      ScaleWidth      =   1020
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   0
      Width           =   1080
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   60
         Picture         =   "OVAcompanhamento.ctx":39B7
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   540
         Picture         =   "OVAcompanhamento.ctx":3EE9
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8835
      Left            =   45
      TabIndex        =   50
      Top             =   255
      Width           =   16755
      _ExtentX        =   29554
      _ExtentY        =   15584
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Orçamentos de Venda"
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
Attribute VB_Name = "OVAcompanhamentoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iAlteradoGrid As Integer
Dim iAlteradoTab As Integer
Dim iFrameAtual As Integer

Dim bAbrindoTela As Boolean

Dim gbTrazendoDados As Boolean

Dim objGridOV As AdmGrid
Dim iGrid_NumOV_Col As Integer
Dim iGrid_Emissao_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_Cliente_Col As Integer
Dim iGrid_Filial_Col As Integer
Dim iGrid_Projeto_Col As Integer
Dim iGrid_Status_Col As Integer
Dim iGrid_DataProx_Col As Integer
Dim iGrid_DataPrev_Col As Integer

Dim iStatus_ListIndex_Padrao As Integer

Dim gobjOVAcompAnt As ClassOVAcomp

Dim gcolOVs As Collection
Dim gcolClientes As Collection
Dim gcolFiliais As Collection
Dim gcolEnderecos As Collection
Dim gcolProjetos As Collection

Dim gcolCRM As Collection
Dim giNumCRM As Integer

Const TAB_Selecao = 1
Const TAB_OV = 2

Const FRAMED_FAIXA = 0
Const FRAMED_APENAS = 1
Const FRAMED_ENTRE = 2

Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Orçamentos de Venda - Acompanhamento"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "OVAcompanhamento"

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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
    
    End If
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty(True, UserControl.Enabled, True)
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
    'Call TelaIndice_Preenche(Me)

End Sub
    
Public Sub Form_Deactivate()
    
    'gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload
    
    Set objGridOV = Nothing
    Set objEventoVendedor = Nothing
    Set gobjOVAcompAnt = Nothing
    
    Set gcolOVs = Nothing
    Set gcolClientes = Nothing
    Set gcolFiliais = Nothing
    Set gcolEnderecos = Nothing
    Set gcolProjetos = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182210)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long
Dim colVend As New Collection
Dim objVend As ClassVendedor

On Error GoTo Erro_Form_Load

    gbTrazendoDados = False

    Set objGridOV = New AdmGrid
    Set objEventoVendedor = New AdmEvento
    Set gobjOVAcompAnt = New ClassOVAcomp
    
    Set gcolOVs = New Collection
    Set gcolClientes = New Collection
    Set gcolFiliais = New Collection
    Set gcolEnderecos = New Collection
    Set gcolProjetos = New Collection
    Set gcolCRM = New Collection
    
    lErro = Inicializa_GridOV(objGridOV)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call FrameD_Enabled(FrameD1, FRAMED_FAIXA)
    Call FrameD_Enabled(FrameD3, FRAMED_FAIXA)
    Call FrameD_Enabled(FrameD4, FRAMED_FAIXA)
    
    Call DateParaMasked(DataEmiDe, gdtDataHoje)
    
    Call Carrega_Status(Status)
    
    If gobjFAT.iOVAcompFiltraVend = MARCADO Then
    
        lErro = CF("VendedorAtivo_Le_Todos", colVend)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        For Each objVend In colVend
            If objVend.sCodUsuario = gsUsuario Then
                If gobjFAT.iOVAcompFiltroVendObr = MARCADO Then FrameVendedor.Enabled = False
                Vendedor.Text = CStr(objVend.iCodigo)
                Call Vendedor_Validate(bSGECancelDummy)
                Exit For
            End If
        Next
        
    End If
    
    iFrameAtual = TAB_Selecao
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO
   
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182212)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Move_Selecao_Memoria(ByVal objOVAcomp As ClassOVAcomp) As Long

Dim lErro As Long
Dim dtDataDe As Date
Dim dtDataAte As Date
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer

On Error GoTo Erro_Move_Selecao_Memoria

    If SoOVEmPV.Value = vbUnchecked Then
        objOVAcomp.iSoEmPV = DESMARCADO
    Else
        objOVAcomp.iSoEmPV = MARCADO
    End If

    If SoOVNaoFaturado.Value = vbUnchecked Then
        objOVAcomp.iSoNaoFaturado = DESMARCADO
    Else
        objOVAcomp.iSoNaoFaturado = MARCADO
    End If

    If SoOVNaoPerdido.Value = vbUnchecked Then
        objOVAcomp.iSoNaoPerdido = DESMARCADO
    Else
        objOVAcomp.iSoNaoPerdido = MARCADO
    End If
    
    If FaixaDataEmi.Value Then
        objOVAcomp.dtDataEmiDe = StrParaDate(DataEmiDe.Text)
        objOVAcomp.dtDataEmiAte = StrParaDate(DataEmiAte.Text)
    End If

    If ApenasEmi.Value Then
    
        If ApenasQualifEmi.ListIndex = -1 Then gError 182213
    
        Call Datas_Trata_Apenas(dtDataDe, dtDataAte, StrParaInt(ApenasDiasEmi.Text), ApenasQualifEmi.ItemData(ApenasQualifEmi.ListIndex))
        
        objOVAcomp.dtDataEmiDe = dtDataDe
        objOVAcomp.dtDataEmiAte = dtDataAte
    End If
    
    If EntreEmi.Value Then
        
        If EntreQualifEmiAte.ListIndex = -1 Then gError 182214
        If EntreQualifEmiAte.ListIndex = -1 Then gError 182215
        
        Call Datas_Trata_Entre(dtDataDe, dtDataAte, StrParaInt(EntreDiasEmiDe), EntreQualifEmiDe.ItemData(EntreQualifEmiDe.ListIndex), StrParaInt(EntreDiasEmiAte), EntreQualifEmiAte.ItemData(EntreQualifEmiAte.ListIndex))
        
        objOVAcomp.dtDataEmiDe = dtDataDe
        objOVAcomp.dtDataEmiAte = dtDataAte
    End If

    If FaixaDataProx.Value Then
        objOVAcomp.dtDataProxDe = StrParaDate(DataProxDe.Text)
        objOVAcomp.dtDataProxAte = StrParaDate(DataProxAte.Text)
    End If

    If ApenasProx.Value Then
    
        If ApenasQualifProx.ListIndex = -1 Then gError 182213
    
        Call Datas_Trata_Apenas(dtDataDe, dtDataAte, StrParaInt(ApenasDiasProx.Text), ApenasQualifProx.ItemData(ApenasQualifProx.ListIndex))
        
        objOVAcomp.dtDataProxDe = dtDataDe
        objOVAcomp.dtDataProxAte = dtDataAte
    End If
    
    If EntreProx.Value Then
        
        If EntreQualifProxAte.ListIndex = -1 Then gError 182214
        If EntreQualifProxAte.ListIndex = -1 Then gError 182215
        
        Call Datas_Trata_Entre(dtDataDe, dtDataAte, StrParaInt(EntreDiasProxDe), EntreQualifProxDe.ItemData(EntreQualifProxDe.ListIndex), StrParaInt(EntreDiasProxAte), EntreQualifProxAte.ItemData(EntreQualifProxAte.ListIndex))
        
        objOVAcomp.dtDataProxDe = dtDataDe
        objOVAcomp.dtDataProxAte = dtDataAte
    End If
    
    If FaixaDataPrev.Value Then
        objOVAcomp.dtDataPrevDe = StrParaDate(DataPrevDe.Text)
        objOVAcomp.dtDataPrevAte = StrParaDate(DataPrevAte.Text)
    End If

    If ApenasPrev.Value Then
    
        If ApenasQualifPrev.ListIndex = -1 Then gError 182219

        Call Datas_Trata_Apenas(dtDataDe, dtDataAte, StrParaInt(ApenasDiasPrev.Text), ApenasQualifPrev.ItemData(ApenasQualifPrev.ListIndex))
        
        objOVAcomp.dtDataPrevDe = dtDataDe
        objOVAcomp.dtDataPrevAte = dtDataAte
    End If
    
    If EntrePrev.Value Then
        
        If EntreQualifPrevDe.ListIndex = -1 Then gError 182220
        If EntreQualifPrevAte.ListIndex = -1 Then gError 182221
        
        Call Datas_Trata_Entre(dtDataDe, dtDataAte, StrParaInt(EntreDiasPrevDe), EntreQualifPrevDe.ItemData(EntreQualifPrevDe.ListIndex), StrParaInt(EntreDiasPrevAte.Text), EntreQualifPrevAte.ItemData(EntreQualifPrevAte.ListIndex))
        
        objOVAcomp.dtDataPrevDe = dtDataDe
        objOVAcomp.dtDataPrevAte = dtDataAte
    End If
    
    If objOVAcomp.dtDataPrevAte <> DATA_NULA And objOVAcomp.dtDataPrevDe <> DATA_NULA Then
        If objOVAcomp.dtDataPrevDe > objOVAcomp.dtDataPrevAte Then gError 182222
    End If
    
    If objOVAcomp.dtDataProxDe <> DATA_NULA And objOVAcomp.dtDataProxAte <> DATA_NULA Then
        If objOVAcomp.dtDataProxDe > objOVAcomp.dtDataProxAte Then gError 182223
    End If
    
    If objOVAcomp.dtDataEmiDe <> DATA_NULA And objOVAcomp.dtDataEmiAte <> DATA_NULA Then
        If objOVAcomp.dtDataEmiDe > objOVAcomp.dtDataEmiAte Then gError 182223
    End If
    
    objOVAcomp.iVendedor = Codigo_Extrai(Vendedor.Text)
       
    Move_Selecao_Memoria = SUCESSO

    Exit Function

Erro_Move_Selecao_Memoria:

    Move_Selecao_Memoria = gErr

    Select Case gErr
    
        Case 182222 To 182224
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)
            
        Case 182213 To 182221
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_TRATA_TIPO", gErr)
            
        Case 182225
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 182226
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 182227

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182228)

    End Select

    Exit Function

End Function

Function Datas_Trata_Entre(dtDataDe As Date, dtDataAte As Date, ByVal iNumDias1 As Integer, ByVal iData1 As Integer, ByVal iNumDias2 As Integer, ByVal iData2 As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Datas_Trata_Entre

    Select Case iData1
    
        Case DATA_AFRENTE
            dtDataDe = DateAdd("d", iNumDias1, gdtDataAtual)
        
        Case DATA_ATRAS
            dtDataDe = DateAdd("d", -iNumDias1, gdtDataAtual)
            
        Case Else
            gError 182229

    End Select
    
    Select Case iData2
    
        Case DATA_AFRENTE
            dtDataAte = DateAdd("d", iNumDias2, gdtDataAtual)
        
        Case DATA_ATRAS
            dtDataAte = DateAdd("d", -iNumDias2, gdtDataAtual)
        
        Case Else
            gError 182230

    End Select
   
    Datas_Trata_Entre = SUCESSO

    Exit Function

Erro_Datas_Trata_Entre:

    Datas_Trata_Entre = gErr

    Select Case gErr
    
        Case 182229, 182230
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_TRATA_TIPO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182231)

    End Select

    Exit Function

End Function

Function Datas_Trata_Apenas(dtDataDe As Date, dtDataAte As Date, ByVal iNumDias As Integer, ByVal iData As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Datas_Trata_Apenas

    Select Case iData
    
        Case DATA_AFRENTE
            dtDataDe = gdtDataAtual
            dtDataAte = DateAdd("d", iNumDias, gdtDataAtual)
        
        Case DATA_ATRAS
            dtDataDe = DateAdd("d", -iNumDias, gdtDataAtual)
            dtDataAte = gdtDataAtual
            
        Case Else
            gError 182232

    End Select
   
    Datas_Trata_Apenas = SUCESSO

    Exit Function

Erro_Datas_Trata_Apenas:

    Datas_Trata_Apenas = gErr

    Select Case gErr
    
        Case 182232
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_TRATA_TIPO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182233)

    End Select

    Exit Function

End Function

Function Trata_Selecao(ByVal objOVAcomp As ClassOVAcomp) As Long

Dim lErro As Long
Dim colOVs As New Collection

On Error GoTo Erro_Trata_Selecao

    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = CF("OVAcomp_Le", objOVAcomp, colOVs)
    If lErro <> SUCESSO Then gError 182234
        
        
    gbTrazendoDados = True
    lErro = Preenche_GridOV(colOVs)
    gbTrazendoDados = False
    If lErro <> SUCESSO Then gError 182236
    
    If colOVs.Count = 0 Then gError 182235
    GridOV.Row = 1
    Call Trata_RowColChange
    
    GL_objMDIForm.MousePointer = vbDefault
   
    Trata_Selecao = SUCESSO

    Exit Function

Erro_Trata_Selecao:

    GL_objMDIForm.MousePointer = vbDefault

    Trata_Selecao = gErr

    Select Case gErr
    
        Case 182234, 182236
        
        Case 182235
            Call Rotina_Erro(vbOKOnly, "ERRO_SELECAO_OVACOMP_SEM_PARCELAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182237)

    End Select

    Exit Function

End Function

Function Preenche_GridOV(ByVal colOVs As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objOV As ClassOrcamentoVenda
Dim objcliente As ClassCliente
Dim objFilialCliente As ClassFilialCliente
Dim objEndereco As ClassEndereco
Dim objProjeto As ClassProjetos
Dim sProjeto As String

On Error GoTo Erro_Preenche_GridOV

    Set gcolOVs = New Collection
    Set gcolClientes = New Collection
    Set gcolFiliais = New Collection
    Set gcolEnderecos = New Collection
    Set gcolProjetos = New Collection

    Call Grid_Limpa(objGridOV)
    
    NumOV.Caption = ""
    Versao.Caption = ""
    Emissao.Caption = ""
    Valor.Caption = ""
    Cliente.Caption = ""
    Projeto.Caption = ""
    Contato.Caption = ""
    Telefone1.Caption = ""
    Telefone2.Caption = ""
    NumIntOV.Caption = ""
    ValorTotal.Caption = ""
    Responsavel.Caption = ""
    
    'Aumenta o número de linhas do grid se necessário
    If colOVs.Count >= objGridOV.objGrid.Rows Then
        Call Refaz_Grid(objGridOV, colOVs.Count)
    End If

    iIndice = 0
    For Each objOV In colOVs
    
        iIndice = iIndice + 1
    
        Set objcliente = New ClassCliente
        Set objFilialCliente = New ClassFilialCliente
        Set objEndereco = New ClassEndereco
        Set objProjeto = New ClassProjetos
        
        gcolClientes.Add objcliente
        gcolFiliais.Add objFilialCliente
        gcolEnderecos.Add objEndereco
        gcolProjetos.Add objProjeto
        
        objcliente.lCodigo = objOV.lCliente
        objFilialCliente.lCodCliente = objOV.lCliente
        objFilialCliente.iCodFilial = objOV.iFilial
    
        lErro = CF("Cliente_Le", objcliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError ERRO_SEM_MENSAGEM
        
        If lErro = SUCESSO Then
                
            lErro = CF("FilialCliente_Le", objFilialCliente)
            If lErro <> SUCESSO And lErro <> 12567 Then gError ERRO_SEM_MENSAGEM
    
            objEndereco.lCodigo = objFilialCliente.lEndereco
            
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        Set objOV.objProjetoInfo = New ClassProjetoInfo
        
        objOV.objProjetoInfo.iTipoOrigem = PRJ_CR_TIPO_OV
        objOV.objProjetoInfo.lNumIntDocOrigem = objOV.lNumIntDoc
        objOV.objProjetoInfo.sCodigoOP = ""
        objOV.objProjetoInfo.iFilialEmpresa = objOV.iFilialEmpresa
    
        'Le as associação gravadas no BD para esse tipo de documento
        lErro = CF("ProjetoInfo_Le", objOV.objProjetoInfo)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objProjeto.lNumIntDoc = objOV.objProjetoInfo.lNumIntDocPRJ
        
        If objProjeto.lNumIntDoc <> 0 Then
            
            lErro = CF("Projetos_Le_NumIntDoc", objProjeto)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        
            Call Retorno_Projeto_Tela2(objProjeto.sCodigo, sProjeto)
        
            GridOV.TextMatrix(iIndice, iGrid_Projeto_Col) = sProjeto & SEPARADOR & objProjeto.sDescricao
        
        End If
                   
        GridOV.TextMatrix(iIndice, iGrid_Cliente_Col) = CStr(objcliente.lCodigo) & SEPARADOR & objcliente.sNomeReduzido
        GridOV.TextMatrix(iIndice, iGrid_Filial_Col) = CStr(objFilialCliente.iCodFilial) & SEPARADOR & objFilialCliente.sNome
                
        GridOV.TextMatrix(iIndice, iGrid_Valor_Col) = Format(objOV.dValorTotal, "STANDARD")
        GridOV.TextMatrix(iIndice, iGrid_NumOV_Col) = CStr(objOV.lCodigo)
        GridOV.TextMatrix(iIndice, iGrid_Emissao_Col) = Format(objOV.dtDataEmissao, "dd/mm/yyyy")

        Call Combo_Seleciona_ItemData(Status, objOV.lStatus)

        GridOV.TextMatrix(iIndice, iGrid_Status_Col) = Status.Text

        If objOV.dtDataPrevReceb <> DATA_NULA Then
            GridOV.TextMatrix(iIndice, iGrid_DataPrev_Col) = Format(objOV.dtDataPrevReceb, "dd/mm/yyyy")
        End If
        
        If objOV.dtDataProxCobr <> DATA_NULA Then
            GridOV.TextMatrix(iIndice, iGrid_DataProx_Col) = Format(objOV.dtDataProxCobr, "dd/mm/yyyy")
        End If

    Next
        
    objGridOV.iLinhasExistentes = iIndice
    
    Call Grid_Refresh_Checkbox(objGridOV)
    
    Set gcolOVs = colOVs
       
    Call Soma_Coluna_Grid(objGridOV, iGrid_Valor_Col, ValorTotal, False)
   
    Preenche_GridOV = SUCESSO

    Exit Function

Erro_Preenche_GridOV:

    Preenche_GridOV = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182241)

    End Select

    Exit Function

End Function

Function Trata_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
   
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182242)

    End Select

    bAbrindoTela = False
    
    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objRelacCli As ClassRelacClientes) As Long

Dim lErro As Long
Dim iCodFilial As Integer
Dim lCliente As Long

On Error GoTo Erro_Move_Tela_Memoria

    objRelacCli.lCliente = LCodigo_Extrai(Cliente.Caption)
    objRelacCli.iFilialCliente = Codigo_Extrai(Filial.Caption)
    objRelacCli.dtData = Date
    objRelacCli.dtDataPrevReceb = StrParaDate(DataPrev.Text)
    objRelacCli.dtDataProxCobr = StrParaDate(DataProx.Text)
    objRelacCli.dtDataFim = DATA_NULA
    objRelacCli.iFilialEmpresa = giFilialEmpresa
    'Guarda no obj a primeira parte do assunto
    objRelacCli.sAssunto1 = left(Assunto.Text, STRING_BUFFER_MAX_TEXTO - 1)
    'Guarda no obj a segunda parte do assunto
    objRelacCli.sAssunto2 = Mid(Assunto.Text, STRING_BUFFER_MAX_TEXTO)
    
    objRelacCli.lNumIntDocOrigem = StrParaLong(NumIntOV.Caption)
    objRelacCli.iTipoDoc = RELACCLI_TIPODOC_OV
    
    objRelacCli.iOrigem = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA
    objRelacCli.dtHora = Time
    objRelacCli.iContato = 0
    objRelacCli.lTipo = TIPO_RELACIONAMENTO_OVACOMP
    
    If Status.ListIndex <> -1 Then
        objRelacCli.lStatusTipoDoc = Status.ItemData(Status.ListIndex)
    Else
        objRelacCli.lStatusTipoDoc = 0
    End If
   
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 182280

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182243)

    End Select

    Exit Function

End Function

Function Trata_TabClick(Optional ByVal bAtualiza As Boolean = False) As Long

Dim lErro As Long
Dim objOVAcomp As New ClassOVAcomp

On Error GoTo Erro_Trata_TabClick

    'Valida a seleção
    lErro = Move_Selecao_Memoria(objOVAcomp)
    If lErro <> SUCESSO Then gError 182244
    
    If (objOVAcomp.dtDataPrevAte <> gobjOVAcompAnt.dtDataPrevAte Or _
        objOVAcomp.dtDataPrevDe <> gobjOVAcompAnt.dtDataPrevDe Or _
        objOVAcomp.dtDataProxAte <> gobjOVAcompAnt.dtDataProxAte Or _
        objOVAcomp.dtDataProxDe <> gobjOVAcompAnt.dtDataProxDe Or _
        objOVAcomp.dtDataEmiAte <> gobjOVAcompAnt.dtDataEmiAte Or _
        objOVAcomp.dtDataEmiDe <> gobjOVAcompAnt.dtDataEmiDe Or _
        objOVAcomp.iSoEmPV <> gobjOVAcompAnt.iSoEmPV Or _
        objOVAcomp.iSoNaoFaturado <> gobjOVAcompAnt.iSoNaoFaturado Or _
        objOVAcomp.iSoNaoPerdido <> gobjOVAcompAnt.iSoNaoPerdido Or _
        objOVAcomp.iVendedor <> gobjOVAcompAnt.iVendedor) Or _
        bAtualiza Then
    
        lErro = Trata_Selecao(objOVAcomp)
        If lErro <> SUCESSO Then
            Set gobjOVAcompAnt = New ClassOVAcomp
            gError 182245
        End If
        
        Set gobjOVAcompAnt = objOVAcomp
        
    End If
   
    Trata_TabClick = SUCESSO

    Exit Function

Erro_Trata_TabClick:

    Trata_TabClick = gErr

    Select Case gErr
    
        Case 182244, 182245

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182246)

    End Select

    Exit Function

End Function

Function Trata_RowColChange() As Long

Dim lErro As Long
Dim objOV As ClassOrcamentoVenda
Dim objcliente As ClassCliente
Dim objFilialCliente As ClassFilialCliente
Dim objEndereco As ClassEndereco
Dim objProjeto As ClassProjetos
Dim sProjeto As String

On Error GoTo Erro_Trata_RowColChange

    If GridOV.Row <> 0 And Not gbTrazendoDados Then
    
        'Se não é na abertura da tela
        If gcolOVs.Count <> 0 Then
    
            Set objOV = gcolOVs.Item(GridOV.Row)
            Set objcliente = gcolClientes.Item(GridOV.Row)
            Set objFilialCliente = gcolFiliais.Item(GridOV.Row)
            Set objEndereco = gcolEnderecos.Item(GridOV.Row)
            Set objProjeto = gcolProjetos.Item(GridOV.Row)
            
            'Se mudou a parcela
            If StrParaLong(NumIntOV.Caption) <> objOV.lNumIntDoc Then
            
                'Testa se deseja salvar mudanças
                lErro = Teste_Salva(Me, iAlterado)
                If lErro <> SUCESSO Then gError 182285
                                
                Call Limpa_Relacionamento
            
                NumOV.Caption = CStr(objOV.lCodigo)
                Versao.Caption = CStr(objOV.iVersao)
                Emissao.Caption = Format(objOV.dtDataEmissao, "dd/mm/yyyy")
                Valor.Caption = Format(objOV.dValorTotal, "STANDARD")
                Cliente.Caption = CStr(objcliente.lCodigo)
                ClienteNome.Caption = objcliente.sNomeReduzido
                Filial.Caption = CStr(objFilialCliente.iCodFilial) & SEPARADOR & objFilialCliente.sNome
                
                If Len(Trim(objProjeto.sCodigo)) > 0 Then
                    Call Retorno_Projeto_Tela2(objProjeto.sCodigo, sProjeto)
                    Projeto.Caption = sProjeto
                    ProjetoDesc.Caption = objProjeto.sDescricao
                Else
                    Projeto.Caption = ""
                    ProjetoDesc.Caption = ""
                End If
                                
                Contato.Caption = objEndereco.sContato
                Telefone1.Caption = objEndereco.sTelefone1
                Telefone2.Caption = objEndereco.sTelefone2
                NumIntOV.Caption = CStr(objOV.lNumIntDoc)
                Responsavel.Caption = objProjeto.sResponsavel
                
                DataPrev.PromptInclude = False
                If objOV.dtDataPrevReceb <> DATA_NULA Then
                    DataPrev.Text = Format(objOV.dtDataPrevReceb, "dd/mm/yy")
                Else
                    DataPrev.Text = ""
                End If
                DataPrev.PromptInclude = True
                
                DataProx.PromptInclude = False
                If objOV.dtDataProxCobr <> DATA_NULA Then
                    DataProx.Text = Format(objOV.dtDataProxCobr, "dd/mm/yy")
                Else
                    DataProx.Text = ""
                End If
                DataProx.PromptInclude = True
                
                Call Combo_Seleciona_ItemData(Status, objOV.lStatus)
                
                lErro = Trata_HistoricoCRM(objOV)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                iAlterado = 0
                
            End If
            
        End If
    
    End If
   
    Trata_RowColChange = SUCESSO

    Exit Function

Erro_Trata_RowColChange:

    Trata_RowColChange = gErr

    Select Case gErr
    
        Case 182285, ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182281)

    End Select

    Exit Function

End Function

Private Sub TabStrip1_BeforeClick(Cancel As Integer)

Dim lErro As Long
On Error GoTo Erro_TabStrip1_BeforeClick

    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
    
    'Se estava no tab de seleção e está passando para outro tab
    If iFrameAtual = TAB_Selecao Then
    
        lErro = Trata_TabClick
        If lErro <> SUCESSO Then gError 182247
    
    End If

    Exit Sub

Erro_TabStrip1_BeforeClick:

    Cancel = True

    Select Case gErr
    
        Case 182247
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182248)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
        
        If iFrameAtual = TAB_OV Then GridOV.SetFocus
        
    End If

End Sub

Private Function Inicializa_GridOV(objGrid As AdmGrid) As Long

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Número")
    objGrid.colColuna.Add ("Cliente")
    objGrid.colColuna.Add ("Filial")
    objGrid.colColuna.Add ("Projeto")
    objGrid.colColuna.Add ("Emissão")
    objGrid.colColuna.Add ("Valor")
    objGrid.colColuna.Add ("Status")
    objGrid.colColuna.Add ("Próx.Contato")
    objGrid.colColuna.Add ("Fechamento")
    objGrid.colColuna.Add ("Observação")
    
    'Controles que participam do Grid
    objGrid.colCampo.Add (NumOVGrid.Name)
    objGrid.colCampo.Add (ClienteGrid.Name)
    objGrid.colCampo.Add (FilialGrid.Name)
    objGrid.colCampo.Add (ProjetoGrid.Name)
    objGrid.colCampo.Add (EmissaoGrid.Name)
    objGrid.colCampo.Add (ValorGrid.Name)
    objGrid.colCampo.Add (StatusGrid.Name)
    objGrid.colCampo.Add (DataProxGrid.Name)
    objGrid.colCampo.Add (DataPrevGrid.Name)
    objGrid.colCampo.Add (OBSGrid.Name)

    'Colunas do Grid
    iGrid_NumOV_Col = 1
    iGrid_Cliente_Col = 2
    iGrid_Filial_Col = 3
    iGrid_Projeto_Col = 4
    iGrid_Emissao_Col = 5
    iGrid_Valor_Col = 6
    iGrid_Status_Col = 7
    iGrid_DataProx_Col = 8
    iGrid_DataPrev_Col = 9

    objGrid.objGrid = GridOV

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 100 + 1

    objGrid.iLinhasVisiveis = 14

    'Largura da primeira coluna
    GridOV.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    Call Grid_Inicializa(objGrid)

    Inicializa_GridOV = SUCESSO

End Function

Private Sub GridOV_Click()

Dim iExecutaEntradaCelula As Integer
Dim colcolColecoes As New Collection

    Call Grid_Click(objGridOV, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridOV, iAlteradoGrid)
    End If
    
    colcolColecoes.Add gcolOVs
    colcolColecoes.Add gcolClientes
    colcolColecoes.Add gcolFiliais
    colcolColecoes.Add gcolEnderecos
    colcolColecoes.Add gcolProjetos
    
    Call Ordenacao_ClickGrid(objGridOV, , colcolColecoes)

End Sub

Private Sub GridOV_GotFocus()
    Call Grid_Recebe_Foco(objGridOV)
End Sub

Private Sub GridOV_EnterCell()
    Call Grid_Entrada_Celula(objGridOV, iAlteradoGrid)
End Sub

Private Sub GridOV_LeaveCell()
    Call Saida_Celula(objGridOV)
End Sub

Private Sub GridOV_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridOV, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridOV, iAlteradoGrid)
    End If

End Sub

Private Sub GridOV_RowColChange()
    Call Grid_RowColChange(objGridOV)
    
    If objGridOV.iLinhaAntiga <> objGridOV.objGrid.Row Then
        Call Trata_RowColChange
        objGridOV.iLinhaAntiga = objGridOV.objGrid.Row
    End If
End Sub

Private Sub GridOV_Scroll()
    Call Grid_Scroll(objGridOV)
End Sub

Private Sub GridOV_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridOV)
End Sub

Private Sub GridOV_LostFocus()
    Call Grid_Libera_Foco(objGridOV)
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 182274

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr
        
        Case 182274
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182275)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 182249

    'Limpa a Tela
    lErro = Limpa_Tela_OVAcomp
    If lErro <> SUCESSO Then gError 182250

    iAlterado = 0
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 182249, 182250

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182251)

    End Select

End Sub

Function Limpa_Tela_OVAcomp() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_OVAcomp
   
    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    iAlterado = 0
    
    Call Grid_Limpa(objGridOV)
    
    Set gobjOVAcompAnt = New ClassOVAcomp
    Set gcolOVs = New Collection
    Set gcolClientes = New Collection
    Set gcolFiliais = New Collection
    Set gcolEnderecos = New Collection
    Set gcolProjetos = New Collection
    
    SoOVEmPV.Value = vbUnchecked
    SoOVNaoFaturado = vbChecked
    SoOVNaoPerdido = vbChecked
    
    ApenasQualifProx.ListIndex = -1
    EntreQualifProxDe.ListIndex = -1
    EntreQualifProxAte.ListIndex = -1
    ApenasQualifPrev.ListIndex = -1
    EntreQualifPrevDe.ListIndex = -1
    EntreQualifPrevAte.ListIndex = -1
    ApenasQualifEmi.ListIndex = -1
    EntreQualifEmiDe.ListIndex = -1
    EntreQualifEmiAte.ListIndex = -1
    
    Call FrameD_Enabled(FrameD1, FRAMED_FAIXA)
    Call FrameD_Enabled(FrameD3, FRAMED_FAIXA)
    Call FrameD_Enabled(FrameD4, FRAMED_FAIXA)

    NumOV.Caption = ""
    Versao.Caption = ""
    Emissao.Caption = ""
    Valor.Caption = ""
    Cliente.Caption = ""
    Projeto.Caption = ""
    Contato.Caption = ""
    Telefone1.Caption = ""
    Telefone2.Caption = ""
    NumIntOV.Caption = ""
    ValorTotal.Caption = ""
    Responsavel.Caption = ""
    ProjetoDesc.Caption = ""
    ClienteNome.Caption = ""
    
    LabelNumCRM.Caption = "0 de 0"
    LblDataCRM.Caption = ""
    LblHistoricoCRM.Caption = ""
    giNumCRM = 0
    
    Call Ordenacao_Limpa(objGridOV)
    
    If iFrameAtual <> TAB_Selecao Then
        'Torna Frame atual invisível
        Frame1(TabStrip1.SelectedItem.Index).Visible = False
        'Torna Frame atual visível
        Frame1(TAB_Selecao).Visible = True
        TabStrip1.Tabs.Item(TAB_Selecao).Selected = True
        iFrameAtual = TAB_Selecao
    End If
    
    iAlterado = 0

    Limpa_Tela_OVAcomp = SUCESSO

    Exit Function

Erro_Limpa_Tela_OVAcomp:

    Limpa_Tela_OVAcomp = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182252)

    End Select

    Exit Function

End Function

Private Sub UpDownData_DownClick(objDataMask As MaskEdBox)

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    objDataMask.SetFocus

    If Len(objDataMask.ClipText) > 0 Then

        sData = objDataMask.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 182253

        objDataMask.Text = sData

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 182253

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182254)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick(objDataMask As MaskEdBox)

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_UpClick

    objDataMask.SetFocus

    If Len(Trim(objDataMask.ClipText)) > 0 Then

        sData = objDataMask.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 182255

        objDataMask.Text = sData

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 182255

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182256)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataPrevAte_DownClick()
    Call UpDownData_DownClick(DataPrevAte)
End Sub

Private Sub UpDownDataPrevAte_UpClick()
    Call UpDownData_UpClick(DataPrevAte)
End Sub

Private Sub UpDownDataPrevDe_DownClick()
    Call UpDownData_DownClick(DataPrevDe)
End Sub

Private Sub UpDownDataPrevDe_UpClick()
    Call UpDownData_UpClick(DataPrevDe)
End Sub

Private Sub UpDownDataProxAte_DownClick()
    Call UpDownData_DownClick(DataProxAte)
End Sub

Private Sub UpDownDataProxAte_UpClick()
    Call UpDownData_UpClick(DataProxAte)
End Sub

Private Sub UpDownDataProxDe_DownClick()
    Call UpDownData_DownClick(DataProxDe)
End Sub

Private Sub UpDownDataProxDe_UpClick()
    Call UpDownData_UpClick(DataProxDe)
End Sub

Private Sub UpDownDataEmiAte_DownClick()
    Call UpDownData_DownClick(DataEmiAte)
End Sub

Private Sub UpDownDataEmiAte_UpClick()
    Call UpDownData_UpClick(DataEmiAte)
End Sub

Private Sub UpDownDataEmiDe_DownClick()
    Call UpDownData_DownClick(DataEmiDe)
End Sub

Private Sub UpDownDataEmiDe_UpClick()
    Call UpDownData_UpClick(DataEmiDe)
End Sub

Private Sub FrameD_Enabled(objFrame As Object, ByVal iIndice As Integer)

Dim iIndiceAux As Integer

    For iIndiceAux = 0 To 2
        If iIndiceAux = iIndice Then
            objFrame(iIndiceAux).Enabled = True
        Else
            objFrame(iIndiceAux).Enabled = False
        End If
    Next

End Sub

Private Sub Limpa_FrameD1(ByVal iIndice As Integer)

    If iIndice <> FRAMED_FAIXA Then
        DataProxDe.PromptInclude = False
        DataProxDe.Text = ""
        DataProxDe.PromptInclude = True
    
        DataProxAte.PromptInclude = False
        DataProxAte.Text = ""
        DataProxAte.PromptInclude = True
    End If

    If iIndice <> FRAMED_APENAS Then
        ApenasQualifProx.ListIndex = -1
        
        ApenasDiasProx.PromptInclude = False
        ApenasDiasProx.Text = ""
        ApenasDiasProx.PromptInclude = True
    End If

    If iIndice <> FRAMED_ENTRE Then
    
        EntreDiasProxDe.PromptInclude = False
        EntreDiasProxDe.Text = ""
        EntreDiasProxDe.PromptInclude = True
        
        EntreQualifProxDe.ListIndex = -1
    
        EntreDiasProxAte.PromptInclude = False
        EntreDiasProxAte.Text = ""
        EntreDiasProxAte.PromptInclude = True
    
        EntreQualifProxAte.ListIndex = -1
    
    End If

End Sub

Private Sub Limpa_FrameD3(ByVal iIndice As Integer)

    If iIndice <> FRAMED_FAIXA Then
        DataPrevDe.PromptInclude = False
        DataPrevDe.Text = ""
        DataPrevDe.PromptInclude = True
    
        DataPrevAte.PromptInclude = False
        DataPrevAte.Text = ""
        DataPrevAte.PromptInclude = True
    End If

    If iIndice <> FRAMED_APENAS Then
        ApenasQualifPrev.ListIndex = -1
        
        ApenasDiasPrev.PromptInclude = False
        ApenasDiasPrev.Text = ""
        ApenasDiasPrev.PromptInclude = True
    End If

    If iIndice <> FRAMED_ENTRE Then
    
        EntreDiasPrevDe.PromptInclude = False
        EntreDiasPrevDe.Text = ""
        EntreDiasPrevDe.PromptInclude = True
        
        EntreQualifPrevDe.ListIndex = -1
    
        EntreDiasPrevAte.PromptInclude = False
        EntreDiasPrevAte.Text = ""
        EntreDiasPrevAte.PromptInclude = True
    
        EntreQualifPrevAte.ListIndex = -1
    
    End If

End Sub

Private Sub Limpa_FrameD4(ByVal iIndice As Integer)

    If iIndice <> FRAMED_FAIXA Then
        DataEmiDe.PromptInclude = False
        DataEmiDe.Text = ""
        DataEmiDe.PromptInclude = True
    
        DataEmiAte.PromptInclude = False
        DataEmiAte.Text = ""
        DataEmiAte.PromptInclude = True
    End If

    If iIndice <> FRAMED_APENAS Then
        ApenasQualifEmi.ListIndex = -1
        
        ApenasDiasEmi.PromptInclude = False
        ApenasDiasEmi.Text = ""
        ApenasDiasEmi.PromptInclude = True
    End If

    If iIndice <> FRAMED_ENTRE Then
    
        EntreDiasEmiDe.PromptInclude = False
        EntreDiasEmiDe.Text = ""
        EntreDiasEmiDe.PromptInclude = True
        
        EntreQualifEmiDe.ListIndex = -1
    
        EntreDiasEmiAte.PromptInclude = False
        EntreDiasEmiAte.Text = ""
        EntreDiasEmiAte.PromptInclude = True
    
        EntreQualifEmiAte.ListIndex = -1
    
    End If

End Sub

Private Sub ApenasProx_Click()
    Call FrameD_Enabled(FrameD1, FRAMED_APENAS)
    Call Limpa_FrameD1(FRAMED_APENAS)
End Sub

Private Sub EntreProx_Click()
    Call FrameD_Enabled(FrameD1, FRAMED_ENTRE)
    Call Limpa_FrameD1(FRAMED_ENTRE)
End Sub

Private Sub FaixaDataProx_Click()
    Call FrameD_Enabled(FrameD1, FRAMED_FAIXA)
    Call Limpa_FrameD1(FRAMED_FAIXA)
End Sub

Private Sub ApenasPrev_Click()
    Call FrameD_Enabled(FrameD3, FRAMED_APENAS)
    Call Limpa_FrameD3(FRAMED_APENAS)
End Sub

Private Sub EntrePrev_Click()
    Call FrameD_Enabled(FrameD3, FRAMED_ENTRE)
    Call Limpa_FrameD3(FRAMED_ENTRE)
End Sub

Private Sub FaixaDataPrev_Click()
    Call FrameD_Enabled(FrameD3, FRAMED_FAIXA)
    Call Limpa_FrameD3(FRAMED_FAIXA)
End Sub

Private Sub ApenasEmi_Click()
    Call FrameD_Enabled(FrameD4, FRAMED_APENAS)
    Call Limpa_FrameD4(FRAMED_APENAS)
End Sub

Private Sub EntreEmi_Click()
    Call FrameD_Enabled(FrameD4, FRAMED_ENTRE)
    Call Limpa_FrameD4(FRAMED_ENTRE)
End Sub

Private Sub FaixaDataEmi_Click()
    Call FrameD_Enabled(FrameD4, FRAMED_FAIXA)
    Call Limpa_FrameD4(FRAMED_FAIXA)
End Sub

Private Sub Data_Validate(objDataMask As MaskEdBox, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    'Verifica se Data está preenchida
    If Len(Trim(objDataMask.ClipText)) <> 0 Then

        'Critica a Data
        lErro = Data_Critica(objDataMask.Text)
        If lErro <> SUCESSO Then gError 182257
        
    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case 182257
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182258)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 182358
    
    Call Limpa_Relacionamento
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case gErr
    
        Case 182358
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182359)

    End Select
    
End Sub

Private Sub Limpa_Relacionamento()

'    DataPrev.PromptInclude = False
'    DataPrev.Text = ""
'    DataPrev.PromptInclude = True
'
'    DataProx.PromptInclude = False
'    DataProx.Text = ""
'    DataProx.PromptInclude = True
    
    Assunto.Text = ""
    
'    Status.ListIndex = iStatus_ListIndex_Padrao
    
    iAlterado = 0

End Sub

Private Sub DataProxDe_Validate(Cancel As Boolean)
    Call Data_Validate(DataProxDe, Cancel)
End Sub

Private Sub DataProxAte_Validate(Cancel As Boolean)
    Call Data_Validate(DataProxAte, Cancel)
End Sub

Private Sub DataPrevDe_Validate(Cancel As Boolean)
    Call Data_Validate(DataPrevDe, Cancel)
End Sub

Private Sub DataPrevAte_Validate(Cancel As Boolean)
    Call Data_Validate(DataPrevAte, Cancel)
End Sub

Private Sub DataEmiDe_Validate(Cancel As Boolean)
    Call Data_Validate(DataEmiDe, Cancel)
End Sub

Private Sub DataEmiAte_Validate(Cancel As Boolean)
    Call Data_Validate(DataEmiAte, Cancel)
End Sub

Private Sub DataPrev_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataProx_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataPrev_Validate(Cancel As Boolean)
    Call Data_Validate(DataPrev, Cancel)
End Sub

Private Sub DataProx_Validate(Cancel As Boolean)
    Call Data_Validate(DataProx, Cancel)
End Sub

Private Sub UpDownDataPrev_DownClick()
    Call UpDownData_DownClick(DataPrev)
End Sub

Private Sub UpDownDataPrev_UpClick()
    Call UpDownData_UpClick(DataPrev)
End Sub

Private Sub UpDownDataProx_DownClick()
    Call UpDownData_DownClick(DataProx)
End Sub

Private Sub UpDownDataProx_UpClick()
    Call UpDownData_UpClick(DataProx)
End Sub

Private Sub Assunto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub BotaoRelac_Click()

Dim objRelacionamentoCli As New ClassRelacClientes
Dim colSelecao As New Collection

    If GridOV.Row <> 0 Then

        colSelecao.Add RELACCLI_TIPODOC_OV
        colSelecao.Add StrParaLong(NumIntOV.Caption)
        
        Call Chama_Tela("RelacionamentoClientes_Lista", colSelecao, objRelacionamentoCli, Nothing, "TipoDoc = ? AND NumIntDocOrigem = ? ")

    End If

End Sub

Public Sub BotaoDocOriginal_Click()

Dim lErro As Long
Dim objOV As New ClassOrcamentoVenda

On Error GoTo Erro_BotaoDocOriginal_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridOV.Row = 0 Then gError 182282
        
    'Se foi selecionada uma linha que está preenchida
    If GridOV.Row <= objGridOV.iLinhasExistentes Then
        
        objOV.lNumIntDoc = StrParaLong(NumIntOV.Caption)
        objOV.lCodigo = StrParaLong(NumOV.Caption)
        objOV.iFilialEmpresa = giFilialEmpresa
        
        Call Chama_Tela("OrcamentoVenda", objOV)
    
    End If
        
    Exit Sub
    
Erro_BotaoDocOriginal_Click:

    Select Case gErr
    
        Case 182282
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182285)

    End Select

    Exit Sub
  
End Sub

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objRelacCli As New ClassRelacClientes
Dim objOV As ClassOrcamentoVenda
Dim sStatus As String

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    If StrParaLong(NumIntOV.Caption) = 0 Then gError 182342
    
    sStatus = Status.Text
    
    lErro = Move_Tela_Memoria(objRelacCli)
    If lErro <> SUCESSO Then gError 182278
    
    lErro = CF("RelacionamentoClientes_Grava", objRelacCli, True, gsUsuario)
    If lErro <> SUCESSO Then gError 182283
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_GRAVACAO_COM_SUCESSO")
    
    iLinha = 0
    
    For Each objOV In gcolOVs
        iLinha = iLinha + 1
        If objOV.lNumIntDoc = StrParaLong(NumIntOV.Caption) Then
            Exit For
        End If
    Next
    
    If StrParaDate(DataPrev.Text) <> DATA_NULA Then
        GridOV.TextMatrix(iLinha, iGrid_DataPrev_Col) = Format(DataPrev.Text, "dd/mm/yyyy")
    Else
        GridOV.TextMatrix(iLinha, iGrid_DataPrev_Col) = ""
    End If
    objOV.dtDataPrevReceb = objRelacCli.dtDataPrevReceb
    
    If StrParaDate(DataProx.Text) <> DATA_NULA Then
        GridOV.TextMatrix(iLinha, iGrid_DataProx_Col) = Format(DataProx.Text, "dd/mm/yyyy")
    Else
        GridOV.TextMatrix(iLinha, iGrid_DataProx_Col) = ""
    End If
    objOV.dtDataProxCobr = objRelacCli.dtDataProxCobr
    
    GridOV.TextMatrix(iLinha, iGrid_Status_Col) = sStatus
    objOV.lStatus = objRelacCli.lStatusTipoDoc
    
    lErro = Trata_HistoricoCRM(objOV)
    If lErro <> SUCESSO Then gError 182283
        
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 182278, 182327, 182283
        
        Case 182342
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_GRID_NAO_SELECIONADA", gErr)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182259)
        
    End Select
    
End Function

Private Function Carrega_Status(ByVal objComboBox As ComboBox) As Long
'Carrega a combo de Tipo

Dim lErro As Long

On Error GoTo Erro_Carrega_Status

    'carregar tipos de desconto
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_STATUSOV, objComboBox)
    If lErro <> SUCESSO Then gError 141371

    objComboBox.AddItem ""
    objComboBox.ItemData(objComboBox.NewIndex) = 0
    
    iStatus_ListIndex_Padrao = objComboBox.ListIndex

    Carrega_Status = SUCESSO

    Exit Function

Erro_Carrega_Status:

    Carrega_Status = gErr

    Select Case gErr
    
        Case 141371

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157851)

    End Select

    Exit Function

End Function

Private Sub Status_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub LabelVendedor_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection
   
    If Len(Trim(Vendedor.Text)) > 0 Then
        'Preenche com o Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(Vendedor.Text)
    End If
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1
    
    Vendedor.Text = CStr(objVendedor.iCodigo)
    Call Vendedor_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Vendedor_Validate

    If Len(Trim(Vendedor.Text)) > 0 Then
   
        'Tenta ler o vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(Vendedor, objVendedor, 0)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167058)

    End Select

End Sub

Public Sub DataPrevDe_GotFocus()
     Call MaskEdBox_TrataGotFocus(DataPrevDe, iAlterado)
End Sub

Public Sub DataPrevAte_GotFocus()
     Call MaskEdBox_TrataGotFocus(DataPrevAte, iAlterado)
End Sub

Public Sub DataProxDe_GotFocus()
     Call MaskEdBox_TrataGotFocus(DataProxDe, iAlterado)
End Sub

Public Sub DataProxAte_GotFocus()
     Call MaskEdBox_TrataGotFocus(DataProxAte, iAlterado)
End Sub

Public Sub DataEmiDe_GotFocus()
     Call MaskEdBox_TrataGotFocus(DataEmiDe, iAlterado)
End Sub

Public Sub DataEmiAte_GotFocus()
     Call MaskEdBox_TrataGotFocus(DataEmiAte, iAlterado)
End Sub

Private Sub ApenasDiasEmi_GotFocus()
     Call MaskEdBox_TrataGotFocus(ApenasDiasEmi, iAlterado)
End Sub

Private Sub EntreDiasEmiDe_GotFocus()
     Call MaskEdBox_TrataGotFocus(EntreDiasEmiDe, iAlterado)
End Sub

Private Sub EntreDiasEmiAte_GotFocus()
     Call MaskEdBox_TrataGotFocus(EntreDiasEmiAte, iAlterado)
End Sub

Private Sub ApenasDiasProx_GotFocus()
     Call MaskEdBox_TrataGotFocus(ApenasDiasProx, iAlterado)
End Sub

Private Sub EntreDiasProxDe_GotFocus()
     Call MaskEdBox_TrataGotFocus(EntreDiasProxDe, iAlterado)
End Sub

Private Sub EntreDiasProxAte_GotFocus()
     Call MaskEdBox_TrataGotFocus(EntreDiasProxAte, iAlterado)
End Sub

Private Sub ApenasDiasPrev_GotFocus()
     Call MaskEdBox_TrataGotFocus(ApenasDiasPrev, iAlterado)
End Sub

Private Sub EntreDiasPrevDe_GotFocus()
     Call MaskEdBox_TrataGotFocus(EntreDiasPrevDe, iAlterado)
End Sub

Private Sub EntreDiasPrevAte_GotFocus()
     Call MaskEdBox_TrataGotFocus(EntreDiasPrevAte, iAlterado)
End Sub

Function Trata_HistoricoCRM(ByVal objOV As ClassOrcamentoVenda) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_HistoricoCRM

    Set gcolCRM = New Collection

    lErro = CF("RelacCli_Le_TipoDoc2", RELACCLI_TIPODOC_OV, StrParaLong(NumIntOV.Caption), gcolCRM)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If gcolCRM.Count = 0 Then
        giNumCRM = 0
    Else
        giNumCRM = 1
    End If
    Call Traz_HistoricoCRM
    
    Trata_HistoricoCRM = SUCESSO

    Exit Function

Erro_Trata_HistoricoCRM:

    Trata_HistoricoCRM = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182281)

    End Select

    Exit Function

End Function

Function Traz_HistoricoCRM() As Long

Dim lErro As Long
Dim objRelacCli As ClassRelacClientes

On Error GoTo Erro_Traz_HistoricoCRM

    If giNumCRM > 0 And giNumCRM <= gcolCRM.Count Then
        Set objRelacCli = gcolCRM(giNumCRM)

        LabelNumCRM.Caption = CStr(giNumCRM) & " de " & CStr(gcolCRM.Count)
        LblCodCRM.Caption = CStr(objRelacCli.lCodigo)
        LblDataCRM.Caption = Format(objRelacCli.dtData, "dd/mm/yyyy")
        If objRelacCli.dtDataPrevReceb <> DATA_NULA Then
            LblDataFCRM.Caption = Format(objRelacCli.dtDataPrevReceb, "dd/mm/yyyy")
        Else
            LblDataFCRM.Caption = ""
        End If
        If objRelacCli.dtDataProxCobr <> DATA_NULA Then
            LblDataCCRM.Caption = Format(objRelacCli.dtDataProxCobr, "dd/mm/yyyy")
        Else
            LblDataCCRM.Caption = ""
        End If
        LblHistoricoCRM.Caption = objRelacCli.sAssunto1 & objRelacCli.sAssunto2
    Else
        LabelNumCRM.Caption = "0 de 0"
        LblCodCRM.Caption = ""
        LblDataCRM.Caption = ""
        LblDataFCRM.Caption = ""
        LblDataCCRM.Caption = ""
        LblHistoricoCRM.Caption = ""
    End If
    
    Traz_HistoricoCRM = SUCESSO

    Exit Function

Erro_Traz_HistoricoCRM:

    Traz_HistoricoCRM = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182281)

    End Select

    Exit Function

End Function

Private Sub BtnCRMAnterior_Click()
    If gcolCRM.Count > 0 Then
        If giNumCRM > 1 Then
            giNumCRM = giNumCRM - 1
            Call Traz_HistoricoCRM
        End If
    End If
End Sub

Private Sub BtnCRMPrimeiro_Click()
    If gcolCRM.Count > 0 Then
        If giNumCRM > 1 Then
            giNumCRM = 1
            Call Traz_HistoricoCRM
        End If
    End If
End Sub

Private Sub BtnCRMProximo_Click()
    If gcolCRM.Count > 0 Then
        If giNumCRM < gcolCRM.Count Then
            giNumCRM = giNumCRM + 1
            Call Traz_HistoricoCRM
        End If
    End If
End Sub

Private Sub BtnCRMUltimo_Click()
    If gcolCRM.Count > 0 Then
        If giNumCRM <> gcolCRM.Count Then
            giNumCRM = gcolCRM.Count
            Call Traz_HistoricoCRM
        End If
    End If
End Sub

Private Sub BotaoImportarCRM_Click()

Dim lErro As Long
'Dim colTabelaPrecoItem As New Collection
Dim objRelacCli As New ClassRelacClientes

On Error GoTo Erro_BotaoImportarCRM_Click

'    GL_objMDIForm.MousePointer = vbHourglass
'
    
    Call Chama_Tela_Modal("ImportarDadosArq", objRelacCli)
    
    Call Trata_TabClick(True)
    
    Exit Sub

Erro_BotaoImportarCRM_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 131015
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", gErr)
        
        Case 131016
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_VIGENCIA_NAO_PREENCHIDA", gErr)
        
        Case 131017
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_VIGENCIA_MENOR_DATA_ATUAL", gErr, Date)

        Case 131018 To 131019

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158286)

    End Select

    Exit Sub

End Sub
