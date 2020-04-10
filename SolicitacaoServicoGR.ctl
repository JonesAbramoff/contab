VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl SolicitacaoServico 
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   9510
   Begin VB.CommandButton BotaoProxNum 
      Height          =   315
      Left            =   2805
      Picture         =   "SolicitacaoServicoGR.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Numeração Automática"
      Top             =   240
      Width           =   345
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6795
      ScaleHeight     =   495
      ScaleWidth      =   2565
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   75
      Width           =   2625
      Begin VB.CommandButton BotaoImprimir 
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
         Left            =   60
         Picture         =   "SolicitacaoServicoGR.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   555
         Picture         =   "SolicitacaoServicoGR.ctx":01EC
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   1065
         Picture         =   "SolicitacaoServicoGR.ctx":0346
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1590
         Picture         =   "SolicitacaoServicoGR.ctx":04D0
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2070
         Picture         =   "SolicitacaoServicoGR.ctx":0A02
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox MaskNumero 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   225
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "########"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MaskCodTabPreco 
      Height          =   300
      Left            =   5460
      TabIndex        =   2
      Top             =   225
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   8
      Mask            =   "########"
      PromptChar      =   " "
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4635
      Index           =   2
      Left            =   105
      TabIndex        =   59
      Top             =   1005
      Visible         =   0   'False
      Width           =   9300
      Begin VB.Frame Frame3 
         Caption         =   "Navio"
         Height          =   1215
         Left            =   390
         TabIndex        =   69
         Top             =   195
         Width           =   8685
         Begin MSMask.MaskEdBox MaskIdentProgNavio 
            Height          =   315
            Left            =   1380
            TabIndex        =   110
            Top             =   315
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   10
            Mask            =   "##########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskBooking 
            Height          =   315
            Left            =   6750
            TabIndex        =   111
            Top             =   315
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskPorto 
            Height          =   315
            Left            =   2430
            TabIndex        =   116
            Top             =   735
            Width           =   5850
            _ExtentX        =   10319
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   40
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Porto de Desembarque:"
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
            Index           =   7
            Left            =   330
            TabIndex        =   117
            Top             =   795
            Width           =   2010
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Booking:"
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
            Index           =   6
            Left            =   5955
            TabIndex        =   115
            Top             =   375
            Width           =   765
         End
         Begin VB.Label LabelNavio 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000007&
            Height          =   315
            Left            =   3420
            TabIndex        =   114
            Top             =   315
            Width           =   2205
         End
         Begin VB.Label LabelIdentProgNavio 
            AutoSize        =   -1  'True
            Caption         =   "Id. Viagem:"
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
            Left            =   345
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   113
            Top             =   375
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Navio:"
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
            Left            =   2790
            TabIndex        =   112
            Top             =   375
            Width           =   570
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Previsão Término do Serviço"
         Height          =   780
         Left            =   390
         TabIndex        =   62
         Top             =   2415
         Width           =   8685
         Begin MSMask.MaskEdBox MaskDataFim 
            Height          =   300
            Left            =   3345
            TabIndex        =   16
            Top             =   270
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskHoraFim 
            Height          =   300
            Left            =   6270
            TabIndex        =   17
            Top             =   285
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDown4 
            Height          =   330
            Left            =   4410
            TabIndex        =   101
            Top             =   285
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
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
            Index           =   11
            Left            =   5700
            TabIndex        =   66
            Top             =   330
            Width           =   480
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   13
            Left            =   2745
            TabIndex        =   65
            Top             =   330
            Width           =   480
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Previsão Início do Serviço"
         Height          =   780
         Left            =   390
         TabIndex        =   61
         Top             =   1545
         Width           =   8685
         Begin MSMask.MaskEdBox MaskDataInicio 
            Height          =   300
            Left            =   3300
            TabIndex        =   14
            Top             =   270
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskHoraInicio 
            Height          =   300
            Left            =   6240
            TabIndex        =   15
            Top             =   270
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "hh:mm:ss"
            Mask            =   "##:##:##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDown3 
            Height          =   330
            Left            =   4380
            TabIndex        =   100
            Top             =   255
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
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
            Index           =   10
            Left            =   5685
            TabIndex        =   64
            Top             =   315
            Width           =   480
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   12
            Left            =   2730
            TabIndex        =   63
            Top             =   315
            Width           =   480
         End
      End
      Begin VB.TextBox TextObs 
         Height          =   1290
         Left            =   1545
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   3330
         Width           =   7530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Observação:"
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
         Index           =   9
         Left            =   390
         TabIndex        =   60
         Top             =   3360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4635
      Index           =   3
      Left            =   105
      TabIndex        =   67
      Top             =   990
      Visible         =   0   'False
      Width           =   9315
      Begin VB.Frame Frame6 
         Caption         =   " Origem "
         Height          =   4476
         Left            =   108
         TabIndex        =   70
         Top             =   60
         Width           =   9084
         Begin VB.TextBox TextEndereco 
            Height          =   315
            Index           =   0
            Left            =   1500
            MaxLength       =   40
            MultiLine       =   -1  'True
            TabIndex        =   21
            Top             =   468
            Width           =   6345
         End
         Begin MSMask.MaskEdBox MaskBairro 
            Height          =   312
            Index           =   0
            Left            =   1500
            TabIndex        =   22
            Top             =   1260
            Width           =   1572
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskCEP 
            Height          =   312
            Index           =   0
            Left            =   6900
            TabIndex        =   23
            Top             =   1272
            Width           =   924
            _ExtentX        =   1640
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#####-###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskTelefone1 
            Height          =   312
            Index           =   0
            Left            =   1500
            TabIndex        =   24
            Top             =   2964
            Width           =   1236
            _ExtentX        =   2170
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskTelefone2 
            Height          =   315
            Index           =   0
            Left            =   1530
            TabIndex        =   26
            Top             =   3765
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskEmail 
            Height          =   312
            Index           =   0
            Left            =   4260
            TabIndex        =   27
            Top             =   3768
            Width           =   1572
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskContato 
            Height          =   315
            Index           =   0
            Left            =   6990
            TabIndex        =   28
            Top             =   3780
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskFax 
            Height          =   312
            Index           =   0
            Left            =   4260
            TabIndex        =   25
            Top             =   2964
            Width           =   1236
            _ExtentX        =   2170
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            PromptChar      =   " "
         End
         Begin VB.Label LabelCidade 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000007&
            Height          =   330
            Index           =   0
            Left            =   4305
            TabIndex        =   105
            Top             =   1245
            Width           =   1335
         End
         Begin VB.Label LabelEstado 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000007&
            Height          =   285
            Index           =   0
            Left            =   1515
            TabIndex        =   104
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label LabelPais 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000007&
            Height          =   300
            Index           =   0
            Left            =   4305
            TabIndex        =   103
            Top             =   2145
            Width           =   735
         End
         Begin VB.Label Label46 
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
            Height          =   192
            Index           =   0
            Left            =   528
            TabIndex        =   81
            Top             =   528
            Width           =   912
         End
         Begin VB.Label Label57 
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
            Height          =   192
            Index           =   0
            Left            =   3540
            TabIndex        =   80
            Top             =   1320
            Width           =   672
         End
         Begin VB.Label Label63 
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
            Height          =   192
            Index           =   0
            Left            =   756
            TabIndex        =   79
            Top             =   2208
            Width           =   672
         End
         Begin VB.Label Label64 
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
            Height          =   192
            Index           =   0
            Left            =   852
            TabIndex        =   78
            Top             =   1320
            Width           =   588
         End
         Begin VB.Label Label65 
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
            Height          =   192
            Index           =   0
            Left            =   432
            TabIndex        =   77
            Top             =   3024
            Width           =   1008
         End
         Begin VB.Label Label66 
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
            Height          =   192
            Index           =   0
            Left            =   432
            TabIndex        =   76
            Top             =   3828
            Width           =   1008
         End
         Begin VB.Label Label67 
            AutoSize        =   -1  'True
            Caption         =   "Internet:"
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
            Index           =   0
            Left            =   3456
            TabIndex        =   75
            Top             =   3828
            Width           =   768
         End
         Begin VB.Label Label68 
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
            Height          =   192
            Index           =   0
            Left            =   3816
            TabIndex        =   74
            Top             =   3024
            Width           =   408
         End
         Begin VB.Label Label69 
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
            Height          =   192
            Index           =   0
            Left            =   6396
            TabIndex        =   73
            Top             =   1320
            Width           =   468
         End
         Begin VB.Label Label70 
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
            Height          =   192
            Index           =   0
            Left            =   6108
            TabIndex        =   72
            Top             =   3828
            Width           =   756
         End
         Begin VB.Label PaisLabel 
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
            Height          =   192
            Index           =   0
            Left            =   3720
            TabIndex        =   71
            Top             =   2208
            Width           =   492
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4635
      Index           =   1
      Left            =   150
      TabIndex        =   45
      Top             =   990
      Width           =   9300
      Begin VB.Frame Frame4 
         Caption         =   "Cliente"
         Height          =   1095
         Left            =   90
         TabIndex        =   46
         Top             =   120
         Width           =   9165
         Begin MSMask.MaskEdBox MaskDocRefer 
            Height          =   315
            Left            =   1890
            TabIndex        =   3
            Top             =   705
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   300
            Left            =   7200
            TabIndex        =   19
            Top             =   690
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox MaskDataPedido 
            Height          =   315
            Left            =   6240
            TabIndex        =   4
            Top             =   690
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelTextCliente 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   1965
            TabIndex        =   108
            Top             =   285
            Width           =   1620
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
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
            Left            =   1980
            TabIndex        =   102
            Top             =   285
            Width           =   75
         End
         Begin VB.Label LabelCliente 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   1185
            TabIndex        =   49
            Top             =   285
            Width           =   660
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Data Pedido:"
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
            Left            =   5025
            TabIndex        =   48
            Top             =   750
            Width           =   1125
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Número Referência:"
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
            TabIndex        =   47
            Top             =   750
            Width           =   1710
         End
      End
      Begin VB.TextBox TextUM 
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   7680
         MaxLength       =   20
         TabIndex        =   10
         Top             =   3420
         Width           =   915
      End
      Begin VB.CommandButton BotaoServicos 
         Caption         =   "Serviços"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7245
         TabIndex        =   99
         Top             =   4275
         Width           =   1965
      End
      Begin VB.TextBox TextDescProduto 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   1410
         MaxLength       =   50
         TabIndex        =   98
         Top             =   1905
         Width           =   5715
      End
      Begin MSMask.MaskEdBox MaskQuantidade 
         Height          =   225
         Left            =   6900
         TabIndex        =   96
         Top             =   1665
         Width           =   1155
         _ExtentX        =   2037
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
      Begin MSMask.MaskEdBox MaskProduto 
         Height          =   225
         Left            =   150
         TabIndex        =   97
         Top             =   1575
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.TextBox TextMaterial 
         Height          =   315
         Left            =   1380
         MaxLength       =   50
         TabIndex        =   8
         Top             =   3420
         Width           =   1605
      End
      Begin VB.ComboBox ComboTipoEmbalagem 
         Height          =   315
         ItemData        =   "SolicitacaoServicoGR.ctx":0B80
         Left            =   4710
         List            =   "SolicitacaoServicoGR.ctx":0B82
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3870
         Width           =   1560
      End
      Begin VB.ComboBox ComboTipoContainer 
         Height          =   315
         ItemData        =   "SolicitacaoServicoGR.ctx":0B84
         Left            =   7680
         List            =   "SolicitacaoServicoGR.ctx":0B86
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3870
         Width           =   1560
      End
      Begin VB.TextBox TextDespachante 
         Height          =   315
         Left            =   7680
         TabIndex        =   7
         Top             =   2925
         Width           =   1545
      End
      Begin VB.ComboBox ComboTipoOperacao 
         Height          =   315
         ItemData        =   "SolicitacaoServicoGR.ctx":0B88
         Left            =   4725
         List            =   "SolicitacaoServicoGR.ctx":0B98
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2925
         Width           =   1560
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   330
         Left            =   2340
         TabIndex        =   20
         Top             =   2910
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox MaskDataEmissao 
         Height          =   315
         Left            =   1380
         TabIndex        =   5
         Top             =   2925
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskQuantMaterial 
         Height          =   315
         Left            =   4710
         TabIndex        =   9
         Top             =   3420
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskValorMerc 
         Height          =   315
         Left            =   1380
         TabIndex        =   11
         Top             =   3900
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridServicos 
         Height          =   1455
         Left            =   120
         TabIndex        =   95
         Top             =   1305
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   2566
         _Version        =   393216
         Rows            =   100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   58
         Top             =   2985
         Width           =   1230
      End
      Begin VB.Label LabelMaterial 
         AutoSize        =   -1  'True
         Caption         =   "Material:"
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
         Left            =   600
         TabIndex        =   57
         Top             =   3480
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Quant. Material:"
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
         Left            =   3270
         TabIndex        =   56
         Top             =   3480
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "U.M.:"
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
         Index           =   0
         Left            =   7170
         TabIndex        =   55
         Top             =   3480
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor Mercad.:"
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
         Left            =   75
         TabIndex        =   54
         Top             =   3930
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Embalagem:"
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
         Index           =   3
         Left            =   3180
         TabIndex        =   53
         Top             =   3930
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Container:"
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
         Index           =   4
         Left            =   6330
         TabIndex        =   52
         Top             =   3930
         Width           =   1320
      End
      Begin VB.Label LabelDespachante 
         AutoSize        =   -1  'True
         Caption         =   "Despachante:"
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
         Left            =   6450
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   51
         Top             =   2985
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Operação:"
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
         Index           =   8
         Left            =   3045
         TabIndex        =   50
         Top             =   2985
         Width           =   1605
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4635
      Index           =   4
      Left            =   105
      TabIndex        =   68
      Top             =   990
      Visible         =   0   'False
      Width           =   9300
      Begin VB.Frame Frame7 
         Caption         =   " Destino "
         Height          =   4476
         Left            =   108
         TabIndex        =   82
         Top             =   90
         Width           =   9120
         Begin VB.TextBox TextEndereco 
            Height          =   315
            Index           =   1
            Left            =   1530
            MaxLength       =   40
            MultiLine       =   -1  'True
            TabIndex        =   29
            Top             =   468
            Width           =   6345
         End
         Begin MSMask.MaskEdBox MaskBairro 
            Height          =   330
            Index           =   1
            Left            =   1500
            TabIndex        =   30
            Top             =   1290
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskCEP 
            Height          =   312
            Index           =   1
            Left            =   6900
            TabIndex        =   31
            Top             =   1272
            Width           =   924
            _ExtentX        =   1640
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#####-###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskTelefone1 
            Height          =   312
            Index           =   1
            Left            =   1500
            TabIndex        =   32
            Top             =   2964
            Width           =   1236
            _ExtentX        =   2170
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskTelefone2 
            Height          =   312
            Index           =   1
            Left            =   1500
            TabIndex        =   34
            Top             =   3768
            Width           =   1236
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskEmail 
            Height          =   312
            Index           =   1
            Left            =   4260
            TabIndex        =   35
            Top             =   3768
            Width           =   1572
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskContato 
            Height          =   312
            Index           =   1
            Left            =   6900
            TabIndex        =   36
            Top             =   3768
            Width           =   1776
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox MaskFax 
            Height          =   315
            Index           =   1
            Left            =   4290
            TabIndex        =   33
            Top             =   2970
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            PromptChar      =   " "
         End
         Begin VB.Label LabelCidade 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000007&
            Height          =   330
            Index           =   1
            Left            =   4260
            TabIndex        =   109
            Top             =   1260
            Width           =   1335
         End
         Begin VB.Label LabelEstado 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000007&
            Height          =   300
            Index           =   1
            Left            =   1470
            TabIndex        =   107
            Top             =   2175
            Width           =   375
         End
         Begin VB.Label LabelPais 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000007&
            Height          =   300
            Index           =   1
            Left            =   4275
            TabIndex        =   106
            Top             =   2145
            Width           =   735
         End
         Begin VB.Label PaisLabel 
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
            Height          =   192
            Index           =   1
            Left            =   3720
            TabIndex        =   93
            Top             =   2208
            Width           =   492
         End
         Begin VB.Label Label70 
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
            Height          =   192
            Index           =   1
            Left            =   6108
            TabIndex        =   92
            Top             =   3828
            Width           =   756
         End
         Begin VB.Label Label69 
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
            Height          =   192
            Index           =   1
            Left            =   6384
            TabIndex        =   91
            Top             =   1320
            Width           =   468
         End
         Begin VB.Label Label68 
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
            Height          =   192
            Index           =   1
            Left            =   3804
            TabIndex        =   90
            Top             =   3024
            Width           =   408
         End
         Begin VB.Label Label67 
            AutoSize        =   -1  'True
            Caption         =   "Internet:"
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
            Index           =   1
            Left            =   3444
            TabIndex        =   89
            Top             =   3828
            Width           =   768
         End
         Begin VB.Label Label66 
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
            Height          =   192
            Index           =   1
            Left            =   432
            TabIndex        =   88
            Top             =   3828
            Width           =   1008
         End
         Begin VB.Label Label65 
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
            Height          =   192
            Index           =   1
            Left            =   432
            TabIndex        =   87
            Top             =   3024
            Width           =   1008
         End
         Begin VB.Label Label64 
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
            Height          =   192
            Index           =   1
            Left            =   852
            TabIndex        =   86
            Top             =   1320
            Width           =   588
         End
         Begin VB.Label Label63 
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
            Height          =   192
            Index           =   1
            Left            =   768
            TabIndex        =   85
            Top             =   2208
            Width           =   672
         End
         Begin VB.Label Label57 
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
            Height          =   192
            Index           =   1
            Left            =   3540
            TabIndex        =   84
            Top             =   1320
            Width           =   672
         End
         Begin VB.Label Label46 
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
            Height          =   192
            Index           =   1
            Left            =   528
            TabIndex        =   83
            Top             =   528
            Width           =   912
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5040
      Left            =   75
      TabIndex        =   44
      Top             =   675
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   8890
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Complemento"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Origem"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Destino"
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
   Begin VB.Label LabelCodTabPreco 
      AutoSize        =   -1  'True
      Caption         =   "Código Tab. Preço:"
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
      Left            =   3720
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   94
      Top             =   255
      Width           =   1665
   End
   Begin VB.Label LabelNumero 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1245
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   43
      Top             =   270
      Width           =   720
   End
End
Attribute VB_Name = "SolicitacaoServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Dim iFrameAtual As Integer
Dim iAlterado As Integer
Dim giTabPrecoAlterado As Integer

'Constantes
Private Const TAB_DadosPrincipais = 1
Private Const TAB_Complemento = 2
Private Const TAB_Origem = 3
Private Const TAB_Destino = 4
Private Const TIPO_ORIGEM = 0
Private Const TIPO_DESTINO = 1
'Private Const PAIS_BRASIL = "Brasil"
'Private Const PAIS_CODIGO_BRASIL = 0

Const NUM_MAX_SERVICOS = 100

'Itens Grid
Dim objGridServicoss As AdmGrid
Dim iGrid_Produto_Col As Integer
Dim iGrid_DescProduto_Col As Integer
Dim iGrid_Quantidade_Col As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

'Eventos da tela
Private WithEvents objEventoSolicitacao As AdmEvento
Attribute objEventoSolicitacao.VB_VarHelpID = -1
Private WithEvents objEventoTabPreco As AdmEvento
Attribute objEventoTabPreco.VB_VarHelpID = -1
Private WithEvents objEventoDespachante As AdmEvento
Attribute objEventoDespachante.VB_VarHelpID = -1
Private WithEvents objEventoProgNavio As AdmEvento
Attribute objEventoProgNavio.VB_VarHelpID = -1
Private WithEvents objEventoServico As AdmEvento
Attribute objEventoServico.VB_VarHelpID = -1

Private Sub LabelIdentProgNavio_Click()

Dim objProgNavio As New ClassProgNavio
Dim colSelecao As Collection
Dim lErro As Long

On Error GoTo Erro_LabelIdentProgNavio_Click

    'Preenche com o Código do Prognavio da tela
    If Len(Trim(MaskIdentProgNavio.Text)) > 0 Then objProgNavio.lCodigo = StrParaLong(MaskIdentProgNavio.Text)

    'Chama Tela PrognavioLista
    Call Chama_Tela("ProgNavioLista", colSelecao, objProgNavio, objEventoProgNavio)
    
    Exit Sub
    
Erro_LabelIdentProgNavio_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Sub

End Sub

Private Sub LabelNumero_Click()

Dim objSolicitacaoServico As New ClassSolicitacaoServico
Dim colSelecao As Collection

On Error GoTo Erro_LabelNumero_Click

    'Preenche com o número da tela
    If Len(Trim(MaskNumero.Text)) > 0 Then objSolicitacaoServico.lNumero = StrParaLong(MaskNumero.Text)
    
    'Chama Tela SolicitacaoLista
    Call Chama_Tela("SolicitacaoLista", colSelecao, objSolicitacaoServico, objEventoSolicitacao)
    
    Exit Sub
    
Erro_LabelNumero_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Sub

End Sub

Private Sub LabelCodTabPreco_Click()

Dim objTabPreco As New ClassTabPreco
Dim colSelecao As New Collection
Dim lErro As Long
Dim sSelecaoSQL As String

On Error GoTo Erro_LabelCodTabPreco_Click

    'Se a data de emissão não está preenchida --> erro
    If Len(Trim(MaskDataEmissao.ClipText)) = 0 Then gError 98300
    
    'Preenche com o Código da Tabela de Preço da tela
    If Len(Trim(MaskCodTabPreco.Text)) > 0 Then objTabPreco.lCodigo = StrParaLong(MaskCodTabPreco.Text)
        
    colSelecao.Add StrParaDate(MaskDataEmissao.Text)
    
    'Tratamento de browse dinâmico.
    sSelecaoSQL = "DataVigencia<=?"
    
    'Passagem da data no último parâmetro do chama_tela
    'Chama Tela tabPrecoLista
    Call Chama_Tela("TabPrecoLista", colSelecao, objTabPreco, objEventoTabPreco, sSelecaoSQL)

    Exit Sub
    
Erro_LabelCodTabPreco_Click:

    Select Case gErr
        
        Case 98300
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Sub

End Sub

Private Sub LabelDespachante_Click()

Dim objDespachante As New ClassDespachante
Dim colSelecao As Collection

On Error GoTo Erro_LabelDespachante_Click

    'Preenche com o Nome Reduzido do Despachante da tela
    If Len(Trim(TextDespachante.Text)) > 0 Then objDespachante.sNomeReduzido = TextDespachante.Text

    'Chama Tela DespachanteLista
    Call Chama_Tela("DespachanteLista", colSelecao, objDespachante, objEventoDespachante)
    
    Exit Sub
    
Erro_LabelDespachante_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoProgNavio_evSelecao(obj1 As Object)

Dim objProgNavio As ClassProgNavio
Dim lErro As Long

On Error GoTo Erro_objEventoProgNavio_evSelecao

    Set objProgNavio = obj1

    'Coloca o Código na tela
    MaskIdentProgNavio.Text = objProgNavio.lCodigo
    LabelNavio.Caption = objProgNavio.sNavio
    
    Me.Show

    Exit Sub
    
Erro_objEventoProgNavio_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoDespachante_evSelecao(obj1 As Object)

Dim objDespachante As ClassDespachante
Dim lErro As Long

On Error GoTo Erro_objEventoDespachante_evSelecao

    Set objDespachante = obj1

    'Coloca o Nome Reduzido do Despachante na tela
    TextDespachante.Text = objDespachante.sNomeReduzido
    
    Me.Show

    Exit Sub
    
Erro_objEventoDespachante_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoSolicitacao_evSelecao(obj1 As Object)

Dim objSolicitacaoServico As ClassSolicitacaoServico
Dim lErro As Long

On Error GoTo Erro_objEventoSolicitacao_evSelecao

    Set objSolicitacaoServico = obj1
    
    'Preenche com a Filial que será utilizada na Traz Tela
    objSolicitacaoServico.iFilialEmpresa = giFilialEmpresa
    
    'Move os dados da tabela SolicitacaoServico para a tela
    lErro = Traz_SolicitacaoServico_Tela(objSolicitacaoServico)
    If lErro <> SUCESSO And lErro <> 98087 Then gError 98113
        
    'Não achou a Solicitação de Serviço --> Erro
    If lErro = 98087 Then gError 98114
       
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0
    
    Me.Show

    Exit Sub
    
Erro_objEventoSolicitacao_evSelecao:

    Select Case gErr

        Case 98113
            
        Case 98114
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SOLICITACAO_NAO_CADASTRADA", gErr, objSolicitacaoServico.iFilialEmpresa, objSolicitacaoServico.lNumero)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoTabPreco_evSelecao(obj1 As Object)

Dim objTabPreco As ClassTabPreco
Dim lErro As Long

On Error GoTo Erro_objEventoTabPreco_evSelecao

    Set objTabPreco = obj1

    'Move o Nome Reduzido do Cliente e os campos referentes a Cidade, Estado e País de Origem e Destino para a tela
    lErro = Traz_TabPreco_Tela(objTabPreco)
    If lErro <> SUCESSO Then gError 98118
    
    Me.Show

    Exit Sub
    
Erro_objEventoTabPreco_evSelecao:

    Select Case gErr

        Case 98118
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoServico_evSelecao(obj1 As Object)

Dim objTabPrecoItens As ClassTabPrecoItens
Dim sProduto As String
Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iIndice As Integer
Dim sProduto1 As String
Dim iPreenchido As Integer

On Error GoTo Erro_objEventoServico_evSelecao

    'objTabPrecoItens é o obj de retorno do Browser TabPrecoProduto
    Set objTabPrecoItens = obj1

    'Verifica se alguma linha está selecionada
    If GridServicos.Row < 1 Then Exit Sub
    
    'Verifica se já existe este produto em outra linha do Grid
    For iIndice = 1 To objGridServicoss.iLinhasExistentes
        If iIndice <> GridServicos.Row Then
            
            sProduto1 = GridServicos.TextMatrix(iIndice, iGrid_Produto_Col)
            
            'Formata o produto contido na variável se estiver preenchida
            lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
            If lErro <> SUCESSO Then gError 98380
    
            'Se existir o produto em outra linha do grid --> erro
            If sProduto = objTabPrecoItens.sProduto Then gError 98381
        End If
    Next
       
    MaskProduto.PromptInclude = False
    MaskProduto.Text = objTabPrecoItens.sProduto
    MaskProduto.PromptInclude = True
    
    'Faz o Tratamento do produto
    lErro = Traz_Produto_Tela(objProduto)
    If lErro <> SUCESSO Then gError 98152
    
    If GridServicos.Row - GridServicos.FixedRows = objGridServicoss.iLinhasExistentes Then
        objGridServicoss.iLinhasExistentes = objGridServicoss.iLinhasExistentes + 1
    End If
    
    
    'Verifica se o browser está sendo chamado pelo botão, se for
    'joga no grid a descrição e o produto
    If Not (Me.ActiveControl Is MaskProduto) Then
        GridServicos.TextMatrix(GridServicos.Row, iGrid_DescProduto_Col) = objProduto.sDescricao
        GridServicos.TextMatrix(GridServicos.Row, iGrid_Produto_Col) = MaskProduto.Text
    Else
        GridServicos.TextMatrix(GridServicos.Row, iGrid_Produto_Col) = ""
    End If
        
    Me.Show

    Exit Sub

Erro_objEventoServico_evSelecao:

    Select Case gErr
    
        Case 98152
            GridServicos.TextMatrix(GridServicos.Row, iGrid_Produto_Col) = ""
        
        Case 98380
        
        Case 98381
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE1", gErr, objTabPrecoItens.sProduto, iIndice)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub
       
Public Function Trata_Parametros(Optional objSolicitacaoServico As ClassSolicitacaoServico) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
    
    'Verifica se alguma Solicitação de Serviço foi passada por parâmetro
    If Not (objSolicitacaoServico Is Nothing) Then
        
        objSolicitacaoServico.iFilialEmpresa = giFilialEmpresa
        
        'Tenta ler a Solicitação de Serviço passada por parâmetro
        lErro = Traz_SolicitacaoServico_Tela(objSolicitacaoServico)
        If lErro <> SUCESSO And lErro <> 98087 Then gError 98072

        'Se o Número passado não está cadastrado...
        If lErro = 98087 Then
            
            'Limpa a tela
            Call Limpa_SolicitacaoServico

            'Coloca o Número na tela
            MaskNumero.Text = objSolicitacaoServico.lNumero

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO
        
    Exit Function

Erro_Trata_Parametros:
    
    Trata_Parametros = gErr

    Select Case gErr

        Case 98072

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    iAlterado = 0

    Exit Function
    
End Function

Public Sub Form_Load()
'inicialização da tela

Dim lErro As Long
Dim colEmbalagens As New Collection
Dim colContainers As New Collection
Dim objTipoEmb As ClassTipoEmbalagem
Dim objTipoContainer As ClassTipoContainer

On Error GoTo Erro_Form_Load

    iFrameAtual = TAB_DadosPrincipais
     
    'Inicializa os objeventos
    Set objEventoSolicitacao = New AdmEvento
    Set objEventoTabPreco = New AdmEvento
    Set objEventoDespachante = New AdmEvento
    Set objEventoProgNavio = New AdmEvento
    Set objEventoServico = New AdmEvento
    
    'Inicializa o objGrid
    Set objGridServicoss = New AdmGrid
    
    'Inicializa o Grid
    lErro = Inicializa_GridServicos(objGridServicoss)
    If lErro <> SUCESSO Then gError 97040
    
    'Inicializa a máscara do Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", MaskProduto)
    If lErro <> SUCESSO Then gError 98069
    
    'Inicializando combo TipoEmbalagem com
    'as embalagens lidas da tabela
    lErro = CF("TipoEmbalagem_Le_Todos", colEmbalagens)
    If lErro <> SUCESSO Then gError 98070
    
    For Each objTipoEmb In colEmbalagens
    'Adiciona o item na combo de tipo embalagem e preenche o itemdata
        ComboTipoEmbalagem.AddItem objTipoEmb.iTipo & SEPARADOR & objTipoEmb.sDescricao
        ComboTipoEmbalagem.ItemData(ComboTipoEmbalagem.NewIndex) = objTipoEmb.iTipo
    Next
    
    'Inicializando combo TipoContainer com
    'os containers lidos da tabela
    lErro = CF("TipoContainer_Le_Todos", colContainers)
    If lErro <> SUCESSO And lErro <> 97091 Then gError 98071
    
    For Each objTipoContainer In colContainers
    'Adiciona o item na combo de tipo container e preenche o itemdata
        ComboTipoContainer.AddItem objTipoContainer.iTipo & SEPARADOR & objTipoContainer.sDescricao
        ComboTipoContainer.ItemData(ComboTipoContainer.NewIndex) = objTipoContainer.iTipo
    Next
    
    'Data de Emissão inicializada com a Data Atual
    MaskDataEmissao.PromptInclude = False
    MaskDataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    MaskDataEmissao.PromptInclude = True
    
    'Data de Pedido inicializada com a Data Atual
    MaskDataPedido.PromptInclude = False
    MaskDataPedido.Text = Format(gdtDataAtual, "dd/mm/yy")
    MaskDataPedido.PromptInclude = True
            
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
        
        Case 97040, 98069, 98070, 98071
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
    
    Exit Sub

End Sub

Private Function Inicializa_GridServicos(objGridInt As AdmGrid) As Long

On Error GoTo Erro_Inicializa_GridServicos

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Serviço")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Quantidade")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (MaskProduto.Name)
    objGridInt.colCampo.Add (TextDescProduto.Name)
    objGridInt.colCampo.Add (MaskQuantidade.Name)

    'Colunas do Grid
    iGrid_Produto_Col = 1
    iGrid_DescProduto_Col = 2
    iGrid_Quantidade_Col = 3

    'Grid do Grid
    objGridInt.objGrid = GridServicos

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_SERVICOS + 1
    
    'Usado para que se possa utilizar a Rotina_Enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    
    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 4

    'Largura da primeira coluna
    GridServicos.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridServicos = SUCESSO

    Exit Function
    
Erro_Inicializa_GridServicos:

    Inicializa_GridServicos = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select
    
    Exit Function

End Function

Sub Limpa_SolicitacaoServico()

    'Limpa a tela
    Call Limpa_Tela(Me)

    'Limpa os outros campos da tela
    ComboTipoContainer.ListIndex = -1
    ComboTipoEmbalagem.ListIndex = -1
    ComboTipoOperacao.ListIndex = -1
    LabelEstado(TIPO_ORIGEM).Caption = ""
    LabelEstado(TIPO_DESTINO).Caption = ""
    LabelPais(TIPO_ORIGEM).Caption = ""
    LabelPais(TIPO_DESTINO).Caption = ""
    LabelTextCliente.Caption = ""
    LabelCidade(TIPO_ORIGEM).Caption = ""
    LabelCidade(TIPO_DESTINO).Caption = ""
    LabelNavio.Caption = ""
    
    'Data de Emissão com a Data Atual
    MaskDataEmissao.PromptInclude = False
    MaskDataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    MaskDataEmissao.PromptInclude = True

    'Data de Pedido com a Data Atual
    MaskDataPedido.PromptInclude = False
    MaskDataPedido.Text = Format(gdtDataAtual, "dd/mm/yy")
    MaskDataPedido.PromptInclude = True
      
    'Limpa o grid da tela
    Call Grid_Limpa(objGridServicoss)
    
End Sub

Private Function Move_Tela_Memoria(objSolicitacaoServico As ClassSolicitacaoServico) As Long
'Move os dados da tela para a memória

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objOrigemDestino As New ClassOrigemDestino
Dim objDespachante As New ClassDespachante

On Error GoTo Erro_Move_Tela_Memoria
    
    objSolicitacaoServico.lNumero = StrParaLong(MaskNumero.Text)
    objSolicitacaoServico.lCodTabPreco = StrParaLong(MaskCodTabPreco.Text)
    
    'Move os dados do primeiro Tab para a memória
    objSolicitacaoServico.sNumReferencia = MaskDocRefer.Text
    objSolicitacaoServico.dtDataPedido = StrParaDate(MaskDataPedido.Text)
    objSolicitacaoServico.dQuantMaterial = StrParaDbl(MaskQuantMaterial.Text)
    objSolicitacaoServico.sMaterial = TextMaterial.Text
    objSolicitacaoServico.dValorMercadoria = StrParaDbl(MaskValorMerc.Text)
    objSolicitacaoServico.sUM = TextUM.Text
    objSolicitacaoServico.dtDataEmissao = StrParaDate(MaskDataEmissao.Text)
        
    If Len(Trim(TextDespachante.Text)) > 0 Then
    
        'Faz a leitura do Despachante
        objDespachante.sNomeReduzido = TextDespachante.Text
        
        lErro = CF("Despachante_Le_NomeRed", objDespachante)
        If lErro <> SUCESSO And lErro <> 98516 Then gError 98105
        
        'Se não encontrou despachante --> Erro
        If lErro = 98516 Then gError 98106
        
        objSolicitacaoServico.iDespachante = objDespachante.iCodigo
        
    End If
    
    If Not (ComboTipoEmbalagem.ListIndex < 0) Then
        objSolicitacaoServico.iTipoEmbalagem = ComboTipoEmbalagem.ItemData(ComboTipoEmbalagem.ListIndex)
    End If
    
    If Not (ComboTipoContainer.ListIndex < 0) Then
        objSolicitacaoServico.iTipoContainer = ComboTipoContainer.ItemData(ComboTipoContainer.ListIndex)
    End If
    
    If Not (ComboTipoOperacao.ListIndex < 0) Then
        objSolicitacaoServico.iTipoOperacao = ComboTipoOperacao.ItemData(ComboTipoOperacao.ListIndex)
    End If
    
    'Move os dados do GridServicos para a Memória
    Call Move_GridServicos_Memoria(objSolicitacaoServico)
    
    'Move os dados do Segundo tab para a memória
    objSolicitacaoServico.lCodProgNavio = StrParaLong(MaskIdentProgNavio.Text)
    objSolicitacaoServico.dtDataPrevFim = StrParaDate(MaskDataFim.Text)
    objSolicitacaoServico.dtHoraPrevFim = StrParaDate(MaskHoraFim.Text)
    objSolicitacaoServico.dtDataPrevInicio = StrParaDate(MaskDataInicio.Text)
    objSolicitacaoServico.dtHoraPrevInicio = StrParaDate(MaskHoraInicio.Text)
    objSolicitacaoServico.sObservacao = TextObs.Text
    objSolicitacaoServico.sBooking = MaskBooking.Text
    objSolicitacaoServico.sPorto = MaskPorto.Text
    
    'Inicialização dos endereços
    Set objSolicitacaoServico.objEnderecoOrigem = New ClassEndereco
    Set objSolicitacaoServico.objEnderecoDestino = New ClassEndereco
    
    'Move os dados do terceiro tab para a memória
    Call Move_Endereco_Memoria(objSolicitacaoServico.objEnderecoOrigem, TIPO_ORIGEM)
    
    'Move os dados do quarto tab para a memória
    Call Move_Endereco_Memoria(objSolicitacaoServico.objEnderecoDestino, TIPO_DESTINO)
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
        
        Case 98105
                    
        Case 98106
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESPACHANTE_NAO_CADASTRADO", gErr, TextDespachante.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Sub Move_Endereco_Memoria(objEndereco As ClassEndereco, iTipo As Integer)
'Move os dados do endereço para a memória
'utilizado para os dados da origem e do destino
'Constantes TIPO_ORIGEM ou TIPO_DESTINO são passadas por parâmetro para diferenciar o endereço a ser recolhido

    objEndereco.sEndereco = TextEndereco(iTipo).Text
    objEndereco.sBairro = MaskBairro(iTipo).Text
    objEndereco.sContato = MaskContato(iTipo).Text
    objEndereco.sTelefone1 = MaskTelefone1(iTipo).Text
    objEndereco.sTelefone2 = MaskTelefone2(iTipo).Text
    objEndereco.sFax = MaskFax(iTipo).Text
    objEndereco.sEmail = MaskEmail(iTipo).Text
    objEndereco.sCEP = MaskCEP(iTipo).ClipText
    objEndereco.iCodigoPais = PAIS_BRASIL '??? jones 18/12/07
    
End Sub

Private Sub Move_GridServicos_Memoria(objSolicitacaoServico As ClassSolicitacaoServico)
'Move os dados do GridServicos para a memória

Dim iIndice As Integer
Dim objSolServServicos As ClassServico
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_Move_GridServicos_Memoria

    'Para cada Serviço do grid
    For iIndice = 1 To objGridServicoss.iLinhasExistentes
        
        'inicializa o obj
        Set objSolServServicos = New ClassServico
        
        'Coloca o produto no formato de banco de dados
        lErro = CF("Produto_Formata", Trim(GridServicos.TextMatrix(iIndice, iGrid_Produto_Col)), sProduto, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 98306
        
        'recolhe os dados do grid de Serviços e adiciona na coleção
        objSolServServicos.sProduto = sProduto
        objSolServServicos.dQuant = StrParaDbl(GridServicos.TextMatrix(iIndice, iGrid_Quantidade_Col))
        
        'Adiciona o obj já carregado na coleção
        objSolicitacaoServico.colServico.Add objSolServServicos

    Next
    
    Exit Sub
    
Erro_Move_GridServicos_Memoria:

    Select Case gErr
        
        Case 98306
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Function Traz_SolicitacaoServico_Tela(objSolicitacaoServico As ClassSolicitacaoServico) As Long

'Move os dados da tabela SolicitacaoServico para a tela

Dim lErro As Long
Dim objEndereco As New ClassEndereco

On Error GoTo Erro_Traz_SolicitacaoServico_Tela
                    
    Call Limpa_SolicitacaoServico
    
    'Le os dados da tabela Solicitação de Serviço
    lErro = CF("SolicitacaoServico_Le", objSolicitacaoServico)
    If lErro <> SUCESSO And lErro <> 98085 Then gError 98086
    
    'Se não encontrar --> Erro
    If lErro = 98085 Then gError 98087
    
    'Le os dados do serviço relacionado a Solicitação de Serviço em questão
    lErro = CF("SolServServico_Le", objSolicitacaoServico)
    If lErro <> SUCESSO And lErro <> 98095 Then gError 98088

    'Nao achou os Serviços associados a Solicitação de Serviço --> erro
    If lErro = 98095 Then gError 98089
   
    'carrega a tela com os dados dos endereços passados no obj
    Call Carrega_Endereco(objSolicitacaoServico.objEnderecoDestino, TIPO_DESTINO)
    Call Carrega_Endereco(objSolicitacaoServico.objEnderecoOrigem, TIPO_ORIGEM)
    
    'carrega o tab1 com os dados passados no obj
    Call Carrega_Tela1(objSolicitacaoServico)
        
    'carrega o tab2 com os dados passados no obj
    Call Carrega_Tela2(objSolicitacaoServico)
        
    'carrega o Grid de Serviços com os dados passados no obj
    Call Carrega_GridServicos(objSolicitacaoServico)
        
    'Fecha comando de setas
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Traz_SolicitacaoServico_Tela = SUCESSO

    Exit Function

Erro_Traz_SolicitacaoServico_Tela:

    Traz_SolicitacaoServico_Tela = gErr

    Select Case gErr

        Case 98086, 98087, 98088
        
        Case 98089
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SOLSERVSERVICO_NAO_CADASTRADA", gErr, objSolicitacaoServico.lNumero)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Traz_Produto_Tela(objProduto As ClassProduto) As Long
'Faz o Tratamento do produto

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim iIndice As Integer
Dim sProduto As String

On Error GoTo Erro_Traz_Produto_Tela

    'Critica o Produto verificando se ele existe ou não
    lErro = CF("Produto_Critica_Filial", MaskProduto.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 Then gError 98133
    
    'Se o Produto não existe --> Erro
    If lErro = 51381 Then gError 98134
    
    'Coloca a máscara no prouto
    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 98135
           
    'O controle do produto no grid recebe o produto formatado
    MaskProduto.PromptInclude = False
    MaskProduto.Text = sProduto
    MaskProduto.PromptInclude = True

    'Verifica se já existe este produto em outra linha do Grid
    For iIndice = 1 To objGridServicoss.iLinhasExistentes
        If iIndice <> GridServicos.Row Then
            'Se existir --> erro
            If GridServicos.TextMatrix(iIndice, iGrid_Produto_Col) = MaskProduto.Text Then gError 98136
        End If
    Next

    'Verifica se é de Faturamento, se não for --> Erro
    If objProduto.iFaturamento = PRODUTO_NAO_VENDAVEL Then gError 98137
    
    Traz_Produto_Tela = SUCESSO

    Exit Function

Erro_Traz_Produto_Tela:

    Traz_Produto_Tela = gErr

    Select Case gErr

        Case 98133
        
        Case 98134
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, MaskProduto.Text)
           
        Case 98135
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sNomeReduzido)
                
        Case 98136
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE1", gErr, MaskProduto.Text, iIndice)
        
        Case 98137
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_VENDIDO2", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function
'Mario
Private Function Traz_TabPreco_Tela(objTabPreco As ClassTabPreco) As Long
'Move o Nome Reduzido do Cliente e os campos referentes a Cidade, Estado e País de Origem e Destino para a tela

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objOrigemDestino As New ClassOrigemDestino

On Error GoTo Erro_Traz_TabPreco_Tela

    'Le a tabela de Preços referentes a Solicitação de Serviço
    lErro = CF("TabPrecoServico_Le", objTabPreco)
    If lErro <> SUCESSO And lErro <> 98101 Then gError 98077

    'Nao achou nenhuma Tabela de Preços --> Erro
    If lErro = 98101 Then gError 98078
        
    'Move os dados contidos no obj para a tela
    MaskCodTabPreco.Text = objTabPreco.lCodigo
    LabelTextCliente.Caption = objTabPreco.sClienteNomeRed
    LabelCidade(TIPO_ORIGEM).Caption = objTabPreco.sOrigemCidade
    LabelEstado(TIPO_ORIGEM).Caption = objTabPreco.sOrigemUF
    LabelCidade(TIPO_DESTINO).Caption = objTabPreco.sDestinoCidade
    LabelEstado(TIPO_DESTINO).Caption = objTabPreco.sDestinoUF
    LabelPais(TIPO_DESTINO).Caption = PAIS_BRASIL_NOME
    LabelPais(TIPO_ORIGEM).Caption = PAIS_BRASIL_NOME
    
    'Fecha o comando de setas
    Call ComandoSeta_Fechar(Me.Name)
    
    Traz_TabPreco_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_TabPreco_Tela:
    
    Traz_TabPreco_Tela = gErr
    
    Select Case gErr

        Case 98077
        
        Case 98078
            Call Rotina_Erro(vbOKOnly, "ERRO_TABPRECO_NAO_CADASTRADA1", gErr, objTabPreco.lCodigo)
                           
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Sub Carrega_Endereco(objEndereco As ClassEndereco, iTipo As Integer)
'Move os dados do endereço trazidos no obj para a tela
'utilizado para os dados da origem e do destino
'Constantes TIPO_ORIGEM ou TIPO_DESTINO são passadas por parâmetro para diferenciar o endereço a ser jogado
    
On Error GoTo Erro_Carrega_Endereco
    
    TextEndereco(iTipo).Text = objEndereco.sEndereco
    MaskBairro(iTipo).Text = objEndereco.sBairro
    LabelCidade(iTipo).Caption = objEndereco.sCidade
    MaskContato(iTipo).Text = objEndereco.sContato
    MaskTelefone1(iTipo).Text = objEndereco.sTelefone1
    MaskTelefone2(iTipo).Text = objEndereco.sTelefone2
    MaskFax(iTipo).Text = objEndereco.sFax
    MaskEmail(iTipo).Text = objEndereco.sEmail
    LabelEstado(iTipo).Caption = objEndereco.sSiglaEstado
    
    'Se o código for zero --> o país é o Brasil
    If objEndereco.iCodigoPais = PAIS_BRASIL Then
        LabelPais(iTipo).Caption = PAIS_BRASIL_NOME
    Else
        LabelPais(iTipo).Caption = PAIS_BRASIL_NOME
    End If
    
    MaskCEP(iTipo).PromptInclude = False
    MaskCEP(iTipo).Text = objEndereco.sCEP
    MaskCEP(iTipo).PromptInclude = True
        
    Exit Sub
    
Erro_Carrega_Endereco:
            
    Select Case gErr
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
    
End Sub

Private Sub Carrega_Tela1(objSolicitacaoServico As ClassSolicitacaoServico)
'Move os dados carregados no obj para o Tab1

Dim iIndex As Integer

On Error GoTo Erro_Carrega_Tela1

    MaskNumero.Text = objSolicitacaoServico.lNumero
    
    MaskCodTabPreco.Text = objSolicitacaoServico.lCodTabPreco
    
    LabelTextCliente.Caption = objSolicitacaoServico.sClienteNomeRed
    MaskDocRefer.Text = objSolicitacaoServico.sNumReferencia
    
    MaskDataPedido.PromptInclude = False
    MaskDataPedido.Text = Format(objSolicitacaoServico.dtDataPedido, "dd/mm/yy")
    MaskDataPedido.PromptInclude = True
    
    MaskDataEmissao.PromptInclude = False
    MaskDataEmissao.Text = Format(objSolicitacaoServico.dtDataEmissao, "dd/mm/yy")
    MaskDataEmissao.PromptInclude = True
    
    MaskQuantMaterial.Text = objSolicitacaoServico.dQuantMaterial
    TextMaterial.Text = objSolicitacaoServico.sMaterial
    MaskValorMerc.Text = objSolicitacaoServico.dValorMercadoria
    TextUM.Text = objSolicitacaoServico.sUM
    
    'Por não ser campo obrigatório na tela é necessário verificação de sua existência
    TextDespachante.Text = objSolicitacaoServico.sDespachanteNomeRed
        
    'Busca na combo o Tipo de Operação do obj
    For iIndex = 0 To ComboTipoOperacao.ListCount - 1
        'Quando achar, seleciona este item na combo
        If ComboTipoOperacao.ItemData(iIndex) = objSolicitacaoServico.iTipoOperacao Then
            ComboTipoOperacao.ListIndex = iIndex
            Exit For
        End If
    Next
    
    'Busca na combo o Tipo de Embalagem do obj
    For iIndex = 0 To ComboTipoEmbalagem.ListCount - 1
        'Quando achar, seleciona este item na combo
        If ComboTipoEmbalagem.ItemData(iIndex) = objSolicitacaoServico.iTipoEmbalagem Then
            ComboTipoEmbalagem.ListIndex = iIndex
            Exit For
        End If
    Next
    
    'Busca na combo o Tipo de Container do obj
    For iIndex = 0 To ComboTipoContainer.ListCount - 1
        'Quando achar, seleciona este item na combo
        If ComboTipoContainer.ItemData(iIndex) = objSolicitacaoServico.iTipoContainer Then
            ComboTipoContainer.ListIndex = iIndex
            Exit For
        End If
    Next
            
    Exit Sub
    
Erro_Carrega_Tela1:
            
    Select Case gErr
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub Carrega_Tela2(objSolicitacaoServico As ClassSolicitacaoServico)
'Move os dados carregados no obj para o Tab2
'Por nenhum dado ser campo obrigatório na tela é necessário verificação de sua existência no obj

On Error GoTo Erro_Carrega_Tela2

    If objSolicitacaoServico.lCodProgNavio <> 0 Then
        MaskIdentProgNavio.Text = objSolicitacaoServico.lCodProgNavio
         Call MaskIdentPrognavio_Validate(False)
    End If
    
    If objSolicitacaoServico.dtDataPrevInicio <> DATA_NULA Then
        MaskDataInicio.PromptInclude = False
        MaskDataInicio.Text = Format(objSolicitacaoServico.dtDataPrevInicio, "dd/mm/yy")
        MaskDataInicio.PromptInclude = True
    End If
    

    If objSolicitacaoServico.dtHoraPrevInicio <> DATA_NULA Then
        MaskHoraInicio.PromptInclude = False
        MaskHoraInicio.Text = Format(objSolicitacaoServico.dtHoraPrevInicio, "hh:mm:ss")
        MaskHoraInicio.PromptInclude = True
    End If
    
    If objSolicitacaoServico.dtDataPrevFim <> DATA_NULA Then
        MaskDataFim.PromptInclude = False
        MaskDataFim.Text = Format(objSolicitacaoServico.dtDataPrevFim, "dd/mm/yy")
        MaskDataFim.PromptInclude = True
    End If
    
    If objSolicitacaoServico.dtHoraPrevFim <> DATA_NULA Then
        MaskHoraFim.PromptInclude = False
        MaskHoraFim.Text = Format(objSolicitacaoServico.dtHoraPrevFim, "hh:mm:ss")
        MaskHoraFim.PromptInclude = True
    End If
        
    MaskPorto.Text = objSolicitacaoServico.sPorto
    MaskBooking.Text = objSolicitacaoServico.sBooking
    TextObs.Text = objSolicitacaoServico.sObservacao
        
    Exit Sub
    
Erro_Carrega_Tela2:
            
    Select Case gErr
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
    
End Sub

Private Sub Carrega_GridServicos(objSolicitacaoServico As ClassSolicitacaoServico)
'Carrega os dados do grid no obj para a tela

Dim iLinha As Integer
Dim objSolServServico As New ClassServico
Dim sProdutoEnxuto As String
Dim lErro As Long

On Error GoTo Erro_Carrega_GridServicos

    'Limpa o Grid de Serviços
    Call Grid_Limpa(objGridServicoss)

    iLinha = 0

    'Preenche o grid com os objetos da coleção de Serviços
    For Each objSolServServico In objSolicitacaoServico.colServico

        iLinha = iLinha + 1
        
        'coloca a mascara no produto
        lErro = Mascara_RetornaProdutoEnxuto(objSolServServico.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 98097
        
        'Coloca o produto já com a máscara no controle
        MaskProduto.PromptInclude = False
        MaskProduto.Text = sProdutoEnxuto
        MaskProduto.PromptInclude = True
        
        'Preenche o grid de serviços com os dados
        GridServicos.TextMatrix(iLinha, iGrid_Produto_Col) = MaskProduto.Text
        GridServicos.TextMatrix(iLinha, iGrid_DescProduto_Col) = objSolServServico.sDescricao
        GridServicos.TextMatrix(iLinha, iGrid_Quantidade_Col) = Format(objSolServServico.dQuant, "Standard")
       
    Next

    'Preenche com o número atual de linhas existentes no grid
    objGridServicoss.iLinhasExistentes = iLinha
    
    Exit Sub
    
Erro_Carrega_GridServicos:

    Select Case gErr

        Case 98097
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objSolServServico.sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()
'gera um novo número para Solicitação de Serviço automaticamente

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'retorna em lCodigo o próximo número a ser usado
    lErro = CF("Config_ObterAutomatico", "FatConfig", "NUM_PROX_SOLICITACAO", "SolicitacaoServico", "Numero", lCodigo)
    If lErro <> SUCESSO Then gError 98111
    
    'Joga o Número na tela
    MaskNumero.Text = CStr(lCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 98111

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Function TP_Despachante_Le(objDespachante As ClassDespachante, sDespachante As String) As Long
'Lê o Despachante com Código ou NomeRed

Dim eTipoDespachante As enumTipo
Dim lErro As Long

On Error GoTo TP_Despachante_Le
    
    'Verifica se o parâmetro passado "sdespachante" é um inteiro ou uma string
    eTipoDespachante = Tipo_Despachante(sDespachante)

    Select Case eTipoDespachante
    
    'Se for uma string...
    Case TIPO_STRING
        
        'Joga o conteudo da variável no nome reduzido do obj
        objDespachante.sNomeReduzido = sDespachante
        
        'Tenta encontrar o despachante com o nome reduzido passado
        lErro = CF("Despachante_Le_NomeRed", objDespachante)
        If lErro <> SUCESSO And lErro <> 98516 Then gError 98139
        
        'Se não encontrou --> Erro
        If lErro = 98516 Then gError 98140
                                  
    'Se for um Código...
    Case TIPO_CODIGO
        
        'Joga o conteudo da variável no Código do obj
        objDespachante.iCodigo = StrParaInt(sDespachante)
        
        'Tenta encontrar o despachante com o Código passado
        lErro = CF("Despachante_Le", objDespachante)
        If lErro <> SUCESSO And lErro <> 96679 Then gError 98141
        
        'Se não encontrou --> Erro
        If lErro = 96679 Then gError 98142
   
    Case TIPO_DECIMAL

        gError 98144

    Case TIPO_NAO_POSITIVO

        gError 98145

    End Select

    TP_Despachante_Le = SUCESSO

    Exit Function

TP_Despachante_Le:

    TP_Despachante_Le = gErr

    Select Case gErr
        
        Case 98139, 98141 'Tratados nas rotinas chamadas
        
        Case 98140
            'Envia aviso que Despachante não está cadastrado e pergunta se deseja criar
            lErro = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_DESPACHANTE", objDespachante.sNomeReduzido)
    
                If lErro = vbYes Then
                    'Chama tela de Despachante
                    lErro = Chama_Tela("Despachante", objDespachante)
                End If
                
        Case 98142
            'Envia aviso que Despachante não está cadastrado e pergunta se deseja criar
            lErro = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_DESPACHANTE", objDespachante.iCodigo)
    
                If lErro = vbYes Then
                    'Chama tela de Despachante
                    lErro = Chama_Tela("Despachante", objDespachante)
                End If
   
        Case 98144
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", gErr, sDespachante)

        Case 98145
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", gErr, sDespachante)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    Exit Function

End Function

Private Function Tipo_Despachante(ByVal sText As String) As enumTipo
'Verifica se o parâmetro passado em "stext" é um inteiro ou uma string

    If Not IsNumeric(sText) Or Len(Trim(sText)) > 5 Then
        Tipo_Despachante = TIPO_STRING
    ElseIf Int(CDbl(sText)) <> CDbl(sText) Then
        Tipo_Despachante = TIPO_DECIMAL
    ElseIf CDbl(sText) <= 0 Then
        Tipo_Despachante = TIPO_NAO_POSITIVO
    Else
        Tipo_Despachante = TIPO_CODIGO
    End If

End Function

Public Sub BotaoServicos_Click()
'Chama o browser do TabPrecoProduto
'Este Browser traz os dados do Produto pertencentes a tabela de Preço
'usada nesta solicitação com data de vigencia menor ou igual a
'data de emissão da tela

Dim objTabPrecoItens As New ClassTabPrecoItens
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim colSelecao As New Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoServicos_Click
    
    'Verifica se os campos necessários para a chamada do browser estão preenchidos, senão --> erro
    If Len(Trim(MaskCodTabPreco.ClipText)) = 0 Then gError 98150
    If Len(Trim(MaskDataEmissao.ClipText)) = 0 Then gError 98151
    
    'Verifica se o browser está sendo chamado do controle(F3), se for
    'joga o conteudo do controle numa variável
    If Me.ActiveControl Is MaskProduto Then
    
        sProduto1 = MaskProduto.Text
        
    Else
    'Verifica se o browser está sendo chamado pelo botão, se for
    'joga o conteudo do grid numa variável
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridServicos.Row = 0 Then gError 98148

        sProduto1 = GridServicos.TextMatrix(GridServicos.Row, iGrid_Produto_Col)
        
    End If
    
    'Formata o produto contido na variável se estiver preenchida
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 98149
    
    'Se não estiver --> limpa a variável
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    colSelecao.Add StrParaInt(MaskCodTabPreco.Text)
    colSelecao.Add StrParaDate(MaskDataEmissao.Text)
    
    'Chama a tela de browse TabPrecoProdutoLista
    Call Chama_Tela("TabPrecoProdutoLista", colSelecao, objTabPrecoItens, objEventoServico)

    Exit Sub
        
Erro_BotaoServicos_Click:
    
    Select Case gErr
        
        Case 98148
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 98149 'Tratado na rotina chamada
        
        Case 98150
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABPRECO_NAO_PREENCHIDA", gErr)
            
        Case 98151
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    'Verifica se existe algo para ser salvo antes de sair
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 98184
    
    'Fecha a tela
    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case 98184

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Chama a função de gravação e limpa a tela

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 98185

    'Limpa a Tela
    Call Limpa_SolicitacaoServico
    
    iAlterado = 0
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 98185

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()
'pergunta se o usuário deseja salvar as alterações e limpa a Tela

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 98186

    'Limpa a Tela
    Call Limpa_SolicitacaoServico
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0
    
    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 98186

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()
'Exclui a Solicitação de Serviço do Número passado

Dim lErro As Long
Dim vbMsgRet As VbMsgBoxResult
Dim lCodigo As Long
Dim objSolicitacaoServico As New ClassSolicitacaoServico

On Error GoTo Erro_BotaoExcluir_Click
      
    'Coloca o cursor com formato de ampulheta
    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o campo Número foi informado, senão --> Erro.
    If Len(Trim(MaskNumero.ClipText)) = 0 Then gError 98164
            
    'Carrega o obj com os dados do número e filial
    objSolicitacaoServico.lNumero = StrParaLong(MaskNumero.Text)
    objSolicitacaoServico.iFilialEmpresa = giFilialEmpresa
    
    'Lê a solicitação com o número e filial passados
    lErro = CF("SolicitacaoServico_Le", objSolicitacaoServico)
    If lErro <> SUCESSO And lErro <> 98085 Then gError 98165
    
    'Se não está cadastrado --> Erro
    If lErro = 98085 Then gError 98166

    'Pede confirmação para exclusão ao usuário
    vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_SOLICITACAO", objSolicitacaoServico.iFilialEmpresa, objSolicitacaoServico.lNumero)
    
    'Se confirma
    If vbMsgRet = vbYes Then

        'exclui a Solicitação de Serviço
        lErro = CF("SolicitacaoServico_Exclui", objSolicitacaoServico)
        If lErro <> SUCESSO Then gError 98167
        
        'Fecha o comando das setas se estiver aberto
        Call ComandoSeta_Fechar(Me.Name)
        
        'Limpa a Tela
        Call Limpa_SolicitacaoServico

        iAlterado = 0

    End If
    
    'Retorna o cursor para seu formato default
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    'Retorna o cursor para seu formato default
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
                
        Case 98164
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INFORMADO1", gErr)

        Case 98165, 98167
        
        Case 98166
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SOLICITACAO_NAO_CADASTRADA", gErr, objSolicitacaoServico.iFilialEmpresa, objSolicitacaoServico.lNumero)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select
   
    Exit Sub

End Sub

Private Sub BotaoImprimir_Click()
'Acabar de fazer ainda ...

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objSolicitacaoServico As New ClassSolicitacaoServico
Dim iIndice As Integer

On Error GoTo Erro_BotaoImprimir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se os campos obrigatórios foram informados, senão --> Erro.
    If Len(Trim(MaskNumero.ClipText)) = 0 Then gError 98187
    
    'Preenche o obj com os dados do número e da filial
    objSolicitacaoServico.lNumero = StrParaLong(MaskNumero.Text)
    objSolicitacaoServico.iFilialEmpresa = giFilialEmpresa
    
    'Le a solicitação de serviço com o número e filial passados
    lErro = CF("SolicitacaoServico_Le", objSolicitacaoServico)
    If lErro <> SUCESSO And lErro <> 98085 Then gError 98201
    
    'Se não encontrou --> erro
    If lErro = 98085 Then gError 98202

    lErro = objRelatorio.ExecutarDireto("Solicitação de Serviço", "", 1, "", "NNUMSOLSERV", objSolicitacaoServico.lNumero)
    If lErro <> SUCESSO Then gError 98203

    'Limpa a Tela
    Call Limpa_SolicitacaoServico

    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 98187
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INFORMADO1", gErr)

        Case 98201, 98203
        
        Case 98202
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SOLICITACAO_NAO_CADASTRADA", gErr, objSolicitacaoServico.iFilialEmpresa, objSolicitacaoServico.lNumero)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
    
    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

End Sub

Function Gravar_Registro() As Long
'Chama as funções de recolhimento de dados da tela e Gravação

Dim objSolicitacaoServico As New ClassSolicitacaoServico
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Gravar_Registro

    'Coloca o cursor com formato de ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos obrigatórios foram informados, senão --> Erro.
    If Len(Trim(MaskNumero.ClipText)) = 0 Then gError 98168
    If Len(Trim(MaskCodTabPreco.ClipText)) = 0 Then gError 98169
    If Len(Trim(MaskDocRefer.Text)) = 0 Then gError 98170
    If Len(Trim(MaskDataPedido.ClipText)) = 0 Then gError 98171
    If Len(Trim(MaskDataEmissao.ClipText)) = 0 Then gError 98172
    If ComboTipoOperacao.ListIndex = -1 Then gError 98173
    If Len(Trim(ComboTipoContainer.Text)) = 0 Then gError 98178
                
    'Se não houver pelo menos uma linha do grid preenchida, ERRO.
    If objGridServicoss.iLinhasExistentes <= 0 Then gError 98180

    'Verifica se a quantidade e o produto do grid estão preenchidos, senão --> erro
    For iIndice = 1 To objGridServicoss.iLinhasExistentes
        If Len(Trim(GridServicos.TextMatrix(iIndice, iGrid_Produto_Col))) = 0 Then gError 98284
        If Len(Trim(GridServicos.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 98285
    Next
    
    'Se a Hora do Inicio está preenchida obriga a data de Inicio também estar
    If Len(Trim(MaskDataInicio.ClipText)) = 0 And Len(Trim(MaskHoraInicio.ClipText)) <> 0 Then gError 98373
    'Se a Hora do fim está preenchida obriga a data de fim também estar
    If Len(Trim(MaskDataFim.ClipText)) = 0 And Len(Trim(MaskHoraFim.ClipText)) <> 0 Then gError 98374
    
    'Verificação das datas do tab2
    If Len(Trim(MaskDataInicio.ClipText)) <> 0 And Len(Trim(MaskDataFim.ClipText)) <> 0 Then
        'Verifica se a data de fim é menor que a data de início, se for --> erro
        If StrParaDate(MaskDataFim.Text) < StrParaDate(MaskDataInicio.Text) Then gError 98274
    End If
    
    'Caso as datas sejam iguais --> Verificação das horas
    If MaskDataInicio.ClipText = MaskDataFim.ClipText Then
        'Verifica se a Hora de fim é menor que a hora de início, se for --> erro
        If StrParaDate(MaskHoraFim.Text) < StrParaDate(MaskHoraInicio.Text) Then gError 98328
    End If
          
    'Move os dados da tela para a memória
    lErro = Move_Tela_Memoria(objSolicitacaoServico)
    If lErro <> SUCESSO Then gError 98181
    
    'preenche o obj com a filial
    objSolicitacaoServico.iFilialEmpresa = giFilialEmpresa
    
    'Verifica se o Número da solicitação já existe, se existir manda uma mensagem
    lErro = Trata_Alteracao(objSolicitacaoServico, objSolicitacaoServico.lNumero, objSolicitacaoServico.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 98182
    
    'Verifica se a solicitação já existe e se está associada a um comprovante
    lErro = CF("Verifica_Comprovante", objSolicitacaoServico)
    If lErro <> SUCESSO Then gError 98324
    
    'Grava no BD os dados da Tela
    lErro = CF("SolicitacaoServico_Grava", objSolicitacaoServico)
    If lErro <> SUCESSO Then gError 98183

     'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Retorna o cursor ao formato default
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    'Retorna o cursor ao formato default
    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 98168
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INFORMADO1", gErr)

        Case 98169
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABPRECO_NAO_PREENCHIDA", gErr)

        Case 98170
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOCREFER_NAO_PREENCHIDO", gErr)

        Case 98171
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAPEDIDO_NAO_PREENCHIDA", gErr)
            
        Case 98172
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)

        Case 98173
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOOPERACAO_NAO_SELECIONADO", gErr)

        Case 98178
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCONTAINER_NAO_SELECIONADO", gErr)
        
        Case 98180
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRID_NAO_PREENCHIDO1", gErr)
                
        Case 98181, 98182, 98183, 98324
        
        Case 98274
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAFIM_MENOR_DATAINICIO", gErr, MaskDataFim.Text, MaskDataInicio.Text)
                      
        Case 98284
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_INFORMADO", gErr, iIndice)
                      
        Case 98285
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 98328
            Call Rotina_Erro(vbOKOnly, "ERRO_HORAFIM_MENOR_HORAINICIO", gErr, MaskHoraFim.Text, MaskHoraInicio.Text)
        
        Case 98373, 98374
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA1", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

     End Select

     Exit Function

End Function

Private Sub MaskNumero_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskNumero_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskNumero, iAlterado)

End Sub

Private Sub MaskNumero_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskNumero_Validate
    
    'Verifica se o campo numero foi preenchido
    If Len(Trim(MaskNumero.Text)) = 0 Then Exit Sub
    
    'Faz a crítica do Número
    lErro = Long_Critica(MaskNumero.Text)
    If lErro <> SUCESSO Then gError 98112

    Exit Sub

Erro_MaskNumero_Validate:

    Cancel = True

    Select Case gErr

        Case 98112

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub MaskCodTabPreco_Change()

    iAlterado = REGISTRO_ALTERADO
    giTabPrecoAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub MaskCodTabPreco_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskCodTabPreco, iAlterado)

End Sub

Private Sub MaskCodTabPreco_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTabPreco As New ClassTabPreco

On Error GoTo Erro_MaskCodTabPreco_Validate
    
    'Verifica se o campo codtabpreco foi alterado
    If giTabPrecoAlterado <> REGISTRO_ALTERADO Then Exit Sub
    
    'verifica se o campo codtabpreco foi preenchido
    If Len(Trim(MaskCodTabPreco.Text)) = 0 Then Exit Sub
    
    'Faz a critica do código
    lErro = Long_Critica(MaskCodTabPreco.Text)
    If lErro <> SUCESSO Then gError 98116
    
    'Joga o conteudo deste campo para o obj
    objTabPreco.lCodigo = StrParaLong(MaskCodTabPreco)
    objTabPreco.dtDataVigencia = StrParaDate(MaskDataEmissao.Text)
    
    'Move o Nome Reduzido do Cliente e os campos referentes a Cidade, Estado e País de Origem e Destino para a tela
    lErro = Traz_TabPreco_Tela(objTabPreco)
    If lErro <> SUCESSO Then gError 98117
        
    Exit Sub

Erro_MaskCodTabPreco_Validate:

    Cancel = True

    Select Case gErr

        Case 98116, 98117
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub MaskDocRefer_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskDataPedido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskDataPedido_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskDataPedido, iAlterado)

End Sub

Private Sub MaskDataPedido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskDataPedido_Validate

    'Verifica se a data de Pedido foi digitada
    If Len(Trim(MaskDataPedido.Text)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(MaskDataPedido.Text)
    If lErro <> SUCESSO Then gError 98122
    
    Exit Sub

Erro_MaskDataPedido_Validate:

    Cancel = True

    Select Case gErr

        Case 98122

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDown1_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    'Diminui um dia em DataPedido
    lErro = Data_Up_Down_Click(MaskDataPedido, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 98123

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 98123
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    'Aumenta um dia em DataPedido
    lErro = Data_Up_Down_Click(MaskDataPedido, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 98124

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 98124
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub MaskDataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskDataEmissao_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskDataEmissao, iAlterado)

End Sub

Private Sub MaskDataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskDataEmissao_Validate

    'Verifica se a data de Emissao foi digitada
    If Len(Trim(MaskDataEmissao.Text)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(MaskDataEmissao.Text)
    If lErro <> SUCESSO Then gError 98125
    
    Exit Sub

Erro_MaskDataEmissao_Validate:

    Cancel = True

    Select Case gErr

        Case 98125

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDown2_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub UpDown2_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_DownClick

    'Diminui um dia em DataEmissao
    lErro = Data_Up_Down_Click(MaskDataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 98126

    Exit Sub

Erro_UpDown2_DownClick:

    Select Case gErr

        Case 98126
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDown2_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown2_UpClick

    'Aumenta um dia em DataEmissao
    lErro = Data_Up_Down_Click(MaskDataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 98127

    Exit Sub

Erro_UpDown2_UpClick:

    Select Case gErr

        Case 98127
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub ComboTipoOperacao_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TextDespachante_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TextDespachante_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objDespachante As New ClassDespachante

On Error GoTo Erro_TextDespachante_Validate
    
    'Se não estiver preenchido sai da Sub
    If Len(Trim(TextDespachante.Text)) = 0 Then Exit Sub

    'Faz a leitura do Despachante
    lErro = TP_Despachante_Le(objDespachante, TextDespachante.Text)
    If lErro <> SUCESSO Then gError 98138
    
    'Preenche o Campo Despachante com o nome reduzido
    TextDespachante.Text = objDespachante.sNomeReduzido
                   
    Exit Sub
    
Erro_TextDespachante_Validate:

    Cancel = True

    Select Case gErr

        Case 98138
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub TextMaterial_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub maskQuantMaterial_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub MaskQuantMaterial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskQuantMaterial_Validate
    
    'Verifca se a quantidade do material foi informado
    If Len(Trim(MaskQuantMaterial.Text)) = 0 Then Exit Sub
    
    'Verifica se trata-se de um valor positivo
    lErro = Valor_NaoNegativo_Critica(MaskQuantMaterial.Text)
    If lErro <> SUCESSO Then gError 98146
    
    'Joga na tela com seu formato
    MaskQuantMaterial.Text = Format(MaskQuantMaterial.Text, "Standard")

    Exit Sub

Erro_MaskQuantMaterial_Validate:

    Cancel = True

    Select Case gErr

        Case 98146
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub TextUM_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskValorMerc_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskValorMerc_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorMaterial As Double

On Error GoTo Erro_MaskValorMerc_Validate
    
    'Verifca se o valor do material foi informado
    If Len(Trim(MaskValorMerc.Text)) = 0 Then Exit Sub
    
    'Verifica se trata-se de um valor positivo
    lErro = Valor_NaoNegativo_Critica(MaskValorMerc.Text)
    If lErro <> SUCESSO Then gError 98147
    
    'Joga o valor numa variável
    dValorMaterial = CDbl(MaskValorMerc.Text)
    
    'Coloca no seu formato e joga na tela
    MaskValorMerc.Text = Format(dValorMaterial, "STANDARD")

    Exit Sub

Erro_MaskValorMerc_Validate:

    Cancel = True

    Select Case gErr

        Case 98147
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub ComboTipoEmbalagem_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ComboTipoContainer_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskIdentPrognavio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskIdentPrognavio_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskIdentProgNavio, iAlterado)

End Sub

Private Sub MaskIdentPrognavio_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProgNavio As New ClassProgNavio

On Error GoTo Erro_MaskIdentPrognavio_Validate
    
    'Verifca se o código da prognavio foi informada
     If Len(Trim(MaskIdentProgNavio.Text)) <> 0 Then
    
        'Critica o código
        lErro = Long_Critica(MaskIdentProgNavio.Text)
        If lErro <> SUCESSO Then gError 98155
        
        'Carrega o obj com o valor do código
        objProgNavio.lCodigo = StrParaLong(MaskIdentProgNavio.Text)
            
        'Verifica se existe a programação do navio com este código
        lErro = CF("ProgNavio_Le", objProgNavio)
        If lErro <> SUCESSO And lErro <> 96657 Then gError 98275
            
        'Se não existe --> Erro
        If lErro = 96657 Then gError 98276
        
        LabelNavio.Caption = objProgNavio.sNavio
        
    Else
    
        LabelNavio.Caption = ""
        
    End If
    
    Exit Sub

Erro_MaskIdentPrognavio_Validate:

    Cancel = True

    Select Case gErr

        Case 98155, 98275
        
        Case 98276
            'Envia aviso que a Programação de Navio não está cadastrada e pergunta se deseja criar
            lErro = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PROGNAVIO", objProgNavio.lCodigo)
    
                If lErro = vbYes Then
                    'Chama tela de ProgNavio
                    lErro = Chama_Tela("ProgNavio", objProgNavio)
                End If
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub MaskBooking_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskPorto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub MaskDataInicio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub maskDataInicio_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskDataInicio, iAlterado)

End Sub

Private Sub maskDataInicio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_maskDataInicio_Validate

    'Verifica se a data de Inicio foi digitada
    If Len(Trim(MaskDataInicio.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(MaskDataInicio.Text)
    If lErro <> SUCESSO Then gError 98156
    
    Exit Sub

Erro_maskDataInicio_Validate:

    Cancel = True

    Select Case gErr

        Case 98156
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDown3_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub UpDown3_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown3_DownClick

    'Diminui um dia em DataInicio
    lErro = Data_Up_Down_Click(MaskDataInicio, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 98157

    Exit Sub

Erro_UpDown3_DownClick:

    Select Case gErr

        Case 98157
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDown3_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown3_UpClick

    'Aumenta um dia em DataInicio
    lErro = Data_Up_Down_Click(MaskDataInicio, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 98158

    Exit Sub

Erro_UpDown3_UpClick:

    Select Case gErr

        Case 98158
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub MaskDataFim_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub maskDataFim_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskDataFim, iAlterado)

End Sub

Private Sub maskDataFim_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskDataFim_Validate

    'Verifica se a data de Fim foi digitada
    If Len(Trim(MaskDataFim.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(MaskDataFim.Text)
    If lErro <> SUCESSO Then gError 98159
    
    Exit Sub

Erro_MaskDataFim_Validate:

    Cancel = True

    Select Case gErr

        Case 98159
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDown4_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub UpDown4_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown4_DownClick

    'Diminui um dia em DataFim
    lErro = Data_Up_Down_Click(MaskDataFim, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 98160

    Exit Sub

Erro_UpDown4_DownClick:

    Select Case gErr

        Case 98160
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub UpDown4_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown4_UpClick

    'Aumenta um dia em DataFim
    lErro = Data_Up_Down_Click(MaskDataFim, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 98161

    Exit Sub

Erro_UpDown4_UpClick:

    Select Case gErr

        Case 98161
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub TextObs_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub MaskHoraInicio_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub maskHoraInicio_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskHoraInicio, iAlterado)

End Sub

Private Sub maskHoraInicio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskHoraInicio_Validate
    
    'Se a HoraInicio foi preenchida...
    If Len(Trim(MaskHoraInicio.ClipText)) > 0 Then
        
        'Verifica se é válida
        lErro = Hora_Critica(MaskHoraInicio.Text)
        If lErro <> AD_SQL_SUCESSO Then gError 98162
        
    End If
        
    Exit Sub
    
Erro_MaskHoraInicio_Validate:
            
    Cancel = True
    
    Select Case gErr

        Case 98162
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
  
End Sub

Private Sub MaskHoraFim_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub maskHoraFim_GotFocus()

    Call MaskEdBox_TrataGotFocus(MaskHoraFim, iAlterado)

End Sub

Private Sub maskHoraFim_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MaskHoraFim_Validate
    
    'Se a HoraFim foi preenchida...
    If Len(Trim(MaskHoraFim.ClipText)) > 0 Then
        
        'Verifica se é válida
        lErro = Hora_Critica(MaskHoraFim.Text)
        If lErro <> AD_SQL_SUCESSO Then gError 98163
        
    End If
        
    Exit Sub
    
Erro_MaskHoraFim_Validate:
            
    Cancel = True
    
    Select Case gErr

        Case 98163
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
  
End Sub

Private Sub TextEndereco_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub MaskBairro_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub MaskCEP_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub MaskCEP_GotFocus(Index As Integer)

    Call MaskEdBox_TrataGotFocus(MaskCEP(Index), iAlterado)

End Sub

Private Sub MaskTelefone1_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub MaskTelefone2_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub MaskFax_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub MaskEmail_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub MaskContato_Change(Index As Integer)

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub GridServicos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridServicoss, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridServicoss, iAlterado)
    End If

End Sub

Private Sub GridServicos_EnterCell()

    Call Grid_Entrada_Celula(objGridServicoss, iAlterado)

End Sub

Private Sub GridServicos_GotFocus()

    Call Grid_Recebe_Foco(objGridServicoss)

End Sub

Private Sub GridServicos_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridServicoss)

End Sub

Private Sub GridServicos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridServicoss, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridServicoss, iAlterado)
    End If

End Sub

Private Sub GridServicos_LeaveCell()

    Call Saida_Celula(objGridServicoss)

End Sub

Public Sub GridServicos_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridServicoss)

End Sub

Public Sub GridServicos_RowColChange()

    Call Grid_RowColChange(objGridServicoss)

End Sub

Public Sub GridServicos_Scroll()

    Call Grid_Scroll(objGridServicoss)

End Sub

Private Sub MaskProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub maskProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridServicoss)

End Sub

Private Sub maskProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridServicoss)

End Sub

Private Sub maskProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridServicoss.objControle = MaskProduto
    lErro = Grid_Campo_Libera_Foco(objGridServicoss)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub MaskQuantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub maskQuantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridServicoss)

End Sub

Private Sub maskQuantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridServicoss)

End Sub

Private Sub maskQuantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridServicoss.objControle = MaskQuantidade
    lErro = Grid_Campo_Libera_Foco(objGridServicoss)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

On Error GoTo Erro_Rotina_Grid_Enable

    'Pesquisa a controle da coluna em questão
    Select Case objControl.Name
          
    'Cyntia
    'Código do Produto
    Case MaskProduto.Name
        
        If Len(Trim(GridServicos.TextMatrix(GridServicos.Row, iGrid_Produto_Col))) = 0 Then
            MaskProduto.Enabled = True
        Else
           MaskProduto.Enabled = False
        End If
    Case MaskQuantidade.Name
        
        If Len(Trim(GridServicos.TextMatrix(GridServicos.Row, iGrid_Produto_Col))) = 0 Then
            MaskQuantidade.Enabled = False
        Else
           MaskQuantidade.Enabled = True
        End If

    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

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
            
            'Se for o Produto...
            Case iGrid_Produto_Col
                lErro = Saida_Celula_MaskProduto(objGridInt)
                If lErro <> SUCESSO Then gError 97041
            
            'Se for a quantidade...
            Case iGrid_Quantidade_Col
                lErro = Saida_Celula_MaskQuantidade(objGridInt)
                If lErro <> SUCESSO Then gError 97042
            
         End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 97043

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 97041 To 97043
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MaskQuantidade(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidade que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_MaskQuantidade

    Set objGridInt.objControle = MaskQuantidade
    
    'Verifica se a quantidade foi informada
    If Len(MaskQuantidade.Text) > 0 Then
        
        'Veirfica se é um valor positivo
        lErro = Valor_Positivo_Critica(MaskQuantidade.Text)
        If lErro <> SUCESSO Then gError 98128
        
        'Joga a quantidade no grid com seu formato
        MaskQuantidade.Text = Formata_Estoque(MaskQuantidade.Text)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 98129

    Saida_Celula_MaskQuantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_MaskQuantidade:

    Saida_Celula_MaskQuantidade = gErr

    Select Case gErr

        Case 98128, 98129
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MaskProduto(objGridServicos As AdmGrid) As Long
'Faz a crítica da célula MaskProduto que está deixando de ser a corrente

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim sProdutoEnxuto As String
Dim iIndice As Integer
Dim objSolicitacaoServico As New ClassSolicitacaoServico
Dim sProduto As String

On Error GoTo Erro_Saida_Celula_MaskProduto

    Set objGridServicos.objControle = MaskProduto
    
    'Verifica se o produto foi preenchido
    If Len(Trim(MaskProduto.ClipText)) > 0 Then
        
        'Verifica se os campos que serão necessários para a verificação
        'da validade do produto estão preenchidos, senão--> erro
        If Len(Trim(MaskCodTabPreco.ClipText)) = 0 Then gError 98296
        If Len(Trim(MaskDataEmissao.ClipText)) = 0 Then gError 98297
        
        'Preenche o obj com os campos para a verificação
        objSolicitacaoServico.lCodTabPreco = StrParaLong(MaskCodTabPreco.Text)
        objSolicitacaoServico.dtDataEmissao = StrParaDate(MaskDataEmissao.Text)
                
        'Faz o Tratamento do produto
        lErro = Traz_Produto_Tela(objProduto)
        If lErro <> SUCESSO Then gError 98130
        
        'Faz a verificação da validade do produto
        lErro = CF("TabPrecoServico_Le1", objSolicitacaoServico, objProduto.sCodigo)
        If lErro <> SUCESSO And lErro <> 98292 Then gError 98298
            
        'Se o produto não está associado a tabela de preço passada --> erro
        If lErro = 98292 Then gError 98299
        
        'Atualiza o total de linhas existentes no grid
        If GridServicos.Row - GridServicos.FixedRows = objGridServicoss.iLinhasExistentes Then
            objGridServicoss.iLinhasExistentes = objGridServicoss.iLinhasExistentes + 1
        End If
        
        'Joga a descrição do produto no grid
        GridServicos.TextMatrix(GridServicos.Row, iGrid_DescProduto_Col) = objProduto.sDescricao
    
    Else
        'se o produto não foi preenchido --> limpa a descrição e a quantidade
        GridServicos.TextMatrix(GridServicos.Row, iGrid_DescProduto_Col) = ""
            
    End If
    
    lErro = Grid_Abandona_Celula(objGridServicos)
    If lErro <> SUCESSO Then gError 98132

    Saida_Celula_MaskProduto = SUCESSO

    Exit Function

Erro_Saida_Celula_MaskProduto:

    Saida_Celula_MaskProduto = gErr

    Select Case gErr

        Case 98130, 98132, 98298
            Call Grid_Trata_Erro_Saida_Celula(objGridServicos)

        Case 98296
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABPRECO_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridServicos)
            
        Case 98297
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridServicos)
               
        Case 98299
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_ASSOCIADO_TABPRECO", gErr, MaskProduto.Text, objSolicitacaoServico.lCodTabPreco)
            Call Grid_Trata_Erro_Saida_Celula(objGridServicos)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

'Parei aqui. Mario
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objSolicitacaoServico As New ClassSolicitacaoServico

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "SolicitacaoServico"

    'Move os dados da tela para a memória
    lErro = Move_Tela_Memoria(objSolicitacaoServico)
    If lErro <> SUCESSO Then gError 98073
    
    'Preenche a coleção colCampoValor
    colCampoValor.Add "Numero", objSolicitacaoServico.lNumero, 0, "Numero"
    colCampoValor.Add "CodTabPreco", objSolicitacaoServico.lCodTabPreco, 0, "CodTabPreco"
    colCampoValor.Add "Cliente", objSolicitacaoServico.lCliente, 0, "Cliente"
    colCampoValor.Add "DataEmissao", objSolicitacaoServico.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "NumReferencia", objSolicitacaoServico.sNumReferencia, STRING_SOLICITACAOSERVICO_NUMREFERENCIA, "NumReferencia"
    colCampoValor.Add "DataPedido", objSolicitacaoServico.dtDataPedido, 0, "DataPedido"
    colCampoValor.Add "TipoOperacao", objSolicitacaoServico.iTipoOperacao, 0, "TipoOperacao"
    colCampoValor.Add "Despachante", objSolicitacaoServico.iDespachante, 0, "Despachante"
    colCampoValor.Add "Material", objSolicitacaoServico.sMaterial, STRING_SOLICITACAOSERVICO_MATERIAL, "Material"
    colCampoValor.Add "QuantMaterial", objSolicitacaoServico.dQuantMaterial, 0, "QuantMaterial"
    colCampoValor.Add "UM", objSolicitacaoServico.sUM, STRING_SOLICITACAOSERVICO_UM, "UM"
    colCampoValor.Add "ValorMercadoria", objSolicitacaoServico.dValorMercadoria, 0, "ValorMercadoria"
    colCampoValor.Add "TipoEmbalagem", objSolicitacaoServico.iTipoEmbalagem, 0, "TipoEmbalagem"
    colCampoValor.Add "TipoContainer", objSolicitacaoServico.iTipoContainer, 0, "TipoContainer"
    colCampoValor.Add "CodProgNavio", objSolicitacaoServico.lCodProgNavio, 0, "CodProgNavio"
    colCampoValor.Add "Booking", objSolicitacaoServico.sBooking, STRING_SOLICITACAOSERVICO_BOOKING, "Booking"
    colCampoValor.Add "DataPrevInicio", objSolicitacaoServico.dtDataPrevInicio, 0, "DataPrevInicio"
    'colCampoValor.Add "HoraPrevInicio", objSolicitacaoServico.dtHoraPrevInicio, 0, "HoraPrevInicio"
    colCampoValor.Add "DataPrevFim", objSolicitacaoServico.dtDataPrevFim, 0, "DataPrevFim"
    'colCampoValor.Add "HoraPrevFim", objSolicitacaoServico.dtHoraPrevFim, 0, "HoraPrevFim"
    colCampoValor.Add "Observacao", objSolicitacaoServico.sObservacao, STRING_SOLICITACAOSERVICO_OBSERVACAO, "Observacao"
    colCampoValor.Add "EnderecoOrigem", objSolicitacaoServico.lEnderecoOrigem, 0, "EnderecoOrigem"
    colCampoValor.Add "EnderecoDestino", objSolicitacaoServico.lEnderecoDestino, 0, "EnderecoDestino"
    
    'Utilizado na hora de passar o parâmetro FilialEmpresa para o browser SolicitacaoLista
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr
        
        Case 98073
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objSolicitacaoServico As New ClassSolicitacaoServico

On Error GoTo Erro_Tela_Preenche
        
    'Carrega o obj só com os dados do número e filial
    objSolicitacaoServico.lNumero = colCampoValor.Item("Numero").vValor
    objSolicitacaoServico.iFilialEmpresa = giFilialEmpresa
    
    'Se o número foi informado
    If objSolicitacaoServico.lNumero <> 0 Then

        'Move os dados para a tela
        lErro = Traz_SolicitacaoServico_Tela(objSolicitacaoServico)
        If lErro <> SUCESSO And lErro <> 98087 Then gError 98073
        
        'Não achou a Solicitação de Serviço --> Erro
        If lErro = 98087 Then gError 98074
        
    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 98073
        
        Case 98074
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SOLICITACAO_NAO_CADASTRADA", gErr, objSolicitacaoServico.iFilialEmpresa, objSolicitacaoServico.lNumero)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()

    'Se o Frame atual não corresponde ao TAB clicado
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then
    
        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
        
        'Torna Frame selecionado visível
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        
        'Torna Frame atual invisível
        Frame1(iFrameAtual).Visible = False
        
        'Armazena novo valor de iFramePagBaixaAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
    
    End If

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera as variáveis globais
    Set objEventoTabPreco = Nothing
    Set objEventoSolicitacao = Nothing
    Set objEventoDespachante = Nothing
    Set objEventoServico = Nothing
    Set objEventoProgNavio = Nothing
    
    'Libera o comando de setas
    Call ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_Activate()

   Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        
        'Clique em F2
        Case KEYCODE_PROXIMO_NUMERO
            Call BotaoProxNum_Click
            
        'Clique em F3
        Case KEYCODE_BROWSER
            If Me.ActiveControl Is MaskNumero Then Call LabelNumero_Click
            If Me.ActiveControl Is MaskCodTabPreco Then Call LabelCodTabPreco_Click
            If Me.ActiveControl Is TextDespachante Then Call LabelDespachante_Click
            If Me.ActiveControl Is MaskIdentProgNavio Then Call LabelIdentProgNavio_Click
            If Me.ActiveControl Is MaskProduto Then Call BotaoServicos_Click
            
    End Select

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    ??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Solicitação de Serviço"
    Call Form_Load

End Function

Public Function Name() As String
    
    Name = "SolicitacaoServico"

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
'''    m_Caption = New_Caption
End Property
'***** fim do trecho a ser copiado ******
    
