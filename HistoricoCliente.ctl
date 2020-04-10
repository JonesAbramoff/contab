VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl HistoricoClienteOcx 
   ClientHeight    =   6570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5745
      Index           =   2
      Left            =   90
      TabIndex        =   57
      Top             =   765
      Visible         =   0   'False
      Width           =   9360
      Begin VB.ComboBox Status 
         Height          =   315
         Left            =   5985
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   3630
         Width           =   3210
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
         Left            =   7380
         TabIndex        =   53
         Top             =   4035
         Width           =   1800
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
         Height          =   375
         Left            =   7380
         TabIndex        =   54
         Top             =   4515
         Width           =   1800
      End
      Begin VB.TextBox Assunto 
         Height          =   930
         Left            =   45
         MaxLength       =   510
         MultiLine       =   -1  'True
         TabIndex        =   52
         Top             =   4695
         Width           =   7200
      End
      Begin MSMask.MaskEdBox DataPrevGrid 
         Height          =   225
         Left            =   6840
         TabIndex        =   69
         Top             =   2415
         Width           =   1485
         _ExtentX        =   2619
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
         Left            =   6255
         TabIndex        =   68
         Top             =   1740
         Width           =   1425
         _ExtentX        =   2514
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
      Begin MSMask.MaskEdBox ValorEmAbertoGrid 
         Height          =   225
         Left            =   5235
         TabIndex        =   67
         Top             =   1680
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.TextBox DiasAtrasoGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   7350
         TabIndex        =   66
         Top             =   1140
         Width           =   990
      End
      Begin VB.TextBox DataBaixaGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   6210
         TabIndex        =   65
         Top             =   1155
         Width           =   1080
      End
      Begin MSMask.MaskEdBox DataVenctoGrid 
         Height          =   225
         Left            =   5070
         TabIndex        =   63
         Top             =   1140
         Width           =   1065
         _ExtentX        =   1879
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
         Left            =   4140
         TabIndex        =   64
         Top             =   1170
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox ParcelaGrid 
         Height          =   225
         Left            =   2025
         TabIndex        =   59
         Top             =   1110
         Width           =   825
         _ExtentX        =   1455
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
      Begin MSMask.MaskEdBox NumTituloGrid 
         Height          =   225
         Left            =   1140
         TabIndex        =   60
         Top             =   1095
         Width           =   1020
         _ExtentX        =   1799
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
      Begin VB.TextBox DataEmissaoGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   2985
         TabIndex        =   62
         Top             =   1140
         Width           =   1080
      End
      Begin VB.CommandButton BotaoDocOriginal 
         Height          =   690
         Left            =   7410
         Picture         =   "HistoricoCliente.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   4980
         Width           =   1740
      End
      Begin VB.TextBox TipoGrid 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   315
         TabIndex        =   58
         Top             =   1095
         Width           =   870
      End
      Begin VB.ComboBox Ordenacao 
         Height          =   315
         ItemData        =   "HistoricoCliente.ctx":2F16
         Left            =   1125
         List            =   "HistoricoCliente.ctx":2F20
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   30
         Width           =   2910
      End
      Begin MSFlexGridLib.MSFlexGrid GridParcelas 
         Height          =   2880
         Left            =   60
         TabIndex        =   46
         Top             =   390
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   5080
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin MSComCtl2.UpDown UpDownDataPrev 
         Height          =   300
         Left            =   3375
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   4080
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataPrev 
         Height          =   300
         Left            =   2415
         TabIndex        =   48
         ToolTipText     =   "Informe a data prevista para o recebimento."
         Top             =   4080
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
         Left            =   6945
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   4080
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataProx 
         Height          =   300
         Left            =   5985
         TabIndex        =   50
         ToolTipText     =   "Informe a data prevista para o recebimento."
         Top             =   4080
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Total:"
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
         Left            =   4215
         TabIndex        =   120
         Top             =   3300
         Width           =   1050
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
         Left            =   6765
         TabIndex        =   119
         Top             =   3300
         Width           =   1005
      End
      Begin VB.Label SaldoDevTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5280
         TabIndex        =   118
         Top             =   3270
         Width           =   1410
      End
      Begin VB.Label ValorDevTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7830
         TabIndex        =   117
         Top             =   3285
         Width           =   1410
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
         Left            =   5325
         TabIndex        =   116
         Top             =   3675
         Width           =   615
      End
      Begin VB.Label NumIntTitRec 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6570
         TabIndex        =   114
         Top             =   4395
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label NumIntParcRec 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5715
         TabIndex        =   113
         Top             =   4395
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label LabelParcela 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4620
         TabIndex        =   78
         Top             =   3645
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Parcela:"
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
         Left            =   3840
         TabIndex        =   77
         Top             =   3690
         Width           =   720
      End
      Begin VB.Label LabelTitulo 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2415
         TabIndex        =   76
         Top             =   3645
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Título:"
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
         Left            =   1770
         TabIndex        =   75
         Top             =   3690
         Width           =   585
      End
      Begin VB.Label LabelTipo 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   720
         TabIndex        =   74
         Top             =   3645
         Width           =   930
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
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   73
         Top             =   3690
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Próximo Contato:"
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
         Left            =   4020
         TabIndex        =   72
         Top             =   4140
         Width           =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Prevista Receb:"
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
         Left            =   510
         TabIndex        =   71
         Top             =   4140
         Width           =   1845
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
         Left            =   90
         TabIndex        =   70
         Top             =   4440
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ordenação:"
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
         Left            =   90
         TabIndex        =   61
         Top             =   90
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5745
      Index           =   1
      Left            =   90
      TabIndex        =   56
      Top             =   780
      Width           =   9360
      Begin VB.Frame Frame1 
         Caption         =   "Data Próximo Contato"
         Height          =   1575
         Index           =   4
         Left            =   735
         TabIndex        =   102
         Top             =   600
         Width           =   7485
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
            TabIndex        =   12
            Top             =   1110
            Width           =   780
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
            TabIndex        =   9
            Top             =   690
            Width           =   1050
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
            TabIndex        =   4
            Top             =   270
            Value           =   -1  'True
            Width           =   1620
         End
         Begin VB.Frame FrameD1 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   480
            Index           =   0
            Left            =   2295
            TabIndex        =   109
            Top             =   105
            Width           =   5130
            Begin MSMask.MaskEdBox DataProxAte 
               Height          =   300
               Left            =   2325
               TabIndex        =   7
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
               TabIndex        =   6
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
               TabIndex        =   5
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
               TabIndex        =   8
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
               Index           =   23
               Left            =   1920
               TabIndex        =   111
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
               Index           =   22
               Left            =   90
               TabIndex        =   110
               Top             =   165
               Width           =   315
            End
         End
         Begin VB.Frame FrameD1 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   405
            Index           =   1
            Left            =   1185
            TabIndex        =   107
            Top             =   585
            Width           =   6195
            Begin VB.ComboBox ApenasQualifProx 
               Height          =   315
               ItemData        =   "HistoricoCliente.ctx":2F3F
               Left            =   90
               List            =   "HistoricoCliente.ctx":2F49
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
               TabIndex        =   108
               Top             =   90
               Width           =   480
            End
         End
         Begin VB.Frame FrameD1 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   405
            Index           =   2
            Left            =   1230
            TabIndex        =   103
            Top             =   1065
            Width           =   6195
            Begin VB.ComboBox EntreQualifProxDe 
               Height          =   315
               ItemData        =   "HistoricoCliente.ctx":2F64
               Left            =   1515
               List            =   "HistoricoCliente.ctx":2F6E
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   0
               Width           =   1365
            End
            Begin VB.ComboBox EntreQualifProxAte 
               Height          =   315
               ItemData        =   "HistoricoCliente.ctx":2F83
               Left            =   4755
               List            =   "HistoricoCliente.ctx":2F8D
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   15
               Width           =   1365
            End
            Begin MSMask.MaskEdBox EntreDiasProxDe 
               Height          =   315
               Left            =   45
               TabIndex        =   13
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
               TabIndex        =   15
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
               Index           =   20
               Left            =   720
               TabIndex        =   106
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
               Index           =   19
               Left            =   4095
               TabIndex        =   105
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
               Index           =   12
               Left            =   3090
               TabIndex        =   104
               Top             =   45
               Width           =   120
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Vencimento"
         Height          =   1575
         Index           =   0
         Left            =   735
         TabIndex        =   92
         Top             =   2265
         Width           =   7485
         Begin VB.OptionButton FaixaDataVenc 
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
            TabIndex        =   17
            Top             =   285
            Value           =   -1  'True
            Width           =   1620
         End
         Begin VB.OptionButton ApenasVenc 
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
            TabIndex        =   22
            Top             =   690
            Width           =   1050
         End
         Begin VB.OptionButton EntreVenc 
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
            TabIndex        =   25
            Top             =   1110
            Width           =   780
         End
         Begin VB.Frame FrameD2 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   405
            Index           =   0
            Left            =   2400
            TabIndex        =   99
            Top             =   225
            Width           =   4965
            Begin MSMask.MaskEdBox DataVencAte 
               Height          =   300
               Left            =   2235
               TabIndex        =   20
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
            Begin MSComCtl2.UpDown UpDownDataVencDe 
               Height          =   300
               Left            =   1350
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   0
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataVencDe 
               Height          =   300
               Left            =   360
               TabIndex        =   18
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
            Begin MSComCtl2.UpDown UpDownDataVencAte 
               Height          =   300
               Left            =   3210
               TabIndex        =   21
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
               Index           =   10
               Left            =   0
               TabIndex        =   101
               Top             =   60
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
               Index           =   9
               Left            =   1830
               TabIndex        =   100
               Top             =   60
               Width           =   360
            End
         End
         Begin VB.Frame FrameD2 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   405
            Index           =   1
            Left            =   1275
            TabIndex        =   97
            Top             =   630
            Width           =   4965
            Begin VB.ComboBox ApenasQualifVenc 
               Height          =   315
               ItemData        =   "HistoricoCliente.ctx":2FA2
               Left            =   0
               List            =   "HistoricoCliente.ctx":2FAC
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   0
               Width           =   1365
            End
            Begin MSMask.MaskEdBox ApenasDiasVenc 
               Height          =   315
               Left            =   1485
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
               Index           =   11
               Left            =   2175
               TabIndex        =   98
               Top             =   45
               Width           =   480
            End
         End
         Begin VB.Frame FrameD2 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   405
            Index           =   2
            Left            =   1260
            TabIndex        =   93
            Top             =   1050
            Width           =   6120
            Begin VB.ComboBox EntreQualifVencAte 
               Height          =   315
               ItemData        =   "HistoricoCliente.ctx":2FC7
               Left            =   4725
               List            =   "HistoricoCliente.ctx":2FD1
               Style           =   2  'Dropdown List
               TabIndex        =   29
               Top             =   15
               Width           =   1365
            End
            Begin VB.ComboBox EntreQualifVencDe 
               Height          =   315
               ItemData        =   "HistoricoCliente.ctx":2FE6
               Left            =   1485
               List            =   "HistoricoCliente.ctx":2FF0
               Style           =   2  'Dropdown List
               TabIndex        =   27
               Top             =   0
               Width           =   1365
            End
            Begin MSMask.MaskEdBox EntreDiasVencDe 
               Height          =   315
               Left            =   0
               TabIndex        =   26
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
            Begin MSMask.MaskEdBox EntreDiasVencAte 
               Height          =   315
               Left            =   3360
               TabIndex        =   28
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
               Index           =   8
               Left            =   3060
               TabIndex        =   96
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
               Index           =   6
               Left            =   4065
               TabIndex        =   95
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
               Index           =   7
               Left            =   675
               TabIndex        =   94
               Top             =   45
               Width           =   480
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Previsão Recebimento"
         Height          =   1575
         Index           =   3
         Left            =   720
         TabIndex        =   82
         Top             =   3900
         Width           =   7485
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
            TabIndex        =   30
            Top             =   270
            Value           =   -1  'True
            Width           =   1620
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
            TabIndex        =   35
            Top             =   690
            Width           =   1050
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
            TabIndex        =   38
            Top             =   1110
            Width           =   780
         End
         Begin VB.Frame FrameD3 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   405
            Index           =   0
            Left            =   2445
            TabIndex        =   89
            Top             =   225
            Width           =   4965
            Begin MSMask.MaskEdBox DataPrevAte 
               Height          =   300
               Left            =   2235
               TabIndex        =   33
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
               TabIndex        =   32
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
               TabIndex        =   31
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
               TabIndex        =   34
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
               TabIndex        =   91
               Top             =   60
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
               Index           =   18
               Left            =   1830
               TabIndex        =   90
               Top             =   60
               Width           =   360
            End
         End
         Begin VB.Frame FrameD3 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   405
            Index           =   1
            Left            =   1275
            TabIndex        =   87
            Top             =   630
            Width           =   4965
            Begin VB.ComboBox ApenasQualifPrev 
               Height          =   315
               ItemData        =   "HistoricoCliente.ctx":3005
               Left            =   0
               List            =   "HistoricoCliente.ctx":300F
               Style           =   2  'Dropdown List
               TabIndex        =   36
               Top             =   0
               Width           =   1365
            End
            Begin MSMask.MaskEdBox ApenasDiasPrev 
               Height          =   315
               Left            =   1530
               TabIndex        =   37
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
               TabIndex        =   88
               Top             =   45
               Width           =   480
            End
         End
         Begin VB.Frame FrameD3 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   405
            Index           =   2
            Left            =   1275
            TabIndex        =   83
            Top             =   1035
            Width           =   6105
            Begin VB.ComboBox EntreQualifPrevAte 
               Height          =   315
               ItemData        =   "HistoricoCliente.ctx":302A
               Left            =   4740
               List            =   "HistoricoCliente.ctx":3034
               Style           =   2  'Dropdown List
               TabIndex        =   42
               Top             =   30
               Width           =   1365
            End
            Begin VB.ComboBox EntreQualifPrevDe 
               Height          =   315
               ItemData        =   "HistoricoCliente.ctx":3049
               Left            =   1530
               List            =   "HistoricoCliente.ctx":3053
               Style           =   2  'Dropdown List
               TabIndex        =   40
               Top             =   15
               Width           =   1365
            End
            Begin MSMask.MaskEdBox EntreDiasPrevDe 
               Height          =   315
               Left            =   0
               TabIndex        =   39
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
               TabIndex        =   41
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
               Index           =   13
               Left            =   3105
               TabIndex        =   86
               Top             =   60
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
               Index           =   14
               Left            =   4110
               TabIndex        =   85
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
               Index           =   15
               Left            =   705
               TabIndex        =   84
               Top             =   60
               Width           =   480
            End
         End
      End
      Begin VB.CheckBox CheckTitAberto 
         Caption         =   "Só exibir títulos em aberto"
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
         Left            =   735
         TabIndex        =   3
         Top             =   210
         Width           =   2730
      End
   End
   Begin VB.ComboBox Contato 
      Height          =   315
      Left            =   6285
      TabIndex        =   2
      ToolTipText     =   $"HistoricoCliente.ctx":3068
      Top             =   45
      Width           =   1950
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   8370
      ScaleHeight     =   450
      ScaleWidth      =   1020
      TabIndex        =   112
      TabStop         =   0   'False
      Top             =   45
      Width           =   1080
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   60
         Picture         =   "HistoricoCliente.ctx":30F0
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   540
         Picture         =   "HistoricoCliente.ctx":3622
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin VB.TextBox Cliente 
      Height          =   315
      Left            =   795
      TabIndex        =   0
      ToolTipText     =   "Digite código, nome reduzido, cgc do cliente ou pressione F3 para consulta."
      Top             =   30
      Width           =   2175
   End
   Begin VB.ComboBox Filial 
      Height          =   315
      Left            =   3870
      TabIndex        =   1
      ToolTipText     =   "Digite o nome ou o código da filial do cliente com quem foi feito o relacionamento."
      Top             =   45
      Width           =   1380
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6165
      Left            =   45
      TabIndex        =   81
      Top             =   390
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   10874
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Histórico de Recebimentos"
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
   Begin VB.Label LabelContato 
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5490
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   115
      Top             =   105
      Width           =   735
   End
   Begin VB.Label LabelFilialCliente 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   3270
      TabIndex        =   80
      Top             =   105
      Width           =   465
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
      Left            =   75
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   79
      Top             =   90
      Width           =   660
   End
End
Attribute VB_Name = "HistoricoClienteOcx"
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

Dim lClienteAnt As Long
Dim iFilialAnt As Integer

Dim objGridParcelas As AdmGrid
Dim iGrid_Tipo_Col As Integer
Dim iGrid_Titulo_Col As Integer
Dim iGrid_Parcela_Col As Integer
Dim iGrid_Emissao_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_Vencimento_Col As Integer
Dim iGrid_Baixa_Col As Integer
Dim iGrid_Atraso_Col As Integer
Dim iGrid_AReceber_Col As Integer
Dim iGrid_DataProx_Col As Integer
Dim iGrid_DataPrev_Col As Integer

Dim iStatus_ListIndex_Padrao As Integer

Dim gobjHistCobrClienteAnt As ClassHistoricoCobrSelCli

Dim gcolParcelasRec As Collection
Dim gcolTitulosRec As Collection

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1

Const TAB_SELECAO = 1
Const TAB_Parcelas = 2

Const FRAMED_FAIXA = 0
Const FRAMED_APENAS = 1
Const FRAMED_ENTRE = 2

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Histórico de Cobranças do Cliente"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "HistoricoCliente"

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
        
        If Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        ElseIf Me.ActiveControl Is Contato Then
            Call LabelContato_Click
        End If
    
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
    
    Set objGridParcelas = Nothing
    Set objEventoCliente = Nothing
    Set gobjHistCobrClienteAnt = Nothing
    
    Set gcolParcelasRec = Nothing
    Set gcolTitulosRec = Nothing
    
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

On Error GoTo Erro_Form_Load

    Set objGridParcelas = New AdmGrid
    Set objEventoCliente = New AdmEvento
    Set gobjHistCobrClienteAnt = New ClassHistoricoCobrSelCli
    Set gcolParcelasRec = New Collection
    Set gcolTitulosRec = New Collection
    
    lErro = Inicializa_GridParcelas(objGridParcelas)
    If lErro <> SUCESSO Then gError 182211

    Call FrameD_Enabled(FrameD1, FRAMED_FAIXA)
    Call FrameD_Enabled(FrameD2, FRAMED_FAIXA)
    Call FrameD_Enabled(FrameD3, FRAMED_FAIXA)
    
    Call Carrega_Status(Status)
    
    iFrameAtual = TAB_SELECAO
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO
   
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 182211
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182212)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Move_Selecao_Memoria(ByVal objHistCobrSelCli As ClassHistoricoCobrSelCli) As Long

Dim lErro As Long
Dim dtDataDe As Date
Dim dtDataAte As Date
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer

On Error GoTo Erro_Move_Selecao_Memoria

    If CheckTitAberto.Value = vbUnchecked Then
        objHistCobrSelCli.iTitulosBaixados = DESMARCADO
    Else
        objHistCobrSelCli.iTitulosBaixados = MARCADO
    End If
        

    If FaixaDataProx.Value Then
        objHistCobrSelCli.dtDataProxDe = StrParaDate(DataProxDe.Text)
        objHistCobrSelCli.dtDataProxAte = StrParaDate(DataProxAte.Text)
    End If

    If ApenasProx.Value Then
    
        If ApenasQualifProx.ListIndex = -1 Then gError 182213
    
        Call Datas_Trata_Apenas(dtDataDe, dtDataAte, StrParaInt(ApenasDiasProx.Text), ApenasQualifProx.ItemData(ApenasQualifProx.ListIndex))
        
        objHistCobrSelCli.dtDataProxDe = dtDataDe
        objHistCobrSelCli.dtDataProxAte = dtDataAte
    End If
    
    If EntreProx.Value Then
        
        If EntreQualifProxAte.ListIndex = -1 Then gError 182214
        If EntreQualifProxAte.ListIndex = -1 Then gError 182215
        
        Call Datas_Trata_Entre(dtDataDe, dtDataAte, StrParaInt(EntreDiasProxDe), EntreQualifProxDe.ItemData(EntreQualifProxDe.ListIndex), StrParaInt(EntreDiasProxAte), EntreQualifProxAte.ItemData(EntreQualifProxAte.ListIndex))
        
        objHistCobrSelCli.dtDataProxDe = dtDataDe
        objHistCobrSelCli.dtDataProxAte = dtDataAte
    End If

    If FaixaDataVenc.Value Then
        objHistCobrSelCli.dtDataVencDe = StrParaDate(DataVencDe.Text)
        objHistCobrSelCli.dtDataVencAte = StrParaDate(DataVencAte.Text)
    End If

    If ApenasVenc.Value Then
    
        If ApenasQualifVenc.ListIndex = -1 Then gError 182216

        Call Datas_Trata_Apenas(dtDataDe, dtDataAte, StrParaInt(ApenasDiasVenc.Text), ApenasQualifVenc.ItemData(ApenasQualifVenc.ListIndex))
        
        objHistCobrSelCli.dtDataVencDe = dtDataDe
        objHistCobrSelCli.dtDataVencAte = dtDataAte
    End If
    
    If EntreVenc.Value Then
    
        If EntreQualifVencDe.ListIndex = -1 Then gError 182217
        If EntreQualifVencAte.ListIndex = -1 Then gError 182218

        Call Datas_Trata_Entre(dtDataDe, dtDataAte, StrParaInt(EntreDiasVencDe), EntreQualifVencDe.ItemData(EntreQualifVencDe.ListIndex), StrParaInt(EntreDiasVencAte), EntreQualifVencAte.ItemData(EntreQualifVencAte.ListIndex))
        
        objHistCobrSelCli.dtDataVencDe = dtDataDe
        objHistCobrSelCli.dtDataVencAte = dtDataAte
    End If
    
    If FaixaDataPrev.Value Then
        objHistCobrSelCli.dtDataPrevDe = StrParaDate(DataPrevDe.Text)
        objHistCobrSelCli.dtDataPrevAte = StrParaDate(DataPrevAte.Text)
    End If

    If ApenasPrev.Value Then
    
        If ApenasQualifPrev.ListIndex = -1 Then gError 182219

        Call Datas_Trata_Apenas(dtDataDe, dtDataAte, StrParaInt(ApenasDiasPrev.Text), ApenasQualifPrev.ItemData(ApenasQualifPrev.ListIndex))
        
        objHistCobrSelCli.dtDataPrevDe = dtDataDe
        objHistCobrSelCli.dtDataPrevAte = dtDataAte
    End If
    
    If EntrePrev.Value Then
        
        If EntreQualifPrevDe.ListIndex = -1 Then gError 182220
        If EntreQualifPrevAte.ListIndex = -1 Then gError 182221
        
        Call Datas_Trata_Entre(dtDataDe, dtDataAte, StrParaInt(EntreDiasPrevDe), EntreQualifPrevDe.ItemData(EntreQualifPrevDe.ListIndex), StrParaInt(EntreDiasPrevAte.Text), EntreQualifPrevAte.ItemData(EntreQualifPrevAte.ListIndex))
        
        objHistCobrSelCli.dtDataPrevDe = dtDataDe
        objHistCobrSelCli.dtDataPrevAte = dtDataAte
    End If
    
    If objHistCobrSelCli.dtDataPrevAte <> DATA_NULA And objHistCobrSelCli.dtDataPrevDe <> DATA_NULA Then
        If objHistCobrSelCli.dtDataPrevDe > objHistCobrSelCli.dtDataPrevAte Then gError 182222
    End If
    
    If objHistCobrSelCli.dtDataProxDe <> DATA_NULA And objHistCobrSelCli.dtDataProxAte <> DATA_NULA Then
        If objHistCobrSelCli.dtDataProxDe > objHistCobrSelCli.dtDataProxAte Then gError 182223
    End If
    
    If objHistCobrSelCli.dtDataVencDe <> DATA_NULA And objHistCobrSelCli.dtDataVencAte <> DATA_NULA Then
        If objHistCobrSelCli.dtDataVencDe > objHistCobrSelCli.dtDataVencAte Then gError 182224
    End If
    
    If iFrameAtual = TAB_SELECAO Then
    
        If Len(Trim(Cliente.Text)) = 0 Then gError 182225
        If Len(Trim(Filial.Text)) = 0 Then gError 182226
        
    End If
    
    If Len(Trim(Cliente.Text)) > 0 Then
    
        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then gError 182227

        objHistCobrSelCli.lCliente = objCliente.lCodigo
        objHistCobrSelCli.iFilial = Codigo_Extrai(Filial.Text)
        
    End If
   
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

Function Trata_Selecao(ByVal objHistCobrSelCli As ClassHistoricoCobrSelCli) As Long

Dim lErro As Long
Dim colParcelas As New Collection

On Error GoTo Erro_Trata_Selecao

    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = CF("HistoricoCobrCliente_Le", objHistCobrSelCli, colParcelas)
    If lErro <> SUCESSO Then gError 182234
        
    lErro = Preenche_GridParcelas(colParcelas)
    If lErro <> SUCESSO Then gError 182236
    
    If colParcelas.Count = 0 Then gError 182235

    GL_objMDIForm.MousePointer = vbDefault
   
    Trata_Selecao = SUCESSO

    Exit Function

Erro_Trata_Selecao:

    GL_objMDIForm.MousePointer = vbDefault

    Trata_Selecao = gErr

    Select Case gErr
    
        Case 182234, 182236
        
        Case 182235
            Call Rotina_Erro(vbOKOnly, "ERRO_SELECAO_HISTCOBRANCA_SEM_PARCELAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182237)

    End Select

    Exit Function

End Function

Function Preenche_GridParcelas(ByVal colParcelas As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objParcelasRec As ClassParcelaReceber
Dim objTituloRec As ClassTituloReceber
Dim colTituloRec As New Collection
Dim objParcRecBaixa As ClassBaixaParcRec
Dim colParcRecBaixa As New Collection
Dim colBaixas As New Collection
Dim objBaixaRec As ClassBaixaReceber
Dim iAtraso As Integer
Dim bAchou As Boolean
Dim dtDataBaixa As Date

On Error GoTo Erro_Preenche_GridParcelas

    Call Grid_Limpa(objGridParcelas)
    
    LabelTipo.Caption = ""
    LabelTitulo.Caption = ""
    LabelParcela.Caption = ""
    NumIntParcRec.Caption = ""
    
    'Aumenta o número de linhas do grid se necessário
    If colParcelas.Count >= objGridParcelas.objGrid.Rows Then
        Call Refaz_Grid(objGridParcelas, colParcelas.Count)
    End If

    iIndice = 0
    For Each objParcelasRec In colParcelas
    
        iIndice = iIndice + 1
        
        bAchou = False
        For Each objTituloRec In colTituloRec
        
            If objTituloRec.lNumIntDoc = objParcelasRec.lNumIntTitulo Then
                bAchou = True
                Exit For
            End If
        
        Next
        
        If Not bAchou Then
        
            Set objTituloRec = New ClassTituloReceber
            
            objTituloRec.lNumIntDoc = objParcelasRec.lNumIntTitulo
            
            lErro = CF("TituloReceber_Le", objTituloRec)
            If lErro <> SUCESSO And lErro <> 26061 Then gError 182237
            
            If lErro <> SUCESSO Then
            
                lErro = CF("TituloReceberBaixado_Le", objTituloRec)
                If lErro <> SUCESSO And lErro <> 56570 Then gError 182238
                
            End If
            
            colTituloRec.Add objTituloRec
        
        End If
        
        dtDataBaixa = DATA_NULA
        Set colParcRecBaixa = New Collection
        
        lErro = CF("BaixaParcRec_Le_Parcela", objParcelasRec.lNumIntDoc, colParcRecBaixa)
        If lErro <> SUCESSO Then gError 182239
        
        For Each objParcRecBaixa In colParcRecBaixa
        
            bAchou = False
            For Each objBaixaRec In colBaixas
            
                If objBaixaRec.lNumIntBaixa = objParcRecBaixa.lNumIntBaixa Then
                    bAchou = True
                    Exit For
                End If
            
            Next
            
            If Not bAchou Then
            
                Set objBaixaRec = New ClassBaixaReceber
            
                objBaixaRec.lNumIntBaixa = objParcRecBaixa.lNumIntBaixa
            
                lErro = CF("BaixaRec_Le", objBaixaRec)
                If lErro <> SUCESSO And lErro <> 46234 Then gError 182240
                
                colBaixas.Add objBaixaRec
            
            End If
            
            If dtDataBaixa < objBaixaRec.dtData Then
                dtDataBaixa = objBaixaRec.dtData
            End If
            
        Next
        
        GridParcelas.TextMatrix(iIndice, iGrid_AReceber_Col) = Format(objParcelasRec.dSaldo, "STANDARD")
        
        If objParcelasRec.iStatus <> STATUS_BAIXADO Then
            GridParcelas.TextMatrix(iIndice, iGrid_Baixa_Col) = ""
            iAtraso = DateDiff("d", objParcelasRec.dtDataVencimento, gdtDataAtual)
        Else
            GridParcelas.TextMatrix(iIndice, iGrid_Baixa_Col) = Format(dtDataBaixa, "dd/mm/yyyy")
            iAtraso = DateDiff("d", objParcelasRec.dtDataVencimento, dtDataBaixa)
        End If
        
        If iAtraso <= 0 Then
            GridParcelas.TextMatrix(iIndice, iGrid_Atraso_Col) = ""
        Else
            GridParcelas.TextMatrix(iIndice, iGrid_Atraso_Col) = iAtraso
        End If
        GridParcelas.TextMatrix(iIndice, iGrid_Emissao_Col) = Format(objTituloRec.dtDataEmissao, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iIndice, iGrid_Parcela_Col) = objParcelasRec.iNumParcela
        GridParcelas.TextMatrix(iIndice, iGrid_Tipo_Col) = objTituloRec.sSiglaDocumento
        GridParcelas.TextMatrix(iIndice, iGrid_Titulo_Col) = objTituloRec.lNumTitulo
        GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col) = Format(objParcelasRec.dValor, "STANDARD")
        GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col) = Format(objParcelasRec.dtDataVencimento, "dd/mm/yyyy")

        If objParcelasRec.dtDataPrevReceb <> DATA_NULA Then
            GridParcelas.TextMatrix(iIndice, iGrid_DataPrev_Col) = Format(objParcelasRec.dtDataPrevReceb, "dd/mm/yyyy")
        End If
        
        If objParcelasRec.dtDataProxCobr <> DATA_NULA Then
            GridParcelas.TextMatrix(iIndice, iGrid_DataProx_Col) = Format(objParcelasRec.dtDataProxCobr, "dd/mm/yyyy")
        End If

    Next
        
    objGridParcelas.iLinhasExistentes = iIndice
    
    Call Grid_Refresh_Checkbox(objGridParcelas)
    
    Set gcolParcelasRec = colParcelas
    Set gcolTitulosRec = colTituloRec
    
    Call Ordenacao_Limpa(objGridParcelas, Ordenacao)
    
    Call Combo_Seleciona_ItemData(Ordenacao, -iGrid_Vencimento_Col)
    
    Call Soma_Coluna_Grid(objGridParcelas, iGrid_Valor_Col, ValorDevTotal, False)
    Call Soma_Coluna_Grid(objGridParcelas, iGrid_AReceber_Col, SaldoDevTotal, False)
   
    Preenche_GridParcelas = SUCESSO

    Exit Function

Erro_Preenche_GridParcelas:

    Preenche_GridParcelas = gErr

    Select Case gErr
    
        Case 182237 To 182240

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182241)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional ByVal objHistCobrSelCli As ClassHistoricoCobrSelCli) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    bAbrindoTela = True
    If Not (objHistCobrSelCli Is Nothing) Then
    
        'Torna Frame atual invisível
        Frame1(TabStrip1.SelectedItem.Index).Visible = False
        iFrameAtual = TAB_Parcelas
        'Torna Frame atual visível
        Frame1(iFrameAtual).Visible = True
        TabStrip1.Tabs.Item(iFrameAtual).Selected = True
        
        'CheckTitAberto.Value = vbChecked
        
        Cliente.Text = objHistCobrSelCli.lCliente
        Call Cliente_Validate(bSGECancelDummy)
        
        Filial.Text = objHistCobrSelCli.iFilial
        Call Filial_Validate(bSGECancelDummy)
        
        If objHistCobrSelCli.iContato <> 0 Then
            Contato.Text = objHistCobrSelCli.iContato
            Call Contato_Validate(bSGECancelDummy)
        End If
        
        Call Trata_TabClick
        
    End If
    
    bAbrindoTela = False
    
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

    lErro = Obtem_CodCliente(lCliente)
    If lErro <> SUCESSO Then gError 182280

    objRelacCli.lCliente = lCliente
    objRelacCli.iFilialCliente = Codigo_Extrai(Filial.Text)
    objRelacCli.dtData = gdtDataAtual
    objRelacCli.dtDataPrevReceb = StrParaDate(DataPrev.Text)
    objRelacCli.dtDataProxCobr = StrParaDate(DataProx.Text)
    objRelacCli.iFilialEmpresa = giFilialEmpresa
    'Guarda no obj a primeira parte do assunto
    objRelacCli.sAssunto1 = left(Assunto.Text, STRING_BUFFER_MAX_TEXTO - 1)
    'Guarda no obj a segunda parte do assunto
    objRelacCli.sAssunto2 = Mid(Assunto.Text, STRING_BUFFER_MAX_TEXTO)
    objRelacCli.lNumIntParcRec = StrParaLong(NumIntParcRec.Caption)
    objRelacCli.iOrigem = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA
    objRelacCli.dtHora = Time
    objRelacCli.iContato = Codigo_Extrai(Contato.Text)
    objRelacCli.lTipo = TIPO_RELACIONAMENTO_COBRANCA
    
    If Status.ListIndex <> -1 Then
        objRelacCli.iStatusCG = Status.ItemData(Status.ListIndex)
    Else
        objRelacCli.iStatusCG = 0
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
Dim objHistCobrCliente As New ClassHistoricoCobrSelCli

On Error GoTo Erro_Trata_TabClick

    'Valida a seleção
    lErro = Move_Selecao_Memoria(objHistCobrCliente)
    If lErro <> SUCESSO Then gError 182244
    
    If (objHistCobrCliente.dtDataPrevAte <> gobjHistCobrClienteAnt.dtDataPrevAte Or _
        objHistCobrCliente.dtDataPrevDe <> gobjHistCobrClienteAnt.dtDataPrevDe Or _
        objHistCobrCliente.dtDataProxAte <> gobjHistCobrClienteAnt.dtDataProxAte Or _
        objHistCobrCliente.dtDataProxDe <> gobjHistCobrClienteAnt.dtDataProxDe Or _
        objHistCobrCliente.dtDataVencAte <> gobjHistCobrClienteAnt.dtDataVencAte Or _
        objHistCobrCliente.dtDataVencDe <> gobjHistCobrClienteAnt.dtDataVencDe Or _
        objHistCobrCliente.lCliente <> gobjHistCobrClienteAnt.lCliente Or _
        objHistCobrCliente.iFilial <> gobjHistCobrClienteAnt.iFilial Or _
        objHistCobrCliente.iTitulosBaixados <> gobjHistCobrClienteAnt.iTitulosBaixados) Or _
        bAtualiza Then
    
        lErro = Trata_Selecao(objHistCobrCliente)
        If lErro <> SUCESSO Then
            Set gobjHistCobrClienteAnt = New ClassHistoricoCobrSelCli
            gError 182245
        End If
        
        Set gobjHistCobrClienteAnt = objHistCobrCliente
        
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
Dim objParcelaRec As ClassParcelaReceber
Dim objTituloRec As ClassTituloReceber

On Error GoTo Erro_Trata_RowColChange

    If GridParcelas.Row <> 0 Then
    
        'Se não é na abertura da tela
        If gcolParcelasRec.Count <> 0 Then
    
            Set objParcelaRec = gcolParcelasRec.Item(GridParcelas.Row)
            
            For Each objTituloRec In gcolTitulosRec
            
                If objTituloRec.lNumIntDoc = objParcelaRec.lNumIntTitulo Then
                    Exit For
                End If
            
            Next
            
            'Se mudou a parcela
            If StrParaLong(NumIntParcRec.Caption) <> objParcelaRec.lNumIntDoc Then
            
                'Testa se deseja salvar mudanças
                lErro = Teste_Salva(Me, iAlterado)
                If lErro <> SUCESSO Then gError 182285
                                
                Call Limpa_Relacionamento
            
                LabelTipo.Caption = objTituloRec.sSiglaDocumento
                LabelTitulo.Caption = objTituloRec.lNumTitulo
                LabelParcela.Caption = objParcelaRec.iNumParcela
                NumIntParcRec.Caption = objParcelaRec.lNumIntDoc
                NumIntTitRec.Caption = objParcelaRec.lNumIntTitulo
                
                iAlterado = 0
                
            End If
            
        End If
    
    End If
   
    Trata_RowColChange = SUCESSO

    Exit Function

Erro_Trata_RowColChange:

    Trata_RowColChange = gErr

    Select Case gErr
    
        Case 182285

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
    If iFrameAtual = TAB_SELECAO Then
    
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
        
    End If

End Sub

Private Function Inicializa_GridParcelas(objGrid As AdmGrid) As Long

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Tipo")
    objGrid.colColuna.Add ("Título")
    objGrid.colColuna.Add ("Parcela")
    objGrid.colColuna.Add ("Emissão")
    objGrid.colColuna.Add ("Valor")
    objGrid.colColuna.Add ("Vencimento")
    objGrid.colColuna.Add ("Baixa")
    objGrid.colColuna.Add ("Atraso")
    objGrid.colColuna.Add ("Saldo")
    objGrid.colColuna.Add ("Próximo Contato")
    objGrid.colColuna.Add ("Previsão Receb")
    
    'Atualiza a Parte de Ordenação
    Call Ordenacao_Preeenche(objGrid, Ordenacao)

    'Controles que participam do Grid
    objGrid.colCampo.Add (TipoGrid.Name)
    objGrid.colCampo.Add (NumTituloGrid.Name)
    objGrid.colCampo.Add (ParcelaGrid.Name)
    objGrid.colCampo.Add (DataEmissaoGrid.Name)
    objGrid.colCampo.Add (ValorGrid.Name)
    objGrid.colCampo.Add (DataVenctoGrid.Name)
    objGrid.colCampo.Add (DataBaixaGrid.Name)
    objGrid.colCampo.Add (DiasAtrasoGrid.Name)
    objGrid.colCampo.Add (ValorEmAbertoGrid.Name)
    objGrid.colCampo.Add (DataProxGrid.Name)
    objGrid.colCampo.Add (DataPrevGrid.Name)

    'Colunas do Grid
    iGrid_Tipo_Col = 1
    iGrid_Titulo_Col = 2
    iGrid_Parcela_Col = 3
    iGrid_Emissao_Col = 4
    iGrid_Valor_Col = 5
    iGrid_Vencimento_Col = 6
    iGrid_Baixa_Col = 7
    iGrid_Atraso_Col = 8
    iGrid_AReceber_Col = 9
    iGrid_DataProx_Col = 10
    iGrid_DataPrev_Col = 11

    objGrid.objGrid = GridParcelas

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 100 + 1

    objGrid.iLinhasVisiveis = 8

    'Largura da primeira coluna
    GridParcelas.ColWidth(0) = 250

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    Call Grid_Inicializa(objGrid)

    Inicializa_GridParcelas = SUCESSO

End Function

Private Sub GridParcelas_Click()

Dim iExecutaEntradaCelula As Integer
Dim colcolColecoes As New Collection

    Call Grid_Click(objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlteradoGrid)
    End If
    
    colcolColecoes.Add gcolParcelasRec
    
    Call Ordenacao_ClickGrid(objGridParcelas, Ordenacao, colcolColecoes)

End Sub

Private Sub GridParcelas_GotFocus()
    Call Grid_Recebe_Foco(objGridParcelas)
End Sub

Private Sub GridParcelas_EnterCell()
    Call Grid_Entrada_Celula(objGridParcelas, iAlteradoGrid)
End Sub

Private Sub GridParcelas_LeaveCell()
    Call Saida_Celula(objGridParcelas)
End Sub

Private Sub GridParcelas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlteradoGrid)
    End If

End Sub

Private Sub GridParcelas_RowColChange()
    Call Grid_RowColChange(objGridParcelas)
    
    If objGridParcelas.iLinhaAntiga <> objGridParcelas.objGrid.Row Then
        Call Trata_RowColChange
        objGridParcelas.iLinhaAntiga = objGridParcelas.objGrid.Row
    End If
End Sub

Private Sub GridParcelas_Scroll()
    Call Grid_Scroll(objGridParcelas)
End Sub

Private Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridParcelas)
End Sub

Private Sub GridParcelas_LostFocus()
    Call Grid_Libera_Foco(objGridParcelas)
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
    lErro = Limpa_Tela_HistoricoCobranca
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

Function Limpa_Tela_HistoricoCobranca() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_HistoricoCobranca
   
    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    iAlterado = 0
    
    Call Grid_Limpa(objGridParcelas)
    
    Set gobjHistCobrClienteAnt = New ClassHistoricoCobrSelCli
    Set gcolParcelasRec = New Collection
    Set gcolTitulosRec = New Collection
    
    CheckTitAberto.Value = vbUnchecked
    ApenasQualifProx.ListIndex = -1
    EntreQualifProxDe.ListIndex = -1
    EntreQualifProxAte.ListIndex = -1
    ApenasQualifVenc.ListIndex = -1
    EntreQualifVencDe.ListIndex = -1
    EntreQualifVencAte.ListIndex = -1
    ApenasQualifPrev.ListIndex = -1
    EntreQualifPrevDe.ListIndex = -1
    EntreQualifPrevAte.ListIndex = -1
    
    Filial.Clear
    Contato.Clear
    
    Call FrameD_Enabled(FrameD1, FRAMED_FAIXA)
    Call FrameD_Enabled(FrameD2, FRAMED_FAIXA)
    Call FrameD_Enabled(FrameD3, FRAMED_FAIXA)

    LabelTipo.Caption = ""
    LabelTitulo.Caption = ""
    LabelParcela.Caption = ""
    NumIntParcRec.Caption = ""
    
    Call Ordenacao_Limpa(objGridParcelas, Ordenacao)
    
    If iFrameAtual <> TAB_SELECAO Then
        'Torna Frame atual invisível
        Frame1(TabStrip1.SelectedItem.Index).Visible = False
        'Torna Frame atual visível
        Frame1(TAB_SELECAO).Visible = True
        TabStrip1.Tabs.Item(TAB_SELECAO).Selected = True
        iFrameAtual = TAB_SELECAO
    End If
    
    iAlterado = 0

    Limpa_Tela_HistoricoCobranca = SUCESSO

    Exit Function

Erro_Limpa_Tela_HistoricoCobranca:

    Limpa_Tela_HistoricoCobranca = gErr

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

Private Sub UpDownDataVencAte_DownClick()
    Call UpDownData_DownClick(DataVencAte)
End Sub

Private Sub UpDownDataVencAte_UpClick()
    Call UpDownData_UpClick(DataVencAte)
End Sub

Private Sub UpDownDataVencDe_DownClick()
    Call UpDownData_DownClick(DataVencDe)
End Sub

Private Sub UpDownDataVencDe_UpClick()
    Call UpDownData_UpClick(DataVencDe)
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

Private Sub Limpa_FrameD2(ByVal iIndice As Integer)

    If iIndice <> FRAMED_FAIXA Then
        DataVencDe.PromptInclude = False
        DataVencDe.Text = ""
        DataVencDe.PromptInclude = True
    
        DataVencAte.PromptInclude = False
        DataVencAte.Text = ""
        DataVencAte.PromptInclude = True
    End If

    If iIndice <> FRAMED_APENAS Then
        ApenasQualifVenc.ListIndex = -1
        
        ApenasDiasVenc.PromptInclude = False
        ApenasDiasVenc.Text = ""
        ApenasDiasVenc.PromptInclude = True
    End If

    If iIndice <> FRAMED_ENTRE Then
    
        EntreDiasVencDe.PromptInclude = False
        EntreDiasVencDe.Text = ""
        EntreDiasVencDe.PromptInclude = True
        
        EntreQualifVencDe.ListIndex = -1
    
        EntreDiasVencAte.PromptInclude = False
        EntreDiasVencAte.Text = ""
        EntreDiasVencAte.PromptInclude = True
    
        EntreQualifVencAte.ListIndex = -1
    
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

Private Sub ApenasVenc_Click()
    Call FrameD_Enabled(FrameD2, FRAMED_APENAS)
    Call Limpa_FrameD2(FRAMED_APENAS)
End Sub

Private Sub EntreVenc_Click()
    Call FrameD_Enabled(FrameD2, FRAMED_ENTRE)
    Call Limpa_FrameD2(FRAMED_ENTRE)
End Sub

Private Sub FaixaDataVenc_Click()
    Call FrameD_Enabled(FrameD2, FRAMED_FAIXA)
    Call Limpa_FrameD2(FRAMED_FAIXA)
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

    DataPrev.PromptInclude = False
    DataPrev.Text = ""
    DataPrev.PromptInclude = True

    DataProx.PromptInclude = False
    DataProx.Text = ""
    DataProx.PromptInclude = True
    
    Assunto.Text = ""
    
    Status.ListIndex = iStatus_ListIndex_Padrao
    
    iAlterado = 0

End Sub

Private Sub DataProxDe_Validate(Cancel As Boolean)
    Call Data_Validate(DataProxDe, Cancel)
End Sub

Private Sub DataProxAte_Validate(Cancel As Boolean)
    Call Data_Validate(DataProxAte, Cancel)
End Sub

Private Sub DataVencDe_Validate(Cancel As Boolean)
    Call Data_Validate(DataVencDe, Cancel)
End Sub

Private Sub DataVencAte_Validate(Cancel As Boolean)
    Call Data_Validate(DataVencAte, Cancel)
End Sub

Private Sub DataPrevDe_Validate(Cancel As Boolean)
    Call Data_Validate(DataPrevDe, Cancel)
End Sub

Private Sub DataPrevAte_Validate(Cancel As Boolean)
    Call Data_Validate(DataPrevAte, Cancel)
End Sub

Private Sub Ordenacao_Change()
   
Dim colcolColecoes As New Collection

    colcolColecoes.Add gcolParcelasRec
    Call Ordenacao_Atualiza(objGridParcelas, Ordenacao, colcolColecoes)
    
End Sub

Private Sub Ordenacao_Click()

Dim colcolColecoes As New Collection

    colcolColecoes.Add gcolParcelasRec
    Call Ordenacao_Atualiza(objGridParcelas, Ordenacao, colcolColecoes)
    
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer, bCancel As Boolean
Dim colCodigoNome As New AdmColCodigoNome
Dim iClienteAlterado As Integer

On Error GoTo Erro_Cliente_Validate

    'Verifica se o Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then gError 182261

        lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 182262
        
        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)
            
    'Se não estiver preenchido
    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        'Limpa a Combo de Filiais
        Filial.Clear
        
    End If
    
    If TabStrip1.SelectedItem.Index = TAB_Parcelas And Not bAbrindoTela And lClienteAnt <> objCliente.lCodigo Then
        lErro = Trata_TabClick
        'If lErro <> SUCESSO Then gError 182263
    End If
    
    lClienteAnt = objCliente.lCodigo

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 182261 To 182263

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182264)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) <> 0 Then

        'Verifica se é uma filial selecionada
        'If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub
    
        'Tenta selecionar na combo
        lErro = Combo_Seleciona(Filial, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 182265
    
        'Se não encontrou o CÓDIGO
        If lErro = 6730 Then
    
            'Verifica se o cliente foi digitado
            If Len(Trim(Cliente.Text)) = 0 Then gError 182266
    
            sCliente = Cliente.Text
            objFilialCliente.iCodFilial = iCodigo
    
            'Pesquisa se existe Filial com o código extraído
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
            If lErro <> SUCESSO And lErro <> 17660 Then gError 182267
    
            If lErro = 17660 Then gError 182268
    
            'Coloca na tela a Filial lida
            Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome
        
        End If
    
        'Não encontrou a STRING
        If lErro = 6731 Then gError 182269
        
    End If
    
    If TabStrip1.SelectedItem.Index = TAB_Parcelas And Not bAbrindoTela And iFilialAnt <> Codigo_Extrai(Filial.Text) Then
        lErro = Trata_TabClick
        'If lErro <> SUCESSO Then gError 182270
    End If
    
    Call Trata_Contato
    
    iFilialAnt = Codigo_Extrai(Filial.Text)

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 182265, 182267

        Case 182266
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 182268
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

                If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            Else
            End If

        Case 182269
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case 182270

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182271)

    End Select

    Exit Sub

End Sub

Private Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objCliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente, Cancel As Boolean

    Set objCliente = obj1
    
    'Preenche o Cliente com o Cliente selecionado
    Cliente.Text = objCliente.sNomeReduzido
    Call Cliente_Validate(Cancel)

    Me.Show
    
    Exit Sub

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

    If GridParcelas.Row <> 0 Then

        colSelecao.Add NumIntParcRec.Caption
        
        Call Chama_Tela("RelacionamentoClientes_Lista", colSelecao, objRelacionamentoCli, Nothing, "NumIntParcRec = ? ")

    End If

End Sub

Public Sub BotaoDocOriginal_Click()

Dim lErro As Long
Dim objTituloReceber As New ClassTituloReceber

On Error GoTo Erro_BotaoDocOriginal_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridParcelas.Row = 0 Then gError 182282
        
    'Se foi selecionada uma linha que está preenchida
    If GridParcelas.Row <= objGridParcelas.iLinhasExistentes Then
        
        objTituloReceber.lNumIntDoc = StrParaLong(NumIntTitRec.Caption)
        objTituloReceber.iFilialEmpresa = giFilialEmpresa
        
        'Le os Dados do Titulo
        lErro = CF("TituloReceber_Le", objTituloReceber)
        If lErro <> SUCESSO And lErro <> 26061 Then gError 182283
        
        If lErro <> SUCESSO Then
        
            'Le os Dados do Titulo
            lErro = CF("TituloReceberBaixado_Le", objTituloReceber)
            If lErro <> SUCESSO And lErro <> 56570 Then gError 182283
        
        End If
        
        'Se não encontrou
        If lErro <> SUCESSO Then gError 182284
        
        Call Chama_Tela("TituloReceber_Consulta", objTituloReceber)
    
    End If
        
    Exit Sub
    
Erro_BotaoDocOriginal_Click:

    Select Case gErr
    
        Case 182282
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
         
        Case 182283  'Tratado na rotina chamada
        
        Case 182284
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_REC_INEXISTENTE", gErr)
        
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

Public Function Trata_Contato() As Long

Dim lErro As Long
Dim lCliente As Long
Dim objClienteContatos As New ClassClienteContatos

On Error GoTo Erro_Trata_Contato

    lErro = Critica_Cliente
    If lErro <> SUCESSO Then gError 182331
        
    'Obtém o código do cliente
    lErro = Obtem_CodCliente(lCliente)
    If lErro <> SUCESSO Then gError 182332
    
    'Guarda no objClienteContatos, o código do cliente e da
    objClienteContatos.lCliente = lCliente
    objClienteContatos.iFilialCliente = Codigo_Extrai(Filial.Text)
        
    'Carrega a combo de contatos
    lErro = CF("Carrega_ClienteContatos", Contato, objClienteContatos)
    If lErro <> SUCESSO And lErro <> 102622 Then gError 182333
        
    'Se selecionou o contato padrão =>
    If Len(Trim(Contato.Text)) > 0 Then
    
        'traz o telefone do contato
        Call Contato_Click

    End If
    
    Trata_Contato = SUCESSO

    Exit Function

Erro_Trata_Contato:

    Trata_Contato = gErr

    Select Case gErr

        Case 182331 To 182333
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182334)

    End Select

End Function

Private Sub Contato_Click()

Dim lErro As Long
Dim lCliente As Long
Dim objClienteContatos As New ClassClienteContatos

On Error GoTo Erro_Contato_Click

    'Se o campo contato não foi preenchido => sai da função
    If Contato.ListIndex = -1 Then Exit Sub

    lErro = Critica_Cliente
    If lErro <> SUCESSO Then gError 182330
        
    'Obtém o código do cliente
    lErro = Obtem_CodCliente(lCliente)
    If lErro <> SUCESSO Then gError 182310

    'Guarda o código do cliente e da filial no obj
    objClienteContatos.iFilialCliente = Codigo_Extrai(Filial.Text)
    objClienteContatos.lCliente = lCliente
    objClienteContatos.iCodigo = Codigo_Extrai(Contato.Text)

    'Lê o contato no BD
    lErro = CF("ClienteContatos_Le", objClienteContatos)
    If lErro <> SUCESSO And lErro <> 102653 Then gError 182311
    
    'Se não encontrou o contato => erro
    If lErro = 102653 Then gError 182312
    
    Exit Sub
    
Erro_Contato_Click:

    Select Case gErr

        Case 182310, 182311, 182330
        
        Case 182312
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTECONTATO_NAO_ENCONTRADO", gErr, Contato.Text, objClienteContatos.lCliente, objClienteContatos.iFilialCliente)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182313)

    End Select

End Sub

Private Sub Contato_Validate(Cancel As Boolean)
'Faz a validação da filial do cliente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objClienteContatos As New ClassClienteContatos
Dim iCodigo As Integer
Dim lCliente As Long

On Error GoTo Erro_Contato_Validate

    'Se o contato foi preenchido
    If Len(Trim(Contato.Text)) > 0 Then
    
        lErro = Critica_Cliente
        If lErro <> SUCESSO Then gError 182329
    
        'Obtém o código do cliente
        lErro = Obtem_CodCliente(lCliente)
        If lErro <> SUCESSO Then gError 182314
    
        'Guarda o código do cliente e da filial no obj
        objClienteContatos.iFilialCliente = Codigo_Extrai(Filial.Text)
        objClienteContatos.lCliente = lCliente
    
        'Se o contato foi selecionado na própria combo => sai da função
        If Contato.Text = Contato.List(Contato.ListIndex) Then Exit Sub
    
        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
        lErro = Combo_Seleciona(Contato, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 182315
    
        'Se não encontrou o contato na combo, mas retornou um código
        If lErro = 6730 Then

            objClienteContatos.iCodigo = iCodigo
            
            'Lê o contato a partir dos dados passados
            lErro = CF("ClienteContatos_Le", objClienteContatos)
            If lErro <> SUCESSO And lErro <> 102653 Then gError 182316
            
            'Se não encontrou o contato
            If lErro = 102653 Then gError 182317
            
            'Exibe o contato na tela
            Contato.Text = objClienteContatos.iCodigo & SEPARADOR & objClienteContatos.sContato
        
        End If
        
        'Se foi digitado o nome do contato
        'e esse nome não foi encontrado na combo => erro
        If lErro = 6731 Then

            objClienteContatos.sContato = Contato.Text
        
            'Lê o contato a partir dos dados passados
            lErro = CF("ClienteContatos_Le_Nome", objClienteContatos)
            If lErro <> SUCESSO And lErro <> 178440 Then gError 182318
            
            'Se não encontrou o contato
            If lErro = 178440 Then gError 182319
        
            'Exibe o contato na tela
            Contato.Text = objClienteContatos.iCodigo & SEPARADOR & objClienteContatos.sContato
        
        End If
    
    End If
    
    Exit Sub

Erro_Contato_Validate:

    Cancel = True

    Select Case gErr

        Case 182314, 182315, 182316, 182318, 182329
            
        Case 182317, 182319
            
            'Verifica se o usuário deseja criar um novo contato
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTECONTATO", Trim(Contato.Text), Trim(Cliente.Text), Trim(Filial.Text))

            'Se o usuário respondeu sim
            If vbMsgRes = vbYes Then
                'Chama a tela para cadastro de contatos
                Call Chama_Tela("ClienteContatos", objClienteContatos)
            End If
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182320)

    End Select

    Exit Sub

End Sub

Private Sub LabelContato_Click()

Dim lErro As Long
Dim lCliente As Long
Dim objClienteContatos As New ClassClienteContatos

On Error GoTo Erro_LabelContato_Click

    lErro = Critica_Cliente
    If lErro <> SUCESSO Then gError 182328

    'Obtém o código do cliente
    lErro = Obtem_CodCliente(lCliente)
    If lErro <> SUCESSO Then gError 182321

    'Guarda o código do cliente e da filial no obj
    objClienteContatos.iFilialCliente = Codigo_Extrai(Filial.Text)
    objClienteContatos.lCliente = lCliente
        
    Call Chama_Tela("ClienteContatos", objClienteContatos)
    
    Exit Sub

Erro_LabelContato_Click:

    Select Case gErr
    
        Case 182328

        Case 182321
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 182322)

    End Select

    Exit Sub
    
End Sub

Private Function Obtem_CodCliente(lCliente As Long) As Long
'Obtém o código do cliente e da filial que estão na tela e guarda-os no objClienteContatos

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_Obtem_CodCliente

    'Se o cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then
    
        '*** Leitura do cliente a partir do nome reduzido para obter o seu código ***
        
        'Guarda o nome reduzido do cliente
        objCliente.sNomeReduzido = Trim(Cliente.Text)
        
        'Faz a leitura do cliente
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 182323
        
        'Se não encontrou o cliente => erro
        If lErro = 12348 Then gError 182324
        
        'Devolve o código do cliente
        lCliente = objCliente.lCodigo
        
        '*** Fim da leitura de cliente ***
        
    End If

    Obtem_CodCliente = SUCESSO

    Exit Function

Erro_Obtem_CodCliente:

    Obtem_CodCliente = gErr

    Select Case gErr

        Case 182323

        Case 182324
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182325)

    End Select

End Function

Private Function Critica_Cliente() As Long

Dim lErro As Long

On Error GoTo Erro_Critica_Cliente

    If Len(Trim(Cliente.Text)) = 0 Then gError 182276
    If Len(Trim(Filial.Text)) = 0 Then gError 182277
    
    Critica_Cliente = SUCESSO

    Exit Function

Erro_Critica_Cliente:

    Critica_Cliente = gErr

    Select Case gErr
        
        Case 182276
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 182277
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182326)

    End Select

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objRelacCli As New ClassRelacClientes
Dim objParcRec As ClassParcelaReceber

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    If StrParaLong(NumIntParcRec.Caption) = 0 Then gError 182342
    
    lErro = Critica_Cliente
    If lErro <> SUCESSO Then gError 182327
    
    lErro = Move_Tela_Memoria(objRelacCli)
    If lErro <> SUCESSO Then gError 182278
    
    lErro = CF("RelacionamentoClientes_Grava", objRelacCli, True, gsUsuario)
    If lErro <> SUCESSO Then gError 182283
    
    iLinha = 0
    
    For Each objParcRec In gcolParcelasRec
        iLinha = iLinha + 1
        If objParcRec.lNumIntDoc = StrParaLong(NumIntParcRec.Caption) Then
            Exit For
        End If
    Next
    
    If StrParaDate(DataPrev.Text) <> DATA_NULA Then
        GridParcelas.TextMatrix(iLinha, iGrid_DataPrev_Col) = Format(DataPrev.Text, "dd/mm/yyyy")
    Else
        GridParcelas.TextMatrix(iLinha, iGrid_DataPrev_Col) = ""
    End If
    
    If StrParaDate(DataProx.Text) <> DATA_NULA Then
        GridParcelas.TextMatrix(iLinha, iGrid_DataProx_Col) = Format(DataProx.Text, "dd/mm/yyyy")
    Else
        GridParcelas.TextMatrix(iLinha, iGrid_DataProx_Col) = ""
    End If
    
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
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_STATUSRELACCLI, objComboBox)
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
