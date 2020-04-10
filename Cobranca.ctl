VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl CobrancaOcx 
   ClientHeight    =   6570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   9510
   Begin VB.Frame FrameTab 
      BorderStyle     =   0  'None
      Height          =   5625
      Index           =   1
      Left            =   180
      TabIndex        =   63
      Top             =   840
      Width           =   9180
      Begin VB.Frame Frame4 
         Caption         =   "Vendedores"
         Height          =   615
         Left            =   690
         TabIndex        =   130
         Top             =   4935
         Width           =   7485
         Begin VB.OptionButton OptVendIndir 
            Caption         =   "Vendas Indiretas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1950
            TabIndex        =   43
            Top             =   180
            Width           =   1800
         End
         Begin VB.OptionButton OptVendDir 
            Caption         =   "Vendas Diretas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   90
            TabIndex        =   42
            Top             =   180
            Value           =   -1  'True
            Width           =   1800
         End
         Begin MSMask.MaskEdBox Vendedor 
            Height          =   300
            Left            =   4860
            TabIndex        =   44
            Top             =   210
            Width           =   2520
            _ExtentX        =   4445
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3945
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   131
            Top             =   270
            Width           =   885
         End
      End
      Begin VB.Frame FrameCategoriaCliente 
         Caption         =   "Categoria de Cliente"
         Height          =   855
         Left            =   705
         TabIndex        =   121
         Top             =   375
         Width           =   7485
         Begin VB.ComboBox CategoriaCliente 
            Height          =   315
            Left            =   2730
            TabIndex        =   125
            Top             =   150
            Width           =   4620
         End
         Begin VB.ComboBox CategoriaClienteDe 
            Height          =   315
            Left            =   1245
            Sorted          =   -1  'True
            TabIndex        =   124
            Top             =   495
            Width           =   2760
         End
         Begin VB.CheckBox CategoriaClienteTodas 
            Caption         =   "Todas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   315
            TabIndex        =   123
            Top             =   225
            Width           =   855
         End
         Begin VB.ComboBox CategoriaClienteAte 
            Height          =   315
            Left            =   4575
            Sorted          =   -1  'True
            TabIndex        =   122
            Top             =   480
            Width           =   2760
         End
         Begin VB.Label Label4 
            Caption         =   "Label5"
            Height          =   15
            Left            =   360
            TabIndex        =   129
            Top             =   720
            Width           =   30
         End
         Begin VB.Label LabelCategoriaClienteAte 
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
            Left            =   4170
            TabIndex        =   128
            Top             =   540
            Width           =   360
         End
         Begin VB.Label LabelCategoriaClienteDe 
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
            Left            =   885
            TabIndex        =   127
            Top             =   540
            Width           =   315
         End
         Begin VB.Label LabelCategoriaCliente 
            Caption         =   "Categoria:"
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
            Left            =   1815
            TabIndex        =   126
            Top             =   195
            Width           =   855
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
         Left            =   5400
         TabIndex        =   109
         Top             =   45
         Value           =   1  'Checked
         Width           =   2730
      End
      Begin VB.ComboBox ComboCobrador 
         Height          =   315
         Left            =   2385
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   0
         Width           =   2775
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Previsão Recebimento"
         Height          =   1215
         Index           =   3
         Left            =   690
         TabIndex        =   79
         Top             =   3735
         Width           =   7485
         Begin VB.Frame FrameD3 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   360
            Index           =   2
            Left            =   1275
            TabIndex        =   105
            Top             =   825
            Width           =   6105
            Begin VB.ComboBox EntreQualifPrevDe 
               Height          =   315
               ItemData        =   "Cobranca.ctx":0000
               Left            =   1530
               List            =   "Cobranca.ctx":000A
               Style           =   2  'Dropdown List
               TabIndex        =   39
               Top             =   15
               Width           =   1365
            End
            Begin VB.ComboBox EntreQualifPrevAte 
               Height          =   315
               ItemData        =   "Cobranca.ctx":001F
               Left            =   4740
               List            =   "Cobranca.ctx":0029
               Style           =   2  'Dropdown List
               TabIndex        =   41
               Top             =   30
               Width           =   1365
            End
            Begin MSMask.MaskEdBox EntreDiasPrevDe 
               Height          =   315
               Left            =   0
               TabIndex        =   38
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
               TabIndex        =   40
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
               TabIndex        =   108
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
               TabIndex        =   107
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
               TabIndex        =   106
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
            TabIndex        =   103
            Top             =   480
            Width           =   4965
            Begin VB.ComboBox ApenasQualifPrev 
               Height          =   315
               ItemData        =   "Cobranca.ctx":003E
               Left            =   0
               List            =   "Cobranca.ctx":0048
               Style           =   2  'Dropdown List
               TabIndex        =   35
               Top             =   0
               Width           =   1365
            End
            Begin MSMask.MaskEdBox ApenasDiasPrev 
               Height          =   315
               Left            =   1530
               TabIndex        =   36
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
               TabIndex        =   104
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
            TabIndex        =   100
            Top             =   150
            Width           =   4965
            Begin MSMask.MaskEdBox DataPrevAte 
               Height          =   300
               Left            =   2235
               TabIndex        =   32
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
               TabIndex        =   31
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
               TabIndex        =   30
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
               TabIndex        =   33
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
               TabIndex        =   102
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
               TabIndex        =   101
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
            Height          =   195
            Left            =   240
            TabIndex        =   37
            Top             =   900
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
            TabIndex        =   34
            Top             =   540
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
            Left            =   255
            TabIndex        =   29
            Top             =   195
            Value           =   -1  'True
            Width           =   1620
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Vencimento"
         Height          =   1245
         Index           =   2
         Left            =   705
         TabIndex        =   78
         Top             =   2475
         Width           =   7485
         Begin VB.Frame FrameD2 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   345
            Index           =   2
            Left            =   1260
            TabIndex        =   96
            Top             =   855
            Width           =   6120
            Begin VB.ComboBox EntreQualifVencDe 
               Height          =   315
               ItemData        =   "Cobranca.ctx":0063
               Left            =   1485
               List            =   "Cobranca.ctx":006D
               Style           =   2  'Dropdown List
               TabIndex        =   26
               Top             =   0
               Width           =   1365
            End
            Begin VB.ComboBox EntreQualifVencAte 
               Height          =   315
               ItemData        =   "Cobranca.ctx":0082
               Left            =   4725
               List            =   "Cobranca.ctx":008C
               Style           =   2  'Dropdown List
               TabIndex        =   28
               Top             =   15
               Width           =   1365
            End
            Begin MSMask.MaskEdBox EntreDiasVencDe 
               Height          =   315
               Left            =   0
               TabIndex        =   25
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
               Index           =   7
               Left            =   675
               TabIndex        =   99
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
               Index           =   3
               Left            =   4065
               TabIndex        =   98
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
               Index           =   2
               Left            =   3060
               TabIndex        =   97
               Top             =   45
               Width           =   120
            End
         End
         Begin VB.Frame FrameD2 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   405
            Index           =   1
            Left            =   1275
            TabIndex        =   94
            Top             =   495
            Width           =   4965
            Begin VB.ComboBox ApenasQualifVenc 
               Height          =   315
               ItemData        =   "Cobranca.ctx":00A1
               Left            =   0
               List            =   "Cobranca.ctx":00AB
               Style           =   2  'Dropdown List
               TabIndex        =   22
               Top             =   0
               Width           =   1365
            End
            Begin MSMask.MaskEdBox ApenasDiasVenc 
               Height          =   315
               Left            =   1485
               TabIndex        =   23
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
               TabIndex        =   95
               Top             =   45
               Width           =   480
            End
         End
         Begin VB.Frame FrameD2 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   315
            Index           =   0
            Left            =   2400
            TabIndex        =   91
            Top             =   150
            Width           =   4965
            Begin MSMask.MaskEdBox DataVencAte 
               Height          =   300
               Left            =   2235
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
            Begin MSComCtl2.UpDown UpDownDataVencDe 
               Height          =   300
               Left            =   1350
               TabIndex        =   18
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
               TabIndex        =   17
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
               TabIndex        =   20
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
               Index           =   9
               Left            =   1830
               TabIndex        =   93
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
               Index           =   8
               Left            =   0
               TabIndex        =   92
               Top             =   60
               Width           =   315
            End
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
            TabIndex        =   24
            Top             =   915
            Width           =   780
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
            TabIndex        =   21
            Top             =   555
            Width           =   1050
         End
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
            TabIndex        =   16
            Top             =   180
            Value           =   -1  'True
            Width           =   1620
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Próximo Contato"
         Height          =   1245
         Index           =   0
         Left            =   705
         TabIndex        =   77
         Top             =   1230
         Width           =   7485
         Begin VB.Frame FrameD1 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   360
            Index           =   2
            Left            =   1230
            TabIndex        =   87
            Top             =   840
            Width           =   6195
            Begin VB.ComboBox EntreQualifProxAte 
               Height          =   315
               ItemData        =   "Cobranca.ctx":00C6
               Left            =   4755
               List            =   "Cobranca.ctx":00D0
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   15
               Width           =   1365
            End
            Begin VB.ComboBox EntreQualifProxDe 
               Height          =   315
               ItemData        =   "Cobranca.ctx":00E5
               Left            =   1515
               List            =   "Cobranca.ctx":00EF
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   15
               Width           =   1365
            End
            Begin MSMask.MaskEdBox EntreDiasProxDe 
               Height          =   315
               Left            =   45
               TabIndex        =   12
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
            Begin MSMask.MaskEdBox EntreDiasProxAte 
               Height          =   315
               Left            =   3390
               TabIndex        =   14
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
               Index           =   5
               Left            =   3090
               TabIndex        =   90
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
               Left            =   4095
               TabIndex        =   89
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
               Index           =   4
               Left            =   720
               TabIndex        =   88
               Top             =   45
               Width           =   480
            End
         End
         Begin VB.Frame FrameD1 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   345
            Index           =   1
            Left            =   1185
            TabIndex        =   85
            Top             =   480
            Width           =   6195
            Begin VB.ComboBox ApenasQualifProx 
               Height          =   315
               ItemData        =   "Cobranca.ctx":0104
               Left            =   90
               List            =   "Cobranca.ctx":010E
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   15
               Width           =   1365
            End
            Begin MSMask.MaskEdBox ApenasDiasProx 
               Height          =   315
               Left            =   1560
               TabIndex        =   10
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
               Index           =   10
               Left            =   2265
               TabIndex        =   86
               Top             =   60
               Width           =   480
            End
         End
         Begin VB.Frame FrameD1 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   330
            Index           =   0
            Left            =   2295
            TabIndex        =   82
            Top             =   135
            Width           =   5130
            Begin MSMask.MaskEdBox DataProxAte 
               Height          =   300
               Left            =   2325
               TabIndex        =   6
               Top             =   15
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
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   15
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataProxDe 
               Height          =   300
               Left            =   450
               TabIndex        =   4
               Top             =   15
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
               TabIndex        =   7
               TabStop         =   0   'False
               Top             =   15
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
               Index           =   0
               Left            =   90
               TabIndex        =   84
               Top             =   75
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
               Index           =   1
               Left            =   1920
               TabIndex        =   83
               Top             =   75
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
            Top             =   225
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
            TabIndex        =   8
            Top             =   525
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
            TabIndex        =   11
            Top             =   870
            Width           =   780
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuário Cobrador:"
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
         Index           =   12
         Left            =   765
         TabIndex        =   80
         Top             =   45
         Width           =   1545
      End
   End
   Begin VB.Frame FrameTab 
      BorderStyle     =   0  'None
      Height          =   5685
      Index           =   2
      Left            =   180
      TabIndex        =   64
      Top             =   795
      Visible         =   0   'False
      Width           =   9180
      Begin VB.Frame Frame1 
         Caption         =   "Clientes"
         Height          =   5640
         Index           =   1
         Left            =   15
         TabIndex        =   65
         Top             =   30
         Width           =   9135
         Begin MSMask.MaskEdBox ValorDevidoEmAtraso 
            Height          =   225
            Left            =   510
            TabIndex        =   118
            Top             =   1500
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin VB.Frame Frame2 
            Caption         =   "Ligação"
            Height          =   1380
            Left            =   1665
            TabIndex        =   110
            Top             =   4170
            Width           =   5670
            Begin MSMask.MaskEdBox OperDDD 
               Height          =   315
               Left            =   1830
               TabIndex        =   58
               Top             =   975
               Width           =   420
               _ExtentX        =   741
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   2
               Mask            =   "##"
               PromptChar      =   "_"
            End
            Begin VB.TextBox SenhaTel 
               Enabled         =   0   'False
               Height          =   315
               IMEMode         =   3  'DISABLE
               Left            =   2865
               PasswordChar    =   "*"
               TabIndex        =   57
               Top             =   585
               Width           =   675
            End
            Begin VB.CommandButton BotaoLigar 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   4515
               Picture         =   "Cobranca.ctx":0129
               Style           =   1  'Graphical
               TabIndex        =   54
               Top             =   150
               Width           =   405
            End
            Begin VB.CommandButton BotaoDesligar 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   4980
               Picture         =   "Cobranca.ctx":07E3
               Style           =   1  'Graphical
               TabIndex        =   55
               Top             =   150
               Width           =   405
            End
            Begin VB.ComboBox Contato 
               Height          =   315
               Left            =   915
               TabIndex        =   53
               ToolTipText     =   $"Cobranca.ctx":0E9D
               Top             =   210
               Width           =   1785
            End
            Begin VB.CheckBox PossuiSenha 
               Caption         =   "Telefone com senha"
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
               Left            =   105
               TabIndex        =   56
               Top             =   585
               Width           =   2055
            End
            Begin MSMask.MaskEdBox DigDiscExt 
               Height          =   315
               Left            =   5085
               TabIndex        =   59
               Top             =   975
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   1
               Mask            =   "#"
               PromptChar      =   "_"
            End
            Begin VB.Label Label6 
               Caption         =   "Para obter linha externa discar:"
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
               Left            =   2385
               TabIndex        =   117
               Top             =   1035
               Width           =   2745
            End
            Begin VB.Label Label5 
               Caption         =   "Operadora de DDD:"
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
               Left            =   105
               TabIndex        =   116
               Top             =   1020
               Width           =   1725
            End
            Begin VB.Label UF 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   3960
               TabIndex        =   115
               Top             =   585
               Width           =   495
            End
            Begin VB.Label Label2 
               Caption         =   "UF:"
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
               Left            =   3615
               TabIndex        =   114
               Top             =   645
               Width           =   315
            End
            Begin VB.Label Label3 
               Caption         =   "Senha:"
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
               Left            =   2265
               TabIndex        =   113
               Top             =   660
               Width           =   660
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
               Left            =   105
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   112
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Telefone 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   2700
               TabIndex        =   111
               Top             =   210
               Width           =   1770
            End
         End
         Begin VB.CommandButton BotaoMarcar 
            Caption         =   "Marcar Todas"
            Height          =   555
            Left            =   105
            Picture         =   "Cobranca.ctx":0F25
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   4260
            Width           =   1440
         End
         Begin VB.CommandButton BotaoDesmarcar 
            Caption         =   "Desmarcar Todas"
            Height          =   555
            Left            =   105
            Picture         =   "Cobranca.ctx":1F3F
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   4980
            Width           =   1440
         End
         Begin VB.CommandButton BotaoHistorico 
            Caption         =   "Histórico de Recebimentos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   7440
            TabIndex        =   60
            Top             =   4260
            Width           =   1485
         End
         Begin VB.CommandButton BotaoGravarTela 
            Caption         =   "Gravar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   7440
            TabIndex        =   62
            Top             =   5130
            Width           =   1485
         End
         Begin VB.CommandButton BotaoCliente 
            Caption         =   "Cliente..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   7440
            TabIndex        =   61
            Top             =   4695
            Width           =   1485
         End
         Begin VB.ComboBox Ordenacao 
            Height          =   315
            ItemData        =   "Cobranca.ctx":3121
            Left            =   1170
            List            =   "Cobranca.ctx":3123
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   180
            Width           =   2910
         End
         Begin VB.TextBox HistoricoGrid 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   5160
            TabIndex        =   76
            Top             =   1320
            Width           =   3510
         End
         Begin VB.TextBox Fone2Grid 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   6705
            TabIndex        =   75
            Top             =   255
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox CheckLigacaoRealizada 
            Height          =   240
            Left            =   7080
            TabIndex        =   74
            Top             =   915
            Width           =   1185
         End
         Begin VB.CheckBox CheckLigar 
            Height          =   240
            Left            =   600
            TabIndex        =   73
            Top             =   870
            Width           =   615
         End
         Begin VB.TextBox Fone1Grid 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   5550
            TabIndex        =   71
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox ClienteGrid 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   2265
            TabIndex        =   70
            Top             =   885
            Width           =   2850
         End
         Begin VB.TextBox FilialClienteGrid 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Left            =   1260
            TabIndex        =   69
            Top             =   1140
            Width           =   945
         End
         Begin MSMask.MaskEdBox ContatoGrid 
            Height          =   240
            Left            =   4365
            TabIndex        =   72
            Top             =   240
            Visible         =   0   'False
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            Format          =   "dd/mm/yyyy"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridClientes 
            Height          =   3285
            Left            =   90
            TabIndex        =   50
            Top             =   570
            Width           =   8850
            _ExtentX        =   15610
            _ExtentY        =   5794
            _Version        =   393216
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
            Left            =   6420
            TabIndex        =   120
            Top             =   3900
            Width           =   1005
         End
         Begin VB.Label ValorDevTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   7485
            TabIndex        =   119
            Top             =   3885
            Width           =   1410
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
            Index           =   19
            Left            =   135
            TabIndex        =   81
            Top             =   240
            Width           =   1095
         End
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5625
      Top             =   15
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4890
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   7275
      ScaleHeight     =   450
      ScaleWidth      =   2055
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   45
      Width           =   2115
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "Cobranca.ctx":3125
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "Cobranca.ctx":32A3
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "Cobranca.ctx":37D5
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   585
         Picture         =   "Cobranca.ctx":392F
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
   End
   Begin VB.ComboBox OpcoesTela 
      Height          =   315
      Left            =   900
      TabIndex        =   0
      Top             =   75
      Width           =   2775
   End
   Begin VB.CheckBox OpcaoPadrao 
      Caption         =   "Padrão"
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
      Left            =   3900
      TabIndex        =   1
      Top             =   135
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6030
      Left            =   120
      TabIndex        =   67
      Top             =   480
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   10636
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Relacionamentos"
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
   Begin MSCommLib.MSComm ComDiscar 
      Left            =   6480
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   0   'False
      InBufferSize    =   2000
      OutBufferSize   =   2000
      ParityReplace   =   48
   End
   Begin VB.Label LabelOpcao 
      AutoSize        =   -1  'True
      Caption         =   "Opção:"
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
      TabIndex        =   68
      Top             =   135
      Width           =   630
   End
End
Attribute VB_Name = "CobrancaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim sMsg As String

Public iAtualizaTela As Integer

Dim sOpcaoAnt As String

Dim iAlterado As Integer
Dim iFrameAtual As Integer

Dim objGridClientes As AdmGrid
Dim iGrid_CheckLigar_Col As Integer
Dim iGrid_FilialCliente_Col As Integer
Dim iGrid_Cliente_Col As Integer
Dim iGrid_Contato_Col As Integer
Dim iGrid_Fone1_Col As Integer
Dim iGrid_Fone2_Col As Integer
Dim iGrid_CheckLigacaoRealizada_Col As Integer
Dim iGrid_Historico_Col As Integer
Dim iGrid_ValorDevido_Col As Integer

Dim gobjCobrancaClienteAnt As ClassCobrancaSelCli

Const TAB_SELECAO = 1
Const TAB_CLIENTE = 2

Const FRAMED_FAIXA = 0
Const FRAMED_APENAS = 1
Const FRAMED_ENTRE = 2

Dim gbTrazendoDados As Boolean

Dim iConTimer As Integer

'''' API do WIndows
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Cobrança"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Cobranca"

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

Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload
    
    Set gobjCobrancaClienteAnt = Nothing
    
    If ComDiscar.PortOpen Then
        ComDiscar.PortOpen = False
        BotaoDesligar.Enabled = False
        DoEvents
    End If
    
    Call ComandoSeta_Liberar(Me.Name)

    Set objGridClientes = Nothing

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182088)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long
Dim objTela As Object
Dim sDigDiscExt As String

On Error GoTo Erro_Form_Load

    gbTrazendoDados = False

    Set objGridClientes = New AdmGrid
    
    Set gobjCobrancaClienteAnt = New ClassCobrancaSelCli
    
    lErro = Inicializa_GridClientes(objGridClientes)
    If lErro <> SUCESSO Then gError 182089

    '#################################################
    'Inserido por Wagner
    Call Carrega_ComboCategoriaCliente(CategoriaCliente)
    '#################################################

    Call Carrega_Usuarios
    
    '#####################################
    'Inserido por Wagner
    CategoriaClienteTodas.Value = vbChecked
    CategoriaCliente.Enabled = False
    CategoriaClienteDe.Enabled = False
    CategoriaClienteAte.Enabled = False
    CategoriaClienteDe.ListIndex = -1
    CategoriaClienteAte.ListIndex = -1
    '#####################################

    'Guarda em objTela os dados dessa tela
    Set objTela = Me
    
    lErro = CF("Carrega_OpcoesTela", objTela, True)
    If lErro <> SUCESSO Then gError 182154
    
    Call FrameD_Enabled(FrameD1, FRAMED_FAIXA)
    Call FrameD_Enabled(FrameD2, FRAMED_FAIXA)
    Call FrameD_Enabled(FrameD3, FRAMED_FAIXA)
    
    comboCobrador.Text = gsUsuario
    
    sDigDiscExt = String(128, 0)
    
    Call GetPrivateProfileString("Geral", "DigDiscExt", "", sDigDiscExt, 128, "ADM100.INI")

    sDigDiscExt = Replace(sDigDiscExt, Chr(0), "")

    DigDiscExt.PromptInclude = False
    DigDiscExt.Text = sDigDiscExt
    DigDiscExt.PromptInclude = True
    
    BotaoDesligar.Enabled = False
    
    iFrameAtual = TAB_SELECAO
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO
   
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 182089, 182154
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182090)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182091)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Selecao_Memoria(ByVal objCobrancaSelCli As ClassCobrancaSelCli) As Long

Dim lErro As Long
Dim dtDataDe As Date
Dim dtDataAte As Date

On Error GoTo Erro_Move_Selecao_Memoria

    If Len(Trim(comboCobrador.Text)) = 0 Then gError 182368

    If CheckTitAberto.Value = vbUnchecked Then
        objCobrancaSelCli.iTitulosBaixados = DESMARCADO
    Else
        objCobrancaSelCli.iTitulosBaixados = MARCADO
    End If
    
    objCobrancaSelCli.sCobrador = comboCobrador.Text

    If FaixaDataProx.Value Then
        objCobrancaSelCli.dtDataProxDe = StrParaDate(DataProxDe.Text)
        objCobrancaSelCli.dtDataProxAte = StrParaDate(DataProxAte.Text)
    End If

    If ApenasProx.Value Then
    
        If ApenasQualifProx.ListIndex = -1 Then gError 182175
    
        Call Datas_Trata_Apenas(dtDataDe, dtDataAte, StrParaInt(ApenasDiasProx.Text), ApenasQualifProx.ItemData(ApenasQualifProx.ListIndex))
        
        objCobrancaSelCli.dtDataProxDe = dtDataDe
        objCobrancaSelCli.dtDataProxAte = dtDataAte
    End If
    
    If EntreProx.Value Then
        
        If EntreQualifProxAte.ListIndex = -1 Then gError 182176
        If EntreQualifProxAte.ListIndex = -1 Then gError 182177
        
        Call Datas_Trata_Entre(dtDataDe, dtDataAte, StrParaInt(EntreDiasProxDe), EntreQualifProxDe.ItemData(EntreQualifProxDe.ListIndex), StrParaInt(EntreDiasProxAte), EntreQualifProxAte.ItemData(EntreQualifProxAte.ListIndex))
        
        objCobrancaSelCli.dtDataProxDe = dtDataDe
        objCobrancaSelCli.dtDataProxAte = dtDataAte
    End If

    If FaixaDataVenc.Value Then
        objCobrancaSelCli.dtDataVencDe = StrParaDate(DataVencDe.Text)
        objCobrancaSelCli.dtDataVencAte = StrParaDate(DataVencAte.Text)
    End If

    If ApenasVenc.Value Then
    
        If ApenasQualifVenc.ListIndex = -1 Then gError 182178

        Call Datas_Trata_Apenas(dtDataDe, dtDataAte, StrParaInt(ApenasDiasVenc.Text), ApenasQualifVenc.ItemData(ApenasQualifVenc.ListIndex))
        
        objCobrancaSelCli.dtDataVencDe = dtDataDe
        objCobrancaSelCli.dtDataVencAte = dtDataAte
    End If
    
    If EntreVenc.Value Then
    
        If EntreQualifVencDe.ListIndex = -1 Then gError 182179
        If EntreQualifVencAte.ListIndex = -1 Then gError 182180

        Call Datas_Trata_Entre(dtDataDe, dtDataAte, StrParaInt(EntreDiasVencDe), EntreQualifVencDe.ItemData(EntreQualifVencDe.ListIndex), StrParaInt(EntreDiasVencAte), EntreQualifVencAte.ItemData(EntreQualifVencAte.ListIndex))
        
        objCobrancaSelCli.dtDataVencDe = dtDataDe
        objCobrancaSelCli.dtDataVencAte = dtDataAte
    End If
    
    If FaixaDataPrev.Value Then
        objCobrancaSelCli.dtDataPrevDe = StrParaDate(DataPrevDe.Text)
        objCobrancaSelCli.dtDataPrevAte = StrParaDate(DataPrevAte.Text)
    End If

    If ApenasPrev.Value Then
    
        If ApenasQualifPrev.ListIndex = -1 Then gError 182181

        Call Datas_Trata_Apenas(dtDataDe, dtDataAte, StrParaInt(ApenasDiasPrev.Text), ApenasQualifPrev.ItemData(ApenasQualifPrev.ListIndex))
        
        objCobrancaSelCli.dtDataPrevDe = dtDataDe
        objCobrancaSelCli.dtDataPrevAte = dtDataAte
    End If
    
    If EntrePrev.Value Then
        
        If EntreQualifPrevDe.ListIndex = -1 Then gError 182182
        If EntreQualifPrevAte.ListIndex = -1 Then gError 182183
        
        Call Datas_Trata_Entre(dtDataDe, dtDataAte, StrParaInt(EntreDiasPrevDe), EntreQualifPrevDe.ItemData(EntreQualifPrevDe.ListIndex), StrParaInt(EntreDiasPrevAte.Text), EntreQualifPrevAte.ItemData(EntreQualifPrevAte.ListIndex))
        
        objCobrancaSelCli.dtDataPrevDe = dtDataDe
        objCobrancaSelCli.dtDataPrevAte = dtDataAte
    End If
    
    If objCobrancaSelCli.dtDataPrevAte <> DATA_NULA And objCobrancaSelCli.dtDataPrevDe <> DATA_NULA Then
        If objCobrancaSelCli.dtDataPrevDe > objCobrancaSelCli.dtDataPrevAte Then gError 182096
    End If
    
    If objCobrancaSelCli.dtDataProxDe <> DATA_NULA And objCobrancaSelCli.dtDataProxAte <> DATA_NULA Then
        If objCobrancaSelCli.dtDataProxDe > objCobrancaSelCli.dtDataProxAte Then gError 182097
    End If
    
    If objCobrancaSelCli.dtDataVencDe <> DATA_NULA And objCobrancaSelCli.dtDataVencAte <> DATA_NULA Then
        If objCobrancaSelCli.dtDataVencDe > objCobrancaSelCli.dtDataVencAte Then gError 182098
    End If
    
    If CategoriaClienteTodas.Value = vbChecked Then
        objCobrancaSelCli.sCategoria = ""
        objCobrancaSelCli.sCategoriaDe = ""
        objCobrancaSelCli.sCategoriaAte = ""
    Else
        If CategoriaCliente.Text = "" Then gError 198620
        objCobrancaSelCli.sCategoria = CategoriaCliente.Text
        objCobrancaSelCli.sCategoriaDe = CategoriaClienteDe.Text
        objCobrancaSelCli.sCategoriaAte = CategoriaClienteAte.Text
    End If
    
    If OptVendDir.Value Then
        objCobrancaSelCli.iTipoVend = VENDEDOR_DIRETO
    Else
        objCobrancaSelCli.iTipoVend = VENDEDOR_INDIRETO
    End If
    
    objCobrancaSelCli.iVendedor = Codigo_Extrai(Vendedor.Text)
    
    If objCobrancaSelCli.sCategoriaDe > objCobrancaSelCli.sCategoriaAte Then gError 198621
   
    Move_Selecao_Memoria = SUCESSO

    Exit Function

Erro_Move_Selecao_Memoria:

    Move_Selecao_Memoria = gErr

    Select Case gErr
    
        Case 182096 To 182098
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)
            
        Case 182175 To 182183
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_TRATA_TIPO", gErr)
            
        Case 182368
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIOCOBRADOR_NAO_PREENCHIDO", gErr)
            
        Case 198620
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTE_NAO_INFORMADA", gErr)
            
        Case 198621
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTE_ITEM_INICIAL_MAIOR", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182093)

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
            gError 182102

    End Select
    
    Select Case iData2
    
        Case DATA_AFRENTE
            dtDataAte = DateAdd("d", iNumDias2, gdtDataAtual)
        
        Case DATA_ATRAS
            dtDataAte = DateAdd("d", -iNumDias2, gdtDataAtual)
        
        Case Else
            gError 182103

    End Select
   
    Datas_Trata_Entre = SUCESSO

    Exit Function

Erro_Datas_Trata_Entre:

    Datas_Trata_Entre = gErr

    Select Case gErr
    
        Case 182102, 182103
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_TRATA_TIPO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182094)

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
            gError 182102

    End Select
   
    Datas_Trata_Apenas = SUCESSO

    Exit Function

Erro_Datas_Trata_Apenas:

    Datas_Trata_Apenas = gErr

    Select Case gErr
    
        Case 182102
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_TRATA_TIPO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182095)

    End Select

    Exit Function

End Function

Function Trata_Selecao(ByVal objCobrancaSelCli As ClassCobrancaSelCli) As Long

Dim lErro As Long
Dim colFiliais As New Collection
Dim colEnderecos As New Collection

On Error GoTo Erro_Trata_Selecao

    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = CF("CobrancaCliente_Le", objCobrancaSelCli, colFiliais, colEnderecos)
    If lErro <> SUCESSO Then gError 182099
    
    If colFiliais.Count = 0 Then gError 182100
    
    lErro = Preenche_GridCliente(colFiliais, colEnderecos)
    If lErro <> SUCESSO Then gError 182101
    
    GL_objMDIForm.MousePointer = vbDefault
   
    Trata_Selecao = SUCESSO

    Exit Function

Erro_Trata_Selecao:

    GL_objMDIForm.MousePointer = vbDefault

    Trata_Selecao = gErr

    Select Case gErr
    
        Case 182099, 182101
        
        Case 182100
            Call Rotina_Erro(vbOKOnly, "ERRO_SELECAO_COBRANCA_SEM_CLIENTES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174254)

    End Select

    Exit Function

End Function

Function Preenche_GridCliente(ByVal colFiliais As Collection, ByVal colEnderecos As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objFilial As ClassFilialCliente
Dim objEnderecos As ClassEndereco
Dim colCliente As New Collection
Dim objCliente As ClassCliente
Dim bAchou As Boolean
Dim objFilialContatoData As New ClassFilialContatoData
Dim objClieEst As ClassFilialClienteEst
Dim colEstatisticas As New Collection
Dim iPos As Integer

On Error GoTo Erro_Preenche_GridCliente

    gbTrazendoDados = True

    Call Grid_Limpa(objGridClientes)
    
    'Aumenta o número de linhas do grid se necessário
    If colFiliais.Count >= objGridClientes.objGrid.Rows Then
        Call Refaz_Grid(objGridClientes, colFiliais.Count)
    End If

    iIndice = 0
    For Each objFilial In colFiliais
    
        iIndice = iIndice + 1
        
        Set objEnderecos = colEnderecos(iIndice)
        
        bAchou = False
        iPos = 0
        For Each objCliente In colCliente
            iPos = iPos + 1
            If objCliente.lCodigo = objFilial.lCodCliente Then
                bAchou = True
                Set objClieEst = colEstatisticas.Item(iPos)
                Exit For
            End If
        Next
        
        If Not bAchou Then
        
            Set objCliente = New ClassCliente
            Set objClieEst = New ClassFilialClienteEst
            
            objCliente.lCodigo = objFilial.lCodCliente
        
            'le o cliente
            lErro = CF("Cliente_Le", objCliente)
            If lErro <> SUCESSO And lErro <> 12293 Then gError 182105
            
            objClieEst.lCodCliente = objFilial.lCodCliente
            objClieEst.iCodFilial = objFilial.iCodFilial
            
            If gobjCP.iFilialCentralizadora = giFilialEmpresa Or giFilialEmpresa = EMPRESA_TODA Then
                objClieEst.iFilialEmpresa = EMPRESA_TODA
            Else
                objClieEst.iFilialEmpresa = giFilialEmpresa
            End If
            
            lErro = CF("Cliente_Le_Estatistica_Atraso", objClieEst)
            If lErro <> SUCESSO Then gError 185940
        
        End If
        
        objFilialContatoData.lCliente = objFilial.lCodCliente
        objFilialContatoData.iFilial = objFilial.iCodFilial
        objFilialContatoData.dtData = gdtDataAtual
        
        'Busca dados de Ligação, histórico, Ligações efetuadas naquela data
        lErro = CF("FilialContatoData_Le", objFilialContatoData)
        If lErro <> SUCESSO Then gError 182106

        GridClientes.TextMatrix(iIndice, iGrid_CheckLigacaoRealizada_Col) = objFilialContatoData.iLigacaoEfetuada
        GridClientes.TextMatrix(iIndice, iGrid_CheckLigar_Col) = objFilialContatoData.iLigar
        GridClientes.TextMatrix(iIndice, iGrid_Cliente_Col) = objCliente.lCodigo & SEPARADOR & objCliente.sNomeReduzido
        GridClientes.TextMatrix(iIndice, iGrid_FilialCliente_Col) = objFilial.iCodFilial & SEPARADOR & objFilial.sNome
        GridClientes.TextMatrix(iIndice, iGrid_Historico_Col) = objFilialContatoData.sHistorico
'        GridClientes.TextMatrix(iIndice, iGrid_Fone1_Col) = objEnderecos.sTelefone1
'        GridClientes.TextMatrix(iIndice, iGrid_Fone2_Col) = objEnderecos.sTelefone2
'        GridClientes.TextMatrix(iIndice, iGrid_Contato_Col) = objEnderecos.sContato
        GridClientes.TextMatrix(iIndice, iGrid_ValorDevido_Col) = Format(objClieEst.dSaldoAtrasados, "STANDARD")
    
    Next
           
    objGridClientes.iLinhasExistentes = iIndice
    
    Call Grid_Refresh_Checkbox(objGridClientes)
    
    Call Ordenacao_Limpa(objGridClientes, Ordenacao)
    
    Call Combo_Seleciona_ItemData(Ordenacao, -iGrid_ValorDevido_Col)

    Call Soma_Coluna_Grid(objGridClientes, iGrid_ValorDevido_Col, ValorDevTotal, False)

    gbTrazendoDados = False

    Preenche_GridCliente = SUCESSO

    Exit Function

Erro_Preenche_GridCliente:

    gbTrazendoDados = False

    Preenche_GridCliente = gErr

    Select Case gErr
    
        Case 182105, 182106, 185940, 185940

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182107)

    End Select

    Exit Function

End Function

Private Sub TabStrip1_BeforeClick(Cancel As Integer)

Dim lErro As Long
Dim objCobrancaCliente As New ClassCobrancaSelCli

On Error GoTo Erro_TabStrip1_BeforeClick

    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
    
    'Se estava no tab de seleção e está passando para outro tab
    If iFrameAtual = TAB_SELECAO Then
    
        'Valida a seleção
        lErro = Move_Selecao_Memoria(objCobrancaCliente)
        If lErro <> SUCESSO Then gError 182108
        
        If objCobrancaCliente.dtDataPrevAte <> gobjCobrancaClienteAnt.dtDataPrevAte Or _
            objCobrancaCliente.dtDataPrevDe <> gobjCobrancaClienteAnt.dtDataPrevDe Or _
            objCobrancaCliente.dtDataProxAte <> gobjCobrancaClienteAnt.dtDataProxAte Or _
            objCobrancaCliente.dtDataProxDe <> gobjCobrancaClienteAnt.dtDataProxDe Or _
            objCobrancaCliente.dtDataVencAte <> gobjCobrancaClienteAnt.dtDataVencAte Or _
            objCobrancaCliente.dtDataVencDe <> gobjCobrancaClienteAnt.dtDataVencDe Or _
            objCobrancaCliente.sCobrador <> gobjCobrancaClienteAnt.sCobrador Or _
            objCobrancaCliente.iTitulosBaixados <> gobjCobrancaClienteAnt.iTitulosBaixados Or _
            objCobrancaCliente.sCategoria <> gobjCobrancaClienteAnt.sCategoria Or _
            objCobrancaCliente.sCategoriaDe <> gobjCobrancaClienteAnt.sCategoriaDe Or _
            objCobrancaCliente.sCategoriaAte <> gobjCobrancaClienteAnt.sCategoriaAte Or _
            objCobrancaCliente.iVendedor <> gobjCobrancaClienteAnt.iVendedor Or _
            objCobrancaCliente.iTipoVend <> gobjCobrancaClienteAnt.iTipoVend Then
        
            lErro = Trata_Selecao(objCobrancaCliente)
            If lErro <> SUCESSO Then gError 182109
            
            Set gobjCobrancaClienteAnt = objCobrancaCliente
            
        End If
    
    End If

    Exit Sub

Erro_TabStrip1_BeforeClick:

    Cancel = True

    Select Case gErr

        Case 182108, 182109

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182110)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        FrameTab(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        FrameTab(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
        
    End If

End Sub

Private Function Inicializa_GridClientes(objGrid As AdmGrid) As Long

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Ligar")
    objGrid.colColuna.Add ("Cliente")
    objGrid.colColuna.Add ("Filial")
    objGrid.colColuna.Add ("Vlr Devido")
'    objGrid.colColuna.Add ("Contato")
'    objGrid.colColuna.Add ("Telefone 1")
'    objGrid.colColuna.Add ("Telefone 2")
    objGrid.colColuna.Add ("Lig. Efetuada")
    objGrid.colColuna.Add ("Histórico")
    
    'Atualiza a Parte de Ordenação
    Call Ordenacao_Preeenche(objGrid, Ordenacao)
    
    'Controles que participam do Grid
    objGrid.colCampo.Add (CheckLigar.Name)
    objGrid.colCampo.Add (ClienteGrid.Name)
    objGrid.colCampo.Add (FilialClienteGrid.Name)
    objGrid.colCampo.Add (ValorDevidoEmAtraso.Name)
'    objGrid.colCampo.Add (ContatoGrid.Name)
'    objGrid.colCampo.Add (Fone1Grid.Name)
'    objGrid.colCampo.Add (Fone2Grid.Name)
    objGrid.colCampo.Add (CheckLigacaoRealizada.Name)
    objGrid.colCampo.Add (HistoricoGrid.Name)

    'Colunas do Grid
    iGrid_CheckLigar_Col = 1
    iGrid_Cliente_Col = 2
    iGrid_FilialCliente_Col = 3
    iGrid_ValorDevido_Col = 4
'    iGrid_Contato_Col = 4
'    iGrid_Fone1_Col = 5
'    iGrid_Fone2_Col = 6
    iGrid_CheckLigacaoRealizada_Col = 5 '7
    iGrid_Historico_Col = 6 '8

    objGrid.objGrid = GridClientes

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 100 + 1

    objGrid.iLinhasVisiveis = 10

    'Largura da primeira coluna
    GridClientes.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    Call Grid_Inicializa(objGrid)

    Inicializa_GridClientes = SUCESSO

End Function

Private Sub GridClientes_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridClientes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridClientes, iAlterado)
    End If
    
    If Not gbTrazendoDados Then Call Ordenacao_ClickGrid(objGridClientes, Ordenacao)

End Sub

Private Sub GridClientes_GotFocus()
    Call Grid_Recebe_Foco(objGridClientes)
End Sub

Private Sub GridClientes_EnterCell()
    Call Grid_Entrada_Celula(objGridClientes, iAlterado)
End Sub

Private Sub GridClientes_LeaveCell()
    Call Saida_Celula(objGridClientes)
End Sub

Private Sub GridClientes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridClientes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridClientes, iAlterado)
    End If

End Sub

Private Sub GridClientes_RowColChange()

Dim iLinhaAnt As Integer

    Call Grid_RowColChange(objGridClientes)
    
    If Not gbTrazendoDados Then
        If objGridClientes.iLinhaAntiga <> objGridClientes.objGrid.Row Then
            Call Trata_Contato
            objGridClientes.iLinhaAntiga = objGridClientes.objGrid.Row
        End If
    End If
    
End Sub

Private Sub GridClientes_Scroll()
    Call Grid_Scroll(objGridClientes)
End Sub

Private Sub GridClientes_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridClientes)
End Sub

Private Sub GridClientes_LostFocus()
    Call Grid_Libera_Foco(objGridClientes)
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then
        
        'OperacaoInsumos
        If objGridInt.objGrid.Name = GridClientes.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                Case iGrid_CheckLigar_Col

                    lErro = Saida_Celula_Ligar(objGridInt)
                    If lErro <> SUCESSO Then gError 182111

                Case iGrid_CheckLigacaoRealizada_Col

                    lErro = Saida_Celula_LigacaoReal(objGridInt)
                    If lErro <> SUCESSO Then gError 182112
                    
                Case iGrid_Historico_Col

                    lErro = Saida_Celula_Historico(objGridInt)
                    If lErro <> SUCESSO Then gError 182113

            End Select
                    
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 182114

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 182111 To 182113

        Case 182114
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182115)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Ligar(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Ligar do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Ligar

    Set objGridInt.objControle = CheckLigar

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 182116

    Saida_Celula_Ligar = SUCESSO

    Exit Function

Erro_Saida_Celula_Ligar:

    Saida_Celula_Ligar = gErr

    Select Case gErr
        
        Case 182116
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182117)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_LigacaoReal(objGridInt As AdmGrid) As Long
'Faz a crítica da célula LigacaoReal do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_LigacaoReal

    Set objGridInt.objControle = CheckLigacaoRealizada

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 182118

    Saida_Celula_LigacaoReal = SUCESSO

    Exit Function

Erro_Saida_Celula_LigacaoReal:

    Saida_Celula_LigacaoReal = gErr

    Select Case gErr
        
        Case 182118
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182119)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Historico(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Historico do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Historico

    Set objGridInt.objControle = HistoricoGrid

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 182120

    Saida_Celula_Historico = SUCESSO

    Exit Function

Erro_Saida_Celula_Historico:

    Saida_Celula_Historico = gErr

    Select Case gErr
        
        Case 182120
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182121)

    End Select

    Exit Function

End Function

Private Sub BotaoDesligar_Click()
'Desconecta a porta do modem

On Error GoTo Erro_BotaoLigar_Click

    If ComDiscar.PortOpen Then
       ComDiscar.PortOpen = False
        BotaoDesligar.Enabled = False
        DoEvents
    End If

    Exit Sub

Erro_BotaoLigar_Click:

    Select Case gErr

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182122)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoLigar_Click()
'Faz a ligação telefonica usando o modem

Dim sCOM As String
Dim sTelefone As String
Dim sTelefoneCompl As String
Dim iCont As Integer
Dim sOperDDDViaPABX As String

On Error GoTo Erro_BotaoLigar_Click

    iConTimer = 0
    Timer1.Enabled = True
    Timer2.Enabled = True
    Timer1.Interval = 100
    Timer2.Interval = 1

    If GridClientes.Row = 0 Then gError 182123
    If ComDiscar.PortOpen Then gError 182124
    If Len(Trim(Telefone.Caption)) = 0 Then gError 182345
    
    sCOM = String(128, 0)
    sOperDDDViaPABX = String(128, 0)
    Call GetPrivateProfileString("Geral", "modem", "3", sCOM, 128, "ADM100.INI")
    Call GetPrivateProfileString("Geral", "OperDDDViaPABX", "0", sOperDDDViaPABX, 128, "ADM100.INI")
    sOperDDDViaPABX = Replace(sOperDDDViaPABX, Chr(0), "")
    sCOM = Replace(sCOM, Chr(0), "")
       
    ComDiscar.CommPort = CInt(sCOM)
    ComDiscar.Settings = "9600,N,8,1"
    ComDiscar.PortOpen = True
    
    BotaoDesligar.Enabled = True
    DoEvents
    
    sTelefone = Telefone.Caption

    sTelefone = Replace(sTelefone, "(", "")
    sTelefone = Replace(sTelefone, ")", "")
    sTelefone = Replace(sTelefone, "-", "")
    sTelefone = Replace(sTelefone, " ", "")
    
    'Se tem Senha
    If PossuiSenha.Value = vbChecked Then
        sTelefoneCompl = sTelefoneCompl & "#" & Trim(SenhaTel.Text) & ","
    End If
    
    'Se tem dígito para ligações externas
    If Len(Trim(DigDiscExt.Text)) > 0 Then
        sTelefoneCompl = sTelefoneCompl & Trim(DigDiscExt.Text) & ","
    End If
    
    'Se é um DDD
    If Len(sTelefone) > 9 Then
    
        'Coloca o 0 na frente
        sTelefoneCompl = sTelefoneCompl & "0"
        
        'Se a operadora de DDD não é obtida do PABX
        If StrParaInt(sOperDDDViaPABX) = DESMARCADO Then
        
            'Coloca o código da operadora de DDD
            sTelefoneCompl = sTelefoneCompl & OperDDD.Text
        End If
    End If
    
    sTelefone = sTelefoneCompl & sTelefone
    
    ComDiscar.Output = "ATDT" & sTelefone & vbCr
    
    Call BotaoHistorico_Click
    
    Exit Sub

Erro_BotaoLigar_Click:

    If ComDiscar.PortOpen Then
        ComDiscar.PortOpen = False
        BotaoDesligar.Enabled = False
        DoEvents
    End If

    Select Case gErr
    
        Case 182123
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case 182124, 182346
             Call Rotina_Erro(vbOKOnly, "ERRO_SEM_SINAL_LINHA", gErr)
             
        Case 182345
             Call Rotina_Erro(vbOKOnly, "ERRO_TELEFONE_NAO_PREENCHIDO", gErr)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182125)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoGravarTela_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objFilialContatoData As ClassFilialContatoData
Dim colFilialContatoData As New Collection

On Error GoTo Erro_BotaoGravarTela_Click

    GL_objMDIForm.MousePointer = vbHourglass

    For iIndice = 1 To objGridClientes.iLinhasExistentes
    
        Set objFilialContatoData = New ClassFilialContatoData
    
        objFilialContatoData.dtData = gdtDataAtual
        objFilialContatoData.iLigacaoEfetuada = StrParaInt(GridClientes.TextMatrix(iIndice, iGrid_CheckLigacaoRealizada_Col))
        objFilialContatoData.iLigar = StrParaInt(GridClientes.TextMatrix(iIndice, iGrid_CheckLigar_Col))
        objFilialContatoData.lCliente = LCodigo_Extrai(GridClientes.TextMatrix(iIndice, iGrid_Cliente_Col))
        objFilialContatoData.iFilial = Codigo_Extrai(GridClientes.TextMatrix(iIndice, iGrid_FilialCliente_Col))
        objFilialContatoData.sHistorico = GridClientes.TextMatrix(iIndice, iGrid_Historico_Col)
        objFilialContatoData.sCodUsuario = comboCobrador.Text
    
        colFilialContatoData.Add objFilialContatoData
    
    Next
    
    lErro = CF("FilialContatoData_Grava", colFilialContatoData)
    If lErro <> SUCESSO Then gError 182148
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_GRAVACAO_COM_SUCESSO")
    
    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoGravarTela_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 182148

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182126)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 182127

    'Limpa a Tela
    lErro = Limpa_Tela_Cobranca
    If lErro <> SUCESSO Then gError 182128

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 182127, 182128

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182129)

    End Select

End Sub

Function Limpa_Tela_Cobranca() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Cobranca
        
    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    Call Grid_Limpa(objGridClientes)
    
    Set gobjCobrancaClienteAnt = New ClassCobrancaSelCli
    
    sOpcaoAnt = ""
    
    OpcoesTela.Text = ""
    comboCobrador.Text = gsUsuario
    ApenasQualifProx.ListIndex = -1
    EntreQualifProxDe.ListIndex = -1
    EntreQualifProxAte.ListIndex = -1
    ApenasQualifVenc.ListIndex = -1
    EntreQualifVencDe.ListIndex = -1
    EntreQualifVencAte.ListIndex = -1
    ApenasQualifPrev.ListIndex = -1
    EntreQualifPrevDe.ListIndex = -1
    EntreQualifPrevAte.ListIndex = -1
    
    '#####################################
    'Inserido por Wagner
    CategoriaClienteTodas.Value = vbChecked
    CategoriaCliente.Enabled = False
    CategoriaClienteDe.Enabled = False
    CategoriaClienteAte.Enabled = False
    CategoriaClienteDe.ListIndex = -1
    CategoriaClienteAte.ListIndex = -1
    '#####################################
    
    CheckTitAberto.Value = vbChecked
    
    FaixaDataPrev.Value = True
    FaixaDataProx.Value = True
    FaixaDataVenc.Value = True
    
    Call FrameD_Enabled(FrameD1, FRAMED_FAIXA)
    Call FrameD_Enabled(FrameD2, FRAMED_FAIXA)
    Call FrameD_Enabled(FrameD3, FRAMED_FAIXA)
    
    Call Ordenacao_Limpa(objGridClientes, Ordenacao)
    
    'Torna Frame atual invisível
    FrameTab(TabStrip1.SelectedItem.Index).Visible = False
    iFrameAtual = TAB_SELECAO
    'Torna Frame atual visível
    FrameTab(iFrameAtual).Visible = True
    TabStrip1.Tabs.Item(iFrameAtual).Selected = True
    
    ValorDevTotal.Caption = ""
    
    iAlterado = 0

    Limpa_Tela_Cobranca = SUCESSO

    Exit Function

Erro_Limpa_Tela_Cobranca:

    Limpa_Tela_Cobranca = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182130)

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
        If lErro <> SUCESSO Then gError 182131

        objDataMask.Text = sData

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 182131

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182132)

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
        If lErro <> SUCESSO Then gError 182133

        objDataMask.Text = sData

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 182133

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182134)

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

Private Function Carrega_Usuarios() As Long
'Carrega a Combo CodUsuarios com todos os usuários do BD

Dim lErro As Long
Dim colUsuarios As New Collection
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_Carrega_Usuarios

    lErro = CF("UsuariosFilialEmpresa_Le_Todos", colUsuarios)
    If lErro <> SUCESSO Then gError 182135

    For Each objUsuarios In colUsuarios
        comboCobrador.AddItem objUsuarios.sCodUsuario
    Next

    Carrega_Usuarios = SUCESSO

    Exit Function

Erro_Carrega_Usuarios:

    Carrega_Usuarios = gErr

    Select Case gErr

        Case 182135

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182136)

    End Select

    Exit Function

End Function

Private Sub Marca_Desmarca(ByVal bFlag As Boolean)

Dim iIndice As Integer

    For iIndice = 1 To objGridClientes.iLinhasExistentes
    
        If bFlag Then
            GridClientes.TextMatrix(iIndice, iGrid_CheckLigar_Col) = MARCADO
        Else
            GridClientes.TextMatrix(iIndice, iGrid_CheckLigar_Col) = DESMARCADO
        End If
    
    Next
    
    Call Grid_Refresh_Checkbox(objGridClientes)

End Sub

Private Sub BotaoDesmarcar_Click()
    Call Marca_Desmarca(False)
End Sub

Private Sub BotaoMarcar_Click()
    Call Marca_Desmarca(True)
End Sub

Private Sub OpcoesTela_Validate(Cancel As Boolean)
    'Se a opção não foi selecionada na combo => chama a função OpcoesTela_Click
    If OpcoesTela.ListIndex = -1 Then Call OpcoesTela_Click
End Sub

Public Function Gravar_Registro() As Long

Dim objTela As Object
Dim lErro As Long

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    Set objTela = Me
    
    lErro = CF("OpcoesTelas_Grava", objTela)
    If lErro <> SUCESSO Then gError 182149
    
    Call Limpa_Tela_Cobranca
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 182149
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182150)
        
    End Select
    
End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTela As Object

On Error GoTo Erro_BotaoExcluir_Click

    Set objTela = Me
    
    lErro = CF("OpcoesTelas_Exclui", objTela)
    If lErro <> SUCESSO Then gError 182151
    
    Call Limpa_Tela_Cobranca

    Exit Sub
    
Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 182151
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182152)

    End Select

End Sub

Private Sub OpcoesTela_Click()
    
Dim lErro As Long
Dim objTela As Object
Dim iCancel As Integer

On Error GoTo Erro_OpcoesTela_Click

    Set objTela = Me
    
    If sOpcaoAnt <> OpcoesTela.Text Then
    
        iAtualizaTela = MARCADO
        
        'Trata o evento click da combo opções
        lErro = CF("OpcoesTela_Click", objTela)
        If lErro <> SUCESSO Then gError 182156
        
        sOpcaoAnt = OpcoesTela.Text
        
        'Se Frame selecionado foi o de seleção e é para atualizar o grid
        If TabStrip1.SelectedItem.Index = TAB_CLIENTE Then
        
            iCancel = bSGECancelDummy
        
            Call TabStrip1_BeforeClick(iCancel)
        
        End If
        
    End If
    
    Exit Sub

Erro_OpcoesTela_Click:

    Select Case gErr

        Case 182156
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182157)

    End Select

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
        If lErro <> SUCESSO Then gError 182161
        
    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case 182161
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182162)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 182158
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case gErr

        Case 182158
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182159)

    End Select
    
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

Private Sub BotaoHistorico_Click()

Dim objHistCobr As New ClassHistoricoCobrSelCli

On Error GoTo Erro_BotaoHistorico_Click

    If GridClientes.Row = 0 Then gError 182163

    objHistCobr.iFilial = Codigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_FilialCliente_Col))
    objHistCobr.lCliente = LCodigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_Cliente_Col))
    objHistCobr.iContato = Codigo_Extrai(Contato.Text)

    Call Chama_Tela("HistoricoCliente", objHistCobr)

    Exit Sub

Erro_BotaoHistorico_Click:

    Select Case gErr
    
        Case 182163
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182164)

    End Select

    Exit Sub
    
End Sub

Private Sub Ordenacao_Change()

Dim colcolColecoes As New Collection

    Call Ordenacao_Atualiza(objGridClientes, Ordenacao, colcolColecoes)
    
End Sub

Private Sub Ordenacao_Click()

Dim colcolColecoes As New Collection

    Call Ordenacao_Atualiza(objGridClientes, Ordenacao, colcolColecoes)

End Sub

Public Sub ComboCobrador_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objUsuarios As New ClassUsuarios

On Error GoTo Erro_ComboCobrador_Validate
    
    'Verifica se algum codigo está selecionado
    If comboCobrador.ListIndex = -1 Then Exit Sub
    
    If Len(Trim(comboCobrador.Text)) > 0 Then
    
        'Coloca o código selecionado nos obj's
        objUsuarios.sCodUsuario = comboCobrador.Text
    
        'Le o nome do Usário
        lErro = CF("Usuarios_Le", objUsuarios)
        If lErro <> SUCESSO And lErro <> 40832 Then gError 182172
        
        If lErro <> SUCESSO Then gError 182173
        
    End If
    
    Exit Sub
    
Erro_ComboCobrador_Validate:

    Cancel = True

    Select Case gErr
            
        Case 182172
        
        Case 182173 'O usuário não está na tabela
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", gErr, objUsuarios.sCodUsuario)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182174)
    
    End Select
    
    Exit Sub
    
End Sub

Public Function Trata_Contato() As Long

Dim lErro As Long
Dim objClienteContatos As New ClassClienteContatos
Dim objFilialCliente As New ClassFilialCliente
Dim objEndereco As New ClassEndereco
Dim sOperDDD As String
Dim sFilialMaquina As String
Dim lCliente As Long
Dim iFilial

On Error GoTo Erro_Trata_Contato

    If GridClientes.Row <> 0 Then

        lCliente = LCodigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_Cliente_Col))
        iFilial = Codigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_FilialCliente_Col))

        'Guarda no objClienteContatos, o código do cliente e da
        objClienteContatos.lCliente = lCliente
        objClienteContatos.iFilialCliente = iFilial
        
        'Carrega a combo de contatos
        lErro = CF("Carrega_ClienteContatos", Contato, objClienteContatos)
        If lErro <> SUCESSO And lErro <> 102622 Then gError 182297
        
        'Se selecionou o contato padrão =>
        If Len(Trim(Contato.Text)) > 0 Then
        
            'traz o telefone do contato
            Call Contato_Click
        
        Else
        
            'Limpa o campo telefone
            Telefone.Caption = ""
            
        End If
        
        If lCliente <> 0 And iFilial <> 0 Then
        
            objFilialCliente.lCodCliente = lCliente
            objFilialCliente.iCodFilial = iFilial
        
            'Lê os dados da Filial Cliente
            lErro = CF("FilialCliente_Le", objFilialCliente)
            If lErro <> SUCESSO And lErro <> 12567 Then gError 185058
            
            objEndereco.lCodigo = objFilialCliente.lEndereco
        
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO And lErro <> 12309 Then gError 185059
            
            UF.Caption = objEndereco.sSiglaEstado
            
            sFilialMaquina = String(128, 0)
            
            Call GetPrivateProfileString("Geral", "FilialMaquina", "1", sFilialMaquina, 128, "ADM100.INI")
        
            sFilialMaquina = Replace(sFilialMaquina, Chr(0), "")
            
            lErro = CF("OperadorasDDD_Le", StrParaInt(sFilialMaquina), UF.Caption, sOperDDD)
            If lErro <> SUCESSO Then gError 185064
            
            OperDDD.PromptInclude = False
            OperDDD.Text = sOperDDD
            OperDDD.PromptInclude = True
            
        Else
        
            UF.Caption = ""
        
            OperDDD.PromptInclude = False
            OperDDD.Text = ""
            OperDDD.PromptInclude = True
        
        End If
        
    End If
    
    Trata_Contato = SUCESSO

    Exit Function

Erro_Trata_Contato:

    Trata_Contato = gErr

    Select Case gErr

        Case 182297, 185058, 185059
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182298)

    End Select

End Function

Private Sub Contato_Click()

Dim lErro As Long
Dim objClienteContatos As New ClassClienteContatos

On Error GoTo Erro_Contato_Click

    'Se o campo contato não foi preenchido => sai da função
    If Contato.ListIndex = -1 Then Exit Sub
    
    If GridClientes.Row = 0 Then gError 182297

    'Guarda o código do cliente e da filial no obj
    objClienteContatos.iFilialCliente = Codigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_FilialCliente_Col))
    objClienteContatos.lCliente = LCodigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_Cliente_Col))
    objClienteContatos.iCodigo = Codigo_Extrai(Contato.Text)

    'Lê o contato no BD
    lErro = CF("ClienteContatos_Le", objClienteContatos)
    If lErro <> SUCESSO And lErro <> 102653 Then gError 182298
    
    'Se não encontrou o contato => erro
    If lErro = 102653 Then gError 182299
    
    'Exibe o telefone cadastrado para o contato selecionado
    Telefone.Caption = objClienteContatos.sTelefone
    
    Exit Sub
    
Erro_Contato_Click:

    Select Case gErr

        Case 182297
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 182298
        
        Case 182299
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTECONTATO_NAO_ENCONTRADO", gErr, Contato.Text, objClienteContatos.lCliente, objClienteContatos.iFilialCliente)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182300)

    End Select
    

End Sub

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
End Sub

Private Sub Contato_Validate(Cancel As Boolean)
'Faz a validação da filial do cliente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objClienteContatos As New ClassClienteContatos
Dim iCodigo As Integer

On Error GoTo Erro_Contato_Validate

    'Se o contato foi preenchido
    If Len(Trim(Contato.Text)) > 0 Then
    
        If GridClientes.Row = 0 Then gError 182301
    
        'Guarda o código do cliente e da filial no obj
        objClienteContatos.iFilialCliente = Codigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_FilialCliente_Col))
        objClienteContatos.lCliente = LCodigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_Cliente_Col))
    
        'Se o contato foi selecionado na própria combo => sai da função
        If Contato.Text = Contato.List(Contato.ListIndex) Then Exit Sub
    
        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
        lErro = Combo_Seleciona(Contato, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 182302
    
        'Se não encontrou o contato na combo, mas retornou um código
        If lErro = 6730 Then

            objClienteContatos.iCodigo = iCodigo
            
            'Lê o contato a partir dos dados passados
            lErro = CF("ClienteContatos_Le", objClienteContatos)
            If lErro <> SUCESSO And lErro <> 102653 Then gError 182303
            
            'Se não encontrou o contato
            If lErro = 102653 Then gError 182304
            
            'Exibe o contato na tela
            Contato.Text = objClienteContatos.iCodigo & SEPARADOR & objClienteContatos.sContato
            
            'Exibe o telefone do contato
            Telefone.Caption = objClienteContatos.sTelefone
        
        End If
        
        'Se foi digitado o nome do contato
        'e esse nome não foi encontrado na combo => erro
        If lErro = 6731 Then

            objClienteContatos.sContato = Contato.Text
        
            'Lê o contato a partir dos dados passados
            lErro = CF("ClienteContatos_Le_Nome", objClienteContatos)
            If lErro <> SUCESSO And lErro <> 178440 Then gError 182305
            
            'Se não encontrou o contato
            If lErro = 178440 Then gError 182306
        
            'Exibe o contato na tela
            Contato.Text = objClienteContatos.iCodigo & SEPARADOR & objClienteContatos.sContato
            
            'Exibe o telefone do contato
            Telefone.Caption = objClienteContatos.sTelefone
        
        End If
    
    'Senão
    Else
    
        'Limpa o campo telefone
        Telefone.Caption = ""
    
    End If
    
    Exit Sub

Erro_Contato_Validate:

    Cancel = True

    Select Case gErr

        Case 182301
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 182302, 182303, 182305
            
        Case 182304, 182306
            
            'Verifica se o usuário deseja criar um novo contato
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTECONTATO", Contato.Text, GridClientes.TextMatrix(GridClientes.Row, iGrid_Cliente_Col), GridClientes.TextMatrix(GridClientes.Row, iGrid_Cliente_Col))

            'Se o usuário respondeu sim
            If vbMsgRes = vbYes Then
                'Chama a tela para cadastro de contatos
                Call Chama_Tela("ClienteContatos", objClienteContatos)
            End If
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182307)

    End Select

    Exit Sub

End Sub

Private Sub LabelContato_Click()

Dim lErro As Long
Dim objClienteContatos As New ClassClienteContatos

On Error GoTo Erro_LabelContato_Click

    If GridClientes.Row = 0 Then gError 182308

    'Guarda o código do cliente e da filial no obj
    objClienteContatos.iFilialCliente = Codigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_FilialCliente_Col))
    objClienteContatos.lCliente = LCodigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_Cliente_Col))
    
    Call Chama_Tela("ClienteContatos", objClienteContatos)
    
    Exit Sub

Erro_LabelContato_Click:

    Select Case gErr

        Case 182308
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 182309)

    End Select

    Exit Sub
    
End Sub

Private Sub HistoricoGrid_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub HistoricoGrid_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridClientes)
End Sub

Private Sub HistoricoGrid_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridClientes)
End Sub

Private Sub HistoricoGrid_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridClientes.objControle = HistoricoGrid
    lErro = Grid_Campo_Libera_Foco(objGridClientes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CheckLigar_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CheckLigar_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridClientes)
End Sub

Private Sub CheckLigar_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridClientes)
End Sub

Private Sub CheckLigar_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridClientes.objControle = CheckLigar
    lErro = Grid_Campo_Libera_Foco(objGridClientes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CheckLigacaoRealizada_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CheckLigacaoRealizada_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridClientes)
End Sub

Private Sub CheckLigacaoRealizada_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridClientes)
End Sub

Private Sub CheckLigacaoRealizada_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridClientes.objControle = CheckLigacaoRealizada
    lErro = Grid_Campo_Libera_Foco(objGridClientes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub BotaoCliente_Click()

Dim objCliente As New ClassCliente

On Error GoTo Erro_BotaoCliente_Click

    If GridClientes.Row = 0 Then gError 182163

    objCliente.lCodigo = LCodigo_Extrai(GridClientes.TextMatrix(GridClientes.Row, iGrid_Cliente_Col))

    Call Chama_Tela("Clientes", objCliente)

    Exit Sub

Erro_BotaoCliente_Click:

    Select Case gErr
    
        Case 182163
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182164)

    End Select

    Exit Sub
    
End Sub

Private Sub Timer1_Timer()

On Error GoTo Erro_Timer1_Timer

    If sMsg <> "" Then

        If InStr(1, sMsg, "NO DIALTONE") <> 0 Then gError 182346
'        If InStr(1, sMsg, "NO ANSWER") <> 0 Then gError 182346
        If InStr(1, sMsg, "BUSY") <> 0 Then gError 182705
'        If InStr(1, sMsg, "NO CARRIER") <> 0 Then gError 182346
    
        If InStr(1, sMsg, "RING BACK") <> 0 Then
            Call Rotina_Aviso(vbOKOnly, "AVISO_TEL_TOCANDO")
            Timer1.Interval = 0
            Timer2.Interval = 0
            sMsg = ""
        End If
        
    End If
    
    iConTimer = iConTimer + 1
    
    If iConTimer > 50 Then
        'gError 182346
        
        Call Rotina_Aviso(vbOKOnly, "AVISO_TEL_TOCANDO")
        Timer1.Interval = 0
        Timer2.Interval = 0
        sMsg = ""
    End If

    Exit Sub

Erro_Timer1_Timer:

    sMsg = ""
    Timer1.Interval = 0
    Timer2.Interval = 0
    
    If ComDiscar.PortOpen Then
        ComDiscar.PortOpen = False
        BotaoDesligar.Enabled = False
    End If

    Select Case gErr

        Case 182346
             Call Rotina_Erro(vbOKOnly, "ERRO_SEM_SINAL_LINHA", gErr)
             
        Case 182705
             Call Rotina_Erro(vbOKOnly, "ERRO_TEL_OCUPADO", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182164)

    End Select

    Exit Sub

End Sub

Private Sub Timer2_Timer()

On Error GoTo Erro_Timer2_Timer

    If ComDiscar.PortOpen Then

        sMsg = sMsg + ComDiscar.Input
        
    End If

    Exit Sub

Erro_Timer2_Timer:

    Select Case gErr
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182164)

    End Select

    Exit Sub

End Sub

Private Sub PossuiSenha_Click()

    If PossuiSenha.Value = vbChecked Then
        SenhaTel.Enabled = True
    Else
        SenhaTel.Enabled = False
        SenhaTel.Text = ""
    End If

End Sub


'###################################################################
'Inserido por Wagner
Private Sub CategoriaCliente_Click()

Dim lErro As Long

On Error GoTo Erro_CategoriaCliente_Click

    If Len(Trim(CategoriaCliente.Text)) > 0 Then
        CategoriaClienteDe.Enabled = True
        CategoriaClienteAte.Enabled = True
        Call Carrega_ComboCategoriaItens(CategoriaCliente, CategoriaClienteDe)
        Call Carrega_ComboCategoriaItens(CategoriaCliente, CategoriaClienteAte)
    Else
        CategoriaClienteDe.Enabled = False
        CategoriaClienteAte.Enabled = False
    End If


    Exit Sub

Erro_CategoriaCliente_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168906)

    End Select

    Exit Sub

End Sub

Private Sub Carrega_ComboCategoriaCliente(ByVal objCombo As ComboBox)

Dim lErro As Long
Dim colCategoriaCliente As New Collection
Dim objCategoriaCliente As New ClassCategoriaCliente

On Error GoTo Erro_Carrega_ComboCategoriaCliente

    'Le as categorias de cliente
    lErro = CF("CategoriaCliente_Le_Todos", colCategoriaCliente)
    If lErro <> SUCESSO Then gError 131995

    'Preenche CategoriaCliente
    For Each objCategoriaCliente In colCategoriaCliente

        objCombo.AddItem objCategoriaCliente.sCategoria

    Next
    
    Exit Sub

Erro_Carrega_ComboCategoriaCliente:

    Select Case gErr
    
        Case 131995

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168907)

    End Select

    Exit Sub

End Sub

Private Sub Carrega_ComboCategoriaItens(ByVal objComboCategoria As ComboBox, ByVal objComboItens As ComboBox)

Dim lErro As Long
Dim iIndice As Integer
Dim objCategoriaClienteItem As ClassCategoriaClienteItem
Dim objCategoriaCliente As New ClassCategoriaCliente
Dim colCategoria As New Collection

On Error GoTo Erro_Carrega_ComboCategoriaItens

    'Verifica se a CategoriaCliente foi preenchida
    If objComboCategoria.ListIndex <> -1 Then

        objCategoriaCliente.sCategoria = objComboCategoria.Text

        'Lê os dados de Itens da Categoria do Cliente
        lErro = CF("CategoriaCliente_Le_Itens", objCategoriaCliente, colCategoria)
        If lErro <> SUCESSO Then gError 131994

        objComboItens.Enabled = True

        'Limpa os dados de ItemCategoriaCliente
        objComboItens.Clear

        'Preenche ItemCategoriaCliente
        For Each objCategoriaClienteItem In colCategoria

            objComboItens.AddItem objCategoriaClienteItem.sItem

        Next
        
        CategoriaClienteTodas.Value = vbFalse
    
    Else
        
        'Senão Desablita ItemCategoriaCliente
        objComboItens.ListIndex = -1
        objComboItens.Enabled = False
    
    End If
    
    Exit Sub

Erro_Carrega_ComboCategoriaItens:

    Select Case gErr
    
        Case 131993

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168908)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaCliente_Validate(Cancel As Boolean)

Dim lErro As Long, iCodigo As Integer

On Error GoTo Erro_CategoriaCliente_Validate

    If Len(CategoriaCliente.Text) <> 0 And CategoriaCliente.ListIndex = -1 Then
    
        'pesquisa a categoria na lista
        lErro = Combo_Item_Igual(CategoriaCliente)
        If lErro <> SUCESSO And lErro <> 12253 Then gError 131998
        
        If lErro <> SUCESSO Then gError 131999
    
    End If
    
    'Se a CategoriaCliente estiver em branco desabilita e limpa a combo
    If Len(CategoriaCliente.Text) = 0 Then
        CategoriaClienteDe.Enabled = False
        CategoriaClienteDe.Clear
        CategoriaClienteAte.Enabled = False
        CategoriaClienteAte.Clear
    End If
    
    Exit Sub

Erro_CategoriaCliente_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 131998
         
        Case 131999
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA", gErr, CategoriaCliente.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168909)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaClienteItem_Validate(Cancel As Boolean, objCombo As ComboBox)

Dim lErro As Long

On Error GoTo Erro_CategoriaClienteItem_Validate

    If Len(objCombo.Text) <> 0 Then
    
        'pesquisa o item na lista
        lErro = Combo_Item_Igual(objCombo)
        If lErro <> SUCESSO And lErro <> 12253 Then gError 131996
        
        If lErro <> SUCESSO Then gError 131997
    
    End If

    Exit Sub

Erro_CategoriaClienteItem_Validate:

    Cancel = True

    Select Case gErr

        Case 131996
        
        Case 131997
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTEITEM_INEXISTENTE", gErr, objCombo.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 168910)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaClienteTodas_Click()

Dim lErro As Long

On Error GoTo Erro_CategoriaClienteTodas_Click

    If CategoriaClienteTodas.Value = vbChecked Then
        'Desabilita o combotipo
        CategoriaCliente.ListIndex = -1
        CategoriaCliente.Enabled = False
        CategoriaClienteDe.Clear
        CategoriaClienteAte.Clear
    Else
        CategoriaCliente.Enabled = True
    End If

    Call CategoriaCliente_Click

    Exit Sub

Erro_CategoriaClienteTodas_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 168911)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaClienteAte_Validate(Cancel As Boolean)
    Call CategoriaClienteItem_Validate(Cancel, CategoriaClienteAte)
End Sub


Private Sub CategoriaClienteDe_Validate(Cancel As Boolean)
    Call CategoriaClienteItem_Validate(Cancel, CategoriaClienteDe)
End Sub
'####################################################################
