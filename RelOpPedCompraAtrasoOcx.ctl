VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpPedCompraAtrasoOcx 
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9105
   ScaleHeight     =   4455
   ScaleWidth      =   9105
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2760
      Index           =   2
      Left            =   630
      TabIndex        =   52
      Top             =   1305
      Visible         =   0   'False
      Width           =   7890
      Begin VB.Frame Frame6 
         Caption         =   "Destinatários"
         Height          =   1050
         Left            =   135
         TabIndex        =   58
         Top             =   1440
         Width           =   7515
         Begin VB.Frame Frame7 
            Caption         =   "Tipo"
            Height          =   555
            Left            =   90
            TabIndex        =   64
            Top             =   270
            Width           =   3930
            Begin VB.OptionButton TipoDestino 
               Caption         =   "Fornecedor"
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
               Index           =   2
               Left            =   2580
               TabIndex        =   19
               Top             =   225
               Width           =   1305
            End
            Begin VB.OptionButton TipoDestino 
               Caption         =   "Filial Empresa"
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
               Index           =   1
               Left            =   990
               TabIndex        =   18
               Top             =   225
               Value           =   -1  'True
               Width           =   1515
            End
            Begin VB.OptionButton TipoDestino 
               Caption         =   "Todos"
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
               Index           =   0
               Left            =   105
               TabIndex        =   17
               Top             =   225
               Width           =   870
            End
         End
         Begin VB.Frame FrameTipo 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   780
            Index           =   1
            Left            =   4140
            TabIndex        =   59
            Top             =   180
            Width           =   3255
            Begin VB.ComboBox FilialEmpresa 
               Height          =   315
               ItemData        =   "RelOpPedCompraAtrasoOcx.ctx":0000
               Left            =   765
               List            =   "RelOpPedCompraAtrasoOcx.ctx":0002
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   270
               Width           =   2160
            End
            Begin VB.Label LabelFilialEmpDestino 
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
               Height          =   285
               Left            =   180
               TabIndex        =   60
               Top             =   285
               Width           =   465
            End
         End
         Begin VB.Frame FrameTipo 
            BorderStyle     =   0  'None
            Height          =   750
            Index           =   2
            Left            =   4005
            TabIndex        =   61
            Top             =   180
            Visible         =   0   'False
            Width           =   3420
            Begin VB.ComboBox FilialFornecedor 
               Height          =   315
               Left            =   1230
               TabIndex        =   22
               Top             =   435
               Width           =   2160
            End
            Begin MSMask.MaskEdBox Fornecedor 
               Height          =   300
               Left            =   1215
               TabIndex        =   21
               Top             =   75
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin VB.Label LabelFornDestino 
               AutoSize        =   -1  'True
               Caption         =   "Fornecedor:"
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
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   63
               Top             =   135
               Width           =   1035
            End
            Begin VB.Label LabelFilialFornDestino 
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
               Height          =   195
               Left            =   690
               TabIndex        =   62
               Top             =   495
               Width           =   465
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Filial Empresa"
         Height          =   1110
         Left            =   135
         TabIndex        =   53
         Top             =   180
         Width           =   7500
         Begin MSMask.MaskEdBox CodigoFilialDe 
            Height          =   300
            Left            =   1230
            TabIndex        =   13
            Top             =   285
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodigoFilialAte 
            Height          =   300
            Left            =   4830
            TabIndex        =   14
            Top             =   240
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NomeFilialAte 
            Height          =   300
            Left            =   4830
            TabIndex        =   16
            Top             =   660
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NomeFilialDe 
            Height          =   300
            Left            =   1230
            TabIndex        =   15
            Top             =   690
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin VB.Label LabelNomeAte 
            AutoSize        =   -1  'True
            Caption         =   "Nome Até:"
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
            Left            =   3915
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   57
            Top             =   720
            Width           =   900
         End
         Begin VB.Label LabelCodigoAte 
            AutoSize        =   -1  'True
            Caption         =   "Código Até:"
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
            Left            =   3810
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   56
            Top             =   300
            Width           =   1005
         End
         Begin VB.Label LabelNomeDe 
            AutoSize        =   -1  'True
            Caption         =   "Nome De:"
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
            Left            =   360
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   55
            Top             =   750
            Width           =   855
         End
         Begin VB.Label LabelCodigoDe 
            AutoSize        =   -1  'True
            Caption         =   "Código De:"
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
            Left            =   255
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   54
            Top             =   330
            Width           =   960
         End
      End
   End
   Begin VB.ComboBox ComboOrdenacao 
      Height          =   315
      ItemData        =   "RelOpPedCompraAtrasoOcx.ctx":0004
      Left            =   1575
      List            =   "RelOpPedCompraAtrasoOcx.ctx":0017
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   495
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpPedCompraAtrasoOcx.ctx":005B
      Left            =   1575
      List            =   "RelOpPedCompraAtrasoOcx.ctx":005D
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   3135
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4950
      Picture         =   "RelOpPedCompraAtrasoOcx.ctx":005F
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   135
      Width           =   1635
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6840
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpPedCompraAtrasoOcx.ctx":0161
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpPedCompraAtrasoOcx.ctx":02DF
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpPedCompraAtrasoOcx.ctx":0811
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpPedCompraAtrasoOcx.ctx":099B
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pedidos de Compra"
      Height          =   2565
      Index           =   1
      Left            =   765
      TabIndex        =   32
      Top             =   1395
      Width           =   7545
      Begin VB.Frame FrameCodigo 
         Caption         =   "Código"
         Height          =   675
         Left            =   105
         TabIndex        =   49
         Top             =   195
         Width           =   3270
         Begin MSMask.MaskEdBox CodPCDe 
            Height          =   300
            Left            =   525
            TabIndex        =   2
            Top             =   225
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodPCAte 
            Height          =   300
            Left            =   2100
            TabIndex        =   3
            Top             =   240
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin VB.Label LabelCodPCDe 
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
            Left            =   180
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   51
            Top             =   300
            Width           =   315
         End
         Begin VB.Label LabelCodPCAte 
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
            Left            =   1650
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   50
            Top             =   300
            Width           =   360
         End
      End
      Begin VB.Frame FrameNome 
         Caption         =   "Data de Envio"
         Height          =   705
         Left            =   3465
         TabIndex        =   44
         Top             =   960
         Width           =   3975
         Begin MSComCtl2.UpDown UpDownDataEnvioDe 
            Height          =   315
            Left            =   1680
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   225
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEnvioDe 
            Height          =   315
            Left            =   495
            TabIndex        =   8
            Top             =   240
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataEnvioAte 
            Height          =   315
            Left            =   3585
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   225
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEnvioAte 
            Height          =   315
            Left            =   2400
            TabIndex        =   9
            Top             =   225
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelNomeReqAte 
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
            Left            =   2040
            TabIndex        =   48
            Top             =   315
            Width           =   360
         End
         Begin VB.Label LabelNomeReqDe 
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
            Left            =   165
            TabIndex        =   47
            Top             =   315
            Width           =   315
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Data"
         Height          =   690
         Left            =   3465
         TabIndex        =   39
         Top             =   195
         Width           =   3945
         Begin MSComCtl2.UpDown UpDownDataDe 
            Height          =   315
            Left            =   1665
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataDe 
            Height          =   315
            Left            =   480
            TabIndex        =   4
            Top             =   255
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataAte 
            Height          =   315
            Left            =   3585
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataAte 
            Height          =   315
            Left            =   2400
            TabIndex        =   5
            Top             =   255
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label2 
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
            Left            =   165
            TabIndex        =   43
            Top             =   315
            Width           =   315
         End
         Begin VB.Label Label3 
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
            Left            =   2025
            TabIndex        =   42
            Top             =   315
            Width           =   360
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Fornecedores"
         Height          =   705
         Left            =   105
         TabIndex        =   36
         Top             =   960
         Width           =   3270
         Begin MSMask.MaskEdBox FornecedorDe 
            Height          =   300
            Left            =   525
            TabIndex        =   6
            Top             =   255
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FornecedorAte 
            Height          =   300
            Left            =   2115
            TabIndex        =   7
            Top             =   255
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin VB.Label LabelFornecedorDe 
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
            Left            =   165
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   38
            Top             =   315
            Width           =   315
         End
         Begin VB.Label LabelFornecedorAte 
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
            Left            =   1680
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   37
            Top             =   315
            Width           =   360
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Compradores"
         Height          =   705
         Left            =   105
         TabIndex        =   33
         Top             =   1740
         Width           =   3255
         Begin MSMask.MaskEdBox CompradorDe 
            Height          =   300
            Left            =   525
            TabIndex        =   10
            Top             =   255
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CompradorAte 
            Height          =   300
            Left            =   2115
            TabIndex        =   11
            Top             =   255
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label LabelCompradorDe 
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
            Left            =   165
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   35
            Top             =   315
            Width           =   315
         End
         Begin VB.Label LabelCompradorAte 
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
            Left            =   1680
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   34
            Top             =   315
            Width           =   360
         End
      End
      Begin VB.CheckBox CheckItens 
         Caption         =   "Exibe Item a Item"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3660
         TabIndex        =   12
         Top             =   1980
         Width           =   2070
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3255
      Left            =   540
      TabIndex        =   31
      Top             =   945
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   5741
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pedido"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Continuação"
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
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Ordenados Por:"
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
      TabIndex        =   30
      Top             =   585
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   225
      TabIndex        =   29
      Top             =   150
      Width           =   660
   End
End
Attribute VB_Name = "RelOpPedCompraAtrasoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'RelOpPedCompraAtraso
Const ORD_POR_CODIGO = 0
Const ORD_POR_DATA = 1
Const ORD_POR_FORNECEDOR = 2
Const ORD_POR_COMPRADOR = 3

Private WithEvents objEventoCodPCDe As AdmEvento
Attribute objEventoCodPCDe.VB_VarHelpID = -1
Private WithEvents objEventoCodPCAte As AdmEvento
Attribute objEventoCodPCAte.VB_VarHelpID = -1
Private WithEvents objEventoFornecedorDe As AdmEvento
Attribute objEventoFornecedorDe.VB_VarHelpID = -1
Private WithEvents objEventoFornecedorAte As AdmEvento
Attribute objEventoFornecedorAte.VB_VarHelpID = -1
Private WithEvents objEventoCompradorDe As AdmEvento
Attribute objEventoCompradorDe.VB_VarHelpID = -1
Private WithEvents objEventoCompradorAte As AdmEvento
Attribute objEventoCompradorAte.VB_VarHelpID = -1
Private WithEvents objEventoCodFilialDe As AdmEvento
Attribute objEventoCodFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoCodFilialAte As AdmEvento
Attribute objEventoCodFilialAte.VB_VarHelpID = -1
Private WithEvents objEventoNomeFilialDe As AdmEvento
Attribute objEventoNomeFilialDe.VB_VarHelpID = -1
Private WithEvents objEventoNomeFilialAte As AdmEvento
Attribute objEventoNomeFilialAte.VB_VarHelpID = -1
Private WithEvents objEventoFornDestino As AdmEvento
Attribute objEventoFornDestino.VB_VarHelpID = -1

Dim iFrameAtual As Integer
Dim iAlterado As Integer
Dim giTipoDestinoAtual  As Integer
Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 72234
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 72235

    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 72235
        
        Case 72234
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170826)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Private Sub Limpa_Tela_Rel()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Rel
  
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 68879
    
    ComboOrdenacao.ListIndex = 0
    ComboOpcoes.Text = ""
    ComboOpcoes.SetFocus
    CheckItens.Value = vbUnchecked
    FilialEmpresa.ListIndex = 0
    
    Exit Sub
    
Erro_Limpa_Tela_Rel:
    
    Select Case gErr
    
        Case 68879
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170827)

    End Select

    Exit Sub
   
End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_Rel

End Sub


Public Sub Form_Load()

Dim lErro As Long
Dim sMascaraCcl As String
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Form_Load
    
    Set objEventoCodFilialDe = New AdmEvento
    Set objEventoCodFilialAte = New AdmEvento
        
    Set objEventoFornDestino = New AdmEvento
    
    Set objEventoNomeFilialDe = New AdmEvento
    Set objEventoNomeFilialAte = New AdmEvento
        
    Set objEventoCodPCDe = New AdmEvento
    Set objEventoCodPCAte = New AdmEvento
        
    Set objEventoFornecedorDe = New AdmEvento
    Set objEventoFornecedorAte = New AdmEvento
        
    Set objEventoCompradorDe = New AdmEvento
    Set objEventoCompradorAte = New AdmEvento
        
    'Lê o Código e o NOme de Toda FilialEmpresa do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then gError 72560

    'Carrega a combo de Filial Empresa
    For Each objCodigoNome In colCodigoNome
        FilialEmpresa.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        FilialEmpresa.ItemData(FilialEmpresa.NewIndex) = objCodigoNome.iCodigo
    Next

    giTipoDestinoAtual = 1
    
    iFrameAtual = 1
    
    ComboOrdenacao.ListIndex = 0
    
    FilialEmpresa.ListIndex = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 72560
            'erro tratado na rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170828)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoCodFilialDe = Nothing
    Set objEventoCodFilialAte = Nothing
        
    Set objEventoFornDestino = Nothing
    
    Set objEventoNomeFilialDe = Nothing
    Set objEventoNomeFilialAte = Nothing
        
    Set objEventoCodPCDe = Nothing
    Set objEventoCodPCAte = Nothing
        
    Set objEventoFornecedorDe = Nothing
    Set objEventoFornecedorAte = Nothing
        
    Set objEventoCompradorDe = Nothing
    Set objEventoCompradorAte = Nothing
    
End Sub

Private Sub CodigoFilialAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoFilialAte, iAlterado)
    
End Sub

Private Sub CodigoFilialDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoFilialDe, iAlterado)
    
End Sub

Private Sub CodPCAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodPCAte, iAlterado)
    
End Sub

Private Sub CodPCDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodPCDe, iAlterado)
    
End Sub

Private Sub CompradorAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(CompradorAte, iAlterado)
    
End Sub

Private Sub CompradorDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(CompradorDe, iAlterado)
    
End Sub

Private Sub DataAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)
    
End Sub

Private Sub DataDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)
    
End Sub

Private Sub DataEnvioAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEnvioAte, iAlterado)
    
End Sub

Private Sub DataEnvioDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEnvioDe, iAlterado)
    
End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objFornecedor As New ClassFornecedor
Dim lCodigo As Long

On Error GoTo Erro_Fornecedor_Validate

    'Verifica se Fornec esta preenchido
    If Len(Trim(Fornecedor.Text)) <> 0 Then

        'Le Fornecedor
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
        If lErro <> SUCESSO Then gError 72563

        'Le as filiais do Fornecedor
        lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
        If lErro <> SUCESSO And lErro <> 6698 Then gError 72564

        'Preenche a combo de filiais
        Call CF("Filial_Preenche", FilialFornecedor, colCodigoNome)

        'Seleciona a filial na combo de filiais
        Call CF("Filial_Seleciona", FilialFornecedor, iCodFilial)

    Else
        'Limpa a combobox
        FilialFornecedor.Clear

    End If
  
    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 72563, 72564

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170829)

    End Select

    Exit Sub

End Sub
Private Sub FilialFornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim vbMsgRes As VbMsgBoxResult
Dim sNomeRed As String
Dim objEndereco As New ClassEndereco
Dim objPais As New ClassPais

On Error GoTo Erro_FilialFornecedor_Validate

    'Verifica se FilialFornecedor esta preenchida
    If Len(Trim(FilialFornecedor.Text)) > 0 Then

        'Verifica se FilialFornecedor esta selecionada
        If FilialFornecedor.ListIndex <> -1 Then Exit Sub

        'Seleciona combo box de FilialFornecedor
        lErro = Combo_Seleciona(FilialFornecedor, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 72565

        'Se nao encontra o ítem com o código informado
        If lErro = 6730 Then

            'Verifica de o fornecedor foi digitado
            If Len(Trim(Fornecedor.ClipText)) = 0 Then gError 72566

            sNomeRed = Fornecedor.Text

            objFilialFornecedor.iCodFilial = iCodigo

            'Pesquisa se existe filial com o codigo extraido
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then gError 72567

            If lErro = 18272 Then gError 72568

            'Coloca na tela Codigo e Nome Reduzido de FilialFornec
            FilialFornecedor.Text = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome

        End If

        'Não encontrou valor informado que era STRING
        If lErro = 6731 Then gError 72569

    End If

    Exit Sub

Erro_FilialFornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 72566
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 72567

        Case 72568, 72569
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", gErr, objFilialFornecedor.sNome)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170830)

    End Select

    Exit Sub

End Sub

Private Sub FornecedorAte_GotFocus()

    Call MaskEdBox_TrataGotFocus(FornecedorAte, iAlterado)
    
End Sub

Private Sub FornecedorDe_GotFocus()

    Call MaskEdBox_TrataGotFocus(FornecedorDe, iAlterado)
    
End Sub

Private Sub LabelCodPCAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objPedCompra As New ClassPedidoCompras

On Error GoTo Erro_LabelCodPCAte_Click

    If Len(Trim(CodPCAte.Text)) > 0 Then
        'Preenche com o Pedido de Compra da tela
        objPedCompra.lCodigo = StrParaLong(CodPCAte.Text)
    End If

    'Chama Tela PedComprasTodosLista
    Call Chama_Tela("PedComprasTodosLista", colSelecao, objPedCompra, objEventoCodPCAte)

   Exit Sub

Erro_LabelCodPCAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170831)

    End Select

    Exit Sub

End Sub
Private Sub LabelCodPCDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objPedCompra As New ClassPedidoCompras

On Error GoTo Erro_LabelCodPCDe_Click

    If Len(Trim(CodPCDe.Text)) > 0 Then
        'Preenche com o Pedido de Compra da tela
        objPedCompra.lCodigo = StrParaLong(CodPCDe.Text)
    End If

    'Chama Tela PedComprasTodosLista
    Call Chama_Tela("PedComprasTodosLista", colSelecao, objPedCompra, objEventoCodPCDe)

   Exit Sub

Erro_LabelCodPCDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170832)

    End Select

    Exit Sub

End Sub

Private Sub DataEnvioDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEnvioDe_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataEnvioDe.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataEnvioDe.Text)
    If lErro <> SUCESSO Then gError 72237

    Exit Sub
                   
Erro_DataEnvioDe_Validate:

    Cancel = True

    Select Case gErr

        Case 72237
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170833)

    End Select

    Exit Sub

End Sub

Private Sub DataEnvioAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEnvioAte_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataEnvioAte.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataEnvioAte.Text)
    If lErro <> SUCESSO Then gError 72238

    Exit Sub
                   
Erro_DataEnvioAte_Validate:

    Cancel = True

    Select Case gErr

        Case 72238
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170834)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataAte.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError 72239

    Exit Sub
                   
Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 72239
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170835)

    End Select

    Exit Sub

End Sub

Private Sub LabelFornDestino_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As New Collection

    objFornecedor.sNomeReduzido = Fornecedor.Text

    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornDestino)

    Exit Sub
    
End Sub

Private Sub objEventoFornDestino_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor
Dim bCancel As Boolean

    Set objFornecedor = obj1

    Fornecedor.Text = objFornecedor.sNomeReduzido
    'Dispara Validate de Fornecedor
    bCancel = False
    Fornecedor_Validate (bCancel)

    Me.Show

End Sub

Private Sub TabStrip1_Click()

     'Se frame atual corresponde ao tab selecionado, sai da rotina
    If TabStrip1.SelectedItem.Index = iFrameAtual Then Exit Sub

    'Torna Frame correspondente ao Tab selecionado visivel
    Frame1(TabStrip1.SelectedItem.Index).Visible = True

    'Torna Frame atual invisivel
    Frame1(iFrameAtual).Visible = False

    'Armazena novo valor de iFrameAtual
    iFrameAtual = TabStrip1.SelectedItem.Index


End Sub

Private Sub TipoDestino_Click(Index As Integer)

Dim lErro As Long

On Error GoTo Erro_TipoDestino_Click

    If Index = giTipoDestinoAtual Then Exit Sub

    If Index <> 0 Then
        
        FilialFornecedor.Enabled = True
        Fornecedor.Enabled = True
        LabelFornDestino.Enabled = True
        LabelFilialFornDestino.Enabled = True
        FilialEmpresa.Enabled = True
        LabelFilialEmpDestino.Enabled = True
        FrameTipo(1).Visible = False
        
        'Torna Frame correspondente a Index visivel
        FrameTipo(Index).Visible = True

        'Torna Frame atual invisivel
        If giTipoDestinoAtual <> 0 Then FrameTipo(giTipoDestinoAtual).Visible = False

        'Armazena novo valor de iFrameTipoDestinoAtual
        giTipoDestinoAtual = Index

        If Index <> 1 Then

            FrameTipo(Index - 1).Visible = False
            FrameTipo(Index).Visible = True
            
            'Verifica se o Fornecedor e sua Filial estão preenchidos
            If Len(Trim(Fornecedor.Text)) > 0 And Len(Trim(FilialFornecedor.Text)) > 0 Then
            
                FilialFornecedor_Click

            End If
        Else

            FrameTipo(Index + 1).Visible = False
            FrameTipo(Index).Visible = True
            Call CF("Filial_Seleciona", FilialEmpresa, giFilialEmpresa)

        End If

    End If
    If Index = 0 Then
    
        FilialEmpresa.Enabled = False
        FilialEmpresa.ListIndex = -1
        LabelFilialEmpDestino.Enabled = False
        Fornecedor.Enabled = False
        Fornecedor.Text = ""
        FilialFornecedor.Enabled = False
        FilialFornecedor.Text = ""
        LabelFornDestino.Enabled = False
        LabelFilialFornDestino.Enabled = False
        giTipoDestinoAtual = Index
        
    End If
    
    Exit Sub

Erro_TipoDestino_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170836)

    End Select

    Exit Sub

End Sub

Private Sub FilialFornecedor_Click()

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objEndereco As New ClassEndereco
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialFornecedor_Click

    If FilialFornecedor.ListIndex = -1 Then Exit Sub
    
    objFilialFornecedor.iCodFilial = FilialFornecedor.ItemData(FilialFornecedor.ListIndex)

    'Busca no BD a FilialFornecedor
    lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilialFornecedor)
    If lErro <> SUCESSO And lErro <> 18272 Then gError 72561

    If lErro = 18272 Then gError 72562

    Exit Sub

Erro_FilialFornecedor_Click:

    Select Case gErr

        Case 72561

        Case 72562
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", gErr, FilialFornecedor.Text)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170837)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEnvioDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvioDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEnvioDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 72240

    Exit Sub

Erro_UpDownDataEnvioDe_DownClick:

    Select Case gErr

        Case 72240
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170838)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEnvioDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvioDe_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEnvioDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 72241

    Exit Sub

Erro_UpDownDataEnvioDe_UpClick:

    Select Case gErr

        Case 72241
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170839)

    End Select

    Exit Sub

End Sub
Private Sub UpDownDataEnvioAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvioAte_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEnvioAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 72242

    Exit Sub

Erro_UpDownDataEnvioAte_DownClick:

    Select Case gErr

        Case 72242
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170840)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEnvioAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEnvioAte_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataEnvioAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 72243

    Exit Sub

Erro_UpDownDataEnvioAte_UpClick:

    Select Case gErr

        Case 72243
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170841)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 72244

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 72244
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170842)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 72245

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 72245
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170843)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 72246

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 72246
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170844)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    'Diminui um dia em DataDe
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 72248

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 72248
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 170845)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Verifica se a DataDe está preenchida
    If Len(Trim(DataDe.Text)) = 0 Then Exit Sub

    'Critica a DataDe informada
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then gError 72249

    Exit Sub
                   
Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 72249
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170846)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigoDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelCodigoDe_Click

    If Len(Trim(CodigoFilialDe.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.iCodFilial = StrParaInt(CodigoFilialDe.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoCodFilialDe)

   Exit Sub

Erro_LabelCodigoDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170847)

    End Select

    Exit Sub

End Sub
Private Sub LabelCodigoAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelCodigoAte_Click

    If Len(Trim(CodigoFilialAte.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.iCodFilial = StrParaInt(CodigoFilialAte.Text)
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoCodFilialAte)

   Exit Sub

Erro_LabelCodigoAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170848)

    End Select

    Exit Sub

End Sub

Private Sub LabelFornecedorAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelFornecedorAte_Click

    If Len(Trim(FornecedorAte.Text)) > 0 Then
        'Preenche com o fornecedor da tela
        objFornecedor.lCodigo = StrParaLong(FornecedorAte.Text)
    End If

    'Chama Tela FornecedorLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedorAte)

   Exit Sub

Erro_LabelFornecedorAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170849)

    End Select

    Exit Sub

End Sub
Private Sub LabelFornecedorDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelFornecedorDe_Click

    If Len(Trim(FornecedorDe.Text)) > 0 Then
        'Preenche com o fornecedor da tela
        objFornecedor.lCodigo = StrParaLong(FornecedorDe.Text)
    End If

    'Chama Tela FornecedorLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedorDe)

   Exit Sub

Erro_LabelFornecedorDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170850)

    End Select

    Exit Sub

End Sub

Private Sub LabelCompradorDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objComprador As New ClassComprador

On Error GoTo Erro_LabelCompradorDe_Click

    If Len(Trim(CompradorDe.Text)) > 0 Then
        'Preenche com o comprador da tela
        objComprador.iCodigo = StrParaInt(CompradorDe.Text)
    End If

    'Chama Tela CompradoresLista
    Call Chama_Tela("CompradoresLista", colSelecao, objComprador, objEventoCompradorDe)

   Exit Sub

Erro_LabelCompradorDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170851)

    End Select

    Exit Sub

End Sub
Private Sub LabelCompradorAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objComprador As New ClassComprador

On Error GoTo Erro_LabelCompradorAte_Click

    If Len(Trim(CompradorAte.Text)) > 0 Then
        'Preenche com o comprador da tela
        objComprador.iCodigo = StrParaInt(CompradorAte.Text)
    End If

    'Chama Tela CompradoresLista
    Call Chama_Tela("CompradoresLista", colSelecao, objComprador, objEventoCompradorAte)

   Exit Sub

Erro_LabelCompradorAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170852)

    End Select

    Exit Sub

End Sub

Private Sub LabelNomeDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelNomeDe_Click

    If Len(Trim(NomeFilialDe.Text)) > 0 Then
        'Preenche com o requisitante da tela
        objFilialEmpresa.sNome = NomeFilialDe.Text
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoNomeFilialDe)

   Exit Sub

Erro_LabelNomeDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170853)

    End Select

    Exit Sub

End Sub
Private Sub LabelNomeAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_LabelNomeAte_Click

    If Len(Trim(NomeFilialAte.Text)) > 0 Then
        'Preenche com a FilialEmpresa da tela
        objFilialEmpresa.sNome = NomeFilialAte.Text
    End If

    'Chama Tela FilialEmpresaLista
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialEmpresa, objEventoNomeFilialAte)

   Exit Sub

Erro_LabelNomeAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170854)

    End Select

    Exit Sub

End Sub


Private Sub objEventoCodFilialAte_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    CodigoFilialAte.Text = CStr(objFilialEmpresa.iCodFilial)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNomeFilialDe_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    NomeFilialDe.Text = objFilialEmpresa.sNome

    Me.Show

    Exit Sub

End Sub
Private Sub objEventoNomeFilialAte_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    NomeFilialAte.Text = objFilialEmpresa.sNome

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCodFilialDe_evSelecao(obj1 As Object)

Dim objFilialEmpresa As New AdmFiliais

    Set objFilialEmpresa = obj1

    CodigoFilialDe.Text = CStr(objFilialEmpresa.iCodFilial)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoCodPCAte_evSelecao(obj1 As Object)

Dim objPedCompra As New ClassPedidoCompras

    Set objPedCompra = obj1

    CodPCAte.Text = CStr(objPedCompra.lCodigo)

    Me.Show

End Sub
Private Sub objEventoCodPCDe_evSelecao(obj1 As Object)

Dim objPedCompra As New ClassPedidoCompras

    Set objPedCompra = obj1

    CodPCDe.Text = CStr(objPedCompra.lCodigo)

    Me.Show

End Sub

Private Sub objEventoFornecedorDe_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor

    Set objFornecedor = obj1

    FornecedorDe.Text = CStr(objFornecedor.lCodigo)

    Me.Show

End Sub
Private Sub objEventoFornecedorAte_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor

    Set objFornecedor = obj1

    FornecedorAte.Text = CStr(objFornecedor.lCodigo)

    Me.Show

End Sub

Private Sub objEventoCompradorDe_evSelecao(obj1 As Object)

Dim objComprador As New ClassComprador

    Set objComprador = obj1

    CompradorDe.Text = CStr(objComprador.iCodigo)

    Me.Show

    Exit Sub

End Sub
Private Sub objEventoCompradorAte_evSelecao(obj1 As Object)

Dim objComprador As New ClassComprador

    Set objComprador = obj1

    CompradorAte.Text = CStr(objComprador.iCodigo)

    Me.Show

    Exit Sub

End Sub





Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 72250

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 72251

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 72252
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 72253
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 72250
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 72251 To 72253
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170855)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 72254

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 72255

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call Limpa_Tela_Rel
        
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 72254
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 72255

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170856)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 72256
    
    Select Case ComboOrdenacao.ListIndex

            Case ORD_POR_CODIGO
                
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "PCCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemPedCompra", 1)
            
            Case ORD_POR_DATA

                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "PCData", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "PCCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemPedCompra", 1)
                
            Case ORD_POR_FORNECEDOR
                
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FornecedorCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "PCFilialForn", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "PCCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemPedCompra", 1)

            Case ORD_POR_COMPRADOR
                
                Call gobjRelOpcoes.IncluirOrdenacao(1, "FilialEmpresaCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "CompradorCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "PCCod", 1)
                Call gobjRelOpcoes.IncluirOrdenacao(1, "ItemPedCompra", 1)

            Case Else
                gError 74957

    End Select
    
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 72256, 74957

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170857)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sCodFilial_I As String
Dim sCodFilial_F As String
Dim sNomeFilial_I As String
Dim sNomeFilial_F As String
Dim sCodPC_I As String
Dim sCodPC_F As String
Dim sFornecedor_I As String
Dim sFornecedor_F As String
Dim sComprador_I As String
Dim sComprador_F As String
Dim sCheck As String
Dim sOrdenacaoPor As String
Dim iOrdenacao As Long
Dim sOrd As String

On Error GoTo Erro_PreencherRelOp
    
    lErro = Formata_E_Critica_Parametros(sCodFilial_I, sCodFilial_F, sNomeFilial_I, sNomeFilial_F, sCodPC_I, sCodPC_F, sFornecedor_I, sFornecedor_F, sComprador_I, sComprador_F)
    If lErro <> SUCESSO Then gError 72257

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 72258
         
    lErro = objRelOpcoes.IncluirParametro("NCODFILIALINIC", sCodFilial_I)
    If lErro <> AD_BOOL_TRUE Then gError 72259
         
    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALINIC", NomeFilialDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 72260
    
    lErro = objRelOpcoes.IncluirParametro("NCODPCINIC", sCodPC_I)
    If lErro <> AD_BOOL_TRUE Then gError 72261
    
    lErro = objRelOpcoes.IncluirParametro("NFORNECEDORINIC", sFornecedor_I)
    If lErro <> AD_BOOL_TRUE Then gError 72262
         
    lErro = objRelOpcoes.IncluirParametro("NCOMPRADORINIC", sComprador_I)
    If lErro <> AD_BOOL_TRUE Then gError 72263
    
    'Preenche data inicial
    If Trim(DataDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAINIC", DataDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 72264
    
    'Preenche a data envio inicial
    If Trim(DataEnvioDe.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DENVINIC", DataEnvioDe.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DENVINIC", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 72265
    
    lErro = objRelOpcoes.IncluirParametro("NCODFILIALFIM", sCodFilial_F)
    If lErro <> AD_BOOL_TRUE Then gError 72267
         
    lErro = objRelOpcoes.IncluirParametro("TNOMEFILIALFIM", NomeFilialAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 72268
    
    lErro = objRelOpcoes.IncluirParametro("NCODPCFIM", sCodPC_F)
    If lErro <> AD_BOOL_TRUE Then gError 72269
    
    lErro = objRelOpcoes.IncluirParametro("NFORNECEDORFIM", sFornecedor_F)
    If lErro <> AD_BOOL_TRUE Then gError 72270
         
    lErro = objRelOpcoes.IncluirParametro("NCOMPRADORFIM", sComprador_F)
    If lErro <> AD_BOOL_TRUE Then gError 72271
    
    'Preenche data final
    If Trim(DataAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAFIM", DataAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 72272
    
    'Preenche a data envio final
    If Trim(DataEnvioAte.ClipText) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DENVFIM", DataEnvioAte.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DENVFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 72273
    
    'Verifica o Tipo Destino
    If TipoDestino(1).Value = True Then
        
        lErro = objRelOpcoes.IncluirParametro("NTIPODESTINO", 1)
        If lErro <> AD_BOOL_TRUE Then gError 72743
        
        lErro = objRelOpcoes.IncluirParametro("TDESTINATARIO", FilialEmpresa.Text)
        If lErro <> AD_BOOL_TRUE Then gError 72744
        
        FrameTipo(2).Visible = False
        FrameTipo(1).Visible = True
        
    End If
    
    If TipoDestino(2).Value = True Then
        
        lErro = objRelOpcoes.IncluirParametro("NTIPODESTINO", 2)
        If lErro <> AD_BOOL_TRUE Then gError 72736
    
        lErro = objRelOpcoes.IncluirParametro("NFILIALDESTINO", Codigo_Extrai(FilialFornecedor.Text))
        If lErro <> AD_BOOL_TRUE Then gError 72737
        
        lErro = objRelOpcoes.IncluirParametro("TDESTINATARIO", Fornecedor.Text)
        If lErro <> AD_BOOL_TRUE Then gError 72738
    
        FrameTipo(1).Visible = False
        FrameTipo(2).Visible = True
    End If
    
    If TipoDestino(0).Value = True Then
    
        lErro = objRelOpcoes.IncluirParametro("NTIPODESTINO", 0)
        If lErro <> AD_BOOL_TRUE Then gError 73439
        
        lErro = objRelOpcoes.IncluirParametro("TDESTINATARIO", "0")
        If lErro <> AD_BOOL_TRUE Then gError 73440
        
    End If
    'Exibe Itens
    If CheckItens.Value Then
        sCheck = vbChecked
        gobjRelatorio.sNomeTsk = "pcatrait"
    Else
        sCheck = vbUnchecked
        gobjRelatorio.sNomeTsk = "pcatraso"
    End If

    lErro = objRelOpcoes.IncluirParametro("NITENS", sCheck)
    If lErro <> AD_BOOL_TRUE Then gError 72316
    
    Select Case ComboOrdenacao.ListIndex
        
            Case ORD_POR_CODIGO
            
                sOrdenacaoPor = "CodPC"
                
            Case ORD_POR_DATA
                sOrdenacaoPor = "Data"
            
            Case ORD_POR_FORNECEDOR
                
                sOrdenacaoPor = "Fornecedor"
                
            Case ORD_POR_COMPRADOR
                
                sOrdenacaoPor = "Comprador"
            
            Case Else
                gError 72274
                  
    End Select

    lErro = objRelOpcoes.IncluirParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> AD_BOOL_TRUE Then gError 72275
   
    sOrd = ComboOrdenacao.ListIndex
    lErro = objRelOpcoes.IncluirParametro("NORDENACAO", sOrd)
    If lErro <> AD_BOOL_TRUE Then gError 72276
   
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sCodFilial_I, sCodFilial_F, sNomeFilial_I, sNomeFilial_F, sCodPC_I, sCodPC_F, sFornecedor_I, sFornecedor_F, sComprador_I, sComprador_F, sCheck, sOrdenacaoPor, sOrd)
    If lErro <> SUCESSO Then gError 72277

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 72257 To 72277, 72736, 72737, 72738
        
        Case 72316, 72743, 72744, 73439, 73440
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170858)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sCodFilial_I As String, sCodFilial_F As String, sNomeFilial_I As String, sNomeFilial_F As String, sCodPC_I As String, sCodPC_F As String, sFornecedor_I As String, sFornecedor_F As String, sComprador_I As String, sComprador_F As String) As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long
Dim sCclFormata As String
Dim iCclPreenchida As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros
       
    'critica Codigo da Filial Inicial e Final
    If CodigoFilialDe.Text <> "" Then
        sCodFilial_I = CStr(CodigoFilialDe.Text)
    Else
        sCodFilial_I = ""
    End If
    
    
    If CodigoFilialAte.Text <> "" Then
        sCodFilial_F = CStr(CodigoFilialAte.Text)
    Else
        sCodFilial_F = ""
    End If
                
    If sCodFilial_I <> "" And sCodFilial_F <> "" Then
        
        If StrParaInt(sCodFilial_I) > StrParaInt(sCodFilial_F) Then gError 72278
        
    End If
    
    If NomeFilialDe.Text <> "" Then
        sNomeFilial_I = NomeFilialDe.Text
    Else
        sNomeFilial_I = ""
    End If
    
    If NomeFilialAte.Text <> "" Then
        sNomeFilial_F = NomeFilialAte.Text
    Else
        sNomeFilial_F = ""
    End If
    
    If sNomeFilial_I <> "" And sNomeFilial_F <> "" Then
        If sNomeFilial_I > sNomeFilial_F Then gError 72279
    End If
    
    'critica CodigoPC Inicial e Final
    If CodPCDe.Text <> "" Then
        sCodPC_I = CStr(CodPCDe.Text)
    Else
        sCodPC_I = ""
    End If

    If CodPCAte.Text <> "" Then
        sCodPC_F = CStr(CodPCAte.Text)
    Else
        sCodPC_F = ""
    End If

    If sCodPC_I <> "" And sCodPC_F <> "" Then

        If StrParaLong(sCodPC_I) > StrParaLong(sCodPC_F) Then gError 72280

    End If
    
    'critica Fornecedor Inicial e Final
    If FornecedorDe.Text <> "" Then
        sFornecedor_I = CStr(FornecedorDe.Text)
    Else
        sFornecedor_I = ""
    End If
    
    If FornecedorAte.Text <> "" Then
        sFornecedor_F = CStr(FornecedorAte.Text)
    Else
        sFornecedor_F = ""
    End If
            
    If sFornecedor_I <> "" And sFornecedor_F <> "" Then
        
        If StrParaLong(sFornecedor_I) > StrParaLong(sFornecedor_F) Then gError 72281
        
    End If
    
    'critica Comprador Inicial e Final
    If CompradorDe.Text <> "" Then
        sComprador_I = CStr(CompradorDe.Text)
    Else
        sComprador_I = ""
    End If
    
    If CompradorAte.Text <> "" Then
        sComprador_F = CStr(CompradorAte.Text)
    Else
        sComprador_F = ""
    End If
            
    If sComprador_I <> "" And sComprador_F <> "" Then
        
        If StrParaInt(sComprador_I) > StrParaInt(sComprador_F) Then gError 72282
        
    End If
    
    'data de Envio inicial não pode ser maior que a final
    If Trim(DataEnvioDe.ClipText) <> "" And Trim(DataEnvioAte.ClipText) <> "" Then
    
         If CDate(DataEnvioDe.Text) > CDate(DataEnvioAte.Text) Then gError 72283
    
    End If
    
    'data  inicial não pode ser maior que a data  final
    If Trim(DataDe.ClipText) <> "" And Trim(DataAte.ClipText) <> "" Then
    
         If CDate(DataDe.Text) > CDate(DataAte.Text) Then gError 72284
    
    End If
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
                
        Case 72278
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            CodigoFilialDe.SetFocus
            
        Case 72279
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_INICIAL_MAIOR", gErr)
            NomeFilialDe.SetFocus
            
        Case 72280
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PC_INICIAL_MAIOR", gErr)
            CodPCDe.SetFocus
        
        Case 72281
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INICIAL_MAIOR", gErr)
            FornecedorDe.SetFocus
        
        Case 72282
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPRADOR_INICIAL_MAIOR", gErr)
            CompradorDe.SetFocus
        
        Case 72283
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAENVIO_INICIAL_MAIOR", gErr)
            DataEnvioDe.SetFocus
            
        Case 72284
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            DataDe.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170859)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sCodFilial_I As String, sCodFilial_F As String, sNomeFilial_I As String, sNomeFilial_F As String, sCodPC_I As String, sCodPC_F As String, sFornecedor_I As String, sFornecedor_F As String, sComprador_I As String, sComprador_F As String, sCheck As String, sOrdenacaoPor As String, sOrd As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Monta_Expressao_Selecao


   If sCodFilial_I <> "" Then sExpressao = "FilEmpCodInic"

   If sCodFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpCodFim"

    End If

   If sNomeFilial_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpNomeInic"

    End If
    
    If sNomeFilial_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilEmpNomeFim"

    End If
 
    If sCodPC_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "PCCod >= " & Forprint_ConvLong(StrParaLong(sCodPC_I))

    End If
   
    If sCodPC_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "PCCod <= " & Forprint_ConvLong(StrParaLong(sCodPC_F))

    End If
   
    If sFornecedor_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FornCod >= " & Forprint_ConvLong(StrParaLong(sFornecedor_I))

    End If
   
    If sFornecedor_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FornCod <= " & Forprint_ConvLong(StrParaLong(sFornecedor_F))

    End If
   
    If sComprador_I <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CompCod >= " & Forprint_ConvInt(StrParaInt(sComprador_I))

    End If
   
    If sComprador_F <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CompCod <= " & Forprint_ConvInt(StrParaInt(sComprador_F))

    End If
    
   If Trim(DataEnvioDe.ClipText) <> "" Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DataEnvioInic"
        
    End If
    
    If Trim(DataEnvioAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "DataEnvioFim"

    End If
        
    If Trim(DataDe.ClipText) <> "" Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "PCDataInic"

    End If
    
    If Trim(DataAte.ClipText) <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "PCDataFim"

    End If
    
    'Se a opção para Tipo Destino = FilialEmpresa estiver selecionada
    If TipoDestino(1).Value = True Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TipoDestino = " & Forprint_ConvInt(TIPO_DESTINO_EMPRESA)
        sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialDestino = " & Forprint_ConvInt(Codigo_Extrai(FilialEmpresa.Text))
        
    End If
        
    'Se a opção para Tipo Destino = Fornecedor estiver selecionada
    If TipoDestino(2).Value = True Then
        
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "TipoDestino = " & Forprint_ConvInt(TIPO_DESTINO_FORNECEDOR)
        sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialDestino = " & Forprint_ConvInt(Codigo_Extrai(FilialFornecedor.Text))
        sExpressao = sExpressao & " E "
        
        objFornecedor.sNomeReduzido = Fornecedor.Text
        
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 74997
        
        sExpressao = sExpressao & "FornCliDestino = " & Forprint_ConvLong(objFornecedor.lCodigo)
        
    End If
    
    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case 74997
            'erro tratado na rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170860)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long, iTipoOrd As Integer, iAscendente As Integer
Dim sParam As String
Dim sTipoCliente As String, iTipo As Integer
Dim sOrdenacaoPor As String
Dim sCclMascarado As String
Dim iIndice  As Integer
Dim bCancel As Boolean

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 72285
   
    'pega Codigo Fililial inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 72286
    
    CodigoFilialDe.Text = sParam
    Call CodigoFilialDe_Validate(bSGECancelDummy)
    
    'pega  Codigo Filial final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 72287
    
    CodigoFilialAte.Text = sParam
    Call CodigoFilialAte_Validate(bSGECancelDummy)
                
    'pega  Nome Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALINIC", sParam)
    If lErro <> SUCESSO Then gError 72288
                   
    NomeFilialDe.Text = sParam
    Call NomeFilialDe_Validate(bSGECancelDummy)
    
    'pega  Nome Final e exibe
    lErro = objRelOpcoes.ObterParametro("TNOMEFILIALFIM", sParam)
    If lErro <> SUCESSO Then gError 72289
                   
    NomeFilialAte.Text = sParam
    Call NomeFilialAte_Validate(bSGECancelDummy)
                        
    'pega  Codigo PC inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCODPCINIC", sParam)
    If lErro <> SUCESSO Then gError 72290
                   
    CodPCDe.Text = sParam
                                        
    'pega  Codigo PC final e exibe
    lErro = objRelOpcoes.ObterParametro("NCODPCFIM", sParam)
    If lErro <> SUCESSO Then gError 72291
                   
    CodPCAte.Text = sParam
    
    'pega  Fornecedor Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NFORNECEDORINIC", sParam)
    If lErro <> SUCESSO Then gError 72292
                   
    FornecedorDe.Text = sParam
    Call FornecedorDe_Validate(bSGECancelDummy)
    
    'pega  Fornecedor Final e exibe
    lErro = objRelOpcoes.ObterParametro("NFORNECEDORFIM", sParam)
    If lErro <> SUCESSO Then gError 72293
                   
    FornecedorAte.Text = sParam
    Call FornecedorAte_Validate(bSGECancelDummy)
                        
    'pega  Comprador Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("NCOMPRADORINIC", sParam)
    If lErro <> SUCESSO Then gError 72294
                   
    CompradorDe.Text = sParam
    Call CompradorDe_Validate(bSGECancelDummy)
    
    'pega  comprador Final e exibe
    lErro = objRelOpcoes.ObterParametro("NCOMPRADORFIM", sParam)
    If lErro <> SUCESSO Then gError 72295
                   
    CompradorAte.Text = sParam
    Call CompradorAte_Validate(bSGECancelDummy)
                                   
    'pega DataEnvio inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DENVINIC", sParam)
    If lErro <> SUCESSO Then gError 72296
    
    Call DateParaMasked(DataEnvioDe, CDate(sParam))
    
    'pega data de envio final e exibe
    lErro = objRelOpcoes.ObterParametro("DENVFIM", sParam)
    If lErro <> SUCESSO Then gError 72297

    Call DateParaMasked(DataEnvioAte, CDate(sParam))

    'pega data  inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAINIC", sParam)
    If lErro <> SUCESSO Then gError 72298

    Call DateParaMasked(DataDe, CDate(sParam))
       
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAFIM", sParam)
    If lErro <> SUCESSO Then gError 72299
    
    Call DateParaMasked(DataAte, CDate(sParam))
       
    lErro = objRelOpcoes.ObterParametro("NITENS", sParam)
    If lErro <> SUCESSO Then gError 72317

    If sParam = "1" Then
        CheckItens.Value = vbChecked
    Else
        CheckItens.Value = vbUnchecked
    End If
   
    'Tipo Destino
    lErro = objRelOpcoes.ObterParametro("NTIPODESTINO", sParam)
    If lErro <> SUCESSO Then gError 72739
    
    If sParam = "0" Then
        TipoDestino(0).Value = vbChecked
    Else
        iIndice = Codigo_Extrai(sParam)
        TipoDestino_Click (iIndice)
        
        If iIndice = 2 Then
        
            lErro = objRelOpcoes.ObterParametro("TDESTINATARIO", sParam)
            If lErro <> SUCESSO Then gError 72740
        
            Fornecedor.Text = sParam
            Call Fornecedor_Validate(bSGECancelDummy)
        
            lErro = objRelOpcoes.ObterParametro("NFILIALDESTINO", sParam)
            If lErro <> SUCESSO Then gError 72741
        
            FilialFornecedor.Text = sParam
            Call FilialFornecedor_Validate(bSGECancelDummy)
            FrameTipo(1).Visible = False
            FrameTipo(2).Visible = True
            TipoDestino(2).Value = vbChecked
            
        ElseIf iIndice = 1 Then
            
            lErro = objRelOpcoes.ObterParametro("TDESTINATARIO", sParam)
            If lErro <> SUCESSO Then gError 72742
            
            FilialEmpresa.Text = sParam
            FrameTipo(2).Visible = False
            FrameTipo(1).Visible = True
            TipoDestino(1).Value = vbChecked
            
        End If
    End If
    
    lErro = objRelOpcoes.ObterParametro("TORDENACAO", sOrdenacaoPor)
    If lErro <> SUCESSO Then gError 72300
    
    Select Case sOrdenacaoPor
        
            Case "CodPC"
            
                ComboOrdenacao.ListIndex = ORD_POR_CODIGO
            
            Case "Data"
            
                ComboOrdenacao.ListIndex = ORD_POR_DATA
            
            Case "Fornecedor"
            
                ComboOrdenacao.ListIndex = ORD_POR_FORNECEDOR
                
            Case "Comprador"
                ComboOrdenacao.ListIndex = ORD_POR_COMPRADOR
                        
            Case Else
                gError 72301
                  
    End Select
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 72285 To 72301
        
        Case 72317
        
        Case 72739 To 72742
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170861)

    End Select

    Exit Function

End Function

Private Sub CompradorDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objComprador As New ClassComprador

On Error GoTo Erro_CompradorDe_Validate

    If Len(Trim(CompradorDe.Text)) > 0 Then

        lErro = CF("TP_Comprador_Le", CompradorDe, objComprador, 0)
        If lErro <> SUCESSO Then gError 72314
        
        CompradorDe.Text = CStr(objComprador.iCodigo)
        
    End If

    Exit Sub

Erro_CompradorDe_Validate:

    Cancel = True

    Select Case gErr

        Case 72314

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170862)

    End Select

    Exit Sub

End Sub

Private Sub CompradorAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objComprador As New ClassComprador

On Error GoTo Erro_CompradorAte_Validate

    If Len(Trim(CompradorAte.Text)) > 0 Then

        'Lê o código informado
        lErro = CF("TP_Comprador_Le", CompradorDe, objComprador, 0)
        If lErro <> SUCESSO Then gError 72315
        
        CompradorDe.Text = CStr(objComprador.iCodigo)
        
    End If

    Exit Sub

Erro_CompradorAte_Validate:

    Cancel = True

    Select Case gErr

        Case 72315

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170863)

    End Select

    Exit Sub

End Sub


Private Sub FornecedorDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FornecedorDe_Validate

    If Len(Trim(FornecedorDe.Text)) > 0 Then

        'Lê o código informado
        objFornecedor.lCodigo = LCodigo_Extrai(FornecedorDe.Text)
        
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then gError 72310
        
        'Se não encontrou o Fornecedor ==> erro
        If lErro = 12729 Then gError 72311
        
    End If

    Exit Sub

Erro_FornecedorDe_Validate:

    Cancel = True

    Select Case gErr

        Case 72310

        Case 72311
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170864)

    End Select

    Exit Sub

End Sub

Private Sub FornecedorAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_FornecedorAte_Validate

    If Len(Trim(FornecedorAte.Text)) > 0 Then

        'Lê o código informado
        objFornecedor.lCodigo = LCodigo_Extrai(FornecedorAte.Text)
        
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then gError 72312
        
        'Se não encontrou o Fornecedor ==> erro
        If lErro = 12729 Then gError 72313
        
    End If

    Exit Sub

Erro_FornecedorAte_Validate:

    Cancel = True

    Select Case gErr

        Case 72312

        Case 72313
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170865)

    End Select

    Exit Sub

End Sub


Private Sub CodigoFilialDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_CodigoFilialDe_Validate

    If Len(Trim(CodigoFilialDe.Text)) > 0 Then

        objFilialEmpresa.iCodFilial = StrParaInt(CodigoFilialDe.Text)
        'Lê o código informado
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 72302
        
        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 72303

    End If
    
    Exit Sub

Erro_CodigoFilialDe_Validate:

    Cancel = True


    Select Case gErr

        Case 72302

        Case 72303
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170866)

    End Select

    Exit Sub

End Sub
Private Sub CodigoFilialAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_CodigoFilialAte_Validate

    If Len(Trim(CodigoFilialAte.Text)) > 0 Then

        objFilialEmpresa.iCodFilial = StrParaInt(CodigoFilialAte.Text)
        'Lê o código informado
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 72304
        
        'Se não encontrou a Filial ==> erro
        If lErro = 27378 Then gError 72305

    End If

    Exit Sub

Erro_CodigoFilialAte_Validate:

    Cancel = True


    Select Case gErr

        Case 72304

        Case 72305
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170867)

    End Select

    Exit Sub

End Sub


Private Sub NomeFilialDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim bAchou As Boolean
Dim colFiliais As New Collection

On Error GoTo Erro_NomeFilialDe_Validate

    bAchou = False
    
    If Len(Trim(NomeFilialDe.Text)) > 0 Then

        lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
        If lErro <> SUCESSO Then gError 72306

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = UCase(NomeFilialDe.Text) Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 72307
        
        NomeFilialDe.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeFilialDe_Validate:

    Cancel = True

    Select Case gErr

        Case 72306

        Case 72307
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeFilialDe.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170868)

    End Select

Exit Sub

End Sub

Private Sub NomeFilialAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim bAchou As Boolean
Dim colFiliais As New Collection

On Error GoTo Erro_NomeFilialAte_Validate

    bAchou = False
    If Len(Trim(NomeFilialAte.Text)) > 0 Then

        lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
        If lErro <> SUCESSO Then gError 72308

        'Carrega a Filial com o Nome informado
        For Each objFilialEmpresa In colFiliais
            If objFilialEmpresa.sNome = UCase(NomeFilialAte.Text) Then
                bAchou = True
                Exit For
            End If
        Next

        'Se não encontrou Filial com o Nome informado ==> erro
        If bAchou = False Then gError 72309

        NomeFilialAte.Text = objFilialEmpresa.sNome

    End If

    Exit Sub

Erro_NomeFilialAte_Validate:

    Cancel = True


    Select Case gErr

        Case 72308

        Case 72309
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, NomeFilialAte.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170869)

    End Select

Exit Sub

End Sub


Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

''    Parent.HelpContextID = IDH_RELOP_REQ
    Set Form_Load_Ocx = Me
    Caption = "Pedidos de Compra Atrasados"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpPedCompraAtraso"
    
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

Public Sub Unload(objme As Object)
    
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
        
        If Me.ActiveControl Is CodPCDe Then
            Call LabelCodPCDe_Click
            
        ElseIf Me.ActiveControl Is CodPCAte Then
            Call LabelCodPCAte_Click
           
        ElseIf Me.ActiveControl Is CodigoFilialDe Then
            Call LabelCodigoDe_Click
        
        ElseIf Me.ActiveControl Is CodigoFilialAte Then
            Call LabelCodigoAte_Click
        
        ElseIf Me.ActiveControl Is NomeFilialDe Then
            Call LabelNomeDe_Click
        
        ElseIf Me.ActiveControl Is NomeFilialAte Then
            Call LabelNomeAte_Click
        
        ElseIf Me.ActiveControl Is FornecedorDe Then
            Call LabelFornecedorDe_Click
        
        ElseIf Me.ActiveControl Is FornecedorAte Then
            Call LabelFornecedorAte_Click
        
        ElseIf Me.ActiveControl Is CompradorDe Then
            Call LabelCompradorDe_Click
        
        ElseIf Me.ActiveControl Is CompradorAte Then
            Call LabelCompradorAte_Click
        
        End If
    
    End If

End Sub


Private Sub LabelCodigoDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoDe, Source, X, Y)
End Sub

Private Sub LabelCodigoDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoDe, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigoAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigoAte, Source, X, Y)
End Sub

Private Sub LabelCodigoAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigoAte, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LabelFilialEmpDestino_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilialEmpDestino, Source, X, Y)
End Sub

Private Sub LabelFilialEmpDestino_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilialEmpDestino, Button, Shift, X, Y)
End Sub

Private Sub LabelFilialFornDestino_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilialFornDestino, Source, X, Y)
End Sub

Private Sub LabelFilialFornDestino_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilialFornDestino, Button, Shift, X, Y)
End Sub

Private Sub LabelFornDestino_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornDestino, Source, X, Y)
End Sub

Private Sub LabelFornDestino_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornDestino, Button, Shift, X, Y)
End Sub

Private Sub LabelCompradorAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCompradorAte, Source, X, Y)
End Sub

Private Sub LabelCompradorAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCompradorAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCompradorDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCompradorDe, Source, X, Y)
End Sub

Private Sub LabelCompradorDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCompradorDe, Button, Shift, X, Y)
End Sub

Private Sub LabelFornecedorAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornecedorAte, Source, X, Y)
End Sub

Private Sub LabelFornecedorAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornecedorAte, Button, Shift, X, Y)
End Sub

Private Sub LabelFornecedorDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornecedorDe, Source, X, Y)
End Sub

Private Sub LabelFornecedorDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornecedorDe, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeReqDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeReqDe, Source, X, Y)
End Sub

Private Sub LabelNomeReqDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeReqDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeReqAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeReqAte, Source, X, Y)
End Sub

Private Sub LabelNomeReqAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeReqAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodPCAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodPCAte, Source, X, Y)
End Sub

Private Sub LabelCodPCAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodPCAte, Button, Shift, X, Y)
End Sub

Private Sub LabelCodPCDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodPCDe, Source, X, Y)
End Sub

Private Sub LabelCodPCDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodPCDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeDe, Source, X, Y)
End Sub

Private Sub LabelNomeDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeDe, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeAte, Source, X, Y)
End Sub

Private Sub LabelNomeAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeAte, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

