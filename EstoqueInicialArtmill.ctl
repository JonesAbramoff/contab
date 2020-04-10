VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl EstoqueInicial 
   ClientHeight    =   5760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9060
   KeyPreview      =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   9060
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4965
      Index           =   1
      Left            =   90
      TabIndex        =   0
      Top             =   705
      Width           =   8850
      Begin VB.ListBox ListaAlmoxarifado 
         Height          =   4350
         Left            =   5970
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   390
         Width           =   2760
      End
      Begin VB.CommandButton BotaoEstoquesIniciais 
         Caption         =   "Estoques Iniciais"
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
         Left            =   1920
         TabIndex        =   8
         Top             =   4380
         Width           =   1935
      End
      Begin VB.TextBox LocalizacaoFisica 
         Height          =   315
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   5
         Top             =   3015
         Width           =   3840
      End
      Begin VB.CheckBox Padrao 
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
         Left            =   4710
         TabIndex        =   2
         Top             =   375
         Width           =   945
      End
      Begin VB.CheckBox Fixar 
         Caption         =   "Fixar"
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
         Left            =   4740
         TabIndex        =   4
         Top             =   780
         Width           =   795
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   3120
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3930
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComctlLib.TreeView TvwContas 
         Height          =   4350
         Left            =   5970
         TabIndex        =   11
         Top             =   390
         Visible         =   0   'False
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   7673
         _Version        =   393217
         Indentation     =   453
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   315
         Left            =   1935
         TabIndex        =   7
         Top             =   3915
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Produto 
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Top             =   765
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ContaContabil 
         Height          =   315
         Left            =   1920
         TabIndex        =   6
         Top             =   3450
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Almoxarifado 
         Height          =   315
         Left            =   1935
         TabIndex        =   1
         Top             =   315
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSComctlLib.TreeView TvwProdutos 
         Height          =   4350
         Left            =   5970
         TabIndex        =   10
         Top             =   390
         Visible         =   0   'False
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   7673
         _Version        =   393217
         Indentation     =   453
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin VB.Label Detalhe 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1920
         TabIndex        =   81
         Top             =   2550
         Width           =   2100
      End
      Begin VB.Label Cor 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1920
         TabIndex        =   80
         Top             =   2100
         Width           =   2100
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cor:"
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
         Left            =   1470
         TabIndex        =   79
         Top             =   2145
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Detalhe:"
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
         Left            =   1095
         TabIndex        =   78
         Top             =   2595
         Width           =   735
      End
      Begin VB.Label UnidMed 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1920
         TabIndex        =   45
         Top             =   1665
         Width           =   1635
      End
      Begin VB.Label DescricaoProduto 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1920
         TabIndex        =   44
         Top             =   1215
         Width           =   3840
      End
      Begin VB.Label ContaContabilLabel 
         AutoSize        =   -1  'True
         Caption         =   "Conta Estoque:"
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
         Left            =   540
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   42
         Top             =   3510
         Width           =   1320
      End
      Begin VB.Label AlmoxarifadoLabel 
         AutoSize        =   -1  'True
         Caption         =   "Almoxarifado:"
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
         Left            =   660
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   41
         Top             =   375
         Width           =   1155
      End
      Begin VB.Label ProdutoLabel1 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1110
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   40
         Top             =   810
         Width           =   735
      End
      Begin VB.Label Label13 
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
         Left            =   915
         TabIndex        =   39
         Top             =   1275
         Width           =   930
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Unidade de Medida:"
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
         TabIndex        =   38
         Top             =   1725
         Width           =   1725
      End
      Begin VB.Label Label6 
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
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1350
         TabIndex        =   37
         Top             =   3945
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Localização Física:"
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
         Left            =   165
         TabIndex        =   36
         Top             =   3075
         Width           =   1680
      End
      Begin VB.Label LabelAlmoxarifado 
         Caption         =   "Almoxarifados"
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
         Left            =   5970
         TabIndex        =   46
         Top             =   195
         Width           =   2340
      End
      Begin VB.Label LabelContas 
         Caption         =   "Plano de Contas"
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
         Left            =   5970
         TabIndex        =   43
         Top             =   195
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.Label LabelProduto 
         AutoSize        =   -1  'True
         Caption         =   "Produtos"
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
         Left            =   5970
         TabIndex        =   35
         Top             =   195
         Visible         =   0   'False
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4935
      Index           =   2
      Left            =   90
      TabIndex        =   47
      Top             =   720
      Visible         =   0   'False
      Width           =   8760
      Begin VB.CommandButton BotaoRastro 
         Caption         =   "Rastreamento"
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
         Left            =   7065
         TabIndex        =   77
         Top             =   135
         Width           =   1560
      End
      Begin VB.Frame Frame2 
         Caption         =   "Saldos Nosso em Poder de Terceiros"
         Height          =   1755
         Left            =   165
         TabIndex        =   54
         Top             =   1410
         Width           =   8475
         Begin MSMask.MaskEdBox QuantConserto 
            Height          =   255
            Left            =   3180
            TabIndex        =   15
            Top             =   255
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorConserto 
            Height          =   255
            Left            =   4950
            TabIndex        =   16
            Top             =   255
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
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
         Begin MSMask.MaskEdBox QuantConsig 
            Height          =   255
            Left            =   3180
            TabIndex        =   17
            Top             =   540
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorConsig 
            Height          =   255
            Left            =   4950
            TabIndex        =   18
            Top             =   540
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
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
         Begin MSMask.MaskEdBox QuantDemo 
            Height          =   255
            Left            =   3180
            TabIndex        =   19
            Top             =   840
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDemo 
            Height          =   255
            Left            =   4950
            TabIndex        =   20
            Top             =   840
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
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
         Begin MSMask.MaskEdBox QuantOutras 
            Height          =   255
            Left            =   3180
            TabIndex        =   21
            Top             =   1125
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorOutras 
            Height          =   255
            Left            =   4950
            TabIndex        =   22
            Top             =   1125
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
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
         Begin MSMask.MaskEdBox QuantBenef 
            Height          =   255
            Left            =   3180
            TabIndex        =   23
            Top             =   1410
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorBenef 
            Height          =   255
            Left            =   4950
            TabIndex        =   24
            Top             =   1410
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
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
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Em Beneficiamento:"
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
            TabIndex        =   59
            Top             =   1425
            Width           =   1695
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Em Conserto:"
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
            Left            =   1935
            TabIndex        =   58
            Top             =   270
            Width           =   1140
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Em Consignação:"
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
            Left            =   1590
            TabIndex        =   57
            Top             =   570
            Width           =   1485
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Em Demonstração:"
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
            Left            =   1470
            TabIndex        =   56
            Top             =   840
            Width           =   1605
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Outras:"
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
            Left            =   2445
            TabIndex        =   55
            Top             =   1125
            Width           =   630
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Saldos de Terceiros em Nosso Poder"
         Height          =   1725
         Left            =   165
         TabIndex        =   48
         Top             =   3210
         Width           =   8475
         Begin MSMask.MaskEdBox QuantConserto3 
            Height          =   255
            Left            =   3180
            TabIndex        =   25
            Top             =   240
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorConserto3 
            Height          =   255
            Left            =   4950
            TabIndex        =   26
            Top             =   240
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
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
         Begin MSMask.MaskEdBox QuantConsig3 
            Height          =   255
            Left            =   3180
            TabIndex        =   27
            Top             =   540
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorConsig3 
            Height          =   255
            Left            =   4950
            TabIndex        =   28
            Top             =   540
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
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
         Begin MSMask.MaskEdBox QuantDemo3 
            Height          =   255
            Left            =   3180
            TabIndex        =   29
            Top             =   825
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDemo3 
            Height          =   255
            Left            =   4950
            TabIndex        =   30
            Top             =   825
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
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
         Begin MSMask.MaskEdBox QuantOutras3 
            Height          =   255
            Left            =   3180
            TabIndex        =   31
            Top             =   1095
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorOutras3 
            Height          =   255
            Left            =   4950
            TabIndex        =   32
            Top             =   1095
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
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
         Begin MSMask.MaskEdBox QuantBenef3 
            Height          =   255
            Left            =   3180
            TabIndex        =   33
            Top             =   1380
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorBenef3 
            Height          =   255
            Left            =   4950
            TabIndex        =   34
            Top             =   1380
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   450
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
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Em Beneficiamento:"
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
            Left            =   1320
            TabIndex        =   53
            Top             =   1410
            Width           =   1695
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Em Conserto:"
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
            Left            =   1875
            TabIndex        =   52
            Top             =   270
            Width           =   1140
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Em Consignação:"
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
            Left            =   1530
            TabIndex        =   51
            Top             =   570
            Width           =   1485
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Em Demonstração:"
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
            Left            =   1410
            TabIndex        =   50
            Top             =   855
            Width           =   1605
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Outras:"
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
            Left            =   2385
            TabIndex        =   49
            Top             =   1125
            Width           =   630
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Saldo Nosso"
         Height          =   615
         Left            =   165
         TabIndex        =   73
         Top             =   750
         Width           =   8475
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   255
            Left            =   3180
            TabIndex        =   13
            Top             =   240
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   450
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Valor 
            Height          =   255
            Left            =   4950
            TabIndex        =   14
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
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
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Disponível:"
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
            Left            =   2070
            TabIndex        =   74
            Top             =   270
            Width           =   990
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Quantidades"
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
         Left            =   3600
         TabIndex        =   76
         Top             =   570
         Width           =   1080
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Valores"
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
         Left            =   5505
         TabIndex        =   75
         Top             =   570
         Width           =   690
      End
      Begin VB.Label AlmoxarifadoComplementar 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5145
         TabIndex        =   65
         Top             =   150
         Width           =   1785
      End
      Begin VB.Label ProdutoComplementar 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1065
         TabIndex        =   64
         Top             =   165
         Width           =   1260
      End
      Begin VB.Label Label28 
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
         Height          =   195
         Left            =   2490
         TabIndex        =   63
         Top             =   195
         Width           =   480
      End
      Begin VB.Label UMComplementar 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3015
         TabIndex        =   62
         Top             =   150
         Width           =   780
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Almoxarifado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   3960
         TabIndex        =   61
         Top             =   195
         Width           =   1155
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
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
         Height          =   165
         Left            =   255
         TabIndex        =   60
         Top             =   195
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   645
      Left            =   5550
      ScaleHeight     =   585
      ScaleWidth      =   3360
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   30
      Width           =   3420
      Begin VB.CommandButton BotaoExcluir 
         Height          =   540
         Left            =   1890
         Picture         =   "EstoqueInicialArtmill.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Excluir"
         Top             =   30
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   540
         Left            =   1380
         Picture         =   "EstoqueInicialArtmill.ctx":018A
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Gravar"
         Top             =   30
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   540
         Left            =   2400
         Picture         =   "EstoqueInicialArtmill.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   30
         Width           =   405
      End
      Begin VB.CommandButton BotaoConsultar 
         Height          =   540
         Left            =   45
         Picture         =   "EstoqueInicialArtmill.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   30
         Width           =   1245
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   540
         Left            =   2895
         Picture         =   "EstoqueInicialArtmill.ctx":25D8
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Fechar"
         Top             =   30
         Width           =   405
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5340
      Left            =   60
      TabIndex        =   66
      Top             =   390
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   9419
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Estoque Inicial"
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
Attribute VB_Name = "EstoqueInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTEstoqueInicial
Attribute objCT.VB_VarHelpID = -1

Private Sub ProdutoLabel1_Click()
    objCT.ProdutoLabel1_Click
End Sub

Private Sub TabStrip1_Click()
    objCT.TabStrip1_Click
End Sub

Private Sub UserControl_Initialize()
    Set objCT = New CTEstoqueInicial
    Set objCT.objUserControl = Me
    
    '#########################################
    'Artmill
    Set objCT.gobjInfoUsu = New CTEstoqueInicialVGArt
    Set objCT.gobjInfoUsu.gobjTelaUsu = New CTEstoqueInicialArt
    '#########################################
End Sub

Private Sub Almoxarifado_Change()
     Call objCT.Almoxarifado_Change
End Sub

Private Sub Almoxarifado_GotFocus()
     Call objCT.Almoxarifado_GotFocus
End Sub

Private Sub Almoxarifado_Validate(Cancel As Boolean)
     Call objCT.Almoxarifado_Validate(Cancel)
End Sub

Private Sub AlmoxarifadoLabel_Click()
     Call objCT.AlmoxarifadoLabel_Click
End Sub

Private Sub BotaoConsultar_Click()
     Call objCT.BotaoConsultar_Click
End Sub

Private Sub BotaoEstoquesIniciais_Click()
     Call objCT.BotaoEstoquesIniciais_Click
End Sub

Private Sub BotaoExcluir_Click()
     Call objCT.BotaoExcluir_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Private Sub BotaoRastro_Click()
     Call objCT.BotaoRastro_Click
End Sub

Private Sub ContaContabil_GotFocus()
     Call objCT.ContaContabil_GotFocus
End Sub

Private Sub DataInicial_Change()
     Call objCT.DataInicial_Change
End Sub

Private Sub DataInicial_GotFocus()
     Call objCT.DataInicial_GotFocus
End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)
     Call objCT.DataInicial_Validate(Cancel)
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub ContaContabil_Change()
     Call objCT.ContaContabil_Change
End Sub

Private Sub ContaContabil_Validate(Cancel As Boolean)
     Call objCT.ContaContabil_Validate(Cancel)
End Sub

Private Sub LocalizacaoFisica_Change()
     Call objCT.LocalizacaoFisica_Change
End Sub

Private Sub ContaContabilLabel_Click()
     Call objCT.ContaContabilLabel_Click
End Sub

Private Sub Padrao_Click()
     Call objCT.Padrao_Click
End Sub

Private Sub Produto_Change()
     Call objCT.Produto_Change
End Sub

Private Sub Produto_GotFocus()
     Call objCT.Produto_GotFocus
End Sub

Private Sub Produto_Validate(Cancel As Boolean)
     Call objCT.Produto_Validate(Cancel)
End Sub

Private Sub Quantidade_Change()
     Call objCT.Quantidade_Change
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)
     Call objCT.Quantidade_Validate(Cancel)
End Sub

Private Sub TvwContas_Expand(ByVal objNode As MSComctlLib.Node)
     Call objCT.TvwContas_Expand(objNode)
End Sub

Private Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.TvwContas_NodeClick(Node)
End Sub

'Private Sub TvwProdutos_Expand(ByVal objNode As MSComctlLib.Node)
'     Call objCT.TvwProdutos_Expand(objNode)
'End Sub
'
'Private Sub TvwProdutos_NodeClick(ByVal Node As MSComctlLib.Node)
'     Call objCT.TvwProdutos_NodeClick(Node)
'End Sub

Private Sub UpDown1_DownClick()
     Call objCT.UpDown1_DownClick
End Sub

Private Sub UpDown1_UpClick()
     Call objCT.UpDown1_UpClick
End Sub

Private Sub Valor_Change()
     Call objCT.Valor_Change
End Sub

Private Sub Valor_Validate(Cancel As Boolean)
     Call objCT.Valor_Validate(Cancel)
End Sub

Private Sub ListaAlmoxarifado_DblClick()
     Call objCT.ListaAlmoxarifado_DblClick
End Sub

Public Function Trata_Parametros() As Long
     Trata_Parametros = objCT.Trata_Parametros()
End Function

Private Sub QuantConserto_Change()
     Call objCT.QuantConserto_Change
End Sub

Private Sub QuantConserto_Validate(Cancel As Boolean)
     Call objCT.QuantConserto_Validate(Cancel)
End Sub

Private Sub ValorConserto_Change()
     Call objCT.ValorConserto_Change
End Sub

Private Sub ValorConserto_Validate(Cancel As Boolean)
     Call objCT.ValorConserto_Validate(Cancel)
End Sub

Private Sub QuantConsig_Change()
     Call objCT.QuantConsig_Change
End Sub

Private Sub QuantConsig_Validate(Cancel As Boolean)
     Call objCT.QuantConsig_Validate(Cancel)
End Sub

Private Sub ValorConsig_Change()
     Call objCT.ValorConsig_Change
End Sub

Private Sub ValorConsig_Validate(Cancel As Boolean)
     Call objCT.ValorConsig_Validate(Cancel)
End Sub

Private Sub QuantDemo_Change()
     Call objCT.QuantDemo_Change
End Sub

Private Sub QuantDemo_Validate(Cancel As Boolean)
     Call objCT.QuantDemo_Validate(Cancel)
End Sub

Private Sub ValorDemo_Change()
     Call objCT.ValorDemo_Change
End Sub

Private Sub ValorDemo_Validate(Cancel As Boolean)
     Call objCT.ValorDemo_Validate(Cancel)
End Sub

Private Sub QuantOutras_Change()
     Call objCT.QuantOutras_Change
End Sub

Private Sub QuantOutras_Validate(Cancel As Boolean)
     Call objCT.QuantOutras_Validate(Cancel)
End Sub

Private Sub ValorOutras_Change()
     Call objCT.ValorOutras_Change
End Sub

Private Sub ValorOutras_Validate(Cancel As Boolean)
     Call objCT.ValorOutras_Validate(Cancel)
End Sub

Private Sub QuantBenef_Change()
     Call objCT.QuantBenef_Change
End Sub

Private Sub QuantBenef_Validate(Cancel As Boolean)
     Call objCT.QuantBenef_Validate(Cancel)
End Sub

Private Sub ValorBenef_Change()
     Call objCT.ValorBenef_Change
End Sub

Private Sub ValorBenef_Validate(Cancel As Boolean)
     Call objCT.ValorBenef_Validate(Cancel)
End Sub

Private Sub QuantConserto3_Change()
     Call objCT.QuantConserto3_Change
End Sub

Private Sub QuantConserto3_Validate(Cancel As Boolean)
     Call objCT.QuantConserto3_Validate(Cancel)
End Sub

Private Sub ValorConserto3_Change()
     Call objCT.ValorConserto3_Change
End Sub

Private Sub ValorConserto3_Validate(Cancel As Boolean)
     Call objCT.ValorConserto3_Validate(Cancel)
End Sub

Private Sub QuantConsig3_Change()
     Call objCT.QuantConsig3_Change
End Sub

Private Sub QuantConsig3_Validate(Cancel As Boolean)
     Call objCT.QuantConsig3_Validate(Cancel)
End Sub

Private Sub ValorConsig3_Change()
     Call objCT.ValorConsig3_Change
End Sub

Private Sub ValorConsig3_Validate(Cancel As Boolean)
     Call objCT.ValorConsig3_Validate(Cancel)
End Sub

Private Sub QuantDemo3_Change()
     Call objCT.QuantDemo3_Change
End Sub

Private Sub QuantDemo3_Validate(Cancel As Boolean)
     Call objCT.QuantDemo3_Validate(Cancel)
End Sub

Private Sub ValorDemo3_Change()
     Call objCT.ValorDemo3_Change
End Sub

Private Sub ValorDemo3_Validate(Cancel As Boolean)
     Call objCT.ValorDemo3_Validate(Cancel)
End Sub

Private Sub QuantOutras3_Change()
     Call objCT.QuantOutras3_Change
End Sub

Private Sub QuantOutras3_Validate(Cancel As Boolean)
     Call objCT.QuantOutras3_Validate(Cancel)
End Sub

Private Sub ValorOutras3_Change()
     Call objCT.ValorOutras3_Change
End Sub

Private Sub ValorOutras3_Validate(Cancel As Boolean)
     Call objCT.ValorOutras3_Validate(Cancel)
End Sub

Private Sub QuantBenef3_Change()
     Call objCT.QuantBenef3_Change
End Sub

Private Sub QuantBenef3_Validate(Cancel As Boolean)
     Call objCT.QuantBenef3_Validate(Cancel)
End Sub

Private Sub ValorBenef3_Change()
     Call objCT.ValorBenef3_Change
End Sub

Private Sub ValorBenef3_Validate(Cancel As Boolean)
     Call objCT.ValorBenef3_Validate(Cancel)
End Sub

Public Function Form_Load_Ocx() As Object

    Call objCT.Form_Load_Ocx
    Set Form_Load_Ocx = Me

End Function

Public Sub Form_Unload(Cancel As Integer)
    If Not (objCT Is Nothing) Then
        Call objCT.Form_Unload(Cancel)
        If Cancel = False Then
             Set objCT.objUserControl = Nothing
             Set objCT = Nothing
        End If
    End If
End Sub

Private Sub objCT_Unload()
   RaiseEvent Unload
End Sub

Public Function Name() As String
    Name = objCT.Name
End Function

Public Sub Show()
    Call objCT.Show
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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Property Get Caption() As String
    Caption = objCT.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    objCT.Caption = New_Caption
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Call objCT.UserControl_KeyDown(KeyCode, Shift)
End Sub

'Private Sub LabelProduto_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(LabelProduto, Source, X, Y)
'End Sub
'
'Private Sub LabelProduto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelProduto, Button, Shift, X, Y)
'End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub AlmoxarifadoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(AlmoxarifadoLabel, Source, X, Y)
End Sub

Private Sub AlmoxarifadoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(AlmoxarifadoLabel, Button, Shift, X, Y)
End Sub

Private Sub ContaContabilLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaContabilLabel, Source, X, Y)
End Sub

Private Sub ContaContabilLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaContabilLabel, Button, Shift, X, Y)
End Sub

Private Sub LabelContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContas, Source, X, Y)
End Sub

Private Sub LabelContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContas, Button, Shift, X, Y)
End Sub

Private Sub DescricaoProduto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescricaoProduto, Source, X, Y)
End Sub

Private Sub DescricaoProduto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescricaoProduto, Button, Shift, X, Y)
End Sub

Private Sub UnidMed_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(UnidMed, Source, X, Y)
End Sub

Private Sub UnidMed_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(UnidMed, Button, Shift, X, Y)
End Sub

Private Sub LabelAlmoxarifado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAlmoxarifado, Source, X, Y)
End Sub

Private Sub LabelAlmoxarifado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAlmoxarifado, Button, Shift, X, Y)
End Sub

Private Sub Label34_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label34, Source, X, Y)
End Sub

Private Sub Label34_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label34, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label35_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label35, Source, X, Y)
End Sub

Private Sub Label35_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label35, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub

Private Sub Label18_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label18, Source, X, Y)
End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label18, Button, Shift, X, Y)
End Sub

Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label26_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label26, Source, X, Y)
End Sub

Private Sub Label26_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label26, Button, Shift, X, Y)
End Sub

Private Sub AlmoxarifadoComplementar_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(AlmoxarifadoComplementar, Source, X, Y)
End Sub

Private Sub AlmoxarifadoComplementar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(AlmoxarifadoComplementar, Button, Shift, X, Y)
End Sub

Private Sub ProdutoComplementar_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ProdutoComplementar, Source, X, Y)
End Sub

Private Sub ProdutoComplementar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutoComplementar, Button, Shift, X, Y)
End Sub

Private Sub Label28_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label28, Source, X, Y)
End Sub

Private Sub Label28_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label28, Button, Shift, X, Y)
End Sub

Private Sub UMComplementar_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(UMComplementar, Source, X, Y)
End Sub

Private Sub UMComplementar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(UMComplementar, Button, Shift, X, Y)
End Sub

Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label25, Source, X, Y)
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label25, Button, Shift, X, Y)
End Sub

Private Sub Label24_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label24, Source, X, Y)
End Sub

Private Sub Label24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label24, Button, Shift, X, Y)
End Sub

Private Sub ProdutoLabel1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ProdutoLabel1, Source, X, Y)
End Sub

Private Sub ProdutoLabel1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutoLabel1, Button, Shift, X, Y)
End Sub

