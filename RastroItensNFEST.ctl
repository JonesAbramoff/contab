VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.UserControl RastroItensNFEST 
   ClientHeight    =   5160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9090
   KeyPreview      =   -1  'True
   ScaleHeight     =   5160
   ScaleWidth      =   9090
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame17"
      Height          =   4050
      Index           =   4
      Left            =   150
      TabIndex        =   41
      Top             =   915
      Visible         =   0   'False
      Width           =   8790
      Begin VB.CommandButton BotaoSerie 
         Caption         =   "Séries"
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
         Left            =   45
         TabIndex        =   65
         Top             =   3645
         Width           =   1665
      End
      Begin VB.CommandButton BotaoLotes 
         Caption         =   "Lotes"
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
         Left            =   6915
         TabIndex        =   53
         Top             =   3630
         Width           =   1665
      End
      Begin VB.Frame Frame17 
         Caption         =   "Rastreamento do Produto"
         Height          =   3525
         Left            =   45
         TabIndex        =   42
         Top             =   30
         Width           =   8550
         Begin VB.ComboBox ProdutoRastro 
            Height          =   315
            ItemData        =   "RastroItensNFEST.ctx":0000
            Left            =   915
            List            =   "RastroItensNFEST.ctx":000D
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   885
            Width           =   1740
         End
         Begin VB.ComboBox EscaninhoRastro 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "RastroItensNFEST.ctx":002C
            Left            =   3840
            List            =   "RastroItensNFEST.ctx":0036
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin MSMask.MaskEdBox UMRastro 
            Height          =   240
            Left            =   3240
            TabIndex        =   44
            Top             =   210
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ItemNFRastro 
            Height          =   225
            Left            =   300
            TabIndex        =   45
            Top             =   765
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
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
         Begin MSMask.MaskEdBox AlmoxRastro 
            Height          =   240
            Left            =   1815
            TabIndex        =   46
            Top             =   210
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantRastro 
            Height          =   225
            Left            =   1995
            TabIndex        =   47
            Top             =   405
            Width           =   1575
            _ExtentX        =   2778
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
         Begin MSMask.MaskEdBox LoteRastro 
            Height          =   225
            Left            =   2970
            TabIndex        =   48
            Top             =   405
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox LoteDataRastro 
            Height          =   255
            Left            =   5730
            TabIndex        =   49
            Top             =   405
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialOPRastro 
            Height          =   225
            Left            =   4125
            TabIndex        =   50
            Top             =   405
            Width           =   1575
            _ExtentX        =   2778
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
         Begin MSMask.MaskEdBox QuantLoteRastro 
            Height          =   225
            Left            =   5730
            TabIndex        =   51
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
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
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridRastro 
            Height          =   3030
            Left            =   135
            TabIndex        =   52
            Top             =   300
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   5345
            _Version        =   393216
            Rows            =   51
            Cols            =   7
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4050
      Index           =   1
      Left            =   150
      TabIndex        =   12
      Top             =   915
      Width           =   8790
      Begin VB.Frame Frame2 
         Caption         =   "Nota Fiscal"
         Height          =   870
         Left            =   210
         TabIndex        =   35
         Top             =   2670
         Width           =   8325
         Begin VB.Label LblValor 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5610
            TabIndex        =   13
            Top             =   330
            Width           =   1275
         End
         Begin VB.Label Label10 
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
            Left            =   5055
            TabIndex        =   40
            Top             =   375
            Width           =   510
         End
         Begin VB.Label LblFilial 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3225
            TabIndex        =   39
            Top             =   330
            Width           =   1560
         End
         Begin VB.Label LabelFilial 
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
            Left            =   2715
            TabIndex        =   38
            Top             =   375
            Width           =   465
         End
         Begin VB.Label LblDataEmissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1110
            TabIndex        =   37
            Top             =   330
            Width           =   1275
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
            Left            =   315
            TabIndex        =   36
            Top             =   375
            Width           =   765
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Dados do Cliente/Fornecedor"
         Height          =   870
         Index           =   0
         Left            =   225
         TabIndex        =   18
         Top             =   1590
         Width           =   8325
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5640
            TabIndex        =   7
            Top             =   405
            Width           =   1605
         End
         Begin MSMask.MaskEdBox Fornecedor 
            Height          =   315
            Left            =   2100
            TabIndex        =   6
            Top             =   405
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   315
            Left            =   2100
            TabIndex        =   5
            Top             =   405
            Visible         =   0   'False
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label FornecedorLabel 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   1020
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   21
            Top             =   450
            Width           =   1035
         End
         Begin VB.Label Label2 
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
            Left            =   5100
            TabIndex        =   20
            Top             =   450
            Width           =   465
         End
         Begin VB.Label ClienteLabel 
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
            Left            =   1380
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   19
            Top             =   450
            Visible         =   0   'False
            Width           =   660
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Identificação"
         Height          =   1260
         Index           =   0
         Left            =   225
         TabIndex        =   14
         Top             =   105
         Width           =   8325
         Begin VB.CheckBox EletronicaFed 
            Caption         =   "Eletrônica Federal"
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
            Left            =   4935
            TabIndex        =   64
            Top             =   855
            Width           =   2070
         End
         Begin VB.CommandButton BotaoNFiscal 
            Caption         =   "Traz Dados Nota Fiscal"
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
            Left            =   4920
            TabIndex        =   4
            Top             =   315
            Width           =   2670
         End
         Begin VB.ComboBox TipoNFiscal 
            Height          =   315
            ItemData        =   "RastroItensNFEST.ctx":0052
            Left            =   870
            List            =   "RastroItensNFEST.ctx":0054
            TabIndex        =   1
            Top             =   315
            Width           =   3435
         End
         Begin VB.ComboBox Serie 
            Height          =   315
            Left            =   870
            TabIndex        =   2
            Top             =   810
            Width           =   765
         End
         Begin MSMask.MaskEdBox Numero 
            Height          =   300
            Left            =   3525
            TabIndex        =   3
            Top             =   810
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin VB.Label Label8 
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
            Left            =   390
            TabIndex        =   17
            Top             =   345
            Width           =   450
         End
         Begin VB.Label LblNumero 
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
            Left            =   2730
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   16
            Top             =   855
            Width           =   720
         End
         Begin VB.Label LblSerie 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   300
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   15
            Top             =   855
            Width           =   510
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4050
      Index           =   3
      Left            =   150
      TabIndex        =   54
      Top             =   915
      Visible         =   0   'False
      Width           =   8790
      Begin VB.Frame Frame11 
         Caption         =   "Distribuição dos Produtos"
         Height          =   3465
         Left            =   300
         TabIndex        =   55
         Top             =   255
         Width           =   8370
         Begin MSMask.MaskEdBox UMDist 
            Height          =   225
            Left            =   4425
            TabIndex        =   56
            Top             =   135
            Visible         =   0   'False
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoAlmoxDist 
            Height          =   225
            Left            =   1740
            TabIndex        =   57
            Top             =   135
            Visible         =   0   'False
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AlmoxDist 
            Height          =   225
            Left            =   3060
            TabIndex        =   58
            Top             =   135
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantDist 
            Height          =   225
            Left            =   6540
            TabIndex        =   59
            Top             =   105
            Width           =   1470
            _ExtentX        =   2593
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
         Begin MSMask.MaskEdBox ItemNFDist 
            Height          =   225
            Left            =   1005
            TabIndex        =   60
            Top             =   105
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridDist 
            Height          =   2910
            Left            =   360
            TabIndex        =   61
            Top             =   345
            Width           =   7665
            _ExtentX        =   13520
            _ExtentY        =   5133
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox QuantItemNFDist 
            Height          =   225
            Left            =   5025
            TabIndex        =   62
            Top             =   150
            Visible         =   0   'False
            Width           =   1470
            _ExtentX        =   2593
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
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4050
      Index           =   2
      Left            =   150
      TabIndex        =   22
      Top             =   915
      Visible         =   0   'False
      Width           =   8790
      Begin VB.Frame Frame3 
         Caption         =   "Itens"
         Height          =   3735
         Left            =   60
         TabIndex        =   23
         Top             =   135
         Width           =   8475
         Begin VB.TextBox DescricaoItem 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   6285
            MaxLength       =   50
            TabIndex        =   25
            Top             =   735
            Width           =   2295
         End
         Begin VB.ComboBox UnidadeMed 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1605
            TabIndex        =   24
            Text            =   "UnidadeMed"
            Top             =   285
            Width           =   645
         End
         Begin MSMask.MaskEdBox Ccl 
            Height          =   225
            Left            =   4140
            TabIndex        =   26
            Top             =   1125
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Almoxarifado 
            Height          =   225
            Left            =   5775
            TabIndex        =   27
            Top             =   330
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Desconto 
            Height          =   225
            Left            =   4995
            TabIndex        =   28
            Top             =   765
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercentDesc 
            Height          =   225
            Left            =   3810
            TabIndex        =   29
            Top             =   825
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
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
         Begin MSMask.MaskEdBox ValorUnitario 
            Height          =   225
            Left            =   3390
            TabIndex        =   30
            Top             =   345
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   225
            Left            =   2340
            TabIndex        =   31
            Top             =   315
            Width           =   990
            _ExtentX        =   1746
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
         Begin MSMask.MaskEdBox Produto 
            Height          =   225
            Left            =   315
            TabIndex        =   32
            Top             =   300
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorTotal 
            Height          =   225
            Left            =   4575
            TabIndex        =   33
            Top             =   360
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   1860
            Left            =   180
            TabIndex        =   34
            Top             =   300
            Width           =   8070
            _ExtentX        =   14235
            _ExtentY        =   3281
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7380
      ScaleHeight     =   495
      ScaleWidth      =   1560
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   1620
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "RastroItensNFEST.ctx":0056
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   570
         Picture         =   "RastroItensNFEST.ctx":01B0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1065
         Picture         =   "RastroItensNFEST.ctx":06E2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4485
      Left            =   105
      TabIndex        =   11
      Top             =   570
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   7911
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Inicial"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Distribuição"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Rastro"
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
Attribute VB_Name = "RastroItensNFEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'IDHs para Help
Const IDH_RASTRO_ITENSNFEST = 0

'distribuicao
Public gobjDistribuicao As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variáveis Globais
Public iAlterado As Integer

'Dim gcolItensNF As Collection
Dim iFrameAtual As Integer
Dim iFornecedorAlterado As Integer
Dim iClienteAlterado As Integer

'GridItens
Public objGridItens As AdmGrid
Public iGrid_Produto_Col As Integer
Public iGrid_DescProduto_Col As Integer
Public iGrid_UnidadeMed_Col As Integer
Public iGrid_Quantidade_Col As Integer
Public iGrid_Almoxarifado_Col As Integer
Public iGrid_Ccl_Col As Integer
Public iGrid_ValorUnitario_Col As Integer
Public iGrid_PercDesc_Col As Integer
Public iGrid_Desconto_Col As Integer
Public iGrid_PrecoTotal_Col As Integer

Dim WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1
Dim WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Dim WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Dim WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1

Public gobjRastreamento As ClassRastreamento
Public gobjNFiscal As ClassNFiscal

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim objNFEntrada As New ClassNFiscal

On Error GoTo Erro_BotaoGravar_Click

    'Inicia gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 75804

    'Limpa a tela de nota fiscal
    Call Limpa_Tela_RastroItens

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 75804

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165939)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long, lTransacao As Long
Dim objMovEstoque As New ClassMovEstoque
Dim objNFEntrada As New ClassNFiscal
Dim objCliente As New ClassCliente
Dim objFornecedor As New ClassFornecedor
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'verifica se todos os campos estao preenchidos ,se nao estiverem => erro
    If Len(Trim(Serie.Text)) = 0 Then gError 75805
    If Len(Trim(Numero.ClipText)) = 0 Then gError 75806
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 75843

    objTipoDocInfo.iCodigo = TipoNFiscal.ItemData(TipoNFiscal.ListIndex)
    
    'Lê o Tipo da NF
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 75845

    'Se não encontrou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 75846

    'de acordo com a sua Origem verifica se o Cliente ou Fornecedor estão preenchidos
    If objTipoDocInfo.iOrigem = DOCINFO_CLIENTE Then
        
        If Len(Trim(Cliente.ClipText)) = 0 Then gError 75830
    
        objCliente.sNomeReduzido = Cliente.Text
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 76000
        
        'Não encontrou o cliente
        If lErro = 12348 Then gError 76001
        
        objNFEntrada.lCliente = objCliente.lCodigo
        objNFEntrada.iFilialCli = Codigo_Extrai(Filial.Text)
    
    ElseIf objTipoDocInfo.iOrigem = DOCINFO_FORNECEDOR Then
        
        If Len(Trim(Fornecedor.ClipText)) = 0 Then gError 75831
        
        objFornecedor.sNomeReduzido = Fornecedor.Text
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 76002
        
        'Se não encontrou o fornecedor, erro
        If lErro = 6681 Then gError 76003
        
        objNFEntrada.lFornecedor = objFornecedor.lCodigo
        objNFEntrada.iFilialForn = Codigo_Extrai(Filial.Text)
    
    End If

    If Len(Trim(Filial.Text)) = 0 Then gError 75832
    
    'Lê a nota fiscal de entrada
    objNFEntrada.iTipoNFiscal = objTipoDocInfo.iCodigo
    objNFEntrada.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    objNFEntrada.lNumNotaFiscal = CLng(Numero.Text)
    lErro = CF("NFiscal_Le_NumFornCli", objNFEntrada)
    If lErro <> SUCESSO And lErro <> 35279 Then gError 75808

    'Nota fiscal não cadastrada
    If lErro <> SUCESSO Then gError 75809

    'Verifica se a nota já está cancelada
    If objNFEntrada.iStatus = STATUS_CANCELADO Then gError 75810
    If objNFEntrada.iFilialEmpresa <> giFilialEmpresa Then gError 75811

    'Lê os itens da nota fiscal
    lErro = CF("NFiscalItens_Le", objNFEntrada)
    If lErro <> SUCESSO Then gError 75812

    lErro = gobjRastreamento.Valida_Rastreamento(objTipoDocInfo)
    If lErro <> SUCESSO Then gError 83421

    Set objNFEntrada.objRastreamento = gobjRastreamento

    'Move os dados do Movimento de Estoque
    lErro = Move_Tela_Memoria(objNFEntrada, objMovEstoque)
    If lErro <> SUCESSO Then gError 75807

    'mover a parte do rastreamento
    lErro = gobjRastreamento.Move_Rastro_Memoria(objNFEntrada)
    If lErro <> SUCESSO Then gError 83422

    'altera os dados de rastreamento
    lErro = NFiscal_Grava_Rastro(objNFEntrada)
    If lErro <> SUCESSO Then gError 83424
    
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr
        
        Case 75805
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 75806
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)

        Case 75807, 75808, 75812, 75813, 75845, 75837, 76000, 76002, 83421, 83422, 83424

        Case 75809
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", gErr, Numero.Text)

        Case 75810
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_CANCELADA", gErr, Serie.Text, Numero.Text)

        Case 75811
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_OUTRA_FILIAL", gErr)
        
        Case 75830
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 75831
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
        
        Case 75832
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 75843
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)
                        
        Case 75846
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)
        
        Case 76001
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.sNomeReduzido)
            
        Case 76003
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165940)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

End Function

Private Sub FornecedorLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'recolhe o Nome Reduzido da tela
    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Chama a Tela de browse Fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

    Exit Sub

End Sub

Public Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor
Dim bCancel As Boolean

    Set objFornecedor = obj1

    'Coloca o Fornecedor na tela
    Fornecedor.Text = objFornecedor.lCodigo

    'Executa o Validate
    Call Fornecedor_Validate(bCancel)

    Me.Show

End Sub

Public Sub ClienteLabel_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    'Preenche NomeReduzido com o cliente da tela
    objCliente.sNomeReduzido = Cliente.Text

    'Chama Tela ClienteLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Public Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche campo Cliente
    Cliente.Text = objCliente.sNomeReduzido

    'Executa o Validate
    Call Cliente_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela_RastroItens

End Sub

Sub Limpa_Tela_RastroItens()

    Call Limpa_Tela(Me)

    'Limpa Grid
    Call Grid_Limpa(objGridItens)

    'Limpa restante dos campos
    LblFilial.Caption = ""
    LblDataEmissao.Caption = ""
    LblValor.Caption = ""
    Serie.Text = ""
    
    TipoNFiscal.ListIndex = -1
    
    EletronicaFed.Value = vbUnchecked
    
'    Set gcolItensNF = New Collection
    Filial.Clear
    
    'distribuicao
    Call gobjDistribuicao.Limpa_Tela_Distribuicao
    
    'Limpa o Frame de Rastreamento
    Call gobjRastreamento.Limpa_Tela_Rastreamento
    
    iAlterado = 0
    
End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable
        
    'Rastreamento
    lErro = gobjRastreamento.Rotina_Grid_Enable(iLinha, objControl, iLocalChamada)
    If lErro <> SUCESSO Then gError 83419
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 83419

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165941)

    End Select

    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then
        
        'Rastreamento
        If objGridInt.objGrid.Name = GridRastro.Name Then
    
            lErro = gobjRastreamento.Saida_Celula()
            If lErro <> SUCESSO Then gError 83420
    
        End If
    
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 83421

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 83421
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 83420
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165942)

    End Select

    Exit Function

End Function

Public Function Trata_Parametros(Optional objNFEntrada As ClassNFiscal) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma nota foi passada por parametro
    If Not (objNFEntrada Is Nothing) Then

        'Tenta ler a nota Fiscal passada por parametro
        lErro = CF("NFiscal_Le_NumFornCli", objNFEntrada)
        If lErro <> SUCESSO And lErro <> 35279 Then gError 75814

        'Nota Fiscal não cadastrada
        If lErro = 35279 Then gError 75815

        'Nota Fiscal cancelada
        If objNFEntrada.iStatus = STATUS_CANCELADO Then gError 75816

        'Filial Empresa atual diferente da Filial Empresa da nota fiscal
        If objNFEntrada.iFilialEmpresa <> giFilialEmpresa Then gError 75817

        'Traz a nota para a tela
        lErro = Traz_NFEntrada_Tela(objNFEntrada)
        If lErro <> SUCESSO Then gError 75818

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 75814, 75818

        Case 75815
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", gErr, objNFEntrada.lNumNotaFiscal)
            Call Limpa_Tela_RastroItens
            iAlterado = 0

        Case 75816
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_CANCELADA", gErr, Serie.Text, Numero.Text)

        Case 75817
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_OUTRA_FILIAL", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165943)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Traz_NFEntrada_Tela(objNFEntrada As ClassNFiscal) As Long
'Traz os dados da Nota Fiscal passada em objNFEntrada

Dim lErro As Long
Dim iIndice As Integer
Dim sTipoNF As String
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objItemNF As ClassItemNF
Dim objFilialEmpresa As New AdmFiliais
Dim bCancel As Boolean

On Error GoTo Erro_Traz_NFEntrada_Tela

    Set gobjNFiscal = objNFEntrada

    objTipoDocInfo.iCodigo = objNFEntrada.iTipoNFiscal

    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 75819

    'Tipo não cadastrado
    If lErro = 31415 Then gError 75820

    If (objTipoDocInfo.iTipo <> DOCINFO_NF_INT_ENTRADA And objTipoDocInfo.iTipo <> DOCINFO_NF_EXTERNA) Then gError 83486
    
    If objTipoDocInfo.iRastreavel = TIPODOCINFO_RASTREAVEL_NAO Then gError 83487
    
    gobjRastreamento.iCodigo = objTipoDocInfo.iCodigo

    'Limpa a tela NFicalEntrada
    Call Limpa_Tela_RastroItens

    'Preenche o número da NF
    If objNFEntrada.lNumNotaFiscal > 0 Then
        Numero.PromptInclude = False
        Numero.Text = CStr(objNFEntrada.lNumNotaFiscal)
        Numero.PromptInclude = True
    End If

    'preenche a serie da NF
    Serie.Text = Desconverte_Serie_Eletronica(objNFEntrada.sSerie)
    If ISSerieEletronica(objNFEntrada.sSerie) Then
        EletronicaFed.Value = vbChecked
    Else
        EletronicaFed.Value = vbUnchecked
    End If
    
    TipoNFiscal.Text = objTipoDocInfo.iCodigo
    Call TipoNFiscal_Validate(bCancel)
     
    'De acordo com a Origem do tipo Coloca o Cliente ou o fornecedor na tela
    If objTipoDocInfo.iOrigem = DOCINFO_CLIENTE Then
        Call Habilita_Cliente
        Cliente.Text = objNFEntrada.lCliente
        Call Cliente_Validate(bCancel)
        Filial.Text = objNFEntrada.iFilialCli
    ElseIf objTipoDocInfo.iOrigem = DOCINFO_FORNECEDOR Then
        Call Habilita_Fornecedor
        Fornecedor.Text = objNFEntrada.lFornecedor
        Call Fornecedor_Validate(bCancel)
        Filial.Text = objNFEntrada.iFilialForn
    End If

    Call Filial_Validate(bCancel)
    
    'Lê Filial Empresa
    objFilialEmpresa.iCodFilial = objNFEntrada.iFilialEmpresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then gError 75908

    'Filial Empresa não cadastrada
    If lErro = 27378 Then gError 75909
    
    LblFilial.Caption = CStr(objFilialEmpresa.iCodFilial) & SEPARADOR & objFilialEmpresa.sNome
    
    'Se a data não for nula coloca na Tela
    If objNFEntrada.dtDataEmissao <> DATA_NULA Then
        LblDataEmissao.Caption = Format(objNFEntrada.dtDataEmissao, "dd/mm/yyyy")
    Else
        LblDataEmissao.Caption = Format("", "dd/mm/yy")
    End If

    'Preenche o valor total da NF
    If objNFEntrada.dValorTotal > 0 Then
        LblValor.Caption = Format(objNFEntrada.dValorTotal, "Fixed")
    Else
        LblValor.Caption = Format(0, "Fixed")
    End If

    'Lê os itens da nota fiscal
    lErro = CF("NFiscalItens_Le", objNFEntrada)
    If lErro <> SUCESSO Then gError 75821

    'distribuicao
    'Lê a Distribuição dos itens da Nota Fiscal
    lErro = CF("AlocacoesNF_Le", objNFEntrada)
    If lErro <> SUCESSO Then gError 92137

    'Preenche GridItens
    lErro = Preenche_GridItens(objNFEntrada)
    If lErro <> SUCESSO Then gError 75822

    'distribuicao
    'Preenche o Grid com as Distribuições dos itens da Nota Fiscal
    lErro = gobjDistribuicao.Preenche_GridDistribuicao(objNFEntrada)
    If lErro <> SUCESSO Then gError 92136

    If objTipoDocInfo.iTipoMovtoEstoque <> 0 Then
        'Carrega ItensNF com Rastreamentos
        lErro = gobjRastreamento.Carrega_RastroItensNF(objNFEntrada)
        If lErro <> SUCESSO Then gError 75847
    End If

    iAlterado = 0
    
    Traz_NFEntrada_Tela = SUCESSO

    Exit Function

Erro_Traz_NFEntrada_Tela:

    Traz_NFEntrada_Tela = gErr

    Select Case gErr

        Case 75819, 75821, 75822, 75847, 75908, 92136

        Case 75820
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, CStr(objTipoDocInfo.iTipo))
        
        Case 75909
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)
        
        Case 83486
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_ENTRADA", gErr, objTipoDocInfo.iCodigo)
        
        Case 83487
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_RASTRO", gErr, CStr(objTipoDocInfo.iCodigo))
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165944)

    End Select

    Exit Function

End Function

'Function Carrega_RastroItensNF(objNFiscal As ClassNFiscal) As Long
''Carrega rastreamento dos itens da nota fiscal
'
'Dim lErro As Long
'Dim objItemNF As ClassItemNF
'Dim colRastroMovto As New Collection
'Dim objRastroMovto As ClassRastreamentoMovto
'Dim objRastroItemNF As ClassRastroItemNF
'Dim objAlmoxarifado As New ClassAlmoxarifado
'Dim objRastreamentoLote As New ClassRastreamentoLote
'Dim objItemMovEstoque As New ClassItemMovEstoque
'
'On Error GoTo Erro_Carrega_RastroItensNF
'
'    'Para cada item da nota fiscal
'    For Each objItemNF In objNFiscal.ColItensNF
'
'        'Lê o Almoxarifado
'        objAlmoxarifado.iCodigo = objItemNF.iAlmoxarifado
'        lErro = CF("Almoxarifado_Le",objAlmoxarifado)
'        If lErro <> SUCESSO And lErro <> 25056 Then gError 75849
'
'        'Se não encontrou Almoxarifado --> Erro
'        If lErro = 25056 Then gError 75850
'
'        'Lê item de movimento de estoque
'        objItemMovEstoque.lNumIntDocOrigem = objItemNF.lNumIntDoc
'        objItemMovEstoque.iTipoNumIntDocOrigem = TIPO_ORIGEM_ITEMNF
'        objItemMovEstoque.iFilialEmpresa = giFilialEmpresa
'        lErro = CF("MovEstoque_Le_ItemNF",objItemMovEstoque)
'        If lErro <> SUCESSO And lErro <> 75796 Then gError 75921
'
'        'Se não encontrou, erro
'        If lErro = 75796 Then gError 75922
'
'        'Lê movimentos de rastreamento vinculados ao itemNF passado ao ItemNF
'        Set colRastroMovto = New Collection
'        lErro = CF("RastreamentoMovto_Le_DocOrigem",objItemMovEstoque.lNumIntDoc, TIPO_ORIGEM_ITEMNF, colRastroMovto)
'        If lErro <> SUCESSO Then gError 75848
'
'        Set objItemNF.colRastreamento = New Collection
'
'        'Guarda as quantidades alocadas dos lotes
'        For Each objRastroMovto In colRastroMovto
'
'            Set objRastroItemNF = New ClassRastroItemNF
'
'            objRastroItemNF.dLoteQdtAlocada = objRastroMovto.dQuantidade
'            objRastroItemNF.sLote = objRastroMovto.sLote
'            objRastroItemNF.iAlmoxCodigo = objItemNF.iAlmoxarifado
'            objRastroItemNF.sAlmoxNomeRed = objAlmoxarifado.sNomeReduzido
'            objRastroItemNF.dAlmoxQtdAlocada = objItemNF.dQuantidade
'
'            'procura RastreamentoLote
'            objRastreamentoLote.sProduto = objItemNF.sProduto
'            objRastreamentoLote.iFilialOP = objRastroMovto.iFilialOP
'            objRastreamentoLote.sCodigo = objRastroMovto.sLote
'            lErro = CF("RastreamentoLote_Le",objRastreamentoLote)
'            If lErro <> SUCESSO And lErro <> 75710 Then gError 75851
'
'            'Se não encontrou, erro
'            If lErro = 75710 Then gError 75852
'
'            objRastroItemNF.dtLoteData = objRastreamentoLote.dtDataEntrada
'            objRastroItemNF.iLoteFilialOP = objRastreamentoLote.iFilialOP
'
'            'Adiciona na coleção de rastreamento
'            objItemNF.colRastreamento.Add objRastroItemNF
'
'        Next
'
'    Next
'
'    Carrega_RastroItensNF = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_RastroItensNF:
'
'    Carrega_RastroItensNF = gErr
'
'    Select Case gErr
'
'        Case 75848, 75849, 75851, 75921
'
'        Case 75850
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", gErr, objAlmoxarifado.iCodigo)
'
'        Case 75852
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_RASTREAMENTOLOTE_NAO_CADASTRADO1", gErr, objRastreamentoLote.sProduto, objRastreamentoLote.sCodigo, objRastreamentoLote.iFilialOP)
'
'        Case 75922
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_CADASTRADO", gErr, objItemNF.lNumIntDoc)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165945)
'
'    End Select
'
'    Exit Function
'
'End Function

Function Preenche_GridItens(objNFiscal As ClassNFiscal) As Long
'Preenche o Grid com os ítens da Nota Fiscal

Dim lErro As Long
Dim iIndice As Integer
Dim objItemNF As ClassItemNF
Dim sProdutoEnxuto As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sCclMascarado As String

On Error GoTo Erro_Preenche_GridItens

    iIndice = 0

    'Para cada ítem da Coleção
    For Each objItemNF In objNFiscal.ColItensNF

        iIndice = iIndice + 1

        'Formata o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objItemNF.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 75823

        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True

        'Formata Ccl
        If Trim(objItemNF.sCcl) <> "" Then

            sCclMascarado = String(STRING_CCL, 0)

            lErro = Mascara_MascararCcl(objItemNF.sCcl, sCclMascarado)
            If lErro <> SUCESSO Then gError 75826

        Else
            sCclMascarado = ""
        End If

        GridItens.TextMatrix(iIndice, iGrid_Ccl_Col) = sCclMascarado

        'Preenche o Grid
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col) = objItemNF.sDescricaoItem
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemNF.sUnidadeMed
        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemNF.dQuantidade)
        GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col) = Format(objItemNF.dPrecoUnitario, "Standard")
        If objItemNF.dPercDesc <> 0 Then GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(objItemNF.dPercDesc, "Percent")
        If objItemNF.dValorDesconto <> 0 Then GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(objItemNF.dValorDesconto, "Standard")
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col) = Format(objItemNF.dValorTotal, "Standard")

        If objItemNF.iAlmoxarifado > 0 Then

            objAlmoxarifado.iCodigo = objItemNF.iAlmoxarifado

            'Lê o Almoxarifado
            lErro = CF("Almoxarifado_Le", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 75824

            'Se não encontrou Almoxarifado --> Erro
            If lErro = 25056 Then gError 75825

            GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
            objItemNF.sAlmoxarifadoNomeRed = objAlmoxarifado.sNomeReduzido
        End If

    Next

    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = iIndice

    Preenche_GridItens = SUCESSO

    Exit Function

Erro_Preenche_GridItens:

    Preenche_GridItens = gErr

    Select Case gErr

        Case 75823
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItemNF.sProduto)

        Case 75824

        Case 75825
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.iCodigo)

        Case 75826
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", gErr, objItemNF.sCcl)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165946)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoSerie = New AdmEvento
    Set objEventoNumero = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoCliente = New AdmEvento
    
    Set objGridItens = New AdmGrid

'    Set gcolItensNF = New Collection
    
    iFrameAtual = 1

    'nao pode entrar como EMPRESA_TODA
    If giFilialEmpresa = EMPRESA_TODA Then gError 75827

    'Carrega os Tipos de Documentos relacionadas à tela
    lErro = Carrega_TiposDocInfo()
    If lErro <> SUCESSO Then gError 75969
    
    'Carrega séries
    lErro = Carrega_Serie()
    If lErro <> SUCESSO Then gError 75828
        
    'Inicializa máscara do produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 75916
    
    'Inicializa Grid de Itens
    lErro = Inicializa_Grid_Itens(objGridItens)
    If lErro <> SUCESSO Then gError 75716
    
    'distribuicao
    Set gobjDistribuicao = CreateObject("RotinasMat.ClassMATDist")
    Set gobjDistribuicao.objTela = Me
    Set gobjDistribuicao.objRastreamento = gobjRastreamento
    gobjDistribuicao.bTela = True
    
    'Rastreamento
    Set gobjRastreamento = New ClassRastreamento
    Set gobjRastreamento.objTela = Me
    gobjRastreamento.bTelaManutencao = True
    

    'Inicializa o grid de Rastreamento
    lErro = gobjRastreamento.Inicializa_Grid_Rastreamento()
    If lErro <> SUCESSO Then gError 83418

    'Inicializa o grid de Distribuicao
    lErro = gobjDistribuicao.Inicializa_GridDist()
    If lErro <> SUCESSO Then gError 92135
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr

        Case 75827
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EMPRESA_INVALIDA", gErr)

        Case 75828, 75716, 75969, 83418, 92135

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165947)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Function Carrega_TiposDocInfo() As Long
'Carrega na os Tipo de Documentos relacionados com a tela de Nota Fiscal de Entrada

Dim lErro As Long
Dim colTipoDocInfo As New colTipoDocInfo
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Carrega_TiposDocInfo

    Set colTipoDocInfo = gobjCRFAT.colTiposDocInfo

    'Carrega na combo só os Tipos ligados essa tela
    For Each objTipoDocInfo In colTipoDocInfo
            
        'Se o tipo de nota fiscal for de entrada
        If (objTipoDocInfo.iTipo = DOCINFO_NF_INT_ENTRADA Or objTipoDocInfo.iTipo = DOCINFO_NF_EXTERNA) And objTipoDocInfo.iRastreavel = TIPODOCINFO_RASTREAVEL_SIM Then
            
            TipoNFiscal.AddItem CStr(objTipoDocInfo.iCodigo) & SEPARADOR & objTipoDocInfo.sNomeReduzido
            TipoNFiscal.ItemData(TipoNFiscal.NewIndex) = objTipoDocInfo.iCodigo
        
        End If
    
    Next

    Carrega_TiposDocInfo = SUCESSO

    Exit Function

Erro_Carrega_TiposDocInfo:

    Carrega_TiposDocInfo = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165948)

    End Select

    Exit Function

End Function

Private Function Carrega_Serie() As Long
'Carrega a combo de Séries com as séries lidas do BD

Dim lErro As Long
Dim colSerie As New colSerie
Dim objSerie As ClassSerie
Dim sSerieAnt As String

On Error GoTo Erro_Carrega_Serie

    'Lê as séries
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then gError 75970

    'Carrega na combo
    For Each objSerie In colSerie
        If UCase(sSerieAnt) <> UCase(Desconverte_Serie_Eletronica(objSerie.sSerie)) Then Serie.AddItem Desconverte_Serie_Eletronica(objSerie.sSerie)
        sSerieAnt = Desconverte_Serie_Eletronica(objSerie.sSerie)
    Next

    Carrega_Serie = SUCESSO

    Exit Function

Erro_Carrega_Serie:

    Carrega_Serie = gErr

    Select Case gErr

        Case 75970

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165949)

    End Select

    Exit Function

End Function

Private Sub BotaoNFiscal_Click()

Dim lErro As Long
Dim objNFEntrada As New ClassNFiscal
Dim objCliente As New ClassCliente
Dim objFornecedor As New ClassFornecedor
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_BotaoNFiscal_Click
    
    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 75988
    If Len(Trim(Serie.Text)) = 0 Then gError 75989
    If Len(Trim(Numero.Text)) = 0 Then gError 75990
    
    objTipoDocInfo.iCodigo = TipoNFiscal.ItemData(TipoNFiscal.ListIndex)
    
    'Lê o Tipo da NF
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 75991

    'Se não encontrou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 75992

    'de acordo com a sua Origem verifica se o Cliente ou Fornecedor estão preenchidos
    If objTipoDocInfo.iOrigem = DOCINFO_CLIENTE Then
                
        If Len(Trim(Cliente.ClipText)) = 0 Then gError 75993
            
        objCliente.sNomeReduzido = Cliente.Text
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 75841
        
        'Não encontrou o cliente
        If lErro = 12348 Then gError 75833
        
        objNFEntrada.lCliente = objCliente.lCodigo
        objNFEntrada.iFilialCli = Codigo_Extrai(Filial.Text)

    ElseIf objTipoDocInfo.iOrigem = DOCINFO_FORNECEDOR Then
        
        If Len(Trim(Fornecedor.ClipText)) = 0 Then gError 75994
    
        objFornecedor.sNomeReduzido = Fornecedor.Text
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 75834
        
        'Se não encontrou o fornecedor, erro
        If lErro = 6681 Then gError 75836
        objNFEntrada.lFornecedor = objFornecedor.lCodigo
        objNFEntrada.iFilialForn = Codigo_Extrai(Filial.Text)
    
    End If
    
    'Filial
    If Len(Trim(Filial.Text)) = 0 Then gError 76004
    
    objNFEntrada.iTipoNFiscal = objTipoDocInfo.iCodigo
    objNFEntrada.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    objNFEntrada.lNumNotaFiscal = CLng(Numero.Text)
    
    'Lê a nota fiscal
    lErro = CF("NFiscal_Le_NumFornCli", objNFEntrada)
    If lErro <> SUCESSO And lErro <> 35279 Then gError 62004
    
    'Se não encontrou a nota fiscal, erro
    If lErro = 35279 Then gError 75837
    
    objNFEntrada.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    objNFEntrada.lNumNotaFiscal = CLng(Numero.Text)
    objNFEntrada.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)
    
    'Traz nota fiscal para a tela
    lErro = Traz_NFEntrada_Tela(objNFEntrada)
    If lErro <> SUCESSO Then gError 75840
    
    Exit Sub
    
Erro_BotaoNFiscal_Click:
    
    Select Case gErr
                        
        Case 75833
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.sNomeReduzido)
            
        Case 75836
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
        
        Case 75837
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", gErr, objNFEntrada.lNumNotaFiscal)
        
        Case 75988
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)
            
        Case 75989
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
        
        Case 75990
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)
        
        Case 62004, 75840, 75991, 75841
        
        Case 75992
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, CStr(objTipoDocInfo.iCodigo))
        
        Case 75993
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
                
        Case 75994
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
        
        Case 76004
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165950)
        
    End Select
    
    Exit Sub

End Sub

Private Sub TipoNFiscal_Click()

Dim objTipoDocInfo As New ClassTipoDocInfo
Dim lErro As Long
Dim iDocInfoAux As Integer

On Error GoTo Erro_TipoNFiscal_Click
    
    'Se não preencheu tipo, sai da rotina
    If Len(Trim(TipoNFiscal.Text)) = 0 Then Exit Sub
    
    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)

    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 75971

    'se não estiver cadastrado ==> erro
    If lErro = 31415 Then gError 75972
    
    'Habilita cliente ou fornecedor
    If objTipoDocInfo.iOrigem = DOCINFO_CLIENTE Then
                
        Call Habilita_Cliente
    
    ElseIf objTipoDocInfo.iOrigem = DOCINFO_FORNECEDOR Then
                    
        Call Habilita_Fornecedor
    
    End If
    
    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub

Erro_TipoNFiscal_Click:

    Select Case gErr

        Case 75971

        Case 75972
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, CStr(objTipoDocInfo.iCodigo))

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165951)

    End Select

    Exit Sub

End Sub

Private Sub Habilita_Cliente()
'Desabilita o Fornecededor e Habilita o Cliente

Dim bCancel As Boolean
        
    ClienteLabel.Visible = True
    Cliente.Visible = True
    FornecedorLabel.Visible = False
    Fornecedor.Visible = False
    iClienteAlterado = REGISTRO_ALTERADO
    Call Cliente_Validate(bCancel)

End Sub

Private Sub Habilita_Fornecedor()
'Desabilita o Cliente e habilita o Fornecedor

Dim bCancel As Boolean
    
    FornecedorLabel.Visible = True
    Fornecedor.Visible = True
    Cliente.Visible = False
    ClienteLabel.Visible = False
    iFornecedorAlterado = REGISTRO_ALTERADO
    Call Fornecedor_Validate(bCancel)

End Sub

Private Sub TipoNFiscal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_TipoNFiscal_Validate

    'Verifica se o tipo está preenchido
    If Len(Trim(TipoNFiscal.Text)) = 0 Then Exit Sub
    
    'Verifica se foi selecionado
    If TipoNFiscal.List(TipoNFiscal.ListIndex) = TipoNFiscal.Text Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(TipoNFiscal, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 75973
    
    If lErro <> SUCESSO Then gError 75975 'Não conseguiu

    Exit Sub

Erro_TipoNFiscal_Validate:

    Cancel = True

    Select Case gErr

        Case 75973

        Case 75975
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_RASTRO", gErr, TipoNFiscal.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165952)

    End Select

    Exit Sub

End Sub

Private Sub TipoNFiscal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    If iFornecedorAlterado = 1 Then

        If Len(Trim(Fornecedor.Text)) > 0 Then

            'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le3(Fornecedor, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then gError 75976

            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then gError 75977

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)

            If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then
            
                If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ
                
                'Seleciona filial na Combo Filial
                Call CF("Filial_Seleciona", Filial, iCodFilial)
                
            End If

        ElseIf Len(Trim(Fornecedor.Text)) = 0 Then

            Filial.Clear

        End If

        iFornecedorAlterado = 0

    End If

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 75976, 75977

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165953)

    End Select

    Exit Sub

End Sub

Public Sub Fornecedor_Change()

    iAlterado = REGISTRO_ALTERADO
    iFornecedorAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Cliente_Change()

    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    If iClienteAlterado = 1 Then

        If Len(Trim(Cliente.Text)) > 0 Then

            lErro = TP_Cliente_Le3(Cliente, objCliente, iCodFilial)
            If lErro <> SUCESSO Then gError 75978

            lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
            If lErro <> SUCESSO Then gError 75979

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)

            'Seleciona filial na Combo Filial
            If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then
            
                If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ
                
                Call CF("Filial_Seleciona", Filial, iCodFilial)
                        
            End If
            
        ElseIf Len(Trim(Cliente.Text)) = 0 Then

            Filial.Clear

        End If

        iClienteAlterado = 0

    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 75978, 75979

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165954)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objFilialCliente As New ClassFilialCliente
Dim sNomeRed As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 75982

    'Se não encontrou o ítem com o código informado
    If lErro = 6730 Then

        If Fornecedor.Visible = True Then

            'Verifica se o Fornecedor foi preenchido
            If Len(Trim(Fornecedor.Text)) = 0 Then gError 75984

            sNomeRed = Fornecedor.Text

            objFilialFornecedor.iCodFilial = iCodigo

            'Pesquisa se existe a Filial do Fornecedor
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sNomeRed, objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then gError 75980

            'Se não encontrou a Filial do Fornecedor --> erro
            If lErro = 18272 Then gError 75983

            'Coloca a Filial do Fornecedor na tela
            Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

        Else

            'Verifica se Cliente está preenchido
            If Len(Trim(Cliente.ClipText)) = 0 Then gError 75985

            sNomeRed = Cliente.Text

            'Lê a Filial do Cliente
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sNomeRed, objFilialCliente)
            If lErro <> SUCESSO And lErro <> 17660 Then gError 75981

            'Se não encontrou a Filial do Cliente --> erro
            If lErro = 17660 Then gError 75986

            'Coloca a Filial do Fornecedor
            Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome
        
        End If

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 75987

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 75980, 75981, 75982

        Case 75983
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 75984
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 75985
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 75986
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 75987
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165955)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filial_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Ccl")
    objGridInt.colColuna.Add ("Valor Unitário")
    objGridInt.colColuna.Add ("% Desc.")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Valor Total")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoItem.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (Ccl.Name)
    objGridInt.colCampo.Add (ValorUnitario.Name)
    objGridInt.colCampo.Add (PercentDesc.Name)
    objGridInt.colCampo.Add (Desconto.Name)
    objGridInt.colCampo.Add (ValorTotal.Name)

    'Colunas da Grid
    iGrid_Produto_Col = 1
    iGrid_DescProduto_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_Quantidade_Col = 4
    iGrid_Almoxarifado_Col = 5
    iGrid_Ccl_Col = 6
    iGrid_ValorUnitario_Col = 7
    iGrid_PercDesc_Col = 8
    iGrid_Desconto_Col = 9
    iGrid_PrecoTotal_Col = 10
    
    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_NF + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 8

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 500

    'Proibido incluir e excluir linhas
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_NAO_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Function Move_Tela_Memoria(objNFEntrada As ClassNFiscal, objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long
Dim objItemNF As New ClassItemNF
Dim objItemMovEstoque As New ClassItemMovEstoque
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iTipoMovtoEstoque As Integer
Dim sDocOrigem As String
Dim objRastroItemNF As ClassRastroItemNF
Dim objRastroMovto As ClassRastreamentoMovto

On Error GoTo Erro_Move_Tela_Memoria

    'Lê TipoDocInfo da nota fiscal
    objTipoDocInfo.iCodigo = objNFEntrada.iTipoNFiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 75854
    
    'Se não encontrou Tipo, erro
    If lErro = 31415 Then gError 75855
    
    iTipoMovtoEstoque = objTipoDocInfo.iTipoMovtoEstoque
    sDocOrigem = objTipoDocInfo.sSigla

    If iTipoMovtoEstoque > 0 Then

        Set objMovEstoque = New ClassMovEstoque

        'Guarda dados do Movimento de Estoque
        objMovEstoque.iFilialEmpresa = giFilialEmpresa
        objMovEstoque.iTipoMov = iTipoMovtoEstoque
        objMovEstoque.lCliente = objNFEntrada.lCliente
        objMovEstoque.lFornecedor = objNFEntrada.lFornecedor
        objMovEstoque.sDocOrigem = sDocOrigem & " " & objNFEntrada.sSerie & " " & CStr(objNFEntrada.lNumNotaFiscal)
                                
'        'Adiciona itens ao Movimento
'        For Each objItemNF In gcolItensNF
'
'            If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
'
'                'Lê Numintdoc e código do movimento estoque a partir de um item da nota fiscal
'                objItemMovEstoque.lNumIntDocOrigem = objItemNF.lNumIntDoc
'                objItemMovEstoque.iTipoNumIntDocOrigem = TIPO_ORIGEM_ITEMNF
'                objItemMovEstoque.iFilialEmpresa = giFilialEmpresa
'                lErro = CF("MovEstoque_Le_ItemNF",objItemMovEstoque)
'                If lErro <> SUCESSO And lErro <> 75796 Then gError 75856
'
'                'Se não encontrou, erro
'                If lErro = 75796 Then gError 75857
'
'                'Guarda código e data do Movimento de estoque
'                objMovEstoque.lCodigo = objItemMovEstoque.lCodigo
'                objMovEstoque.dtData = objItemMovEstoque.dtData
'
'                objItemMovEstoque.sProduto = objItemNF.sProduto
'                objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
'                objItemMovEstoque.sSiglaUM = objItemNF.sUnidadeMed
'                objItemMovEstoque.sSiglaUMEst = objItemNF.sUMEstoque
'
'                Set objItemMovEstoque.colRastreamentoMovto = New Collection
'
'                'Guarda o Rastreamento dos ItensNF
'                For Each objRastroItemNF In objItemNF.colRastreamento
'
'                    Set objRastroMovto = New ClassRastreamentoMovto
'                    objRastroMovto.dQuantidade = objRastroItemNF.dLoteQdtAlocada
'                    objRastroMovto.iTipoDocOrigem = TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE
'                    objRastroMovto.sLote = objRastroItemNF.sLote
'                    objRastroMovto.lNumIntDocOrigem = objItemMovEstoque.lNumIntDoc
'                    objRastroMovto.sProduto = objItemNF.sProduto
'
'                    'Adiciona objRastroMovto na coleção de Rastreamento
'                    objItemMovEstoque.colRastreamentoMovto.Add objRastroMovto
'
'                Next
'
'                'Adiciona na coleção
'                Call objMovEstoque.colItens.Add(objItemMovEstoque.lNumIntDoc, objItemMovEstoque.iTipoMov, 0, 0, objItemMovEstoque.sProduto, "", objItemMovEstoque.sSiglaUM, 0, 0, "", objItemMovEstoque.lNumIntDocOrigem, "", 0, "", "", "", "", 0, objItemMovEstoque.colRastreamentoMovto, objItemMovEstoque.colApropriacaoInsumo, DATA_NULA)
'
'            End If
'        Next
                
    End If
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
        
        Case 75854, 75856
        
        Case 75855
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, CStr(objTipoDocInfo.iCodigo))
        
        Case 75857
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_CADASTRADO", gErr, objItemNF.lNumIntDoc)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165956)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    'Libera as variáveis globais da tela
    Set objEventoSerie = Nothing
    Set objEventoNumero = Nothing
    Set objEventoCliente = Nothing
    Set objEventoFornecedor = Nothing
    
    'distribuicao
    Set gobjDistribuicao = Nothing
    
    Set gobjRastreamento = Nothing
    
    Set objGridItens = Nothing
    Set gobjNFiscal = Nothing

'    Set gcolItensNF = Nothing

End Sub

Public Sub BotaoLotes_Click()
'Chama a tela de Lote de Rastreamento

Dim lErro As Long

On Error GoTo Erro_BotaoLotes_Click
    
    Call gobjRastreamento.BotaoLotes_Click
                    
    Exit Sub

Erro_BotaoLotes_Click:

    Select Case gErr
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165957)
    
    End Select
    
    Exit Sub

End Sub

Public Sub ItemNFRastro_Change()
'Rastreamento

    Call gobjRastreamento.ItemNFRastro_Change

End Sub

Public Sub ItemNFRastro_GotFocus()
'Rastreamento

    Call gobjRastreamento.ItemNFRastro_GotFocus

End Sub

Public Sub ItemNFRastro_KeyPress(KeyAscii As Integer)
'Rastreamento

    Call gobjRastreamento.ItemNFRastro_KeyPress(KeyAscii)

End Sub

Public Sub ItemNFRastro_Validate(Cancel As Boolean)
'Rastreamento

    Call gobjRastreamento.ItemNFRastro_Validate(Cancel)

End Sub

Public Sub LoteRastro_Change()
'Rastreamento

    Call gobjRastreamento.LoteRastro_Change

End Sub

Public Sub LoteRastro_GotFocus()
'Rastreamento

    Call gobjRastreamento.LoteRastro_GotFocus

End Sub

Public Sub LoteRastro_KeyPress(KeyAscii As Integer)
'Rastreamento

    Call gobjRastreamento.LoteRastro_KeyPress(KeyAscii)

End Sub

Public Sub LoteRastro_Validate(Cancel As Boolean)
'Rastreamento

    Call gobjRastreamento.LoteRastro_Validate(Cancel)

End Sub

Public Sub FilialOPRastro_Change()
'Rastreamento

    Call gobjRastreamento.FilialOPRastro_Change

End Sub

Public Sub FilialOPRastro_GotFocus()
'Rastreamento

    Call gobjRastreamento.FilialOPRastro_GotFocus

End Sub

Public Sub FilialOPRastro_KeyPress(KeyAscii As Integer)
'Rastreamento

    Call gobjRastreamento.FilialOPRastro_KeyPress(KeyAscii)

End Sub

Public Sub FilialOPRastro_Validate(Cancel As Boolean)
'Rastreamento

    Call gobjRastreamento.FilialOPRastro_Validate(Cancel)

End Sub

Public Sub QuantLoteRastro_Change()
'Rastreamento

    Call gobjRastreamento.QuantLoteRastro_Change

End Sub

Public Sub QuantLoteRastro_GotFocus()
'Rastreamento

    Call gobjRastreamento.QuantLoteRastro_GotFocus

End Sub

Public Sub QuantLoteRastro_KeyPress(KeyAscii As Integer)
'Rastreamento

    Call gobjRastreamento.QuantLoteRastro_KeyPress(KeyAscii)

End Sub

Public Sub QuantLoteRastro_Validate(Cancel As Boolean)
'Rastreamento

    Call gobjRastreamento.QuantLoteRastro_Validate(Cancel)

End Sub


'Private Sub BotaoRastreamento_Click()
'
'    Call Chama_Tela("RastroProdNFEST", gcolItensNF)
'
'End Sub

Private Sub LblNumero_Click()

Dim lErro As Long
Dim objNFEntrada As New ClassNFiscal
Dim colSelecao As Collection

On Error GoTo Erro_LblNumero_Click

    objNFEntrada.lNumNotaFiscal = StrParaLong(Numero.Text)
    
    Call Chama_Tela("NFiscalEntradaTodasLista", colSelecao, objNFEntrada, objEventoNumero)

    Exit Sub

Erro_LblNumero_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165958)

    End Select

    Exit Sub

End Sub

Private Sub LblSerie_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objSerie As New ClassSerie
Dim colSelecao As Collection

On Error GoTo Erro_LblSerie_Click

    'transfere a série da tela p\ o objSerie
    objSerie.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)

    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerie)

    Exit Sub

Erro_LblSerie_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165959)

    End Select

    Exit Sub

End Sub

Private Sub MSFlexGrid1_Click()

End Sub

Private Sub Numero_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Numero_GotFocus()

    Call MaskEdBox_TrataGotFocus(Numero, iAlterado)

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

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFEntrada As ClassNFiscal

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objNFEntrada = obj1
    
    'Lê dados da Nota Fiscal
    lErro = CF("NFiscal_Le_NumFornCli", objNFEntrada)
    If lErro <> SUCESSO And lErro <> 35279 Then gError 75836
    
    'Se não encontrou a nota fiscal, erro
    If lErro = 35279 Then gError 75837
    
    If objNFEntrada.iStatus = STATUS_CANCELADO Then gError 75838
    If objNFEntrada.iFilialEmpresa <> giFilialEmpresa Then gError 75839

    'Traz a NotaFiscal de Entrada para a a tela
    lErro = Traz_NFEntrada_Tela(objNFEntrada)
    If lErro <> SUCESSO Then gError 75835
    
    Me.Show

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case gErr

        Case 75835, 75836
        
        Case 75837
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", gErr, objNFEntrada.lNumNotaFiscal)
        
        Case 75838
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_CANCELADA", gErr, objNFEntrada.sSerie, objNFEntrada.lNumNotaFiscal)

        Case 75839
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_OUTRA_FILIAL", gErr)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165960)

    End Select

    Exit Sub

End Sub

Private Sub objEventoSerie_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objSerie As ClassSerie
Dim bCancel As Boolean

On Error GoTo Erro_objEventoSerie_evSelecao

    Set objSerie = obj1

    Serie.Text = Desconverte_Serie_Eletronica(objSerie.sSerie)
    If ISSerieEletronica(objSerie.sSerie) Then
        EletronicaFed.Value = vbChecked
    Else
        EletronicaFed.Value = vbUnchecked
    End If
    Call Serie_Validate(bCancel)

    Me.Show

    Exit Sub

Erro_objEventoSerie_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165961)

    End Select

    Exit Sub

End Sub

Private Sub Serie_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Serie_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objNFEntrada As New ClassNFiscal
Dim objSerie As New ClassSerie

On Error GoTo Erro_Serie_Validate

    'Verifica se a série está preenchida
    If Len(Trim(Serie.Text)) > 0 Then

        objSerie.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)

        lErro = CF("Serie_Le", objSerie)
        If lErro <> SUCESSO And lErro <> 22202 Then gError 75842
        If lErro = 22202 Then gError 75844

    End If
        
    Exit Sub

Erro_Serie_Validate:

    Cancel = True

    Select Case gErr

        Case 75842

        Case 75844
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objSerie.sSerie)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165962)

    End Select

    Exit Sub

End Sub

'Tratamento do Grid
Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGridItens, iAlterado)

End Sub

Private Sub GridItens_GotFocus()

    Call Grid_Recebe_Foco(objGridItens)

End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItens)

End Sub

Private Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridItens)

End Sub

Private Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub

Public Sub GridRastro_Click()
'Rastreamento

    Call gobjRastreamento.GridRastro_Click

End Sub

Public Sub GridRastro_EnterCell()
'Rastreamento

    Call gobjRastreamento.GridRastro_EnterCell

End Sub

Public Sub GridRastro_GotFocus()
'Rastreamento

    Call gobjRastreamento.GridRastro_GotFocus

End Sub

Public Sub GridRastro_KeyPress(KeyAscii As Integer)
'Rastreamento

    Call gobjRastreamento.GridRastro_KeyPress(KeyAscii)

End Sub

Public Sub GridRastro_LeaveCell()
'Rastreamento

    Call gobjRastreamento.GridRastro_LeaveCell
    
End Sub

Public Sub GridRastro_Validate(Cancel As Boolean)
'Rastreamento

    Call gobjRastreamento.GridRastro_Validate(Cancel)

End Sub

Public Sub GridRastro_RowColChange()
'Rastreamento

    Call gobjRastreamento.GridRastro_RowColChange

End Sub

Public Sub GridRastro_Scroll()
'Rastreamento

    Call gobjRastreamento.GridRastro_Scroll

End Sub

Public Sub GridRastro_KeyDown(KeyCode As Integer, Shift As Integer)
'Rastreamento

    Call gobjRastreamento.GridRastro_KeyDown(KeyCode, Shift)

End Sub



'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RASTRO_ITENSNFEST
    Set Form_Load_Ocx = Me
    Caption = "Rastreamento de Itens de Nota Fiscais de Entrada"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RastroItensNFEST"

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

        If Me.ActiveControl Is Serie Then
            Call LblSerie_Click
        ElseIf Me.ActiveControl Is Numero Then
            Call LblNumero_Click
        ElseIf Me.ActiveControl Is Fornecedor Then
            Call FornecedorLabel_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call ClienteLabel_Click
        Else
            Call gobjRastreamento.objUserControl_KeyDown(KeyCode, Shift)
        End If

    End If

End Sub

Private Sub LblDataEmissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblDataEmissao, Source, X, Y)
End Sub

Private Sub LblDataEmissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblDataEmissao, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LblFilial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblFilial, Source, X, Y)
End Sub

Private Sub LblFilial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblFilial, Button, Shift, X, Y)
End Sub

Private Sub LabelFilial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilial, Source, X, Y)
End Sub

Private Sub LabelFilial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilial, Button, Shift, X, Y)
End Sub

Private Sub LblSerie_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblSerie, Source, X, Y)
End Sub

Private Sub LblSerie_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblSerie, Button, Shift, X, Y)
End Sub

Private Sub LblNumero_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblNumero, Source, X, Y)
End Sub

Private Sub LblNumero_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblNumero, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LblValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblValor, Source, X, Y)
End Sub

Private Sub LblValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblValor, Button, Shift, X, Y)
End Sub

Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub

Private Sub ClienteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ClienteLabel, Source, X, Y)
End Sub

Private Sub ClienteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ClienteLabel, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

''********************** copiada de outras telas*********************************
''copiada de Produção Entrada
'Function RastreamentoMovto_Le_DocOrigem(lNumIntDocOrigem As Long, iTipoDocOrigem As Integer, colRastreamentoMovto As Collection) As Long
''Lê a tabela de RastreamentoMovto através do Movimento de Estoque
'
'Dim lErro As Long
'Dim tRastreamentoMovto As typeRastreamentoMovto
'Dim objRastreamentoMovto As New ClassRastreamentoMovto
'Dim lComando As Long
'
'On Error GoTo Erro_RastreamentoMovto_Le_DocOrigem
'
'    'Abertura de comando
'    lComando = Comando_Abrir
'    If lComando = 0 Then gError 78411
'
'    tRastreamentoMovto.sProduto = String(STRING_PRODUTO, 0)
'    tRastreamentoMovto.sLote = String(STRING_OPCODIGO, 0)
'
'    'Lê o Rastreamento Movto
'    lErro = Comando_Executar(lComando, "SELECT RastreamentoMovto.NumIntDoc, RastreamentoMovto.TipoDocOrigem, RastreamentoMovto.NumIntDocOrigem, RastreamentoMovto.Produto, RastreamentoMovto.Quantidade, RastreamentoLote.Lote, RastreamentoLote.FilialOP FROM RastreamentoMovto, RastreamentoLote WHERE RastreamentoMovto.NumIntDocLote = RastreamentoLote.NumIntDoc AND TipoDocOrigem = ? AND NumIntDocOrigem = ?" _
'        , tRastreamentoMovto.lNumIntDoc, tRastreamentoMovto.iTipoDocOrigem, tRastreamentoMovto.lNumIntDocOrigem, tRastreamentoMovto.sProduto, tRastreamentoMovto.dQuantidade, tRastreamentoMovto.sLote, tRastreamentoMovto.iFilialOP, TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE, lNumIntDocOrigem)
'    If lErro <> AD_SQL_SUCESSO Then gError 78412
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78413
'
'    Do While lErro = AD_SQL_SUCESSO
'
'        Set objRastreamentoMovto = New ClassRastreamentoMovto
'
'        'passa para o objeto
'        objRastreamentoMovto.lNumIntDoc = tRastreamentoMovto.lNumIntDoc
'        objRastreamentoMovto.iTipoDocOrigem = tRastreamentoMovto.iTipoDocOrigem
'        objRastreamentoMovto.lNumIntDocOrigem = tRastreamentoMovto.lNumIntDocOrigem
'        objRastreamentoMovto.sProduto = tRastreamentoMovto.sProduto
'        objRastreamentoMovto.dQuantidade = tRastreamentoMovto.dQuantidade
'        objRastreamentoMovto.sLote = tRastreamentoMovto.sLote
'        objRastreamentoMovto.iFilialOP = tRastreamentoMovto.iFilialOP
'
'        colRastreamentoMovto.Add objRastreamentoMovto
'
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78421
'
'    Loop
'
'    Call Comando_Fechar(lComando)
'
'    RastreamentoMovto_Le_DocOrigem = SUCESSO
'
'    Exit Function
'
'Erro_RastreamentoMovto_Le_DocOrigem:
'
'    RastreamentoMovto_Le_DocOrigem = gErr
'
'    Select Case gErr
'
'        Case 78411
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 78412, 78413, 78421
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_RASTREAMENTOMOVTO", gErr)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165963)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function

''Copiada de RastreamentoLote
'Function RastreamentoLote_Le(objRastroLote As ClassRastreamentoLote) As Long
''Lê rastreamento do lote a partir do produto, filialOP e código do lote passados
'
'Dim lErro As Long
'Dim lComando As Long
'Dim tRastroLote As typeRastreamentoLote
'
'On Error GoTo Erro_RastreamentoLote_Le
'
'    'Abertura dos comandos
'    lComando = Comando_Abrir()
'    If lErro <> SUCESSO Then gError 75707
'
'    tRastroLote.sObservacao = String(STRING_RASTRO_OBSERVACAO, 0)
'
'    'Lê dados de RastrementoLote a partir de Produto, FilialOP e Lote
'    lErro = Comando_Executar(lComando, "SELECT DataValidade, DataEntrada, DataFabricacao, Observacao FROM RastreamentoLote WHERE Produto = ? AND Lote = ? AND FilialOP = ?", tRastroLote.dtDataValidade, tRastroLote.dtDataEntrada, tRastroLote.dtDataFabricacao, tRastroLote.sObservacao, objRastroLote.sProduto, objRastroLote.sCodigo, objRastroLote.iFilialOP)
'    If lErro <> AD_SQL_SUCESSO Then gError 75708
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 75709
'
'    'Se não encontrou, erro
'    If lErro = AD_SQL_SEM_DADOS Then gError 75710
'
'    'Fechamento dos comandos
'    Call Comando_Fechar(lComando)
'
'    RastreamentoLote_Le = SUCESSO
'
'    Exit Function
'
'Erro_RastreamentoLote_Le:
'
'    RastreamentoLote_Le = gErr
'
'    Select Case gErr
'
'        Case 75707
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 75708, 75709
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTE", gErr)
'
'        Case 75710 'RastreamentoLote não cadastrado
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 165964)
'
'    End Select
'
'    'Fechamento dos comandos
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
''copiada de RastroItensNFFAT
'Function MovEstoque_Le_ItemNF(objItemMovEstoque As ClassItemMovEstoque) As Long
''Lê o NumIntDoc e Código do MovimentoEstoque a partir do NumIntDoc do ItemNF passado
'
'Dim lErro As Long
'Dim lComando As Long
'Dim lNumIntDoc As Long
'Dim lCodigo As Long
'Dim dtData As Date
'
'On Error GoTo Erro_MovEstoque_Le_ItemNF
'
'    'Abre comandos
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 75793
'
'    'Lê NumIntDoc de MovimentoEstoque
'    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, Data FROM MovimentoEstoque WHERE NumIntDocOrigem = ? AND TipoNumIntDocOrigem = ? AND FilialEmpresa = ?", lNumIntDoc, lCodigo, dtData, objItemMovEstoque.lNumIntDocOrigem, objItemMovEstoque.iTipoNumIntDocOrigem, objItemMovEstoque.iFilialEmpresa)
'    If lErro <> AD_SQL_SUCESSO Then gError 75794
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 75795
'
'    'Se não encontrou o movimento estoque, erro
'    If lErro = AD_SQL_SEM_DADOS Then gError 75796
'
'    objItemMovEstoque.lNumIntDoc = lNumIntDoc
'    objItemMovEstoque.lCodigo = lCodigo
'    objItemMovEstoque.dtData = dtData
'
'    'Fecha comandos
'    Call Comando_Fechar(lComando)
'
'    MovEstoque_Le_ItemNF = SUCESSO
'
'    Exit Function
'
'Erro_MovEstoque_Le_ItemNF:
'
'    MovEstoque_Le_ItemNF = gErr
'
'    Select Case gErr
'
'        Case 75793
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 75794, 75795
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)
'
'        Case 75796
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165965)
'
'    End Select
'
'    'Fecha comandos
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
''***************************************************************************
'
'
    
    
'???? Transferir para RotinasMat
Public Function NFiscal_Grava_Rastro(objNFiscal As ClassNFiscal) As Long
'altera os rastreamentos associados à nota fiscal passada como parametro

Dim lTransacao As Long
Dim lErro As Long

On Error GoTo Erro_NFiscal_Grava_Rastro

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 83425
    
    'Se a nota já existir, permite alterar os rastreamentos
    lErro = objNFiscal.objRastreamento.Altera_Rastreamento(objNFiscal)
    If lErro <> SUCESSO Then gError 83423
    
    lErro = CF("NFiscal_Trata_MsgItem", objNFiscal)
    If lErro <> SUCESSO Then gError 83423
    
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 83426
    
    NFiscal_Grava_Rastro = SUCESSO
    
    Exit Function

Erro_NFiscal_Grava_Rastro:

    NFiscal_Grava_Rastro = gErr

    Select Case gErr

        Case 83423

        Case 83425
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 83426
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165966)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Public Sub BotaoSerie_Click()
'Chama a tela de Lote de Rastreamento

Dim lErro As Long

On Error GoTo Erro_BotaoSerie_Click
    
    Call gobjRastreamento.BotaoSerie_Click
                    
    Exit Sub

Erro_BotaoSerie_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141800)
    
    End Select
    
    Exit Sub

End Sub
