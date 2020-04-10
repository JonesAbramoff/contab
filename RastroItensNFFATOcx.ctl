VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.UserControl RastroItensNFFATOcx 
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9120
   KeyPreview      =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   9120
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame17"
      Height          =   4080
      Index           =   4
      Left            =   255
      TabIndex        =   39
      Top             =   795
      Visible         =   0   'False
      Width           =   8805
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
         Left            =   120
         TabIndex        =   68
         Top             =   3600
         Width           =   1665
      End
      Begin VB.Frame Frame18 
         Caption         =   "Rastreamento do Produto"
         Height          =   3390
         Left            =   120
         TabIndex        =   41
         Top             =   90
         Width           =   8490
         Begin VB.ComboBox ProdutoRastro 
            Height          =   315
            ItemData        =   "RastroItensNFFATOcx.ctx":0000
            Left            =   1305
            List            =   "RastroItensNFFATOcx.ctx":000D
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   945
            Width           =   1740
         End
         Begin VB.ComboBox EscaninhoRastro 
            Height          =   315
            ItemData        =   "RastroItensNFFATOcx.ctx":002C
            Left            =   3690
            List            =   "RastroItensNFFATOcx.ctx":0039
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   285
            Width           =   1215
         End
         Begin MSMask.MaskEdBox UMRastro 
            Height          =   240
            Left            =   3075
            TabIndex        =   43
            Top             =   270
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
            Left            =   135
            TabIndex        =   44
            Top             =   840
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
            Left            =   1650
            TabIndex        =   45
            Top             =   285
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
            Left            =   1845
            TabIndex        =   46
            Top             =   480
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
            Left            =   2820
            TabIndex        =   47
            Top             =   480
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
            Left            =   5580
            TabIndex        =   48
            Top             =   480
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
            Left            =   3960
            TabIndex        =   49
            Top             =   465
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
            Left            =   6735
            TabIndex        =   50
            Top             =   495
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
            Height          =   2895
            Left            =   120
            TabIndex        =   51
            Top             =   255
            Width           =   8145
            _ExtentX        =   14367
            _ExtentY        =   5106
            _Version        =   393216
            Rows            =   51
            Cols            =   7
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
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
         TabIndex        =   40
         Top             =   3615
         Width           =   1665
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4080
      Index           =   1
      Left            =   255
      TabIndex        =   0
      Top             =   795
      Width           =   8805
      Begin VB.Frame Frame1 
         Caption         =   "Identificação"
         Height          =   1260
         Index           =   0
         Left            =   225
         TabIndex        =   22
         Top             =   90
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
            TabIndex        =   67
            Top             =   870
            Width           =   2070
         End
         Begin VB.ComboBox Serie 
            Height          =   315
            Left            =   870
            TabIndex        =   2
            Top             =   825
            Width           =   765
         End
         Begin VB.ComboBox TipoNFiscal 
            Height          =   315
            ItemData        =   "RastroItensNFFATOcx.ctx":0058
            Left            =   885
            List            =   "RastroItensNFFATOcx.ctx":005A
            TabIndex        =   1
            Top             =   315
            Width           =   3435
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
         Begin MSMask.MaskEdBox Numero 
            Height          =   300
            Left            =   3525
            TabIndex        =   3
            Top             =   825
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
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
            TabIndex        =   25
            Top             =   855
            Width           =   510
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
            TabIndex        =   24
            Top             =   855
            Width           =   720
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
            TabIndex        =   23
            Top             =   345
            Width           =   450
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
            TabIndex        =   21
            Top             =   450
            Visible         =   0   'False
            Width           =   660
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
            TabIndex        =   19
            Top             =   450
            Width           =   1035
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Nota Fiscal"
         Height          =   870
         Left            =   210
         TabIndex        =   11
         Top             =   2670
         Width           =   8325
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
            TabIndex        =   17
            Top             =   375
            Width           =   765
         End
         Begin VB.Label LblDataEmissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1110
            TabIndex        =   16
            Top             =   330
            Width           =   1275
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
            TabIndex        =   15
            Top             =   375
            Width           =   465
         End
         Begin VB.Label LblFilial 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3225
            TabIndex        =   14
            Top             =   330
            Width           =   1560
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
            Left            =   5070
            TabIndex        =   13
            Top             =   375
            Width           =   510
         End
         Begin VB.Label LblValor 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5625
            TabIndex        =   12
            Top             =   330
            Width           =   1275
         End
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Controle de Qualidade"
      Height          =   690
      Left            =   210
      TabIndex        =   63
      Top             =   5025
      Width           =   4815
      Begin VB.CommandButton BotaoImprimirLaudo 
         Caption         =   "Imprimir Laudos"
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
         Left            =   2760
         TabIndex        =   65
         Top             =   195
         Width           =   1815
      End
      Begin VB.CheckBox ImprimirAoGravar 
         Caption         =   "Imprimir laudos ao gravar"
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
         TabIndex        =   64
         Top             =   315
         Width           =   2565
      End
   End
   Begin VB.CommandButton BotaoImprimirRotulos 
      Caption         =   "Imprimir Rótulos de Expedição"
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
      Left            =   5130
      TabIndex        =   62
      Top             =   5220
      Width           =   3630
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4080
      Index           =   2
      Left            =   255
      TabIndex        =   26
      Top             =   795
      Visible         =   0   'False
      Width           =   8805
      Begin VB.Frame Frame3 
         Caption         =   "Itens"
         Height          =   3810
         Left            =   165
         TabIndex        =   29
         Top             =   105
         Width           =   8415
         Begin MSMask.MaskEdBox Almoxarifado 
            Height          =   225
            Left            =   5190
            TabIndex        =   61
            Top             =   1005
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.ComboBox UnidadeMed 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1725
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   390
            Width           =   855
         End
         Begin VB.TextBox DescricaoItem 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   3705
            MaxLength       =   50
            TabIndex        =   30
            Top             =   975
            Width           =   1305
         End
         Begin MSMask.MaskEdBox PercentDesc 
            Height          =   225
            Left            =   3795
            TabIndex        =   32
            Top             =   660
            Width           =   1155
            _ExtentX        =   2037
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
         Begin MSMask.MaskEdBox Desconto 
            Height          =   225
            Left            =   5280
            TabIndex        =   33
            Top             =   645
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
         Begin MSMask.MaskEdBox PrecoUnitario 
            Height          =   225
            Left            =   3660
            TabIndex        =   34
            Top             =   375
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
            Left            =   2580
            TabIndex        =   35
            Top             =   420
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   225
            Left            =   195
            TabIndex        =   36
            Top             =   390
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrecoTotal 
            Height          =   225
            Left            =   5310
            TabIndex        =   37
            Top             =   375
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
            Height          =   3195
            Left            =   240
            TabIndex        =   38
            Top             =   240
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   5636
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   4080
      Index           =   3
      Left            =   255
      TabIndex        =   52
      Top             =   795
      Visible         =   0   'False
      Width           =   8805
      Begin VB.Frame Frame7 
         Caption         =   "Localização dos Produtos"
         Height          =   3840
         Left            =   120
         TabIndex        =   53
         Top             =   90
         Width           =   8580
         Begin MSMask.MaskEdBox UnidadeMedEst 
            Height          =   225
            Left            =   7410
            TabIndex        =   54
            Top             =   690
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
         Begin MSMask.MaskEdBox ProdutoAlmox 
            Height          =   225
            Left            =   1635
            TabIndex        =   55
            Top             =   480
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
         Begin MSMask.MaskEdBox Almox 
            Height          =   225
            Left            =   2985
            TabIndex        =   56
            Top             =   480
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
         Begin MSMask.MaskEdBox QuantAlocada 
            Height          =   225
            Left            =   4320
            TabIndex        =   57
            Top             =   480
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
         Begin MSMask.MaskEdBox ItemNFiscal 
            Height          =   225
            Left            =   1080
            TabIndex        =   58
            Top             =   480
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
         Begin MSMask.MaskEdBox QuantVendida 
            Height          =   225
            Left            =   5925
            TabIndex        =   59
            Top             =   675
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
         Begin MSFlexGridLib.MSFlexGrid GridAlocacao 
            Height          =   2910
            Left            =   390
            TabIndex        =   60
            Top             =   360
            Width           =   7770
            _ExtentX        =   13705
            _ExtentY        =   5133
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
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   135
      Width           =   1620
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1065
         Picture         =   "RastroItensNFFATOcx.ctx":005C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   570
         Picture         =   "RastroItensNFFATOcx.ctx":01DA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "RastroItensNFFATOcx.ctx":070C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4485
      Left            =   210
      TabIndex        =   28
      Top             =   465
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
            Caption         =   "Almoxarifado"
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
Attribute VB_Name = "RastroItensNFFATOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'IDHs de Help
Const IDH_RASTRO_ITENSNF = 0

'Property Variables:
Dim m_Caption As String
Event Unload()

'GridItens
Public objGridItens As AdmGrid
Public iGrid_Produto_Col As Integer
Public iGrid_DescProduto_Col As Integer
Public iGrid_UnidadeMed_Col As Integer
Public iGrid_Almoxarifado_Col As Integer
Public iGrid_Quantidade_Col As Integer
Public iGrid_PrecoUnitario_Col As Integer
Public iGrid_PercDesc_Col As Integer
Public iGrid_Desconto_Col As Integer
Public iGrid_PrecoTotal_Col As Integer


'GridAlocacao
Public iGrid_Item_Col As Integer
Public iGrid_ProdutoAloc_Col As Integer
Public iGrid_QuantAloc_Col As Integer
Public iGrid_QuantVend_Col As Integer
Public iGrid_AlmoxAloc_Col As Integer
Public iGrid_UMAloc_Col As Integer

Public objGridAlocacoes As AdmGrid

Public gobjRastreamento As ClassRastreamento
Public gobjNFiscal As ClassNFiscal


'Variáveis Globais da tela
Public iAlterado As Integer
'Dim gcolItensNF As Collection
Dim iFrameAtual As Integer
Dim iClienteAlterado As Integer
Dim iFornecedorAlterado As Integer

'Eventos dos Browses
Dim WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1
Dim WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Dim WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Dim WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Dim WithEvents objEventoRastroLoteSaldo As AdmEvento
Attribute objEventoRastroLoteSaldo.VB_VarHelpID = -1


Public Function Trata_Parametros(Optional objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma nota foi passada por parametro
    If Not (objNFiscal Is Nothing) Then

        'Tenta ler a nota Fiscal passada por parametro
        lErro = CF("NFiscal_Le_NumFornCli", objNFiscal)
        If lErro <> SUCESSO And lErro <> 35279 Then gError 75680
        
        'Se encontrou a nota fiscal
        If lErro = SUCESSO Then
        
            objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
            'Lê o Tipo de Documento com o Código de objNFiscal
            lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
            If lErro <> SUCESSO And lErro <> 31415 Then gError 75682
    
            'Se não encontrar --> Erro
            If lErro = 31415 Then gError 75684
    
            'Traz a nota para a tela
            lErro = Traz_NF_Tela(objNFiscal)
            If lErro <> SUCESSO Then gError 75681
        
        End If
        
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 75680, 75681, 75682

        Case 75684
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165967)

    End Select

    iAlterado = 0

    Exit Function

End Function

Private Sub objEventoRastroLoteSaldo_evSelecao(obj1 As Object)

Dim objRastroLoteSaldo As ClassRastroLoteSaldo
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoMascarado As String

On Error GoTo Erro_objEventoRastroLoteSaldo_evSelecao

    Set objRastroLoteSaldo = obj1
    
        sProdutoMascarado = String(STRING_PRODUTO, 0)

        lErro = Mascara_MascararProduto(objRastroLoteSaldo.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 80301

        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True

        GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = sProdutoMascarado

    Exit Sub

Erro_objEventoRastroLoteSaldo_evSelecao:

    Select Case gErr

        Case 80301
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165968)

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

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Eventos dos Browses
    Set objEventoSerie = New AdmEvento
    Set objEventoNumero = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoRastroLoteSaldo = New AdmEvento

'    Set gcolItensNF = New Collection
    iFrameAtual = 1
    
    'Carrega os Tipos de Documentos relacionadas à tela
    lErro = Carrega_TiposDocInfo()
    If lErro <> SUCESSO Then gError 62000
    
    'Carrega séries
    lErro = Carrega_Serie()
    If lErro <> SUCESSO Then gError 75689
    
    'Inicializa o Grid de itens
    Set objGridItens = New AdmGrid
    lErro = Inicializa_Grid_Itens(objGridItens)
    If lErro <> SUCESSO Then gError 75679

    Set objGridAlocacoes = New AdmGrid
    'Inicializa o Grid de Alocações
    lErro = Inicializa_Grid_Alocacoes(objGridAlocacoes)
    If lErro <> SUCESSO Then gError 83432

    'Inicializa máscara do produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 62001

    'Rastreamento
    Set gobjRastreamento = New ClassRastreamento
    Set gobjRastreamento.objTela = Me
    gobjRastreamento.bSaidaMaterial = True
    gobjRastreamento.bTelaManutencao = True
    
    'Inicializa o grid de Rastreamento
    lErro = gobjRastreamento.Inicializa_Grid_Rastreamento()
    If lErro <> SUCESSO Then gError 83437

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr

        Case 62000, 62001, 75679, 75689, 83432, 83437

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165969)

    End Select

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
            
        'Se o tipo de nota fiscal for de saída
        If objTipoDocInfo.iTipo = DOCINFO_NF_INT_SAIDA And objTipoDocInfo.iRastreavel = TIPODOCINFO_RASTREAVEL_SIM Then
            
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165970)

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
    If lErro <> SUCESSO Then gError 62002

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

        Case 62002

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165971)

    End Select

    Exit Function

End Function

Private Sub BotaoNFiscal_Click()

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim objcliente As New ClassCliente
Dim objFornecedor As New ClassFornecedor
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_BotaoNFiscal_Click
    
    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 62008
    If Len(Trim(Serie.Text)) = 0 Then gError 62009
    If Len(Trim(Numero.Text)) = 0 Then gError 62010
    
    objTipoDocInfo.iCodigo = TipoNFiscal.ItemData(TipoNFiscal.ListIndex)
    
    'Lê o Tipo da NF
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 62012

    'Se não encontrou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 62014

    'de acordo com a sua Origem verifica se o Cliente ou Fornecedor estão preenchidos
    If objTipoDocInfo.iDestinatario = DOCINFO_CLIENTE Then
                
        If Len(Trim(Cliente.ClipText)) = 0 Then gError 62015
            
        objcliente.sNomeReduzido = Cliente.Text
        lErro = CF("Cliente_Le_NomeReduzido", objcliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 62013
        
        'Não encontrou o cliente
        If lErro = 12348 Then gError 62003
        
        objNFiscal.lCliente = objcliente.lCodigo
        objNFiscal.iFilialCli = Codigo_Extrai(Filial.Text)

    ElseIf objTipoDocInfo.iDestinatario = DOCINFO_FORNECEDOR Then
        
        If Len(Trim(Fornecedor.ClipText)) = 0 Then gError 62016
    
        objFornecedor.sNomeReduzido = Fornecedor.Text
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 75834
        
        'Se não encontrou o fornecedor, erro
        If lErro = 6681 Then gError 62005
        objNFiscal.lFornecedor = objFornecedor.lCodigo
        objNFiscal.iFilialForn = Codigo_Extrai(Filial.Text)
    
    End If
    
    'Filial
    If Len(Trim(Filial.Text)) = 0 Then gError 62017
    
    objNFiscal.iTipoNFiscal = objTipoDocInfo.iCodigo
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    objNFiscal.lNumNotaFiscal = CLng(Numero.Text)
    
    'Lê a nota fiscal
    lErro = CF("NFiscal_Le_NumFornCli", objNFiscal)
    If lErro <> SUCESSO And lErro <> 35279 Then gError 62006
    
    'Se não encontrou a nota fiscal, erro
    If lErro = 35279 Then gError 62007
    
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    objNFiscal.lNumNotaFiscal = CLng(Numero.Text)
    objNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)
    
    'Traz nota fiscal para a tela
    lErro = Traz_NF_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 62011
    
    Exit Sub
    
Erro_BotaoNFiscal_Click:
    
    Select Case gErr
                        
        Case 62003
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objcliente.sNomeReduzido)
            
        Case 62005
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
        
        Case 62006, 62011, 62012, 62013
        
        Case 62007
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", gErr, objNFiscal.lNumNotaFiscal)
        
        Case 62008
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)
            
        Case 62009
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
        
        Case 62010
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)
                
        Case 62014
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)
        
        Case 62015
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
                
        Case 62016
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
        
        Case 62017
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165972)
        
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
    If lErro <> SUCESSO And lErro <> 31415 Then gError 62018

    'se não estiver cadastrado ==> erro
    If lErro = 31415 Then gError 62019
    
    'Habilita cliente ou fornecedor
    If objTipoDocInfo.iDestinatario = DOCINFO_CLIENTE Then
                
        Call Habilita_Cliente
    
    ElseIf objTipoDocInfo.iDestinatario = DOCINFO_FORNECEDOR Then
                    
        Call Habilita_Fornecedor
    
    End If
    
    'Alterado por Cyntia em 17/05/2002
    If objTipoDocInfo.iEscaninhoRastro = ESCANINHO_HABILITADO Then
        EscaninhoRastro.Enabled = True
    Else
        EscaninhoRastro.Enabled = False
    End If
    
    'Inicializa o grid de Rastreamento
    gobjRastreamento.iCodigo = objTipoDocInfo.iCodigo
    
    lErro = gobjRastreamento.Inicializa_Grid_Rastreamento()
    If lErro <> SUCESSO Then gError 83456
    
    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub

Erro_TipoNFiscal_Click:

    Select Case gErr

        Case 62018, 83456

        Case 62019
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165973)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 62020
    
    If lErro <> SUCESSO Then gError 62021 'Não conseguiu

    Exit Sub

Erro_TipoNFiscal_Validate:

    Cancel = True

    Select Case gErr

        Case 62020

        Case 62021
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_RASTRO", gErr, TipoNFiscal.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165974)

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
            If lErro <> SUCESSO Then gError 62022

            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then gError 62023

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

        Case 62022, 62023

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165975)

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
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    If iClienteAlterado = 1 Then

        If Len(Trim(Cliente.Text)) > 0 Then

            lErro = TP_Cliente_Le3(Cliente, objcliente, iCodFilial)
            If lErro <> SUCESSO Then gError 62024

            lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
            If lErro <> SUCESSO Then gError 62025

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

        Case 62024, 62025

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165976)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 62028

    'Se não encontrou o ítem com o código informado
    If lErro = 6730 Then

        If Fornecedor.Visible = True Then

            'Verifica se o Fornecedor foi preenchido
            If Len(Trim(Fornecedor.Text)) = 0 Then gError 62030

            sNomeRed = Fornecedor.Text

            objFilialFornecedor.iCodFilial = iCodigo

            'Pesquisa se existe a Filial do Fornecedor
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sNomeRed, objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then gError 62026

            'Se não encontrou a Filial do Fornecedor --> erro
            If lErro = 18272 Then gError 62029

            'Coloca a Filial do Fornecedor na tela
            Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

        Else

            'Verifica se Cliente está preenchido
            If Len(Trim(Cliente.ClipText)) = 0 Then gError 62031

            sNomeRed = Cliente.Text

            'Lê a Filial do Cliente
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sNomeRed, objFilialCliente)
            If lErro <> SUCESSO And lErro <> 17660 Then gError 62027

            'Se não encontrou a Filial do Cliente --> erro
            If lErro = 17660 Then gError 62032

            'Coloca a Filial do Fornecedor
            Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome
        
        End If

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 62033

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 62026, 62027, 62028

        Case 62029
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 62030
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 62031
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 62032
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 62033
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165977)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filial_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera as variáveis globais da tela
    Set objEventoSerie = Nothing
    Set objEventoNumero = Nothing
    Set objEventoCliente = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoRastroLoteSaldo = Nothing

    Set objGridAlocacoes = Nothing

    Set gobjRastreamento = Nothing

    Set gobjNFiscal = Nothing
    Set objGridItens = Nothing
'    Set gcolItensNF = Nothing

    'Fecha o Comando de Setas
    lErro = ComandoSeta_Liberar(Me.Name)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165978)
    
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

Public Sub AlmoxRastro_Change()
'Rastreamento

    Call gobjRastreamento.AlmoxRastro_Change

End Sub

Public Sub AlmoxRastro_GotFocus()
'Rastreamento

    Call gobjRastreamento.AlmoxRastro_GotFocus

End Sub

Public Sub AlmoxRastro_KeyPress(KeyAscii As Integer)
'Rastreamento

    Call gobjRastreamento.AlmoxRastro_KeyPress(KeyAscii)

End Sub

Public Sub AlmoxRastro_Validate(Cancel As Boolean)
'Rastreamento

    Call gobjRastreamento.AlmoxRastro_Validate(Cancel)

End Sub

Public Sub EscaninhoRastro_Change()
'Rastreamento

    Call gobjRastreamento.EscaninhoRastro_Change

End Sub

Public Sub EscaninhoRastro_Click()
'Rastreamento

    Call gobjRastreamento.EscaninhoRastro_Click

End Sub


Public Sub EscaninhoRastro_GotFocus()
'Rastreamento

    Call gobjRastreamento.EscaninhoRastro_GotFocus

End Sub

Public Sub EscaninhoRastro_KeyPress(KeyAscii As Integer)
'Rastreamento

    Call gobjRastreamento.EscaninhoRastro_KeyPress(KeyAscii)

End Sub

Public Sub EscaninhoRastro_Validate(Cancel As Boolean)
'Rastreamento

    Call gobjRastreamento.EscaninhoRastro_Validate(Cancel)

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

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

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
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Total")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoItem.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (PrecoUnitario.Name)
    objGridInt.colCampo.Add (PercentDesc.Name)
    objGridInt.colCampo.Add (Desconto.Name)
    objGridInt.colCampo.Add (PrecoTotal.Name)

    iGrid_Produto_Col = 1
    iGrid_DescProduto_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_Quantidade_Col = 4
    iGrid_Almoxarifado_Col = 5
    iGrid_PrecoUnitario_Col = 6
    iGrid_PercDesc_Col = 7
    iGrid_Desconto_Col = 8
    iGrid_PrecoTotal_Col = 9
    
    Almoxarifado.Enabled = False
     
    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 8

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_Alocacoes(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Alocações

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quant. Alocada")
    objGridInt.colColuna.Add ("Quant. Vendida")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ItemNFiscal.Name)
    objGridInt.colCampo.Add (ProdutoAlmox.Name)
    objGridInt.colCampo.Add (Almox.Name)
    objGridInt.colCampo.Add (UnidadeMedEst.Name)
    objGridInt.colCampo.Add (QuantAlocada.Name)
    objGridInt.colCampo.Add (QuantVendida.Name)

    'Colunas da Grid
    iGrid_Item_Col = 1
    iGrid_ProdutoAloc_Col = 2
    iGrid_AlmoxAloc_Col = 3
    iGrid_UMAloc_Col = 4
    iGrid_QuantAloc_Col = 5
    iGrid_QuantVend_Col = 6

    'Grid do GridInterno
    objGridInt.objGrid = GridAlocacao

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ALOCACOES + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 12

    'Largura da primeira coluna
    GridAlocacao.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Alocacoes = SUCESSO

    Exit Function

End Function

Private Sub BotaoFechar_Click()

     Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 75690

    'Limpa a Tela
    Call BotaoLimpar_Click

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 75690

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165979)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro()

Dim lErro As Long
Dim objMovEstoque As New ClassMovEstoque
Dim objFornecedor As New ClassFornecedor
Dim objcliente As New ClassCliente
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objNFiscal As New ClassNFiscal
Dim lTransacao As Long

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se os campos obrigatórios foram preenchidos
    If Len(Trim(Serie.Text)) = 0 Then gError 75691
    If Len(Trim(Numero.ClipText)) = 0 Then gError 75692
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 62034

    objTipoDocInfo.iCodigo = TipoNFiscal.ItemData(TipoNFiscal.ListIndex)
    
    'Lê o Tipo da NF
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 62035

    'Se não encontrou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 62036

    'de acordo com a sua Origem verifica se o Cliente ou Fornecedor estão preenchidos
    If objTipoDocInfo.iDestinatario = DOCINFO_CLIENTE Then
        
        If Len(Trim(Cliente.ClipText)) = 0 Then gError 62037
    
        objcliente.sNomeReduzido = Cliente.Text
        lErro = CF("Cliente_Le_NomeReduzido", objcliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 62038
        
        'Não encontrou o cliente
        If lErro = 12348 Then gError 62039
        
        objNFiscal.lCliente = objcliente.lCodigo
        objNFiscal.iFilialCli = Codigo_Extrai(Filial.Text)
    
    ElseIf objTipoDocInfo.iDestinatario = DOCINFO_FORNECEDOR Then
        
        If Len(Trim(Fornecedor.ClipText)) = 0 Then gError 62040
        
        objFornecedor.sNomeReduzido = Fornecedor.Text
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 62041
        
        'Se não encontrou o fornecedor, erro
        If lErro = 6681 Then gError 62042
        
        objNFiscal.lFornecedor = objFornecedor.lCodigo
        objNFiscal.iFilialForn = Codigo_Extrai(Filial.Text)
    
    End If
    
    If Len(Trim(Filial.Text)) = 0 Then gError 62043
            
    'Lê a nota fiscal de entrada
    objNFiscal.iTipoNFiscal = objTipoDocInfo.iCodigo
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    objNFiscal.lNumNotaFiscal = CLng(Numero.Text)
    lErro = CF("NFiscal_Le_NumFornCli", objNFiscal)
    If lErro <> SUCESSO And lErro <> 35279 Then gError 75664

    'Nota fiscal não cadastrada
    If lErro <> SUCESSO Then gError 75665

    'Verifica se a nota já está cancelada
    If objNFiscal.iStatus = STATUS_CANCELADO Then gError 75666
    If objNFiscal.iFilialEmpresa <> giFilialEmpresa Then gError 75667

    'Lê os itens da nota fiscal
    lErro = CF("NFiscalItens_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 75668
            
    'Lê as Alocações dos itens da Nota Fiscal
    lErro = CF("AlocacoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 83488
            
    lErro = gobjRastreamento.Valida_Rastreamento(objTipoDocInfo)
    If lErro <> SUCESSO Then gError 83442

    Set objNFiscal.objRastreamento = gobjRastreamento
            
    'Move dados da tela para a memória
    lErro = Move_Tela_Memoria(objNFiscal, objMovEstoque)
    If lErro <> SUCESSO Then gError 75694

    'mover a parte do rastreamento
    lErro = gobjRastreamento.Move_Rastro_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 83443

    'altera os dados de rastreamento
    lErro = NFiscal_Grava_Rastro(objNFiscal)
    If lErro <> SUCESSO Then gError 83444

    'Se a opcao de imprimir o Relatorio estiver marcada
    If ImprimirAoGravar.Value = MARCADO Then
        
        'Gera o(s) Relatorio(s)
        Call Executa_Relatorio(objNFiscal)
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
        
        Case 62034
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)
        
        Case 62035, 62038, 62041, 75664, 75668, 75683, 75694, 83442, 83443, 83444, 83488
        
        Case 62036
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)
        
        Case 62037
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 62039
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objcliente.sNomeReduzido)
        
        Case 62040
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
        
        Case 62042
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case 62043
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
                
        Case 75665
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", gErr, Numero.Text)
            
        Case 75666
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_CANCELADA", gErr, Serie.Text, Numero.Text)
        
        Case 75667
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_OUTRA_FILIAL", gErr)
        
        Case 75691
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 75692
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165980)

    End Select

    Call Transacao_Rollback
    
    Exit Function

End Function

Function Move_Tela_Memoria(objNFiscal As ClassNFiscal, objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long, iIndice As Integer
Dim sDocOrigem As String
Dim objItemNF As ClassItemNF, objItemNFAloc As ClassItemNFAlocacao
Dim objItemMovEstoque As New ClassItemMovEstoque
Dim objRastroMovto As ClassRastreamentoMovto
Dim objRastroItemNF As ClassRastroItemNF
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iTipoMovtoEstoque As Integer

On Error GoTo Erro_Move_Tela_Memoria

    'Lê TipoDocInfo da nota fiscal
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 75698
    
    'Se não encontrou Tipo, erro
    If lErro = 31415 Then gError 75699

    iTipoMovtoEstoque = objTipoDocInfo.iTipoMovtoEstoque
    sDocOrigem = objTipoDocInfo.sSigla

    If iTipoMovtoEstoque > 0 Then

        Set objMovEstoque = New ClassMovEstoque

        'Guarda dados do Movimento de Estoque
        objMovEstoque.iFilialEmpresa = giFilialEmpresa
        objMovEstoque.iTipoMov = iTipoMovtoEstoque
        objMovEstoque.lCliente = objNFiscal.lCliente
        objMovEstoque.lFornecedor = objNFiscal.lFornecedor
        objMovEstoque.sDocOrigem = sDocOrigem & " " & objNFiscal.sSerie & " " & CStr(objNFiscal.lNumNotaFiscal)

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
'                If lErro <> SUCESSO And lErro <> 75796 Then gError 75800
'
'                'Se não encontrou, erro
'                If lErro = 75796 Then gError 75801
'
'                'Guarda código e data do Movimento de estoque
'                objMovEstoque.lCodigo = objItemMovEstoque.lCodigo
'                objMovEstoque.dtData = objItemMovEstoque.dtData
'                objItemMovEstoque.sProduto = objItemNF.sProduto
'                objItemMovEstoque.sSiglaUM = objItemNF.sUnidadeMed
'
'                'Para cada alocação do ItemNF
'                For Each objItemNFAloc In objItemNF.colAlocacoes
'
'                    objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
'                    objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
'                    objItemMovEstoque.iControleEstoque = objItemNF.iControleEstoque
'                    objItemMovEstoque.sSiglaUMEst = objItemNF.sUMEstoque
'                    objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
'
'                    Set objItemMovEstoque.colRastreamentoMovto = New Collection
'
'                    'Guarda o Rastreamento vinculados a essa alocação
'
'                    For Each objRastroItemNF In objItemNF.colRastreamento
'
'                        'Se o Rastreamento faz parte da alocação
'                        If objRastroItemNF.iAlmoxCodigo = objItemNFAloc.iAlmoxarifado And objRastroItemNF.dLoteQdtAlocada > 0 Then
'
'                            Set objRastroMovto = New ClassRastreamentoMovto
'
'                            objRastroMovto.dQuantidade = objRastroItemNF.dLoteQdtAlocada
'                            objRastroMovto.iTipoDocOrigem = TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE
'                            objRastroMovto.sLote = objRastroItemNF.sLote
'                            objRastroMovto.lNumIntDocOrigem = objItemMovEstoque.lNumIntDoc
'                            objRastroMovto.sProduto = objItemNF.sProduto
'                            objRastroMovto.iFilialOP = objRastroItemNF.iLoteFilialOP
'
'                            'Adiciona objRastroMovto na coleção de Rastreamento
'                            objItemMovEstoque.colRastreamentoMovto.Add objRastroMovto
'                        End If
'                    Next
'
'                    'Adiciona na coleção de Movimentos de Estoque
'                    Call objMovEstoque.colItens.Add(objItemMovEstoque.lNumIntDoc, objItemMovEstoque.iTipoMov, objItemNF.dValorTotal, APROPR_CUSTO_INFORMADO, objItemNF.sProduto, objItemNF.sDescricaoItem, objItemNF.sUnidadeMed, objItemNFAloc.dQuantidade, objItemNFAloc.iAlmoxarifado, objItemNFAloc.sAlmoxarifado, objItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, objItemMovEstoque.colRastreamentoMovto, objItemMovEstoque.colApropriacaoInsumo, DATA_NULA)
'
'                Next
'
'            End If
'
'        Next
        
    End If
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
        
        Case 75698, 75800
                        
        Case 75699
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)
        
        Case 75801
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_CADASTRADO", gErr, objItemNF.lNumIntDoc)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165981)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()
    
    Call Limpa_Tela_NF

End Sub

'Private Sub BotaoRastreamento_Click()
'
'Dim objRastroLoteSaldo As New ClassRastroLoteSaldo
'
'    'Chama a tela de Rastreamento de produtos
'    Call Chama_Tela("RastroLoteSaldoLista", gcolItensNF, objRastroLoteSaldo, objEventoRastroLoteSaldo)
'
'End Sub

Private Sub Numero_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Numero_GotFocus()

    Call MaskEdBox_TrataGotFocus(Numero, iAlterado)

End Sub

Private Sub Serie_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Serie_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim objSerie As New ClassSerie

On Error GoTo Erro_Serie_Validate

    'Verifica se a série está preenchida
    If Len(Trim(Serie.Text)) > 0 Then

        objSerie.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)

        lErro = CF("Serie_Le", objSerie)
        If lErro <> SUCESSO And lErro <> 22202 Then gError 75662

        'Série não cadastrada
        If lErro = 22202 Then gError 75663
        
    End If
        
    Exit Sub

Erro_Serie_Validate:

    Cancel = True

    Select Case gErr

        Case 75662

        Case 75663
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objSerie.sSerie)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165982)

    End Select

    Exit Sub

End Sub

Function Traz_NF_Tela(objNFiscal As ClassNFiscal) As Long
'Traz os dados da Nota Fiscal passada em objNFiscal

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objItemNF As ClassItemNF
Dim objAlocacao As ClassItemNFAlocacao
Dim objRastroItemNF As ClassRastroItemNF
Dim objFilialEmpresa As New AdmFiliais
Dim bCancel As Boolean

On Error GoTo Erro_Traz_NF_Tela

    Set gobjNFiscal = objNFiscal

    'Lê o Tipo de Documento
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 75669
    If lErro = 31415 Then gError 75671 'Tipo não cadastrado

    'Se o tipo de nota fiscal não for de saída
    If objTipoDocInfo.iTipo <> DOCINFO_NF_INT_SAIDA Then gError 83484
    
    'Se o tipo da nota não é rastreavel
    If objTipoDocInfo.iRastreavel = TIPODOCINFO_RASTREAVEL_NAO Then gError 83485

    'Limpa a tela NFicalSaida
    Call Limpa_Tela_NF

    'Lê os Itens da Nota Fiscal
    lErro = CF("NFiscalItens_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 75672

    'Preenche o número da NF
    If objNFiscal.lNumNotaFiscal > 0 Then
        Numero.PromptInclude = False
        Numero.Text = CStr(objNFiscal.lNumNotaFiscal)
        Numero.PromptInclude = True
    End If

    'preenche a serie da NF
'    Serie.Text = objNFiscal.sSerie
    Serie.Text = Desconverte_Serie_Eletronica(objNFiscal.sSerie)
    If ISSerieEletronica(objNFiscal.sSerie) Then
        EletronicaFed.Value = vbChecked
    Else
        EletronicaFed.Value = vbUnchecked
    End If
    Call Serie_Validate(bSGECancelDummy)

    TipoNFiscal.Text = objTipoDocInfo.iCodigo
    Call TipoNFiscal_Validate(bCancel)

    'De acordo com a Origem do tipo Coloca o Cliente ou o fornecedor na tela
    If objTipoDocInfo.iDestinatario = DOCINFO_CLIENTE Then
        Call Habilita_Cliente
        Cliente.Text = objNFiscal.lCliente
        Call Cliente_Validate(bCancel)
        Filial.Text = objNFiscal.iFilialCli
    ElseIf objTipoDocInfo.iDestinatario = DOCINFO_FORNECEDOR Then
        Call Habilita_Fornecedor
        Fornecedor.Text = objNFiscal.lFornecedor
        Call Fornecedor_Validate(bCancel)
        Filial.Text = objNFiscal.iFilialForn
    End If

    Call Filial_Validate(bCancel)
    
    'Lê FilialEmpresa
    objFilialEmpresa.iCodFilial = objNFiscal.iFilialEmpresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then gError 75916

    'Se não encontrou a FilialEmpresa
    If lErro = 27378 Then gError 75917
    
    LblFilial.Caption = CStr(objFilialEmpresa.iCodFilial) & SEPARADOR & objFilialEmpresa.sNome
    
    'Se a data não for nula coloca na Tela
    If objNFiscal.dtDataEmissao <> DATA_NULA Then
        LblDataEmissao.Caption = Format(objNFiscal.dtDataEmissao, "dd/mm/yyyy")
    Else
        LblDataEmissao.Caption = Format("", "dd/mm/yy")
    End If

    'Preenche o valor total da NF
    If objNFiscal.dValorTotal > 0 Then
        LblValor.Caption = Format(objNFiscal.dValorTotal, "Fixed")
    Else
        LblValor.Caption = Format(0, "Fixed")
    End If

    'Lê as Alocações dos itens da Nota Fiscal
    lErro = CF("AlocacoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 75651
    
    'Preenche Grid de Itens da Nota Fiscal
    lErro = Preenche_GridItens(objNFiscal)
    If lErro <> SUCESSO Then gError 75670
        
    'Preenche o Grid com as Alocações dos itens da Nota Fiscal
    lErro = Preenche_GridAlocacoes(objNFiscal)
    If lErro <> SUCESSO Then gError 83433

    If objTipoDocInfo.iTipoMovtoEstoque <> 0 Then
        'Carrega ItensNF com Rastreamentos
        lErro = gobjRastreamento.Carrega_RastroItensNF(objNFiscal)
        If lErro <> SUCESSO Then gError 83438
    End If

        
'    'Para cada item da nota fiscal
'    For Each objItemNF In objNFiscal.ColItensNF
'
'        'Guarda nos rastreamentos as quantidades alocadas nos almoxarifados
'        For Each objAlocacao In objItemNF.colAlocacoes
'            For Each objRastroItemNF In objItemNF.colRastreamento
'                If objAlocacao.iAlmoxarifado = objRastroItemNF.iAlmoxCodigo Then
'                    objRastroItemNF.dAlmoxQtdAlocada = objRastroItemNF.dAlmoxQtdAlocada + objAlocacao.dQuantidade
'                End If
'            Next
'        Next
'
'    Next
        
'    Set gcolItensNF = New Collection
'
'    'Para cada item da nota fiscal
'    For Each objItemNF In objNFiscal.ColItensNF
'
'        'Guarda oitensNF na coleção global de itens
'        gcolItensNF.Add objItemNF
'
'    Next
        
    Traz_NF_Tela = SUCESSO

    Exit Function

Erro_Traz_NF_Tela:

    Traz_NF_Tela = gErr

    Select Case gErr

        Case 75651, 75669, 75670, 75672, 75916, 83433, 83438

        Case 75671
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iTipo)
        
        Case 75917
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)
        
        Case 83484
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_SAIDA", gErr, objTipoDocInfo.iCodigo)
        
        Case 83485
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_RASTRO", gErr, CStr(objTipoDocInfo.iCodigo))
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165983)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim objUM As ClassUnidadeDeMedida
Dim sUM As String

On Error GoTo Erro_Rotina_Grid_Enable

    'Rastreamento
    lErro = gobjRastreamento.Rotina_Grid_Enable(iLinha, objControl, iCaminho)
    If lErro <> SUCESSO Then gError 83439

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 83439

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165984)

    End Select

    Exit Sub

End Sub

Function Preenche_GridItens(objNFiscal As ClassNFiscal) As Long
'Preenche o Grid com os itens da Nota Fiscal

Dim lErro As Long
Dim iIndice As Integer
Dim objItemNF As ClassItemNF
Dim sProdutoEnxuto As String
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Preenche_GridItens

    iIndice = 0

    'Para cada item da Coleção
    For Each objItemNF In objNFiscal.ColItensNF

        iIndice = iIndice + 1

        'Formata o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objItemNF.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 75673

        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True

        'Preenche o Grid
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col) = objItemNF.sDescricaoItem
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemNF.sUnidadeMed
        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemNF.dQuantidade)
        If EscaninhoRastro.Enabled = False Then
            GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objItemNF.sAlmoxarifadoNomeRed
        End If
        GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col) = Format(objItemNF.dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
        GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(objItemNF.dPercDesc, "Percent")
        GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(objItemNF.dValorDesconto, "Standard")
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col) = Format(objItemNF.dValorTotal, "Standard")

        'Se o almoxarifado estiver preenchido
        If objItemNF.iAlmoxarifado > 0 And EscaninhoRastro.Enabled = False Then
            
            'Lê o almoxarifado
            objAlmoxarifado.iCodigo = objItemNF.iAlmoxarifado
            
            lErro = CF("Almoxarifado_Le", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 83445
            
            If lErro = 25056 Then gError 83446
                
            'COloca o almoxarifado na tela
            GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
        
        End If

    Next

    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = iIndice

    Preenche_GridItens = SUCESSO

    Exit Function

Erro_Preenche_GridItens:

    Preenche_GridItens = gErr

    Select Case gErr

        Case 75673
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItemNF.sProduto)

        Case 83445
        
        Case 83446
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objItemNF.iAlmoxarifado)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165985)

    End Select

    Exit Function

End Function

Private Function Preenche_GridAlocacoes(objNFiscal As ClassNFiscal) As Long
'Preenche o Grid com as Alocações da Nota Fiscal

Dim objItemAloc As ClassItemNFAlocacao
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim lErro As Long
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim dQuantidade As Double
Dim iNumCasasDec As Integer
Dim dAcrescimo As Double
Dim iContador As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Preenche_GridAlocacoes
    
    'Limpa o grid de alocações
    Call Grid_Limpa(objGridAlocacoes)
    
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 109030
    
    If lErro = 31415 Then gError 109031

    'If EscaninhoRastro.Enabled = True Then
    If objTipoDocInfo.bExibeGridAlocacao = True Then

        iIndice1 = 0
        'Para cada item da NF
        For iIndice = 1 To objNFiscal.ColItensNF.Count
                    
            If objNFiscal.ColItensNF.Item(iIndice).colItensRomaneioGrade.Count = 0 Then
                    
                Call AlocacoesNF_Agrupa(objNFiscal.ColItensNF.Item(iIndice).colAlocacoes)
                        
                'Para cada alocação do Item de NF
                For Each objItemAloc In objNFiscal.ColItensNF.Item(iIndice).colAlocacoes
        
                    iIndice1 = iIndice1 + 1
                
                    objProduto.sCodigo = objNFiscal.ColItensNF(iIndice).sProduto
                    'Lê o Produto
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 83434
                    If lErro <> SUCESSO Then gError 83435
                
                    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objNFiscal.ColItensNF(iIndice).sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
                    If lErro <> SUCESSO Then gError 83436
                    
                    'Coloca os dados da alocação na tela
                    GridAlocacao.TextMatrix(iIndice1, iGrid_Item_Col) = objNFiscal.ColItensNF(iIndice).iItem
                    GridAlocacao.TextMatrix(iIndice1, iGrid_ProdutoAloc_Col) = GridItens.TextMatrix(iIndice, iGrid_Produto_Col)
                    GridAlocacao.TextMatrix(iIndice1, iGrid_AlmoxAloc_Col) = objItemAloc.sAlmoxarifado
                    GridAlocacao.TextMatrix(iIndice1, iGrid_UMAloc_Col) = objProduto.sSiglaUMEstoque
                    GridAlocacao.TextMatrix(iIndice1, iGrid_QuantAloc_Col) = Formata_Estoque(objItemAloc.dQuantidade)
                    dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)) * dFator
    
                    dQuantidade = Arredonda_Estoque(dQuantidade)
    
                    GridAlocacao.TextMatrix(iIndice1, iGrid_QuantVend_Col) = Formata_Estoque(dQuantidade)
                
                    objGridAlocacoes.iLinhasExistentes = objGridAlocacoes.iLinhasExistentes + 1

                Next
                
            Else
            
                Call Atualiza_Grid_Alocacao(objNFiscal.ColItensNF.Item(iIndice))
                
                iIndice1 = objGridAlocacoes.iLinhasExistentes
   
            End If
   
        Next
    
    End If

    Preenche_GridAlocacoes = SUCESSO

    Exit Function

Erro_Preenche_GridAlocacoes:

    Preenche_GridAlocacoes = gErr
    
    Select Case gErr
    
        Case 83434, 83436, 109030
        
        Case 83435
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, Produto.Text)
            
        Case 109031
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165986)
            
    End Select

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
        
        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Rastreamento
            Case GridRastro.Name

                lErro = gobjRastreamento.Saida_Celula()
                If lErro <> SUCESSO Then gError 83440

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 83441

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 83440

        Case 83441
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165987)

    End Select

    Exit Function

End Function

Private Sub AlocacoesNF_Agrupa(colAlocacoes As ColAlocacoesItemNF)

Dim iIndice As Integer
Dim iIndice1 As Integer

    If colAlocacoes.Count <= 1 Then Exit Sub

    'se a alocação se referir a uma transferencia ==> não leva em consideração
    For iIndice = colAlocacoes.Count To 1 Step -1
        If colAlocacoes.Item(iIndice).iTransferencia = TIPOMOV_EST_TRANSFERENCIA Then
            colAlocacoes.Remove iIndice
        End If
    Next

    For iIndice = colAlocacoes.Count To 2 Step -1
        
        For iIndice1 = 1 To iIndice - 1
            If (colAlocacoes.Item(iIndice).iAlmoxarifado = colAlocacoes.Item(iIndice1).iAlmoxarifado) Then
                colAlocacoes.Item(iIndice1).dQuantidade = colAlocacoes.Item(iIndice1).dQuantidade + colAlocacoes.Item(iIndice).dQuantidade
                colAlocacoes.Remove iIndice
                Exit For
            End If
        Next
    
    Next

End Sub

Sub Limpa_Tela_NF()
    
    Call Limpa_Tela(Me)

    'Limpa Grid
    Call Grid_Limpa(objGridItens)

    'Limpa restante dos campos
    LblFilial.Caption = ""
    LblDataEmissao.Caption = ""
    LblValor.Caption = ""
    Serie.Text = ""
    EletronicaFed.Value = vbUnchecked
    
    TipoNFiscal.ListIndex = -1
    
'    Set gcolItensNF = New Collection
    Filial.Clear
    
    'Limpa o Frame de Rastreamento
    Call gobjRastreamento.Limpa_Tela_Rastreamento
    
    iAlterado = 0

End Sub

Private Sub LblNumero_Click()

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim colSelecao As Collection

On Error GoTo Erro_LblNumero_Click

    'verifica se a Serie e o Número da NF de saída estão preenchidos
    If Len(Trim(Numero.ClipText)) > 0 Then objNFiscal.lNumNotaFiscal = CLng(Numero.Text)
    If Len(Trim(Serie.Text)) > 0 Then objNFiscal.sSerie = Serie.Text

    Call Chama_Tela("NFiscalSaidaTodasLista", colSelecao, objNFiscal, objEventoNumero)

    Exit Sub

Erro_LblNumero_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165988)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objNFiscal = obj1

    'Lê a nota fiscal selecionada
    lErro = CF("NFiscal_Le_NumFornCli", objNFiscal)
    If lErro <> SUCESSO And lErro <> 35279 Then gError 75675
    If lErro = 35279 Then gError 75676

    'Verifica se  a nota esta cancelada ou pertence a outra filial empresa
    If objNFiscal.iStatus = STATUS_CANCELADO Then gError 75677
    If objNFiscal.iFilialEmpresa <> giFilialEmpresa Then gError 75678

    'Traz a NotaFiscal para a a tela
    lErro = Traz_NF_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 75674

    Me.Show

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case gErr

        Case 75674, 75675

        Case 75676
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", gErr, objNFiscal.lNumNotaFiscal)
            Call Limpa_Tela_NF
            iAlterado = 0

        Case 75677
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_CANCELADA", gErr, Serie.Text, Numero.Text)

        Case 75678
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_OUTRA_FILIAL", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165989)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165990)

    End Select

    Exit Sub

End Sub

Private Sub objEventoSerie_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objSerie As ClassSerie
Dim iIndice As Integer
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165991)

    End Select

    Exit Sub

End Sub

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

Dim objcliente As New ClassCliente
Dim colSelecao As Collection

    'Preenche NomeReduzido com o cliente da tela
    objcliente.sNomeReduzido = Cliente.Text

    'Chama Tela ClienteLista
    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente)

End Sub

Public Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim bCancel As Boolean

    Set objcliente = obj1

    'Preenche campo Cliente
    Cliente.Text = objcliente.sNomeReduzido

    'Executa o Validate
    Call Cliente_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

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

    Parent.HelpContextID = IDH_RASTRO_ITENSNF
    Set Form_Load_Ocx = Me
    Caption = "Rastreamento de Itens da Nota Fiscal"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RastroItensNFFAT"

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

Private Sub LabelFilial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilial, Source, X, Y)
End Sub

Private Sub LabelFilial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilial, Button, Shift, X, Y)
End Sub

Private Sub LblValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblValor, Source, X, Y)
End Sub

Private Sub LblValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblValor, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub LblFilial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblFilial, Source, X, Y)
End Sub

Private Sub LblFilial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblFilial, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub LblDataEmissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblDataEmissao, Source, X, Y)
End Sub

Private Sub LblDataEmissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblDataEmissao, Button, Shift, X, Y)
End Sub

Private Sub ClienteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ClienteLabel, Source, X, Y)
End Sub

Private Sub ClienteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ClienteLabel, Button, Shift, X, Y)
End Sub

Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165995)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Private Function Executa_Relatorio(ByVal objNF As ClassNFiscal) As Long

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_Executa_Relatorio

    lErro = objRelatorio.ExecutarDireto("Laudos do Controle de Qualidade de NFs", "", 0, "", "NNFISCALINIC", CStr(objNF.lNumNotaFiscal), "NNFISCALFIM", CStr(objNF.lNumNotaFiscal), "TSERIE", objNF.sSerie)
    If lErro <> SUCESSO Then gError 130409

    Executa_Relatorio = SUCESSO
     
    Exit Function
    
Erro_Executa_Relatorio:

    Executa_Relatorio = gErr
     
    Select Case gErr
          
        Case 130409
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165996)
     
    End Select
     
    Exit Function

End Function

Private Sub BotaoImprimirLaudo_Click()
    
Dim objNF As New ClassNFiscal

    objNF.lNumNotaFiscal = StrParaLong(Numero.Text)
    objNF.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)

    If objNF.lNumNotaFiscal <> 0 And Len(Trim(objNF.sSerie)) <> 0 Then
    
        Call Executa_Relatorio(objNF)
        
    End If

End Sub

Private Sub BotaoImprimirRotulos_Click()

Dim objRelatorio As New AdmRelatorio
Dim objNF As New ClassNFiscal

    objNF.lNumNotaFiscal = StrParaLong(Numero.Text)
    objNF.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)

    Call objRelatorio.Rel_Menu_Executar("Rótulos de Expedição para Notas Fiscais", objNF)
    
End Sub

Sub Atualiza_Grid_Alocacao(objItemNF As ClassItemNF)

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objReserva As ClassReservaItem
Dim sProdutoMascarado As String
Dim lErro As Long
Dim dFator As Double
Dim dFator2 As Double
Dim objProduto As New ClassProduto
Dim dQuantReservada As Double
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Atualiza_Grid_Alocacao

    For Each objItemRomaneio In objItemNF.colItensRomaneioGrade

'        objProduto.sCodigo = objItemNF.sProduto
'        'Lê o produto
'        lErro = CF("Produto_Le", objProduto)
'        If lErro <> SUCESSO And lErro <> 28030 Then gError 42764
'        If lErro = 28030 Then gError 42765 'Não encontrou
'        'Faz a conversão da unidade do item para a unidade de estoque
'        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemRomaneio.sUMEstoque, objProduto.sSiglaUMEstoque, dFator)
'        If lErro <> SUCESSO Then gError 42766
'
'        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemNF.sUnidadeMed, objProduto.sSiglaUMEstoque, dFator2)
'        If lErro <> SUCESSO Then gError 42766
        
        objProduto.sCodigo = objItemRomaneio.sProduto
        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 42764
        If lErro = 28030 Then gError 42765 'Não encontrou

        'Faz a conversão da unidade do item para a unidade de estoque
        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemRomaneio.sUMEstoque, objProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 42766

        objProduto.sCodigo = objItemNF.sProduto
        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 42764
        If lErro = 28030 Then gError 42765 'Não encontrou
               
        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemNF.sUnidadeMed, objProduto.sSiglaUMEstoque, dFator2)
        If lErro <> SUCESSO Then gError 42766
        
        dQuantReservada = 0
        
        For Each objReserva In objItemRomaneio.colLocalizacao
        
            GridAlocacao.TextMatrix(objGridAlocacoes.iLinhasExistentes + 1, iGrid_Item_Col) = objItemNF.iItem
            
            lErro = Mascara_MascararProduto(objItemRomaneio.sProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 86357
                        
            objAlmoxarifado.iCodigo = objReserva.iAlmoxarifado

            lErro = CF("Almoxarifado_Le", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 46181
            If lErro = 25056 Then gError 46182
                
            objReserva.sAlmoxarifado = objAlmoxarifado.sNomeReduzido
                        
            GridAlocacao.TextMatrix(objGridAlocacoes.iLinhasExistentes + 1, iGrid_ProdutoAloc_Col) = sProdutoMascarado
            GridAlocacao.TextMatrix(objGridAlocacoes.iLinhasExistentes + 1, iGrid_UMAloc_Col) = objItemRomaneio.sUMEstoque
            GridAlocacao.TextMatrix(objGridAlocacoes.iLinhasExistentes + 1, iGrid_AlmoxAloc_Col) = objReserva.sAlmoxarifado
            GridAlocacao.TextMatrix(objGridAlocacoes.iLinhasExistentes + 1, iGrid_QuantVend_Col) = Formata_Estoque((objItemRomaneio.dQuantidade - objItemRomaneio.dQuantCancelada) * dFator2)
            GridAlocacao.TextMatrix(objGridAlocacoes.iLinhasExistentes + 1, iGrid_QuantAloc_Col) = Formata_Estoque(objReserva.dQuantidade * dFator)
            
            objGridAlocacoes.iLinhasExistentes = objGridAlocacoes.iLinhasExistentes + 1
            
        Next
               
    Next

    Exit Sub
    
Erro_Atualiza_Grid_Alocacao:

    Select Case gErr
    
        Case 42764, 42766, 46181, 86357
        
        Case 42765
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 46182
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objAlmoxarifado.iCodigo)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157278)
            
    End Select
    
    Exit Sub

End Sub

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
