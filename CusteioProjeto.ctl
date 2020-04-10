VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl CusteioProjetoOcx 
   ClientHeight    =   5745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5745
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4590
      Index           =   1
      Left            =   105
      TabIndex        =   18
      Top             =   975
      Width           =   9225
      Begin VB.CommandButton BotaoAtualizar 
         Caption         =   "Atualizar"
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
         Left            =   3765
         TabIndex        =   15
         Top             =   4185
         Width           =   1575
      End
      Begin VB.CommandButton BotaoVerKit 
         Caption         =   "Ver Kit"
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
         Left            =   375
         TabIndex        =   13
         Top             =   4185
         Width           =   1545
      End
      Begin VB.CommandButton BotaoVerRoteiro 
         Caption         =   "Ver Roteiro"
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
         Left            =   2070
         TabIndex        =   14
         Top             =   4185
         Width           =   1545
      End
      Begin MSMask.MaskEdBox Projeto 
         Height          =   315
         Left            =   1740
         TabIndex        =   0
         Top             =   60
         Width           =   1390
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Frame Frame2 
         Caption         =   "Itens"
         Height          =   3210
         Index           =   3
         Left            =   195
         TabIndex        =   19
         Top             =   930
         Width           =   8865
         Begin VB.CommandButton BotaoGrade 
            Caption         =   "Grade ..."
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
            Left            =   195
            TabIndex        =   12
            Top             =   2730
            Visible         =   0   'False
            Width           =   1545
         End
         Begin MSMask.MaskEdBox PrecoTotalItens 
            Height          =   315
            Left            =   6345
            TabIndex        =   57
            Top             =   615
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrecoUnitItens 
            Height          =   315
            Left            =   5880
            TabIndex        =   56
            Top             =   1140
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CustoMOItens 
            Height          =   315
            Left            =   4935
            TabIndex        =   55
            Top             =   1140
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CustoInsumosMaqItens 
            Height          =   315
            Left            =   3975
            TabIndex        =   54
            Top             =   1140
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.OptionButton Selecionar 
            Height          =   330
            Left            =   630
            TabIndex        =   53
            Top             =   1590
            Width           =   855
         End
         Begin VB.TextBox DescricaoProdItens 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   3135
            TabIndex        =   50
            Top             =   1620
            Width           =   1770
         End
         Begin VB.ComboBox UMProdItens 
            Height          =   315
            Left            =   4920
            TabIndex        =   49
            Top             =   1605
            Width           =   600
         End
         Begin VB.ComboBox VersaoProdItens 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "CusteioProjeto.ctx":0000
            Left            =   2205
            List            =   "CusteioProjeto.ctx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   1605
            Width           =   930
         End
         Begin MSMask.MaskEdBox CustoMPItens 
            Height          =   315
            Left            =   6960
            TabIndex        =   48
            Top             =   1605
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoItens 
            Height          =   315
            Left            =   1380
            TabIndex        =   51
            Top             =   1320
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantidadeProdItens 
            Height          =   315
            Left            =   5520
            TabIndex        =   52
            Top             =   1605
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   2325
            Left            =   180
            TabIndex        =   2
            Top             =   240
            Width           =   8565
            _ExtentX        =   15108
            _ExtentY        =   4101
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Preço Total:"
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
            Index           =   4
            Left            =   6120
            TabIndex        =   97
            Top             =   2760
            Width           =   1065
         End
         Begin VB.Label PrecoTotalProjeto 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7245
            TabIndex        =   96
            Top             =   2730
            Width           =   1500
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Custo Total:"
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
            Index           =   3
            Left            =   3270
            TabIndex        =   95
            Top             =   2760
            Width           =   1050
         End
         Begin VB.Label CustoTotalProjeto 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4380
            TabIndex        =   94
            Top             =   2730
            Width           =   1500
         End
      End
      Begin VB.CommandButton BotaoCalcularPreco 
         Caption         =   "Recalcular Preço Unitário"
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
         Left            =   5650
         TabIndex        =   16
         Top             =   4185
         Width           =   3300
      End
      Begin MSComCtl2.UpDown UpDownDataCusteio 
         Height          =   300
         Left            =   2910
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   465
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataCusteio 
         Height          =   315
         Left            =   1740
         TabIndex        =   1
         Top             =   465
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelData 
         AutoSize        =   -1  'True
         Caption         =   "Data do Custeio:"
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
         Left            =   195
         TabIndex        =   93
         Top             =   495
         Width           =   1440
      End
      Begin VB.Label LabelDescProjeto 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3195
         TabIndex        =   92
         Top             =   60
         Width           =   5865
      End
      Begin VB.Label LabelProjeto 
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
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   960
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   91
         Top             =   120
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4590
      Index           =   2
      Left            =   90
      TabIndex        =   21
      Top             =   960
      Visible         =   0   'False
      Width           =   9225
      Begin VB.Frame Frame2 
         Caption         =   "Insumos Usados no Kit"
         Height          =   3720
         Index           =   1
         Left            =   210
         TabIndex        =   22
         Top             =   690
         Width           =   8865
         Begin VB.TextBox ObservacaoKit 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   6255
            MaxLength       =   255
            TabIndex        =   66
            Top             =   1185
            Width           =   3540
         End
         Begin MSMask.MaskEdBox CustoTotalKit 
            Height          =   315
            Left            =   5295
            TabIndex        =   65
            Top             =   1185
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox VariacaoKit 
            Height          =   315
            Left            =   6630
            TabIndex        =   64
            Top             =   1605
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   8
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
         Begin MSMask.MaskEdBox CustoUnitKit 
            Height          =   315
            Left            =   5670
            TabIndex        =   63
            Top             =   1605
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.ComboBox UMProdKit 
            Height          =   315
            Left            =   3135
            TabIndex        =   59
            Top             =   1620
            Width           =   600
         End
         Begin VB.TextBox DescricaoProdKit 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   1365
            TabIndex        =   58
            Top             =   1605
            Width           =   1770
         End
         Begin MSMask.MaskEdBox CustoUnitCalculadoKit 
            Height          =   315
            Left            =   4710
            TabIndex        =   60
            Top             =   1605
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoKit 
            Height          =   315
            Left            =   630
            TabIndex        =   61
            Top             =   1605
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantidadeProdKit 
            Height          =   315
            Left            =   3975
            TabIndex        =   62
            Top             =   1605
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridKit 
            Height          =   2820
            Left            =   165
            TabIndex        =   8
            Top             =   315
            Width           =   8565
            _ExtentX        =   15108
            _ExtentY        =   4974
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
         Begin VB.Label CustoTotalInsumosKit 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7215
            TabIndex        =   24
            Top             =   3360
            Width           =   1500
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Custo Total:"
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
            Index           =   0
            Left            =   6105
            TabIndex        =   23
            Top             =   3390
            Width           =   1050
         End
      End
      Begin VB.Frame SSFrame1 
         Height          =   555
         Index           =   1
         Left            =   210
         TabIndex        =   32
         Top             =   90
         Width           =   8865
         Begin VB.Label LabelVersaoInsumosKit 
            Height          =   210
            Left            =   6795
            TabIndex        =   86
            Top             =   210
            Width           =   1305
         End
         Begin VB.Label LabelProdutoInsumosKit 
            Height          =   210
            Left            =   975
            TabIndex        =   85
            Top             =   210
            Width           =   4935
         End
         Begin VB.Label VersaoLabel 
            Height          =   210
            Index           =   1
            Left            =   6735
            TabIndex        =   36
            Top             =   195
            Width           =   1665
         End
         Begin VB.Label Label30 
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
            Height          =   210
            Index           =   0
            Left            =   6060
            TabIndex        =   35
            Top             =   195
            Width           =   810
         End
         Begin VB.Label ProdutoLabel 
            Height          =   210
            Index           =   1
            Left            =   1170
            TabIndex        =   34
            Top             =   180
            Width           =   4665
         End
         Begin VB.Label Label30 
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
            Height          =   210
            Index           =   1
            Left            =   195
            TabIndex        =   33
            Top             =   210
            Width           =   810
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4590
      Index           =   4
      Left            =   90
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   9225
      Begin VB.Frame SSFrame1 
         Height          =   555
         Index           =   2
         Left            =   210
         TabIndex        =   42
         Top             =   90
         Width           =   8865
         Begin VB.Label LabelVersaoMO 
            Height          =   210
            Left            =   6795
            TabIndex        =   90
            Top             =   210
            Width           =   1305
         End
         Begin VB.Label LabelProdutoMO 
            Height          =   210
            Left            =   975
            TabIndex        =   89
            Top             =   210
            Width           =   4935
         End
         Begin VB.Label Label30 
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
            Height          =   210
            Index           =   5
            Left            =   195
            TabIndex        =   46
            Top             =   210
            Width           =   810
         End
         Begin VB.Label ProdutoLabel 
            Height          =   210
            Index           =   3
            Left            =   1155
            TabIndex        =   45
            Top             =   180
            Width           =   4665
         End
         Begin VB.Label Label30 
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
            Height          =   210
            Index           =   4
            Left            =   6060
            TabIndex        =   44
            Top             =   195
            Width           =   810
         End
         Begin VB.Label VersaoLabel 
            Height          =   210
            Index           =   3
            Left            =   6735
            TabIndex        =   43
            Top             =   210
            Width           =   1665
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Mão de Obra Usadas"
         Height          =   3720
         Index           =   5
         Left            =   210
         TabIndex        =   29
         Top             =   690
         Width           =   8865
         Begin VB.TextBox ObservacaoMO 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   6330
            MaxLength       =   255
            TabIndex        =   78
            Top             =   1020
            Width           =   3540
         End
         Begin VB.ComboBox UMMO 
            Height          =   315
            Left            =   3210
            TabIndex        =   77
            Top             =   1455
            Width           =   600
         End
         Begin VB.TextBox DescricaoMO 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   1440
            TabIndex        =   76
            Top             =   1440
            Width           =   1770
         End
         Begin MSMask.MaskEdBox CustoTotalMO 
            Height          =   315
            Left            =   5370
            TabIndex        =   79
            Top             =   1020
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox VariacaoMO 
            Height          =   315
            Left            =   6705
            TabIndex        =   80
            Top             =   1440
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   8
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
         Begin MSMask.MaskEdBox CustoUnitMO 
            Height          =   315
            Left            =   5745
            TabIndex        =   81
            Top             =   1440
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CustoUnitCalculadoMO 
            Height          =   315
            Left            =   4785
            TabIndex        =   82
            Top             =   1440
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox TipoMO 
            Height          =   315
            Left            =   705
            TabIndex        =   83
            Top             =   1440
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   3
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantidadeMO 
            Height          =   315
            Left            =   4050
            TabIndex        =   84
            Top             =   1440
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridMaoDeObra 
            Height          =   2820
            Left            =   165
            TabIndex        =   10
            Top             =   315
            Width           =   8565
            _ExtentX        =   15108
            _ExtentY        =   4974
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Custo Total:"
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
            Left            =   6105
            TabIndex        =   31
            Top             =   3390
            Width           =   1050
         End
         Begin VB.Label CustoTotalMaoDeObra 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7215
            TabIndex        =   30
            Top             =   3360
            Width           =   1500
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4590
      Index           =   3
      Left            =   90
      TabIndex        =   25
      Top             =   960
      Visible         =   0   'False
      Width           =   9225
      Begin VB.Frame Frame2 
         Caption         =   "Insumos Usados na Máquina"
         Height          =   3720
         Index           =   4
         Left            =   210
         TabIndex        =   26
         Top             =   690
         Width           =   8865
         Begin VB.TextBox DescricaoProdMaq 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   1395
            TabIndex        =   72
            Top             =   1650
            Width           =   1770
         End
         Begin VB.ComboBox UMProdMaq 
            Height          =   315
            Left            =   3165
            TabIndex        =   71
            Top             =   1665
            Width           =   600
         End
         Begin VB.TextBox ObservacaoMaq 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   6285
            MaxLength       =   255
            TabIndex        =   67
            Top             =   1230
            Width           =   3540
         End
         Begin MSMask.MaskEdBox CustoTotalMaq 
            Height          =   315
            Left            =   5325
            TabIndex        =   68
            Top             =   1230
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox VariacaoMaq 
            Height          =   315
            Left            =   6660
            TabIndex        =   69
            Top             =   1650
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   8
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
         Begin MSMask.MaskEdBox CustoUnitMaq 
            Height          =   315
            Left            =   5700
            TabIndex        =   70
            Top             =   1650
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CustoUnitCalculadoMaq 
            Height          =   315
            Left            =   4740
            TabIndex        =   73
            Top             =   1650
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ProdutoMaq 
            Height          =   315
            Left            =   660
            TabIndex        =   74
            Top             =   1650
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantidadeProdMaq 
            Height          =   315
            Left            =   4005
            TabIndex        =   75
            Top             =   1650
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridMaquinas 
            Height          =   2820
            Left            =   165
            TabIndex        =   9
            Top             =   316
            Width           =   8565
            _ExtentX        =   15108
            _ExtentY        =   4974
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Custo Total:"
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
            Left            =   6105
            TabIndex        =   28
            Top             =   3390
            Width           =   1050
         End
         Begin VB.Label CustoTotalInsumosMaquina 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7215
            TabIndex        =   27
            Top             =   3360
            Width           =   1500
         End
      End
      Begin VB.Frame SSFrame1 
         Height          =   555
         Index           =   0
         Left            =   210
         TabIndex        =   37
         Top             =   90
         Width           =   8865
         Begin VB.Label LabelVersaoInsumosMaq 
            Height          =   210
            Left            =   6795
            TabIndex        =   88
            Top             =   210
            Width           =   1305
         End
         Begin VB.Label LabelProdutoInsumosMaq 
            Height          =   210
            Left            =   975
            TabIndex        =   87
            Top             =   210
            Width           =   4935
         End
         Begin VB.Label Label30 
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
            Height          =   210
            Index           =   3
            Left            =   195
            TabIndex        =   41
            Top             =   210
            Width           =   810
         End
         Begin VB.Label ProdutoLabel 
            Height          =   210
            Index           =   2
            Left            =   1170
            TabIndex        =   40
            Top             =   180
            Width           =   4665
         End
         Begin VB.Label Label30 
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
            Height          =   210
            Index           =   2
            Left            =   6060
            TabIndex        =   39
            Top             =   195
            Width           =   810
         End
         Begin VB.Label VersaoLabel 
            Height          =   210
            Index           =   2
            Left            =   6720
            TabIndex        =   38
            Top             =   195
            Width           =   1665
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   7275
      ScaleHeight     =   450
      ScaleWidth      =   2055
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   90
      Width           =   2115
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   1590
         Picture         =   "CusteioProjeto.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   1080
         Picture         =   "CusteioProjeto.ctx":0182
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   345
         Left            =   555
         Picture         =   "CusteioProjeto.ctx":06B4
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   345
         Left            =   45
         Picture         =   "CusteioProjeto.ctx":083E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5070
      Left            =   60
      TabIndex        =   17
      Top             =   570
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   8943
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Projeto"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Insumos Kit"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Insumos Máquina"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Mão de Obra"
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
Attribute VB_Name = "CusteioProjetoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iFrameAtual As Integer
Dim iProjetoAlterado As Integer
Dim iGridKitAlterado As Integer
Dim iGridMaqAlterado As Integer
Dim iGridMOAlterado As Integer

'Grid de Itens
Dim objGridItens As AdmGrid
Dim iGrid_Selecionar_Col As Integer
Dim iGrid_ProdutoItens_Col As Integer
Dim iGrid_VersaoProdItens_Col As Integer
Dim iGrid_DescricaoProdItens_Col As Integer
Dim iGrid_UMProdItens_Col As Integer
Dim iGrid_QuantidadeProdItens_Col As Integer
Dim iGrid_CustoMPItens_Col As Integer
Dim iGrid_CustoInsumosMaqItens_Col As Integer
Dim iGrid_CustoMOItens_Col As Integer
Dim iGrid_PrecoUnitItens_Col As Integer
Dim iGrid_PrecoTotalItens_Col As Integer

'Grid de Kits
Dim objGridKit As AdmGrid
Dim iGrid_ProdutoKit_Col As Integer
Dim iGrid_DescricaoProdKit_Col As Integer
Dim iGrid_UMProdKit_Col As Integer
Dim iGrid_QuantidadeProdKit_Col As Integer
Dim iGrid_CustoUnitCalculadoKit_Col As Integer
Dim iGrid_CustoUnitKit_Col As Integer
Dim iGrid_VariacaoKit_Col As Integer
Dim iGrid_CustoTotalKit_Col As Integer
Dim iGrid_ObservacaoKit_Col As Integer

'Grid de Maquinas
Dim objGridMaquinas As AdmGrid
Dim iGrid_ProdutoMaq_Col As Integer
Dim iGrid_DescricaoProdMaq_Col As Integer
Dim iGrid_UMProdMaq_Col As Integer
Dim iGrid_QuantidadeProdMaq_Col As Integer
Dim iGrid_CustoUnitCalculadoMaq_Col As Integer
Dim iGrid_CustoUnitMaq_Col As Integer
Dim iGrid_VariacaoMaq_Col As Integer
Dim iGrid_CustoTotalMaq_Col As Integer
Dim iGrid_ObservacaoMaq_Col As Integer

'Grid de MaoDeObra
Dim objGridMaoDeObra As AdmGrid
Dim iGrid_TipoMO_Col As Integer
Dim iGrid_DescricaoMo_Col As Integer
Dim iGrid_UMMO_Col As Integer
Dim iGrid_QuantidadeMO_Col As Integer
Dim iGrid_CustoUnitCalculadoMO_Col As Integer
Dim iGrid_CustoUnitMO_Col As Integer
Dim iGrid_VariacaoMO_Col As Integer
Dim iGrid_CustoTotalMO_Col As Integer
Dim iGrid_ObservacaoMO_Col As Integer

Private WithEvents objEventoProjeto As AdmEvento
Attribute objEventoProjeto.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoVersao As AdmEvento
Attribute objEventoVersao.VB_VarHelpID = -1

Dim colProjetoCusteioItens As Collection

Private Const TAB_Projeto = 1
Private Const TAB_InsumosKit = 2
Private Const TAB_InsumosMaq = 3
Private Const TAB_MaoDeObra = 4

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Custeio de Projeto"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "CusteioProjeto"

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

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objProjeto As New ClassProjeto
Dim objProjetoCusteio As ClassProjetoCusteio
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'CRITICA DADOS DA TELA
    If Len(Trim(Projeto.Text)) = 0 Then gError 134290

    Set objProjeto = New ClassProjeto
    
    'Le o Projeto para pegar seu NumIntDoc
    lErro = CF("TP_Projeto_Le", Projeto, objProjeto)
    If lErro <> SUCESSO Then gError 134467
    
    'Ajusta a variável alterarada indevidamente pela TP_Projeto_Le
    iProjetoAlterado = 0
    
    Set objProjetoCusteio = New ClassProjetoCusteio

    objProjetoCusteio.lNumIntDocProj = objProjeto.lNumIntDoc

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_PROJETOCUSTEIO", Trim(Projeto.Text))

    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If

    'Exclui o ProjetoCusteio
    lErro = CF("ProjetoCusteio_Exclui", objProjetoCusteio)
    If lErro <> SUCESSO Then gError 134291

    'Limpa Tela
    Call Limpa_Tela_CusteioProjeto

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 134290
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PROJETO_NAO_PREENCHIDO", gErr)

        Case 134291

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158431)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158432)

    End Select

    Exit Sub

End Sub


Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Grava o Projeto
    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 134101

    'Limpa Tela
    Call Limpa_Tela_CusteioProjeto
    
    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 134101

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158433)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 134289
    Call Limpa_Tela_CusteioProjeto

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 134289

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158434)

    End Select

    Exit Sub

End Sub

Private Sub BotaoVerKit_Click()

Dim lErro As Long
Dim objKit As New ClassKit
Dim sProduto As String
Dim sVersao As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoVerKit

    If Me.ActiveControl Is ProdutoItens Then
    
        sProduto = Trim(ProdutoItens.Text)
        sVersao = GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col)
        
    ElseIf Me.ActiveControl Is VersaoProdItens Then
    
        sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
        sVersao = Trim(VersaoProdItens.Text)
            
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 134698
        
        sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
        sVersao = GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col)
        
    End If

    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134756

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        objKit.sProdutoRaiz = sProdutoFormatado
        If Len(sVersao) > 0 Then
        
            objKit.sVersao = sVersao
        
            lErro = CF("Kit_Le", objKit)
            If lErro <> SUCESSO And lErro <> 21826 Then gError 134200
        
            If lErro <> SUCESSO Then gError 134201
            
        Else
        
            lErro = CF("Kit_Le_Padrao", objKit)
            If lErro <> SUCESSO And lErro <> 106304 Then gError 134202
        
            If lErro <> SUCESSO Then gError 134203
        
        End If
            
        Call Chama_Tela("Kit", objKit)
    
    Else
         gError 134757
         
    End If

    Exit Sub
    
Erro_BotaoVerKit:

    Select Case gErr
    
        Case 134200, 134202, 134756
        
        Case 134201, 134203
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_E_KIT", gErr)
    
        Case 134698
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case 134757
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRODUTO_NAO_PREENCHIDO2", gErr)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158435)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoVerRoteiro_Click()

Dim lErro As Long
Dim objRoteirosDeFabricacao As New ClassRoteirosDeFabricacao
Dim sProduto As String
Dim sVersao As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoVerKit

    If Me.ActiveControl Is ProdutoItens Then
    
        sProduto = Trim(ProdutoItens.Text)
        sVersao = GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col)
        
    ElseIf Me.ActiveControl Is VersaoProdItens Then
    
        sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
        sVersao = Trim(VersaoProdItens.Text)
            
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 134698
        
        sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
        sVersao = GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col)
        
    End If

    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134756

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        objRoteirosDeFabricacao.sProdutoRaiz = sProdutoFormatado
        If Len(sVersao) > 0 Then objRoteirosDeFabricacao.sVersao = sVersao
        
        lErro = CF("RoteirosDeFabricacao_Le_Roteiro", objRoteirosDeFabricacao)
        If lErro <> SUCESSO And lErro <> 134622 Then gError 134200
    
        If lErro <> SUCESSO Then gError 134201
            
        Call Chama_Tela("RoteirosDeFabricacao", objRoteirosDeFabricacao)
    
    Else
         gError 134757
         
    End If

    Exit Sub
    
Erro_BotaoVerKit:

    Select Case gErr
            
        Case 134201
            Call Rotina_Erro(vbOKOnly, "ERRO_ROTEIROSDEFABRICACAO_NAO_CADASTRADO", gErr, sProdutoFormatado, sVersao)
    
        Case 134698
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case 134756, 134200
        
        Case 134757
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRODUTO_NAO_PREENCHIDO3", gErr)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158436)
    
    End Select
    
    Exit Sub

End Sub

Private Sub DataCusteio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataCusteio_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataCusteio, iAlterado)
    
End Sub


Private Sub DataCusteio_Validate(Cancel As Boolean)
Dim lErro As Long

On Error GoTo Erro_DataCusteio_Validate

    'Verifica se DataCusteio está preenchida
    If Len(Trim(DataCusteio.Text)) <> 0 Then

        lErro = Data_Critica(DataCusteio.Text)
        If lErro <> SUCESSO Then gError 134752

    End If

    Exit Sub

Erro_DataCusteio_Validate:

    Cancel = True

    Select Case gErr

        Case 134752

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158437)

    End Select

    Exit Sub

End Sub


Private Sub LabelProjeto_Click()

Dim lErro As Long
Dim objProjeto As ClassProjeto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Projeto foi preenchido
    If Len(Trim(Projeto.Text)) <> 0 Then

        Set objProjeto = New ClassProjeto
        
        'Verifica sua existencia
        lErro = CF("TP_Projeto_Le", Projeto, objProjeto)
        If lErro <> SUCESSO Then gError 134467

    End If

    Call Chama_Tela("ProjetoLista", colSelecao, objProjeto, objEventoProjeto)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158438)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String
Dim iLinha As Integer
Dim objProdutoKit As ClassProdutoKit

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1
        
    lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 134708
    
    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158439)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProjeto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProjeto As New ClassProjeto
Dim sProjetoAntigo As String

On Error GoTo Erro_objEventoProjeto_evSelecao

    Set objProjeto = obj1
    
    sProjetoAntigo = Projeto.Text

    Projeto.Text = objProjeto.sNomeReduzido
    
    If sProjetoAntigo = Projeto.Text Then
    
        iProjetoAlterado = 0
    
    End If
    
    Call Projeto_Validate(bSGECancelDummy)
        
    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoProjeto_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158440)

    End Select

    Exit Sub

End Sub

Private Sub objEventoVersao_evSelecao(obj1 As Object)

Dim objKit As ClassKit
Dim lErro As Long

On Error GoTo Erro_objEventoVersao_evSelecao

    Set objKit = obj1

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoVersao_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158441)

    End Select

    Exit Sub
    
End Sub

Private Sub Projeto_Change()

    iAlterado = REGISTRO_ALTERADO
    iProjetoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Projeto_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Projeto, iAlterado)
    
End Sub

Private Sub Projeto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProjeto As ClassProjeto
Dim objProjetoCusteio As ClassProjetoCusteio

On Error GoTo Erro_Projeto_Validate

    If iProjetoAlterado = REGISTRO_ALTERADO Then
    
        LabelDescProjeto.Caption = ""
    
        'Verifica se Projeto está preenchido
        If Len(Trim(Projeto.Text)) > 0 Then
        
            Set objProjeto = New ClassProjeto
            
            'Verifica sua existencia
            lErro = CF("TP_Projeto_Le", Projeto, objProjeto)
            If lErro <> SUCESSO Then gError 134467
            
            'Coloca a descrição na tela
            LabelDescProjeto.Caption = objProjeto.sDescricao
            
            Set objProjetoCusteio = New ClassProjetoCusteio
        
            objProjetoCusteio.lNumIntDocProj = objProjeto.lNumIntDoc
            
            'Traz o Custeio para a tela
            lErro = Traz_ProjetoCusteio_Tela(objProjetoCusteio)
            If lErro <> SUCESSO Then gError 134701
            
        End If
            
    End If
    
    Exit Sub

Erro_Projeto_Validate:

    Cancel = True

    Select Case gErr
    
        Case 134467, 134701
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158442)

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
        
        'Se Frame selecionado foi diferente de Projeto
        If TabStrip1.SelectedItem.Index <> TAB_Projeto Then
                        
            Call Trata_Linha(TabStrip1.SelectedItem.Index)
                
        End If
                
    End If

End Sub

Private Sub UpDownDataCusteio_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataCusteio_DownClick

    DataCusteio.SetFocus

    If Len(DataCusteio.ClipText) > 0 Then

        sData = DataCusteio.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 134750

        DataCusteio.Text = sData

    End If

    Exit Sub

Erro_UpDownDataCusteio_DownClick:

    Select Case gErr

        Case 134750

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158443)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataCusteio_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataCusteio_UpClick

    DataCusteio.SetFocus

    If Len(Trim(DataCusteio.ClipText)) > 0 Then

        sData = DataCusteio.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 134751

        DataCusteio.Text = sData

    End If

    Exit Sub

Erro_UpDownDataCusteio_UpClick:

    Select Case gErr

        Case 134751

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158444)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Projeto Then Call LabelProjeto_Click
        
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
    Call TelaIndice_Preenche(Me)

End Sub
    
Public Sub Form_Deactivate()
    
    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload
    
    Call ComandoSeta_Liberar(Me.Name)

    Set objEventoProjeto = Nothing
    Set objEventoProduto = Nothing
    Set objEventoVersao = Nothing
    
    Set objGridItens = Nothing
    Set objGridKit = Nothing
    Set objGridMaquinas = Nothing
    Set objGridMaoDeObra = Nothing

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158445)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
   
    iFrameAtual = 1
    
    Set objEventoProjeto = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoVersao = New AdmEvento
    
    DataCusteio.PromptInclude = False
    DataCusteio.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataCusteio.PromptInclude = True
    
    LabelDescProjeto.Caption = ""
        
    'Grid Itens
    Set objGridItens = New AdmGrid
    
    'tela em questão
    Set objGridItens.objForm = Me
    
    lErro = Inicializa_GridItens(objGridItens)
    If lErro <> SUCESSO Then gError 134340
            
    'Grid Kit
    Set objGridKit = New AdmGrid
    
    'tela em questão
    Set objGridKit.objForm = Me
    
    lErro = Inicializa_GridKit(objGridKit)
    If lErro <> SUCESSO Then gError 134340
                        
    'Grid Maquinas
    Set objGridMaquinas = New AdmGrid
    
    'tela em questão
    Set objGridMaquinas.objForm = Me
    
    lErro = Inicializa_GridMaquinas(objGridMaquinas)
    If lErro <> SUCESSO Then gError 134340
    
    'Grid MaoDeObra
    Set objGridMaoDeObra = New AdmGrid
    
    'tela em questão
    Set objGridMaoDeObra.objForm = Me
    
    lErro = Inicializa_GridMaoDeObra(objGridMaoDeObra)
    If lErro <> SUCESSO Then gError 134340
    
    iAlterado = 0
    iProjetoAlterado = 0
    iGridKitAlterado = 0
    iGridMaqAlterado = 0
    iGridMOAlterado = 0
            
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 134340, 134200
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158446)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objProjetoCusteio As ClassProjetoCusteio) As Long

Dim lErro As Long
Dim objProjeto As New ClassProjeto
Dim sProjetoAntigo As String

On Error GoTo Erro_Trata_Parametros

    If Not (objProjetoCusteio Is Nothing) Then
    
        objProjeto.lNumIntDoc = objProjetoCusteio.lNumIntDocProj
        
        lErro = CF("Projeto_Le", objProjeto)
        If lErro <> SUCESSO And lErro <> 134449 Then gError 134721
        
        Projeto.Text = objProjeto.sNomeReduzido
        LabelDescProjeto.Caption = objProjeto.sDescricao

        lErro = Traz_ProjetoCusteio_Tela(objProjetoCusteio)
        If lErro <> SUCESSO Then gError 134722

    End If
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case 134721, 134722

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158447)

    End Select

    iAlterado = 0

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
    
        'Verifica se é o GridItens
        If objGridInt.objGrid.Name = GridItens.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_Selecionar_Col

                    lErro = Saida_Celula_Selecionar(objGridInt)
                    If lErro <> SUCESSO Then gError 134370

            End Select
        
        'GridKit
        ElseIf objGridInt.objGrid.Name = GridKit.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_CustoUnitKit_Col

                    lErro = Saida_Celula_CustoUnitKit(objGridInt)
                    If lErro <> SUCESSO Then gError 134373

                Case iGrid_VariacaoKit_Col

                    lErro = Saida_Celula_VariacaoKit(objGridInt)
                    If lErro <> SUCESSO Then gError 134374
                
                Case iGrid_ObservacaoKit_Col

                    lErro = Saida_Celula_ObservacaoKit(objGridInt)
                    If lErro <> SUCESSO Then gError 134374
                    
            End Select
        
        'GridMaquinas
        ElseIf objGridInt.objGrid.Name = GridMaquinas.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_CustoUnitMaq_Col

                    lErro = Saida_Celula_CustoUnitMaq(objGridInt)
                    If lErro <> SUCESSO Then gError 134373

                Case iGrid_VariacaoMaq_Col

                    lErro = Saida_Celula_VariacaoMaq(objGridInt)
                    If lErro <> SUCESSO Then gError 134374
                
                Case iGrid_ObservacaoMaq_Col

                    lErro = Saida_Celula_ObservacaoMaq(objGridInt)
                    If lErro <> SUCESSO Then gError 134374
                    
            End Select
                    
        'GridMaoDeObra
        ElseIf objGridInt.objGrid.Name = GridMaoDeObra.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_CustoUnitMO_Col

                    lErro = Saida_Celula_CustoUnitMO(objGridInt)
                    If lErro <> SUCESSO Then gError 134373

                Case iGrid_VariacaoMO_Col

                    lErro = Saida_Celula_VariacaoMO(objGridInt)
                    If lErro <> SUCESSO Then gError 134374
                
                Case iGrid_ObservacaoMO_Col

                    lErro = Saida_Celula_ObservacaoMO(objGridInt)
                    If lErro <> SUCESSO Then gError 134374
                    
            End Select
                
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 134375

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 134370 To 134374
            'erros tratatos nas rotinas chamadas
        
        Case 134375
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158448)

    End Select

    Exit Function

End Function

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_GotFocus()
    
    Call Grid_Recebe_Foco(objGridItens)

End Sub

Private Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGridItens, iAlterado)

End Sub

Private Sub GridItens_LeaveCell()
    
    Call Saida_Celula(objGridItens)

End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridItens)
            
End Sub

Private Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub

Private Function Inicializa_GridItens(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Selecionar")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Versão")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("U.M.")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("$ M.Prima")
    objGrid.colColuna.Add ("$ Ins.Maq.")
    objGrid.colColuna.Add ("$ M.Obra")
    objGrid.colColuna.Add ("Pço.Unit.")
    objGrid.colColuna.Add ("Total")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Selecionar.Name)
    objGrid.colCampo.Add (ProdutoItens.Name)
    objGrid.colCampo.Add (VersaoProdItens.Name)
    objGrid.colCampo.Add (DescricaoProdItens.Name)
    objGrid.colCampo.Add (UMProdItens.Name)
    objGrid.colCampo.Add (QuantidadeProdItens.Name)
    objGrid.colCampo.Add (CustoMPItens.Name)
    objGrid.colCampo.Add (CustoInsumosMaqItens.Name)
    objGrid.colCampo.Add (CustoMOItens.Name)
    objGrid.colCampo.Add (PrecoUnitItens.Name)
    objGrid.colCampo.Add (PrecoTotalItens.Name)

    'Colunas do Grid
    iGrid_Selecionar_Col = 1
    iGrid_ProdutoItens_Col = 2
    iGrid_VersaoProdItens_Col = 3
    iGrid_DescricaoProdItens_Col = 4
    iGrid_UMProdItens_Col = 5
    iGrid_QuantidadeProdItens_Col = 6
    iGrid_CustoMPItens_Col = 7
    iGrid_CustoInsumosMaqItens_Col = 8
    iGrid_CustoMOItens_Col = 9
    iGrid_PrecoUnitItens_Col = 10
    iGrid_PrecoTotalItens_Col = 11

    objGrid.objGrid = GridItens

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 21

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 250
    
    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    Call Grid_Inicializa(objGrid)

    Inicializa_GridItens = SUCESSO

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iProdutoPreenchidoKit As Integer
Dim iProdutoPreenchidoMaq As Integer

On Error GoTo Erro_Rotina_Grid_Enable
    
    'Guardo o valor do Codigo do Produto do Item do Projeto
    sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
    
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134767
    
    'Guardo o valor do Codigo do Produto de InsumosKit
    sProduto = GridKit.TextMatrix(GridKit.Row, iGrid_ProdutoKit_Col)
    
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchidoKit)
    If lErro <> SUCESSO Then gError 134767
    
    'Guardo o valor do Codigo do Produto de InsumosMaquina
    sProduto = GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_ProdutoMaq_Col)
    
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchidoMaq)
    If lErro <> SUCESSO Then gError 134767
    
    'Grid Itens
    If objControl.Name = "Selecionar" Then
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = True
    
        Else
            objControl.Enabled = False
        
        End If
    
    ElseIf objControl.Name = "ProdutoItens" Or _
            objControl.Name = "VersaoProdItens" Or _
            objControl.Name = "UMProdItens" Or _
            objControl.Name = "QuantidadeProdItens" Or _
            objControl.Name = "DescricaoProdItens" Or _
            objControl.Name = "CustoMPItens" Or _
            objControl.Name = "CustoInsumosMaqItens" Or _
            objControl.Name = "CustoMOItens" Or _
            objControl.Name = "PrecoUnitItens" Or _
            objControl.Name = "PrecoTotalItens" Then

        objControl.Enabled = False
            
    'Grid Kit
    ElseIf objControl.Name = "ProdutoKit" Or _
            objControl.Name = "DescricaoProdKit" Or _
            objControl.Name = "UMProdKit" Or _
            objControl.Name = "QuantidadeProdKit" Or _
            objControl.Name = "CustoUnitCalculadoKit" Or _
            objControl.Name = "CustoTotalKit" Then

        objControl.Enabled = False
    
    ElseIf objControl.Name = "CustoUnitKit" Or _
            objControl.Name = "VariacaoKit" Or _
            objControl.Name = "ObservacaoKit" Then
        
        If iProdutoPreenchidoKit = PRODUTO_PREENCHIDO Then
            objControl.Enabled = True

        Else
            objControl.Enabled = False
        
        End If

    'Grid Maquina
    ElseIf objControl.Name = "ProdutoMaq" Or _
            objControl.Name = "DescricaoProdMaq" Or _
            objControl.Name = "UMProdMaq" Or _
            objControl.Name = "QuantidadeProdMaq" Or _
            objControl.Name = "CustoUnitCalculadoMaq" Or _
            objControl.Name = "CustoTotalMaq" Then

        objControl.Enabled = False
    
    ElseIf objControl.Name = "CustoUnitMaq" Or _
            objControl.Name = "VariacaoMaq" Or _
            objControl.Name = "ObservacaoMaq" Then
        
        If iProdutoPreenchidoMaq = PRODUTO_PREENCHIDO Then
            objControl.Enabled = True

        Else
            objControl.Enabled = False
        
        End If

    'Grid MaoDeObra
    ElseIf objControl.Name = "TipoMO" Or _
            objControl.Name = "DescricaoMO" Or _
            objControl.Name = "UMMO" Or _
            objControl.Name = "QuantidadeMO" Or _
            objControl.Name = "CustoUnitCalculadoMO" Or _
            objControl.Name = "CustoTotalMO" Then

        objControl.Enabled = False
    
    ElseIf objControl.Name = "CustoUnitMO" Or _
            objControl.Name = "VariacaoMO" Or _
            objControl.Name = "ObservacaoMO" Then
        
        If Len(GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_TipoMO_Col)) > 0 Then
            objControl.Enabled = True

        Else
            objControl.Enabled = False
        
        End If

    End If
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
    
        Case 134767 To 134769

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 158449)

    End Select

    Exit Sub

End Sub

Private Sub Selecionar_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub Selecionar_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub Selecionar_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Selecionar
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ProdutoItens_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoItens_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub ProdutoItens_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub ProdutoItens_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ProdutoItens
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescricaoProdItens_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescricaoProdItens_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub DescricaoProdItens_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub DescricaoProdItens_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DescricaoProdItens
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub VersaoProdItens_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub VersaoProdItens_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub VersaoProdItens_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub VersaoProdItens_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = VersaoProdItens
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UMProdItens_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UMProdItens_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub UMProdItens_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub UMProdItens_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = UMProdItens
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantidadeProdItens_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantidadeProdItens_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub QuantidadeProdItens_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub QuantidadeProdItens_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = QuantidadeProdItens
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoMPItens_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoMPItens_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub CustoMPItens_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub CustoMPItens_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = CustoMPItens
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoInsumosMaqItens_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoInsumosMaqItens_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub CustoInsumosMaqItens_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub CustoInsumosMaqItens_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = CustoInsumosMaqItens
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoMOItens_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoMOItens_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub CustoMOItens_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub CustoMOItens_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = CustoMOItens
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PrecoUnitItens_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PrecoUnitItens_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub PrecoUnitItens_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub PrecoUnitItens_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PrecoUnitItens
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PrecoTotalItens_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub PrecoTotalItens_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Private Sub PrecoTotalItens_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Private Sub PrecoTotalItens_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PrecoTotalItens
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_Selecionar(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Selecionar do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Selecionar

    Set objGridInt.objControle = Selecionar
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134394
    
    Saida_Celula_Selecionar = SUCESSO

    Exit Function

Erro_Saida_Celula_Selecionar:

    Saida_Celula_Selecionar = gErr

    Select Case gErr
        
        Case 134393, 134394
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158450)

    End Select

    Exit Function

End Function

Private Sub GridKit_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGridKit, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGridKit, iAlterado)
        End If

End Sub

Private Sub GridKit_GotFocus()
    
    Call Grid_Recebe_Foco(objGridKit)

End Sub

Private Sub GridKit_EnterCell()

    Call Grid_Entrada_Celula(objGridKit, iAlterado)

End Sub

Private Sub GridKit_LeaveCell()
    
    Call Saida_Celula(objGridKit)

End Sub

Private Sub GridKit_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer


    Call Grid_Trata_Tecla(KeyAscii, objGridKit, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridKit, iAlterado)
    End If

End Sub

Private Sub GridKit_RowColChange()

    Call Grid_RowColChange(objGridKit)

End Sub

Private Sub GridKit_Scroll()

    Call Grid_Scroll(objGridKit)

End Sub

Private Function Inicializa_GridKit(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("U.M.")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("Unit.Calc.")
    objGrid.colColuna.Add ("Unitário")
    objGrid.colColuna.Add ("Variação")
    objGrid.colColuna.Add ("Custo Total")
    objGrid.colColuna.Add ("Observação")

    'Controles que participam do Grid
    objGrid.colCampo.Add (ProdutoKit.Name)
    objGrid.colCampo.Add (DescricaoProdKit.Name)
    objGrid.colCampo.Add (UMProdKit.Name)
    objGrid.colCampo.Add (QuantidadeProdKit.Name)
    objGrid.colCampo.Add (CustoUnitCalculadoKit.Name)
    objGrid.colCampo.Add (CustoUnitKit.Name)
    objGrid.colCampo.Add (VariacaoKit.Name)
    objGrid.colCampo.Add (CustoTotalKit.Name)
    objGrid.colCampo.Add (ObservacaoKit.Name)

    'Colunas do Grid
    iGrid_ProdutoKit_Col = 1
    iGrid_DescricaoProdKit_Col = 2
    iGrid_UMProdKit_Col = 3
    iGrid_QuantidadeProdKit_Col = 4
    iGrid_CustoUnitCalculadoKit_Col = 5
    iGrid_CustoUnitKit_Col = 6
    iGrid_VariacaoKit_Col = 7
    iGrid_CustoTotalKit_Col = 8
    iGrid_ObservacaoKit_Col = 9

    objGrid.objGrid = GridKit

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 21

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridKit.ColWidth(0) = 250
    
    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    Call Grid_Inicializa(objGrid)

    Inicializa_GridKit = SUCESSO

End Function

Private Sub ProdutoKit_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoKit_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridKit)

End Sub

Private Sub ProdutoKit_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridKit)

End Sub

Private Sub ProdutoKit_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridKit.objControle = ProdutoKit
    lErro = Grid_Campo_Libera_Foco(objGridKit)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescricaoProdKit_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescricaoProdKit_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridKit)

End Sub

Private Sub DescricaoProdKit_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridKit)

End Sub

Private Sub DescricaoProdKit_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridKit.objControle = DescricaoProdKit
    lErro = Grid_Campo_Libera_Foco(objGridKit)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UMProdKit_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UMProdKit_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridKit)

End Sub

Private Sub UMProdKit_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridKit)

End Sub

Private Sub UMProdKit_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridKit.objControle = UMProdKit
    lErro = Grid_Campo_Libera_Foco(objGridKit)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantidadeProdKit_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantidadeProdKit_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridKit)

End Sub

Private Sub QuantidadeProdKit_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridKit)

End Sub

Private Sub QuantidadeProdKit_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridKit.objControle = QuantidadeProdKit
    lErro = Grid_Campo_Libera_Foco(objGridKit)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoUnitCalculadoKit_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoUnitCalculadoKit_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridKit)

End Sub

Private Sub CustoUnitCalculadoKit_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridKit)

End Sub

Private Sub CustoUnitCalculadoKit_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridKit.objControle = CustoUnitCalculadoKit
    lErro = Grid_Campo_Libera_Foco(objGridKit)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoUnitKit_Change()

    iAlterado = REGISTRO_ALTERADO
    iGridKitAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoUnitKit_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridKit)

End Sub

Private Sub CustoUnitKit_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridKit)

End Sub

Private Sub CustoUnitKit_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridKit.objControle = CustoUnitKit
    lErro = Grid_Campo_Libera_Foco(objGridKit)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub VariacaoKit_Change()

    iAlterado = REGISTRO_ALTERADO
    iGridKitAlterado = REGISTRO_ALTERADO

End Sub

Private Sub VariacaoKit_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridKit)

End Sub

Private Sub VariacaoKit_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridKit)

End Sub

Private Sub VariacaoKit_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridKit.objControle = VariacaoKit
    lErro = Grid_Campo_Libera_Foco(objGridKit)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoTotalKit_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoTotalKit_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridKit)

End Sub

Private Sub CustoTotalKit_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridKit)

End Sub

Private Sub CustoTotalKit_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridKit.objControle = CustoTotalKit
    lErro = Grid_Campo_Libera_Foco(objGridKit)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ObservacaoKit_Change()

    iAlterado = REGISTRO_ALTERADO
    iGridKitAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ObservacaoKit_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridKit)

End Sub

Private Sub ObservacaoKit_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridKit)

End Sub

Private Sub ObservacaoKit_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridKit.objControle = ObservacaoKit
    lErro = Grid_Campo_Libera_Foco(objGridKit)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_CustoUnitKit(objGridInt As AdmGrid) As Long
'Faz a crítica da célula CustoUnitKit do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dVariacao As Double
Dim dCustoTotalKit As Double
Dim dAjusteCustoKit As Double

On Error GoTo Erro_Saida_Celula_CustoUnitKit

    Set objGridInt.objControle = CustoUnitKit

    'Se o campo foi preenchido
    If Len(Trim(CustoUnitKit.Text)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(CustoUnitKit.Text)
        If lErro <> SUCESSO Then gError 134393
    
    End If
    
    If iGridKitAlterado = REGISTRO_ALTERADO Then
    
        'Calcula a Variação
        lErro = Calcula_Variacao(StrParaDbl(GridKit.TextMatrix(GridKit.Row, iGrid_CustoUnitCalculadoKit_Col)), StrParaDbl(CustoUnitKit.Text), dVariacao, iGrid_VariacaoKit_Col)
        If lErro <> SUCESSO Then gError 134393
        
        'Altera a Variação no grid
        If dVariacao <> 0 Then
           GridKit.TextMatrix(GridKit.Row, iGrid_VariacaoKit_Col) = Format(dVariacao, "Percent")
        Else
           GridKit.TextMatrix(GridKit.Row, iGrid_VariacaoKit_Col) = ""
        End If
        
        'Calcula o Custo Total do Item
        dCustoTotalKit = StrParaDbl(GridKit.TextMatrix(GridKit.Row, iGrid_QuantidadeProdKit_Col)) * StrParaDbl(CustoUnitKit.Text)
        
        'Calcula a diferença a ajustar
        dAjusteCustoKit = dCustoTotalKit - StrParaDbl(GridKit.TextMatrix(GridKit.Row, iGrid_CustoTotalKit_Col))
        
        'Altera o Custo Total do Item no grid
        GridKit.TextMatrix(GridKit.Row, iGrid_CustoTotalKit_Col) = Format(dCustoTotalKit, "Standard")
        
        'Altera a Coleção
        lErro = AjustaCustoNecessidade_InsumosKit(CustoUnitKit.Text, iGrid_CustoUnitKit_Col, dAjusteCustoKit)
        If lErro <> SUCESSO Then gError 134200
        
        iGridKitAlterado = 0
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134394

    Saida_Celula_CustoUnitKit = SUCESSO

    Exit Function

Erro_Saida_Celula_CustoUnitKit:

    Saida_Celula_CustoUnitKit = gErr

    Select Case gErr
        
        Case 134393, 134394
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158451)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_VariacaoKit(objGridInt As AdmGrid) As Long
'Faz a crítica da célula VariacaoKit do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dPrecoUnitarioInformado As Double
Dim dVariacao As Double
Dim dCustoTotalKit As Double
Dim dAjusteCustoKit As Double

On Error GoTo Erro_Saida_Celula_VariacaoKit

    Set objGridInt.objControle = VariacaoKit
    
    If iGridKitAlterado = REGISTRO_ALTERADO Then

        'Ajusta e Converte o valor da tela
        If InStr(VariacaoKit.Text, "%") > 0 Then
        
            dVariacao = StrParaDbl(Left(VariacaoKit.Text, Len(VariacaoKit.Text) - 1))
        
        Else
        
            dVariacao = StrParaDbl(VariacaoKit.Text)
        
        End If
        
        'Calcula Variacao retornando o novo Preco Unitario
        lErro = Calcula_Variacao(StrParaDbl(GridKit.TextMatrix(GridKit.Row, iGrid_CustoUnitCalculadoKit_Col)), dPrecoUnitarioInformado, dVariacao, iGrid_CustoUnitKit_Col)
        If lErro <> SUCESSO Then gError 134393
        
        'Altera o Preco Unitario no grid
        If dPrecoUnitarioInformado <> 0 Then
           GridKit.TextMatrix(GridKit.Row, iGrid_CustoUnitKit_Col) = Format(dPrecoUnitarioInformado, gobjFAT.sFormatoPrecoUnitario)
        Else
           GridKit.TextMatrix(GridKit.Row, iGrid_CustoUnitKit_Col) = ""
        End If
        
        'se não tem Variacao ... Limpa no grid
        If dVariacao = 0 Then
           VariacaoKit.Text = ""
        End If
        
        'Calcula o Custo Total do Item
        dCustoTotalKit = StrParaDbl(GridKit.TextMatrix(GridKit.Row, iGrid_QuantidadeProdKit_Col)) * dPrecoUnitarioInformado
        
        'Calcula a diferença a ajustar
        dAjusteCustoKit = dCustoTotalKit - StrParaDbl(GridKit.TextMatrix(GridKit.Row, iGrid_CustoTotalKit_Col))
        
        'Altera o Custo Total no grid
        GridKit.TextMatrix(GridKit.Row, iGrid_CustoTotalKit_Col) = Format(dCustoTotalKit, "Standard")
        
        'Altera a Coleção
        lErro = AjustaCustoNecessidade_InsumosKit(CStr(dPrecoUnitarioInformado), iGrid_CustoUnitKit_Col, dAjusteCustoKit)
        If lErro <> SUCESSO Then gError 134200
        
        iGridKitAlterado = 0
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134394

    Saida_Celula_VariacaoKit = SUCESSO

    Exit Function

Erro_Saida_Celula_VariacaoKit:

    Saida_Celula_VariacaoKit = gErr

    Select Case gErr
        
        Case 134393, 134394
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158452)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ObservacaoKit(objGridInt As AdmGrid) As Long
'Faz a crítica da célula ObservacaoKit do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ObservacaoKit

    Set objGridInt.objControle = ObservacaoKit

    If iGridKitAlterado = REGISTRO_ALTERADO Then
    
        'Altera a Coleção
        lErro = AjustaCustoNecessidade_InsumosKit(ObservacaoKit.Text, iGrid_ObservacaoKit_Col)
        If lErro <> SUCESSO Then gError 134200
            
        iGridKitAlterado = 0
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134394

    Saida_Celula_ObservacaoKit = SUCESSO

    Exit Function

Erro_Saida_Celula_ObservacaoKit:

    Saida_Celula_ObservacaoKit = gErr

    Select Case gErr
        
        Case 134394
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158453)

    End Select

    Exit Function

End Function

Private Sub GridMaquinas_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGridMaquinas, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGridMaquinas, iAlterado)
        End If

End Sub

Private Sub GridMaquinas_GotFocus()
    
    Call Grid_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub GridMaquinas_EnterCell()

    Call Grid_Entrada_Celula(objGridMaquinas, iAlterado)

End Sub

Private Sub GridMaquinas_LeaveCell()
    
    Call Saida_Celula(objGridMaquinas)

End Sub

Private Sub GridMaquinas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer


    Call Grid_Trata_Tecla(KeyAscii, objGridMaquinas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMaquinas, iAlterado)
    End If

End Sub

Private Sub GridMaquinas_RowColChange()

    Call Grid_RowColChange(objGridMaquinas)

End Sub

Private Sub GridMaquinas_Scroll()

    Call Grid_Scroll(objGridMaquinas)

End Sub

Private Function Inicializa_GridMaquinas(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("U.M.")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("Unit.Calc.")
    objGrid.colColuna.Add ("Unitário")
    objGrid.colColuna.Add ("Variação")
    objGrid.colColuna.Add ("Custo Total")
    objGrid.colColuna.Add ("Observação")

    'Controles que participam do Grid
    objGrid.colCampo.Add (ProdutoMaq.Name)
    objGrid.colCampo.Add (DescricaoProdMaq.Name)
    objGrid.colCampo.Add (UMProdMaq.Name)
    objGrid.colCampo.Add (QuantidadeProdMaq.Name)
    objGrid.colCampo.Add (CustoUnitCalculadoMaq.Name)
    objGrid.colCampo.Add (CustoUnitMaq.Name)
    objGrid.colCampo.Add (VariacaoMaq.Name)
    objGrid.colCampo.Add (CustoTotalMaq.Name)
    objGrid.colCampo.Add (ObservacaoMaq.Name)

    'Colunas do Grid
    iGrid_ProdutoMaq_Col = 1
    iGrid_DescricaoProdMaq_Col = 2
    iGrid_UMProdMaq_Col = 3
    iGrid_QuantidadeProdMaq_Col = 4
    iGrid_CustoUnitCalculadoMaq_Col = 5
    iGrid_CustoUnitMaq_Col = 6
    iGrid_VariacaoMaq_Col = 7
    iGrid_CustoTotalMaq_Col = 8
    iGrid_ObservacaoMaq_Col = 9

    objGrid.objGrid = GridMaquinas

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 21

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridMaquinas.ColWidth(0) = 250
    
    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    Call Grid_Inicializa(objGrid)

    Inicializa_GridMaquinas = SUCESSO

End Function

Private Sub ProdutoMaq_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoMaq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub ProdutoMaq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Private Sub ProdutoMaq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = ProdutoMaq
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescricaoProdMaq_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescricaoProdMaq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub DescricaoProdMaq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Private Sub DescricaoProdMaq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = DescricaoProdMaq
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UMProdMaq_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UMProdMaq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub UMProdMaq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Private Sub UMProdMaq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = UMProdMaq
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantidadeProdMaq_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantidadeProdMaq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub QuantidadeProdMaq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Private Sub QuantidadeProdMaq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = QuantidadeProdMaq
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoUnitCalculadoMaq_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoUnitCalculadoMaq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub CustoUnitCalculadoMaq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Private Sub CustoUnitCalculadoMaq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = CustoUnitCalculadoMaq
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoUnitMaq_Change()

    iAlterado = REGISTRO_ALTERADO
    iGridMaqAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoUnitMaq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub CustoUnitMaq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Private Sub CustoUnitMaq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = CustoUnitMaq
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub VariacaoMaq_Change()

    iAlterado = REGISTRO_ALTERADO
    iGridMaqAlterado = REGISTRO_ALTERADO

End Sub

Private Sub VariacaoMaq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub VariacaoMaq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Private Sub VariacaoMaq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = VariacaoMaq
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoTotalMaq_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoTotalMaq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub CustoTotalMaq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Private Sub CustoTotalMaq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = CustoTotalMaq
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ObservacaoMaq_Change()

    iAlterado = REGISTRO_ALTERADO
    iGridMaqAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ObservacaoMaq_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Private Sub ObservacaoMaq_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Private Sub ObservacaoMaq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = ObservacaoMaq
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_CustoUnitMaq(objGridInt As AdmGrid) As Long
'Faz a crítica da célula CustoUnitMaq do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dVariacao As Double
Dim dCustoTotalMaq As Double
Dim dAjusteCustoMaq As Double

On Error GoTo Erro_Saida_Celula_CustoUnitMaq

    Set objGridInt.objControle = CustoUnitMaq

    'Se o campo foi preenchido
    If Len(Trim(CustoUnitMaq.Text)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(CustoUnitMaq.Text)
        If lErro <> SUCESSO Then gError 134393

    End If
    
    If iGridMaqAlterado = REGISTRO_ALTERADO Then
        
        'Calcula a Variação
        lErro = Calcula_Variacao(StrParaDbl(GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_CustoUnitCalculadoMaq_Col)), StrParaDbl(CustoUnitMaq.Text), dVariacao, iGrid_VariacaoMaq_Col)
        If lErro <> SUCESSO Then gError 134393
        
        'Altera a Variação no grid
        If dVariacao <> 0 Then
           GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_VariacaoMaq_Col) = Format(dVariacao, "Percent")
        Else
           GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_VariacaoMaq_Col) = ""
        End If
        
        'Calcula o Custo Total do Item
        dCustoTotalMaq = StrParaDbl(GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_QuantidadeProdMaq_Col)) * StrParaDbl(CustoUnitMaq.Text)
        
        'Calcula a diferença a ajustar
        dAjusteCustoMaq = dCustoTotalMaq - StrParaDbl(GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_CustoTotalMaq_Col))
        
        'Altera o Custo Total do Item no grid
        GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_CustoTotalMaq_Col) = Format(dCustoTotalMaq, "Standard")
        
        'Altera a Coleção
        lErro = AjustaCustoNecessidade_InsumosMaquina(CustoUnitMaq.Text, iGrid_CustoUnitMaq_Col, dAjusteCustoMaq)
        If lErro <> SUCESSO Then gError 134200
        
        iGridMaqAlterado = 0
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134394

    Saida_Celula_CustoUnitMaq = SUCESSO

    Exit Function

Erro_Saida_Celula_CustoUnitMaq:

    Saida_Celula_CustoUnitMaq = gErr

    Select Case gErr
        
        Case 134393, 134394
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158454)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_VariacaoMaq(objGridInt As AdmGrid) As Long
'Faz a crítica da célula VariacaoMaq do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dPrecoUnitarioInformado As Double
Dim dVariacao As Double
Dim dCustoTotalMaq As Double
Dim dAjusteCustoMaq As Double

On Error GoTo Erro_Saida_Celula_VariacaoMaq

    Set objGridInt.objControle = VariacaoMaq

    If iGridMaqAlterado = REGISTRO_ALTERADO Then
    
        'Ajusta e Converte o valor da tela
        If InStr(VariacaoMaq.Text, "%") > 0 Then
        
            dVariacao = StrParaDbl(Left(VariacaoMaq.Text, Len(VariacaoMaq.Text) - 1))
        
        Else
        
            dVariacao = StrParaDbl(VariacaoMaq.Text)
        
        End If
        
        'Calcula Variacao retornando o novo Preco Unitario
        lErro = Calcula_Variacao(StrParaDbl(GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_CustoUnitCalculadoMaq_Col)), dPrecoUnitarioInformado, dVariacao, iGrid_CustoUnitMaq_Col)
        If lErro <> SUCESSO Then gError 134393
        
        'Altera o Preco Unitario no grid
        If dPrecoUnitarioInformado <> 0 Then
           GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_CustoUnitMaq_Col) = Format(dPrecoUnitarioInformado, gobjFAT.sFormatoPrecoUnitario)
        Else
           GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_CustoUnitMaq_Col) = ""
        End If
        
        'se não tem Variacao ... Limpa no grid
        If dVariacao = 0 Then
           VariacaoMaq.Text = ""
        End If
        
        'Calcula o Custo Total do Item
        dCustoTotalMaq = StrParaDbl(GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_QuantidadeProdMaq_Col)) * dPrecoUnitarioInformado
        
        'Calcula a diferença a ajustar
        dAjusteCustoMaq = dCustoTotalMaq - StrParaDbl(GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_CustoTotalMaq_Col))
        
        'Altera o Custo Total no grid
        GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_CustoTotalMaq_Col) = Format(dCustoTotalMaq, "Standard")
        
        'Altera a Coleção
        lErro = AjustaCustoNecessidade_InsumosMaquina(CStr(dPrecoUnitarioInformado), iGrid_CustoUnitMaq_Col, dAjusteCustoMaq)
        If lErro <> SUCESSO Then gError 134200
            
        iGridMaqAlterado = 0
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134394

    Saida_Celula_VariacaoMaq = SUCESSO

    Exit Function

Erro_Saida_Celula_VariacaoMaq:

    Saida_Celula_VariacaoMaq = gErr

    Select Case gErr
        
        Case 134393, 134394
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158455)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ObservacaoMaq(objGridInt As AdmGrid) As Long
'Faz a crítica da célula ObservacaoMaq do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ObservacaoMaq

    Set objGridInt.objControle = ObservacaoMaq

    If iGridMaqAlterado = REGISTRO_ALTERADO Then
    
        'Altera a Coleção
        lErro = AjustaCustoNecessidade_InsumosMaquina(ObservacaoMaq.Text, iGrid_ObservacaoMaq_Col)
        If lErro <> SUCESSO Then gError 134200
            
        iGridMaqAlterado = 0
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134394

    Saida_Celula_ObservacaoMaq = SUCESSO

    Exit Function

Erro_Saida_Celula_ObservacaoMaq:

    Saida_Celula_ObservacaoMaq = gErr

    Select Case gErr
        
        Case 134394
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158456)

    End Select

    Exit Function

End Function

Private Sub GridMaoDeObra_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGridMaoDeObra, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGridMaoDeObra, iAlterado)
        End If

End Sub

Private Sub GridMaoDeObra_GotFocus()
    
    Call Grid_Recebe_Foco(objGridMaoDeObra)

End Sub

Private Sub GridMaoDeObra_EnterCell()

    Call Grid_Entrada_Celula(objGridMaoDeObra, iAlterado)

End Sub

Private Sub GridMaoDeObra_LeaveCell()
    
    Call Saida_Celula(objGridMaoDeObra)

End Sub

Private Sub GridMaoDeObra_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer


    Call Grid_Trata_Tecla(KeyAscii, objGridMaoDeObra, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMaoDeObra, iAlterado)
    End If

End Sub

Private Sub GridMaoDeObra_RowColChange()

    Call Grid_RowColChange(objGridMaoDeObra)

End Sub

Private Sub GridMaoDeObra_Scroll()

    Call Grid_Scroll(objGridMaoDeObra)

End Sub

Private Function Inicializa_GridMaoDeObra(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Tipo")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("U.M.")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("Unit.Calc.")
    objGrid.colColuna.Add ("Unitário")
    objGrid.colColuna.Add ("Variação")
    objGrid.colColuna.Add ("Custo Total")
    objGrid.colColuna.Add ("Observação")

    'Controles que participam do Grid
    objGrid.colCampo.Add (TipoMO.Name)
    objGrid.colCampo.Add (DescricaoMO.Name)
    objGrid.colCampo.Add (UMMO.Name)
    objGrid.colCampo.Add (QuantidadeMO.Name)
    objGrid.colCampo.Add (CustoUnitCalculadoMO.Name)
    objGrid.colCampo.Add (CustoUnitMO.Name)
    objGrid.colCampo.Add (VariacaoMO.Name)
    objGrid.colCampo.Add (CustoTotalMO.Name)
    objGrid.colCampo.Add (ObservacaoMO.Name)

    'Colunas do Grid
    iGrid_TipoMO_Col = 1
    iGrid_DescricaoMo_Col = 2
    iGrid_UMMO_Col = 3
    iGrid_QuantidadeMO_Col = 4
    iGrid_CustoUnitCalculadoMO_Col = 5
    iGrid_CustoUnitMO_Col = 6
    iGrid_VariacaoMO_Col = 7
    iGrid_CustoTotalMO_Col = 8
    iGrid_ObservacaoMO_Col = 9

    objGrid.objGrid = GridMaoDeObra

    'Todas as linhas do grid
    objGrid.objGrid.Rows = 21

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridMaoDeObra.ColWidth(0) = 250
    
    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    Call Grid_Inicializa(objGrid)

    Inicializa_GridMaoDeObra = SUCESSO

End Function

Private Sub TipoMO_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaoDeObra)

End Sub

Private Sub TipoMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaoDeObra)

End Sub

Private Sub TipoMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaoDeObra.objControle = TipoMO
    lErro = Grid_Campo_Libera_Foco(objGridMaoDeObra)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DescricaoMO_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DescricaoMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaoDeObra)

End Sub

Private Sub DescricaoMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaoDeObra)

End Sub

Private Sub DescricaoMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaoDeObra.objControle = DescricaoMO
    lErro = Grid_Campo_Libera_Foco(objGridMaoDeObra)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub UMMO_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UMMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaoDeObra)

End Sub

Private Sub UMMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaoDeObra)

End Sub

Private Sub UMMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaoDeObra.objControle = UMMO
    lErro = Grid_Campo_Libera_Foco(objGridMaoDeObra)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub QuantidadeMO_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QuantidadeMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaoDeObra)

End Sub

Private Sub QuantidadeMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaoDeObra)

End Sub

Private Sub QuantidadeMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaoDeObra.objControle = QuantidadeMO
    lErro = Grid_Campo_Libera_Foco(objGridMaoDeObra)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoUnitCalculadoMO_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoUnitCalculadoMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaoDeObra)

End Sub

Private Sub CustoUnitCalculadoMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaoDeObra)

End Sub

Private Sub CustoUnitCalculadoMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaoDeObra.objControle = CustoUnitCalculadoMO
    lErro = Grid_Campo_Libera_Foco(objGridMaoDeObra)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoUnitMO_Change()

    iAlterado = REGISTRO_ALTERADO
    iGridMOAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoUnitMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaoDeObra)

End Sub

Private Sub CustoUnitMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaoDeObra)

End Sub

Private Sub CustoUnitMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaoDeObra.objControle = CustoUnitMO
    lErro = Grid_Campo_Libera_Foco(objGridMaoDeObra)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub VariacaoMO_Change()

    iAlterado = REGISTRO_ALTERADO
    iGridMOAlterado = REGISTRO_ALTERADO

End Sub

Private Sub VariacaoMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaoDeObra)

End Sub

Private Sub VariacaoMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaoDeObra)

End Sub

Private Sub VariacaoMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaoDeObra.objControle = VariacaoMO
    lErro = Grid_Campo_Libera_Foco(objGridMaoDeObra)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub CustoTotalMO_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoTotalMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaoDeObra)

End Sub

Private Sub CustoTotalMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaoDeObra)

End Sub

Private Sub CustoTotalMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaoDeObra.objControle = CustoTotalMO
    lErro = Grid_Campo_Libera_Foco(objGridMaoDeObra)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ObservacaoMO_Change()

    iAlterado = REGISTRO_ALTERADO
    iGridMOAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ObservacaoMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaoDeObra)

End Sub

Private Sub ObservacaoMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaoDeObra)

End Sub

Private Sub ObservacaoMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaoDeObra.objControle = ObservacaoMO
    lErro = Grid_Campo_Libera_Foco(objGridMaoDeObra)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_CustoUnitMO(objGridInt As AdmGrid) As Long
'Faz a crítica da célula CustoUnitMO do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dVariacao As Double
Dim dCustoTotalMO As Double
Dim dAjusteCustoMO As Double

On Error GoTo Erro_Saida_Celula_CustoUnitMO

    Set objGridInt.objControle = CustoUnitMO

    'Se o campo foi preenchido
    If Len(Trim(CustoUnitMO.Text)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(CustoUnitMO.Text)
        If lErro <> SUCESSO Then gError 134393

    End If
    
    If iGridMOAlterado = REGISTRO_ALTERADO Then
        
        'Calcula a Variação
        lErro = Calcula_Variacao(StrParaDbl(GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_CustoUnitCalculadoMO_Col)), StrParaDbl(CustoUnitMO.Text), dVariacao, iGrid_VariacaoMO_Col)
        If lErro <> SUCESSO Then gError 134393
        
        'Altera a Variação no grid
        If dVariacao <> 0 Then
           GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_VariacaoMO_Col) = Format(dVariacao, "Percent")
        Else
           GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_VariacaoMO_Col) = ""
        End If
        
        'Calcula o Custo Total do Item
        dCustoTotalMO = StrParaDbl(GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_QuantidadeMO_Col)) * StrParaDbl(CustoUnitMO.Text)
        
        'Calcula a diferença a ajustar
        dAjusteCustoMO = dCustoTotalMO - StrParaDbl(GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_CustoTotalMO_Col))
        
        'Altera o Custo Total do Item no grid
        GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_CustoTotalMO_Col) = Format(dCustoTotalMO, "Standard")
        
        'Altera a Coleção
        lErro = AjustaCustoNecessidade_MaoDeObra(CustoUnitMO.Text, iGrid_CustoUnitMO_Col, dAjusteCustoMO)
        If lErro <> SUCESSO Then gError 134200
        
        iGridMOAlterado = 0
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134394

    Saida_Celula_CustoUnitMO = SUCESSO

    Exit Function

Erro_Saida_Celula_CustoUnitMO:

    Saida_Celula_CustoUnitMO = gErr

    Select Case gErr
        
        Case 134393, 134394
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158457)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_VariacaoMO(objGridInt As AdmGrid) As Long
'Faz a crítica da célula VariacaoMO do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dPrecoUnitarioInformado As Double
Dim dVariacao As Double
Dim dCustoTotalMO As Double
Dim dAjusteCustoMO As Double

On Error GoTo Erro_Saida_Celula_VariacaoMO

    Set objGridInt.objControle = VariacaoMO

    If iGridMOAlterado = REGISTRO_ALTERADO Then
        
        'Ajusta e Converte o valor da tela
        If InStr(VariacaoMO.Text, "%") > 0 Then
        
            dVariacao = StrParaDbl(Left(VariacaoMO.Text, Len(VariacaoMO.Text) - 1))
        
        Else
        
            dVariacao = StrParaDbl(VariacaoMO.Text)
        
        End If
        
        'Calcula Variacao retornando o novo Preco Unitario
        lErro = Calcula_Variacao(StrParaDbl(GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_CustoUnitCalculadoMO_Col)), dPrecoUnitarioInformado, dVariacao, iGrid_CustoUnitMO_Col)
        If lErro <> SUCESSO Then gError 134393
        
        'Altera o Preco Unitario no grid
        If dPrecoUnitarioInformado <> 0 Then
           GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_CustoUnitMO_Col) = Format(dPrecoUnitarioInformado, gobjFAT.sFormatoPrecoUnitario)
        Else
           GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_CustoUnitMO_Col) = ""
        End If
        
        'se não tem Variacao ... Limpa no grid
        If dVariacao = 0 Then
           VariacaoMO.Text = ""
        End If
        
        'Calcula o Custo Total do Item
        dCustoTotalMO = StrParaDbl(GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_QuantidadeMO_Col)) * dPrecoUnitarioInformado
        
        'Calcula a diferença a ajustar
        dAjusteCustoMO = dCustoTotalMO - StrParaDbl(GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_CustoTotalMO_Col))
        
        'Altera o Custo Total no grid
        GridMaoDeObra.TextMatrix(GridMaoDeObra.Row, iGrid_CustoTotalMO_Col) = Format(dCustoTotalMO, "Standard")
        
        'Altera a Coleção
        lErro = AjustaCustoNecessidade_MaoDeObra(CStr(dPrecoUnitarioInformado), iGrid_CustoUnitMO_Col, dAjusteCustoMO)
        If lErro <> SUCESSO Then gError 134200
        
        iGridMOAlterado = 0
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134394

    Saida_Celula_VariacaoMO = SUCESSO

    Exit Function

Erro_Saida_Celula_VariacaoMO:

    Saida_Celula_VariacaoMO = gErr

    Select Case gErr
        
        Case 134393, 134394
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158458)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ObservacaoMO(objGridInt As AdmGrid) As Long
'Faz a crítica da célula ObservacaoMO do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ObservacaoMO

    Set objGridInt.objControle = ObservacaoMO

    If iGridMOAlterado = REGISTRO_ALTERADO Then
    
        'Altera a Coleção
        lErro = AjustaCustoNecessidade_MaoDeObra(ObservacaoMO.Text, iGrid_ObservacaoMO_Col)
        If lErro <> SUCESSO Then gError 134200
            
        iGridMOAlterado = 0
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 134394

    Saida_Celula_ObservacaoMO = SUCESSO

    Exit Function

Erro_Saida_Celula_ObservacaoMO:

    Saida_Celula_ObservacaoMO = gErr

    Select Case gErr
        
        Case 134394
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158459)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objProjetoCusteio As New ClassProjetoCusteio

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "ProjetoCusteio"

    'Lê os dados da Tela Custeio de Projeto
    lErro = Move_Tela_Memoria(objProjetoCusteio)
    If lErro <> SUCESSO Then gError 134722

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDocProj", objProjetoCusteio.lNumIntDocProj, 0, "NumIntDocProj"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 134722

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158460)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objProjetoCusteio As New ClassProjetoCusteio

On Error GoTo Erro_Tela_Preenche

    objProjetoCusteio.lNumIntDocProj = colCampoValor.Item("NumIntDocProj").vValor

    If objProjetoCusteio.lNumIntDocProj <> 0 Then
        lErro = Traz_ProjetoCusteio_Tela(objProjetoCusteio)
        If lErro <> SUCESSO Then gError 134723
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 134723

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158461)

    End Select

    Exit Function

End Function

Function Limpa_Tela_CusteioProjeto() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_CusteioProjeto
        
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    'Limpa a Descrição do Projeto
    LabelDescProjeto.Caption = ""
    
    'Limpa o Custo e Preço Total do Projeto
    CustoTotalProjeto.Caption = ""
    PrecoTotalProjeto.Caption = ""
    
    'Limpa o Grid de Itens do Projeto
    Call Grid_Limpa(objGridItens)
    
    'Limpa a Coleção de Itens
    Set colProjetoCusteioItens = New Collection
    
    'Limpa os demais Grids
    Call Trata_LinhaNaoSelecionada
    
    'Coloca a DataAtual como Data do Novo Custeio
    DataCusteio.PromptInclude = False
    DataCusteio.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataCusteio.PromptInclude = True
    
    iAlterado = 0
    iProjetoAlterado = 0
    iGridKitAlterado = 0
    iGridMaqAlterado = 0
    iGridMOAlterado = 0
        
    Limpa_Tela_CusteioProjeto = SUCESSO

    Exit Function

Erro_Limpa_Tela_CusteioProjeto:

    Limpa_Tela_CusteioProjeto = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158462)

    End Select

    Exit Function

End Function
Function Trata_LinhaNaoSelecionada() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_LinhaNaoSelecionada
    
    'Limpa Tab Insumos do Kit
    LabelProdutoInsumosKit.Caption = ""
    LabelVersaoInsumosKit.Caption = ""
    
    Call Grid_Limpa(objGridKit)
    
    CustoTotalInsumosKit.Caption = ""

    'Limpa Tab Insumos da Maquina
    LabelProdutoInsumosMaq.Caption = ""
    LabelVersaoInsumosMaq.Caption = ""
    
    Call Grid_Limpa(objGridMaquinas)
    
    CustoTotalInsumosMaquina.Caption = ""

    'Limpa Tab Mão-de-Obra
    LabelProdutoMO.Caption = ""
    LabelVersaoMO.Caption = ""
    
    Call Grid_Limpa(objGridMaoDeObra)
    
    CustoTotalMaoDeObra.Caption = ""

    Trata_LinhaNaoSelecionada = SUCESSO
    
    Exit Function

Erro_Trata_LinhaNaoSelecionada:

    Trata_LinhaNaoSelecionada = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158463)
    
    End Select
    
    Exit Function

End Function

Function Trata_LinhaSelecionada(ByVal sProduto As String, ByVal sVersao As String, ByVal iTabSelecionada As Integer) As Long

Dim lErro As Long
Dim objProjetoCusteioItens As ClassProjetoCusteioItens

On Error GoTo Erro_Trata_LinhaSelecionada

    Set objProjetoCusteioItens = New ClassProjetoCusteioItens

    'Localiza o item de custeio na coleção
    lErro = Localiza_ProjetoCusteioItens(sProduto, sVersao, objProjetoCusteioItens)
    If lErro <> SUCESSO Then gError 134702

    'Trata o Grid da Tab Selecionada
    Select Case iTabSelecionada
    
        Case Is = TAB_InsumosKit
            
            Call Trata_GridInsumosKit(objProjetoCusteioItens)
        
        Case Is = TAB_InsumosMaq
            
            Call Trata_GridInsumosMaquina(objProjetoCusteioItens)
        
        Case Is = TAB_MaoDeObra
        
            Call Trata_GridMaoDeObra(objProjetoCusteioItens)
            
    End Select
        
    Trata_LinhaSelecionada = SUCESSO
    
    Exit Function

Erro_Trata_LinhaSelecionada:

    Trata_LinhaSelecionada = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158464)
    
    End Select
    
    Exit Function

End Function

Function Trata_GridInsumosKit(ByVal objProjetoCusteioItens As ClassProjetoCusteioItens) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_GridInsumosKit
    
    'preenche labels de produto e versão
    LabelProdutoInsumosKit.Caption = Trim(GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)) & _
                                    SEPARADOR & GridItens.TextMatrix(GridItens.Row, iGrid_DescricaoProdItens_Col)
    LabelVersaoInsumosKit.Caption = GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col)
    
    'Limpa o Grid de Kits
    Call Grid_Limpa(objGridKit)
    
    'preenche GridInsumosKit
    lErro = Preenche_GridInsumosKit(objProjetoCusteioItens.colProjetoInsumosKit)
    If lErro <> SUCESSO Then gError 137999
    
    'Exibe o CustoTotal
    CustoTotalInsumosKit.Caption = Format(objProjetoCusteioItens.dCustoTotalInsumosKit, "Standard")
    
    Trata_GridInsumosKit = SUCESSO
    
    Exit Function

Erro_Trata_GridInsumosKit:

    Trata_GridInsumosKit = gErr
    
    Select Case gErr
    
        Case 137999
            'erro tratado na rotina chamada
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158465)
    
    End Select
    
    Exit Function

End Function

Function Trata_GridMaoDeObra(ByVal objProjetoCusteioItens As ClassProjetoCusteioItens) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_GridMaoDeObra
    
    'preenche labels de produto e versão
    LabelProdutoMO.Caption = Trim(GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)) & _
                            SEPARADOR & GridItens.TextMatrix(GridItens.Row, iGrid_DescricaoProdItens_Col)
    LabelVersaoMO.Caption = GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col)
    
    'Limpa o Grid de MaoDeObra
    Call Grid_Limpa(objGridMaoDeObra)
    
    'preenche GridMaoDeObra
    lErro = Preenche_GridMaoDeObra(objProjetoCusteioItens.colProjetoMaoDeObra)
    If lErro <> SUCESSO Then gError 137999
    
    'Exibe o CustoTotal
    CustoTotalMaoDeObra.Caption = Format(objProjetoCusteioItens.dCustoTotalMaoDeObra, "Standard")
    
    'preenche GridMaoDeObra
        
    Trata_GridMaoDeObra = SUCESSO
    
    Exit Function

Erro_Trata_GridMaoDeObra:

    Trata_GridMaoDeObra = gErr
    
    Select Case gErr
        
        Case 137999
            'erro tratado na rotina chamada
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158466)
    
    End Select
    
    Exit Function

End Function

Function Trata_GridInsumosMaquina(ByVal objProjetoCusteioItens As ClassProjetoCusteioItens) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_GridInsumosMaquina
    
    'preenche labels de produto e versão
    LabelProdutoInsumosMaq.Caption = Trim(GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)) & _
                                    SEPARADOR & GridItens.TextMatrix(GridItens.Row, iGrid_DescricaoProdItens_Col)
    LabelVersaoInsumosMaq.Caption = GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col)
    
    'Limpa o Grid de Maquinas
    Call Grid_Limpa(objGridMaquinas)
    
    'preenche GridInsumosMaquinas
    lErro = Preenche_GridInsumosMaquinas(objProjetoCusteioItens.colProjetoInsumosMaquina)
    If lErro <> SUCESSO Then gError 137999
    
    'Exibe o CustoTotal
    CustoTotalInsumosMaquina.Caption = Format(objProjetoCusteioItens.dCustoTotalInsumosMaq, "Standard")
        
    Trata_GridInsumosMaquina = SUCESSO
    
    Exit Function

Erro_Trata_GridInsumosMaquina:

    Trata_GridInsumosMaquina = gErr
    
    Select Case gErr
    
        Case 137999
            'erro tratado na rotina chamada
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158467)
    
    End Select
    
    Exit Function

End Function

Function CalculaNecessidade_InsumosKit(ByVal objProjetoItens As ClassProjetoItens, ByVal objProjetoCusteioItens As ClassProjetoCusteioItens, iSeq As Integer, dCustoTotalInsumosKit As Double) As Long

Dim lErro As Long
Dim objKit As ClassKit
Dim objProdutoKit As New ClassProdutoKit
Dim objKitIntermediario As ClassKit
Dim sVersaoKit As String
Dim bProdutoNovo As Boolean
Dim objProdutoRaiz As New ClassProduto
Dim objProduto As New ClassProduto
Dim dFatorUMProduto As Double
Dim dFatorUMProjeto As Double
Dim dQuantidade As Double
Dim objProjetoItensIntermediario As ClassProjetoItens
Dim dCustoProduto As Double
Dim objProjetoInsumosKit As ClassProjetoInsumosKit
Dim objAuxProjInsumosKit As New ClassProjetoInsumosKit

On Error GoTo Erro_CalculaNecessidade_InsumosKit
    
    Set objProdutoRaiz = New ClassProduto
   
    objProdutoRaiz.sCodigo = objProjetoItens.sProduto
    
    'Lê o produto para descobrir as unidades de medidas associadas
    lErro = CF("Produto_Le", objProdutoRaiz)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 124205
   
    'Descobre o fator de conversao da UM de ProjetoItens p/UM de estoque do produto
    lErro = CF("UM_Conversao_Trans", objProdutoRaiz.iClasseUM, objProjetoItens.sUMedida, objProdutoRaiz.sSiglaUMEstoque, dFatorUMProjeto)
    If lErro <> SUCESSO Then gError 134976
   
    Set objKit = New ClassKit

    objKit.sProdutoRaiz = objProjetoItens.sProduto
    objKit.sVersao = objProjetoItens.sVersao
    
    'Leio os ProdutosKits que compõem este Kit
    lErro = CF("Kit_Le_Componentes", objKit)
    If lErro <> SUCESSO And lErro <> 21831 Then gError 134703
            
    'para cada produto componente do kit ...
    For Each objProdutoKit In objKit.colComponentes

        'se não é o produto raiz ...
        If objProdutoKit.iNivel <> KIT_NIVEL_RAIZ Then
        
            Set objProduto = New ClassProduto
            
            objProduto.sCodigo = objProdutoKit.sProduto
        
            'se tem composição variável ...
            If objProdutoKit.iComposicao = PRODUTOKIT_COMPOSICAO_VARIAVEL Then
            
                'Lê o produto para descobrir as unidades de medidas associadas
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 124205
            
                'Descobre o fator de conversao da UM do Kit p/UM de estoque do produto
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProdutoKit.sUnidadeMed, objProduto.sSiglaUMEstoque, dFatorUMProduto)
                If lErro <> SUCESSO Then gError 134975
                            
                dQuantidade = (objProdutoKit.dQuantidade * dFatorUMProduto) * (objProjetoItens.dQuantidade * dFatorUMProjeto) / dFatorUMProduto
                                
            Else
                
                'senão... usa a quantidade fixa
                dQuantidade = objProdutoKit.dQuantidade
            
            End If
            
            Set objKitIntermediario = New ClassKit
            
            objKitIntermediario.sProdutoRaiz = objProdutoKit.sProduto
        
            'Le as Versoes Ativas e a Padrao
            lErro = CF("Kit_Le_Padrao", objKitIntermediario)
            If lErro <> SUCESSO And lErro <> 106304 Then gError 134806

            'se encontrou - é outro Kit (Produto Intermediário)
            If lErro = SUCESSO Then
            
                'se a versão está preenchida -> usa ela
                If Len(objProdutoKit.sVersaoKitComp) <> 0 Then
                
                    sVersaoKit = objProdutoKit.sVersaoKitComp
                    
                Else
                    
                    'senão usa a Padrão
                    sVersaoKit = objKitIntermediario.sVersao
                    
                End If
                
                'Cria um novo objProjetoItens
                Set objProjetoItensIntermediario = New ClassProjetoItens
                
                objProjetoItensIntermediario.sProduto = objProdutoKit.sProduto
                objProjetoItensIntermediario.sVersao = sVersaoKit
                objProjetoItensIntermediario.sUMedida = objProdutoKit.sUnidadeMed
                objProjetoItensIntermediario.dQuantidade = dQuantidade
                
                'e chama esta função recursivamente ...
                lErro = CalculaNecessidade_InsumosKit(objProjetoItensIntermediario, objProjetoCusteioItens, iSeq, dCustoTotalInsumosKit)
                If lErro <> SUCESSO Then gError 134201
                
            Else
                
                'senão... vamos incluir na coleção...
                bProdutoNovo = True
                
                'Verifica se já há algum produto do kit na coleção
                For Each objAuxProjInsumosKit In objProjetoCusteioItens.colProjetoInsumosKit
            
                    'se encontrou ...
                    If objAuxProjInsumosKit.sProduto = objProdutoKit.sProduto Then
                        
                        'subtrai o valor anterior do Total do Custo no acumulador
                        dCustoTotalInsumosKit = dCustoTotalInsumosKit - (objAuxProjInsumosKit.dQuantidade * objAuxProjInsumosKit.dCustoUnitarioInformado)
                        
                        'acumula a quantidade
                        objAuxProjInsumosKit.dQuantidade = objAuxProjInsumosKit.dQuantidade + dQuantidade
                        
                        'lança o novo valor Total do Custo no acumulador
                        dCustoTotalInsumosKit = dCustoTotalInsumosKit + (objAuxProjInsumosKit.dQuantidade * objAuxProjInsumosKit.dCustoUnitarioInformado)
                        
                        'avisa que acumulou ...
                        bProdutoNovo = False
                        Exit For
                        
                    End If
                    
                Next
                
                'se não tem o produto ...
                If bProdutoNovo Then
                    
                    'Descobre qual seu custo
                    lErro = CF("Produto_Le_CustoProduto", objProduto, dCustoProduto)
                    If lErro <> SUCESSO Then gError 134202
                                                
                    'incrementa o sequencial
                    iSeq = iSeq + 1
                    
                    'reCria o obj
                    Set objProjetoInsumosKit = New ClassProjetoInsumosKit
                    
                    objProjetoInsumosKit.iSeq = iSeq
                    objProjetoInsumosKit.sProduto = objProdutoKit.sProduto
                    objProjetoInsumosKit.sUMedida = objProdutoKit.sUnidadeMed
                    objProjetoInsumosKit.dQuantidade = dQuantidade
                    objProjetoInsumosKit.dCustoUnitarioCalculado = dCustoProduto
                    objProjetoInsumosKit.dCustoUnitarioInformado = dCustoProduto
                    
                    'lança o valor Total do Custo no acumulador
                    dCustoTotalInsumosKit = dCustoTotalInsumosKit + (dQuantidade * dCustoProduto)
                    
                    'e inclui na coleção.
                    objProjetoCusteioItens.colProjetoInsumosKit.Add objProjetoInsumosKit, "X" & Right$(CStr(100000 + iSeq), 5)
                    
                End If
            
            End If
            
        End If
    
    Next
    
    CalculaNecessidade_InsumosKit = SUCESSO
    
    Exit Function

Erro_CalculaNecessidade_InsumosKit:

    CalculaNecessidade_InsumosKit = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158468)
    
    End Select
    
    Exit Function

End Function

Function CalculaNecessidade_InsumosMaquina(ByVal objProjetoItens As ClassProjetoItens, ByVal objProjetoCusteioItens As ClassProjetoCusteioItens, iSeq As Integer, dCustoTotalInsumosMaq As Double) As Long

Dim lErro As Long
Dim objProduto As ClassProduto
Dim dFator As Double
Dim dFatorQuantidade As Double
Dim objRoteirosDeFabricacao As ClassRoteirosDeFabricacao
Dim objOperacoes As New ClassOperacoes
Dim objPO As ClassPlanoOperacional
Dim objOPOperacoes As ClassOrdemProducaoOperacoes
Dim objPOMaquinas As New ClassPOMaquinas
Dim objMaquinas As ClassMaquinas
Dim objMaquinasInsumos As New ClassMaquinasInsumos
Dim objOperacaoInsumos As New ClassOperacaoInsumos
Dim objProjetoItensIntermediario As ClassProjetoItens
Dim objPMPItens As ClassPMPItens

On Error GoTo Erro_CalculaNecessidade_InsumosMaquina

    Set objRoteirosDeFabricacao = New ClassRoteirosDeFabricacao
    
    objRoteirosDeFabricacao.sProdutoRaiz = objProjetoItens.sProduto
    objRoteirosDeFabricacao.sVersao = objProjetoItens.sVersao
    
    lErro = CF("RoteirosDeFabricacao_Le", objRoteirosDeFabricacao)
    If lErro <> SUCESSO And lErro <> 134617 Then gError 134200
    
    If lErro = SUCESSO Then
    
        Set objProduto = New ClassProduto
       
        objProduto.sCodigo = objProjetoItens.sProduto
    
        'Lê o produto para descobrir as unidades de medidas associadas
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 124205
       
        'Descobre o fator de conversao da UM de ProjetoItens p/UM de RoteirosDeFabricacao
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProjetoItens.sUMedida, objRoteirosDeFabricacao.sUM, dFator)
        If lErro <> SUCESSO Then gError 134976

        dFatorQuantidade = (objProjetoItens.dQuantidade * dFator) / objRoteirosDeFabricacao.dQuantidade
        
        Set objPO = New ClassPlanoOperacional
        
        objPO.sProduto = objProjetoItens.sProduto
        objPO.dQuantidade = objProjetoItens.dQuantidade
        objPO.sUM = objProjetoItens.sUMedida
        
        For Each objOperacoes In objRoteirosDeFabricacao.colOperacoes
        
            Set objOPOperacoes = New ClassOrdemProducaoOperacoes
            
            objOPOperacoes.lNumIntDocCT = objOperacoes.lNumIntDocCT
            objOPOperacoes.lNumIntDocCompet = objOperacoes.lNumIntDocCompet
            objOPOperacoes.iConsideraCarga = MARCADO
            
            Set objOPOperacoes.objOperacoesTempo = objOperacoes.objOperacoesTempo
            
            Set objPMPItens = New ClassPMPItens
                        
            lErro = CF("PlanoOperacional_Calcula_Tempos", objPMPItens, objPO, objOPOperacoes, MRP_ACERTA_POR_DATA_FIM)
            If lErro <> SUCESSO Then gError 134201
            
            For Each objPOMaquinas In objPO.colAlocacaoMaquinas
            
                Set objMaquinas = New ClassMaquinas
            
                objMaquinas.lNumIntDoc = objPOMaquinas.lNumIntDocMaq
                 
                lErro = CF("Maquinas_Le_Itens", objMaquinas)
                If lErro <> SUCESSO Then gError 134202
                
                For Each objMaquinasInsumos In objMaquinas.colProdutos
                
                    lErro = IncluiNecessidade_InsumosMaquina(objMaquinasInsumos, objPOMaquinas, objProjetoCusteioItens, iSeq, dCustoTotalInsumosMaq)
                    If lErro <> SUCESSO Then gError 134203
                
                Next
            
            Next
            
            For Each objOperacaoInsumos In objOperacoes.colOperacaoInsumos
            
                Set objProjetoItensIntermediario = New ClassProjetoItens
                
                objProjetoItensIntermediario.sProduto = objOperacaoInsumos.sProduto
                objProjetoItensIntermediario.sVersao = objOperacaoInsumos.sVersaoKitComp
                objProjetoItensIntermediario.sUMedida = objOperacaoInsumos.sUMProduto
                objProjetoItensIntermediario.dQuantidade = objOperacaoInsumos.dQuantidade * dFatorQuantidade
                
                lErro = CalculaNecessidade_InsumosMaquina(objProjetoItensIntermediario, objProjetoCusteioItens, iSeq, dCustoTotalInsumosMaq)
                If lErro <> SUCESSO Then gError 134204
                
            Next
        
        Next
    
    End If
    
    CalculaNecessidade_InsumosMaquina = SUCESSO
    
    Exit Function

Erro_CalculaNecessidade_InsumosMaquina:

    CalculaNecessidade_InsumosMaquina = gErr
    
    Select Case gErr
    
        Case 134200, 134201, 134202, 134203, 134204
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158469)
    
    End Select
    
    Exit Function

End Function

Function CalculaNecessidade_MaoDeObra(ByVal objProjetoItens As ClassProjetoItens, ByVal objProjetoCusteioItens As ClassProjetoCusteioItens, iSeq As Integer, dCustoTotalMaoDeObra As Double) As Long

Dim lErro As Long
Dim objProduto As ClassProduto
Dim dFator As Double
Dim dFatorQuantidade As Double
Dim objRoteirosDeFabricacao As ClassRoteirosDeFabricacao
Dim objOperacoes As New ClassOperacoes
Dim objPO As ClassPlanoOperacional
Dim objOPOperacoes As ClassOrdemProducaoOperacoes
Dim objPOMaquinas As New ClassPOMaquinas
Dim objMaquinas As ClassMaquinas
Dim objMaquinaOperadores As New ClassMaquinaOperadores
Dim objOperacaoInsumos As New ClassOperacaoInsumos
Dim objProjetoItensIntermediario As ClassProjetoItens
Dim objPMPItens As ClassPMPItens

On Error GoTo Erro_CalculaNecessidade_MaoDeObra

    Set objRoteirosDeFabricacao = New ClassRoteirosDeFabricacao

    objRoteirosDeFabricacao.sProdutoRaiz = objProjetoItens.sProduto
    objRoteirosDeFabricacao.sVersao = objProjetoItens.sVersao

    lErro = CF("RoteirosDeFabricacao_Le", objRoteirosDeFabricacao)
    If lErro <> SUCESSO And lErro <> 134617 Then gError 134200

    If lErro = SUCESSO Then
    
        Set objProduto = New ClassProduto
       
        objProduto.sCodigo = objProjetoItens.sProduto
    
        'Lê o produto para descobrir as unidades de medidas associadas
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 124205
       
        'Descobre o fator de conversao da UM de ProjetoItens p/UM de RoteiroDeFabricacao
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProjetoItens.sUMedida, objRoteirosDeFabricacao.sUM, dFator)
        If lErro <> SUCESSO Then gError 134976
    
        dFatorQuantidade = (objProjetoItens.dQuantidade * dFator) / objRoteirosDeFabricacao.dQuantidade
        
        Set objPO = New ClassPlanoOperacional

        objPO.sProduto = objProjetoItens.sProduto
        objPO.dQuantidade = objProjetoItens.dQuantidade
        objPO.sUM = objProjetoItens.sUMedida

        For Each objOperacoes In objRoteirosDeFabricacao.colOperacoes

            Set objOPOperacoes = New ClassOrdemProducaoOperacoes

            objOPOperacoes.lNumIntDocCT = objOperacoes.lNumIntDocCT
            objOPOperacoes.lNumIntDocCompet = objOperacoes.lNumIntDocCompet
            objOPOperacoes.iConsideraCarga = MARCADO

            Set objOPOperacoes.objOperacoesTempo = objOperacoes.objOperacoesTempo

            Set objPMPItens = New ClassPMPItens

            lErro = CF("PlanoOperacional_Calcula_Tempos", objPMPItens, objPO, objOPOperacoes, MRP_ACERTA_POR_DATA_FIM)
            If lErro <> SUCESSO Then gError 134201

            For Each objPOMaquinas In objPO.colAlocacaoMaquinas

                Set objMaquinas = New ClassMaquinas

                objMaquinas.lNumIntDoc = objPOMaquinas.lNumIntDocMaq

                lErro = CF("Maquinas_Le_Itens", objMaquinas)
                If lErro <> SUCESSO Then gError 134202

                For Each objMaquinaOperadores In objMaquinas.colTipoOperadores

                    lErro = IncluiNecessidade_MaoDeObra(objMaquinaOperadores, objPOMaquinas, objProjetoCusteioItens, iSeq, dCustoTotalMaoDeObra)
                    If lErro <> SUCESSO Then gError 134203

                Next

            Next

            For Each objOperacaoInsumos In objOperacoes.colOperacaoInsumos

                Set objProjetoItensIntermediario = New ClassProjetoItens

                objProjetoItensIntermediario.sProduto = objOperacaoInsumos.sProduto
                objProjetoItensIntermediario.sVersao = objOperacaoInsumos.sVersaoKitComp
                objProjetoItensIntermediario.sUMedida = objOperacaoInsumos.sUMProduto
                objProjetoItensIntermediario.dQuantidade = objOperacaoInsumos.dQuantidade * dFatorQuantidade

                lErro = CalculaNecessidade_MaoDeObra(objProjetoItensIntermediario, objProjetoCusteioItens, iSeq, dCustoTotalMaoDeObra)
                If lErro <> SUCESSO Then gError 134204

            Next

        Next

    End If

    CalculaNecessidade_MaoDeObra = SUCESSO

    Exit Function

Erro_CalculaNecessidade_MaoDeObra:

    CalculaNecessidade_MaoDeObra = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158470)

    End Select

    Exit Function

End Function

Function Traz_ProjetoCusteio_Tela(objProjetoCusteio As ClassProjetoCusteio) As Long

Dim lErro As Long
Dim objProjetoCusteioItens As New ClassProjetoCusteioItens
Dim objProjeto As ClassProjeto
Dim objProjetoItens As New ClassProjetoItens

On Error GoTo Erro_Traz_ProjetoCusteio_Tela

    'Limpa o Grid de Itens do Projeto
    Call Grid_Limpa(objGridItens)
    
    'Limpa a Coleção de Itens
    Set colProjetoCusteioItens = New Collection

    Set objProjeto = New ClassProjeto
    
    objProjeto.lNumIntDoc = objProjetoCusteio.lNumIntDocProj
    
    'Le os Itens do Projeto
    lErro = CF("Projeto_Le_Itens", objProjeto)
    If lErro <> SUCESSO And lErro <> 134453 Then gError 134096
    
    'Se projeto não tem itens -> Erro
    If lErro <> SUCESSO Then gError 134200

    'Verifica se o ProjetoCusteio existe, lendo no BD a partir do NumIntDocProj
    lErro = CF("ProjetoCusteio_Le_NumIntDocProj", objProjetoCusteio)
    If lErro <> SUCESSO And lErro <> 134449 Then gError 134094
    
    'se existe ...
    If lErro = SUCESSO Then
    
        'Exibe Data do Custeio na Tela
        If objProjetoCusteio.dtDataCusteio <> DATA_NULA Then
            DataCusteio.PromptInclude = False
            DataCusteio.Text = Format(objProjetoCusteio.dtDataCusteio, "dd/mm/yy")
            DataCusteio.PromptInclude = True
        End If
                
        'Exibe Custo e Preço Total do Projeto na Tela
        If objProjetoCusteio.dCustoTotalProjeto > 0 Then
        
            CustoTotalProjeto.Caption = Format(objProjetoCusteio.dCustoTotalProjeto, "#0.00")
        
        End If
        
        If objProjetoCusteio.dPrecoTotalProjeto > 0 Then
        
            CustoTotalProjeto.Caption = Format(objProjetoCusteio.dPrecoTotalProjeto, "#0.00")
        
        End If
        
        'Lê os Itens de ProjetoCusteio trazendo a coleção do obj preenchida
        lErro = CF("ProjetoCusteio_Le_Itens", objProjetoCusteio)
        If lErro <> SUCESSO Then gError 134200
        
        'Para cada item da coleção ...
        For Each objProjetoCusteioItens In objProjetoCusteio.colProjetoCusteioItens
        
            'adiciona o obj à coleção de itens
            colProjetoCusteioItens.Add objProjetoCusteioItens
        
        Next
        
    Else
                
        'Para cada item da coleção de ProjetoItens
        For Each objProjetoItens In objProjeto.colProjetoItens
        
            Set objProjetoCusteioItens = New ClassProjetoCusteioItens
        
            objProjetoCusteioItens.lNumIntDocProjetoItem = objProjetoItens.lNumIntDoc
            
            'Calcula as necessidades de Produção do Item
            lErro = CalculaNecessidadesProducao(objProjetoItens, objProjetoCusteioItens)
            If lErro <> SUCESSO Then gError 134200
            
            'adiciona o obj à coleção de itens
            colProjetoCusteioItens.Add objProjetoCusteioItens
        
        Next
        
    End If

    'Preenche o Grid de Itens do Projeto
    lErro = Preenche_GridItens(objProjeto)
    If lErro <> SUCESSO Then gError 134200
        
    iAlterado = 0
    iProjetoAlterado = 0
    iGridKitAlterado = 0
    iGridMaqAlterado = 0
    iGridMOAlterado = 0
            
    Traz_ProjetoCusteio_Tela = SUCESSO

    Exit Function

Erro_Traz_ProjetoCusteio_Tela:

    Traz_ProjetoCusteio_Tela = gErr

    Select Case gErr

        Case 134094, 134096, 134098, 134099, 134100, 134655
            'Erros tratados nas rotinas chamadas
        
        Case 134095, 134097 '134095 = Não encontrou por código; 134097 = Não encontrou por NomeReduzido
            'Erros tratados na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158471)

    End Select

    Exit Function

End Function

Function Trata_Linha(ByVal iTabSelecionada As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim bTemSelecao As Boolean
Dim sProduto As String
Dim sVersao As String

On Error GoTo Erro_Trata_Linha

    'se a linha do grid está selecionada ...
    If Val(GridItens.TextMatrix(GridItens.Row, iGrid_Selecionar_Col)) = MARCADO Then
        
        sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
        sVersao = GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col)
        
        lErro = Trata_LinhaSelecionada(sProduto, sVersao, iTabSelecionada)
        If lErro <> SUCESSO Then gError 137999
    
    Else
    
        'senão
        bTemSelecao = False
        
        'verifica se tem alguma linha selecionada
        For iIndice = 1 To objGridItens.iLinhasExistentes
        
            If Val(GridItens.TextMatrix(iIndice, iGrid_Selecionar_Col)) = MARCADO Then
                bTemSelecao = True
                Exit For
            End If
    
        Next iIndice
        
        'encontrou uma linha selecionada ...
        If bTemSelecao Then
        
            'posiciona na linha selecionada
            GridItens.Row = iIndice
            
            sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
            sVersao = GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col)
            
            lErro = Trata_LinhaSelecionada(sProduto, sVersao, iTabSelecionada)
            If lErro <> SUCESSO Then gError 137998
        
        Else
            
            'senão -> limpa tudo
            Call Trata_LinhaNaoSelecionada
        
        End If
    
    End If
                
    Trata_Linha = SUCESSO
    
    Exit Function
            
Erro_Trata_Linha:

    Trata_Linha = gErr
    
    Select Case gErr
    
        Case 137999, 137998
            'erros tratados na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158472)

    End Select

    Exit Function

End Function

Function Preenche_GridInsumosKit(ByVal colProjetoInsumosKit As Collection) As Long

Dim lErro As Long
Dim objProjetoInsumosKit As New ClassProjetoInsumosKit
Dim objProduto As New ClassProduto
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim iLinha As Integer
Dim dVariacao As Double

On Error GoTo Erro_Preenche_GridInsumosKit

    For Each objProjetoInsumosKit In colProjetoInsumosKit
    
        objProduto.sCodigo = objProjetoInsumosKit.sProduto
        
        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 124205
        
        sProdutoFormatado = objProduto.sCodigo
        
        'Mascara o produto para exibição no Grid
        lErro = Mascara_RetornaProdutoTela(sProdutoFormatado, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 134707
       
        'inicializa linha do grid
        iLinha = objProjetoInsumosKit.iSeq
        
        'inclui linha no grid
        GridKit.TextMatrix(iLinha, iGrid_ProdutoKit_Col) = sProdutoMascarado
        GridKit.TextMatrix(iLinha, iGrid_DescricaoProdKit_Col) = objProduto.sDescricao
        GridKit.TextMatrix(iLinha, iGrid_UMProdKit_Col) = objProjetoInsumosKit.sUMedida
        GridKit.TextMatrix(iLinha, iGrid_QuantidadeProdKit_Col) = Formata_Estoque(objProjetoInsumosKit.dQuantidade)
        
        'se tem custo unitário calculado ...
        If objProjetoInsumosKit.dCustoUnitarioCalculado > 0 Then
            
            GridKit.TextMatrix(iLinha, iGrid_CustoUnitCalculadoKit_Col) = Format(objProjetoInsumosKit.dCustoUnitarioCalculado, gobjFAT.sFormatoPrecoUnitario)
        
        End If
        
        'se tem custo unitario informado
        If objProjetoInsumosKit.dCustoUnitarioInformado > 0 Then
            
            GridKit.TextMatrix(iLinha, iGrid_CustoUnitKit_Col) = Format(objProjetoInsumosKit.dCustoUnitarioInformado, gobjFAT.sFormatoPrecoUnitario)
            
            'inicializa variável
            dVariacao = 0
                    
            'calcula a variação e põe no grid
            Call Calcula_Variacao(objProjetoInsumosKit.dCustoUnitarioCalculado, objProjetoInsumosKit.dCustoUnitarioInformado, dVariacao, iGrid_VariacaoKit_Col)
            If dVariacao <> 0 Then
                GridKit.TextMatrix(iLinha, iGrid_VariacaoKit_Col) = Format(dVariacao, "Percent")
            End If
            
            'exibe o total do custo em função da quantidade
            GridKit.TextMatrix(iLinha, iGrid_CustoTotalKit_Col) = Format(objProjetoInsumosKit.dQuantidade * objProjetoInsumosKit.dCustoUnitarioInformado, "Standard")
        
        End If
        
        'exibe a observacao
        GridKit.TextMatrix(iLinha, iGrid_ObservacaoKit_Col) = objProjetoInsumosKit.sObservacao
        
    Next
    
    'fixa as linhas do grid
    objGridKit.iLinhasExistentes = colProjetoInsumosKit.Count
    
    Preenche_GridInsumosKit = SUCESSO
    
    Exit Function
            
Erro_Preenche_GridInsumosKit:

    Preenche_GridInsumosKit = gErr
    
    Select Case gErr
    
        Case 134200
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158473)

    End Select

    Exit Function

End Function

Function Preenche_GridInsumosMaquinas(ByVal colProjetoInsumosMaquinas As Collection) As Long

Dim lErro As Long
Dim objProjetoInsumosMaquina As New ClassProjetoInsumosMaquina
Dim objProduto As New ClassProduto
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim iLinha As Integer
Dim dVariacao As Double

On Error GoTo Erro_Preenche_GridInsumosMaquinas

    For Each objProjetoInsumosMaquina In colProjetoInsumosMaquinas
    
        objProduto.sCodigo = objProjetoInsumosMaquina.sProduto
        
        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 124205
        
        sProdutoFormatado = objProduto.sCodigo
        
        'Mascara o produto para exibição no Grid
        lErro = Mascara_RetornaProdutoTela(sProdutoFormatado, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 134707
       
        'inicializa linha do grid
        iLinha = objProjetoInsumosMaquina.iSeq
        
        'inclui linha no grid
        GridMaquinas.TextMatrix(iLinha, iGrid_ProdutoMaq_Col) = sProdutoMascarado
        GridMaquinas.TextMatrix(iLinha, iGrid_DescricaoProdMaq_Col) = objProduto.sDescricao
        GridMaquinas.TextMatrix(iLinha, iGrid_UMProdMaq_Col) = objProjetoInsumosMaquina.sUMedida
        GridMaquinas.TextMatrix(iLinha, iGrid_QuantidadeProdMaq_Col) = Formata_Estoque(objProjetoInsumosMaquina.dQuantidade)
        
        'se tem custo unitário calculado ...
        If objProjetoInsumosMaquina.dCustoUnitarioCalculado > 0 Then
            
            GridMaquinas.TextMatrix(iLinha, iGrid_CustoUnitCalculadoMaq_Col) = Format(objProjetoInsumosMaquina.dCustoUnitarioCalculado, gobjFAT.sFormatoPrecoUnitario)
        
        End If
        
        'se tem custo unitario informado
        If objProjetoInsumosMaquina.dCustoUnitarioInformado > 0 Then
            
            GridMaquinas.TextMatrix(iLinha, iGrid_CustoUnitMaq_Col) = Format(objProjetoInsumosMaquina.dCustoUnitarioInformado, gobjFAT.sFormatoPrecoUnitario)
            
            'inicializa variável
            dVariacao = 0
                    
            'calcula a variação e põe no grid
            Call Calcula_Variacao(objProjetoInsumosMaquina.dCustoUnitarioCalculado, objProjetoInsumosMaquina.dCustoUnitarioInformado, dVariacao, iGrid_VariacaoMaq_Col)
            If dVariacao <> 0 Then
                GridMaquinas.TextMatrix(iLinha, iGrid_VariacaoMaq_Col) = Format(dVariacao, "Percent")
            End If
            
            'exibe o total do custo em função da quantidade
            GridMaquinas.TextMatrix(iLinha, iGrid_CustoTotalMaq_Col) = Format(objProjetoInsumosMaquina.dQuantidade * objProjetoInsumosMaquina.dCustoUnitarioInformado, "Standard")
        
        End If
        
        'exibe a observacao
        GridMaquinas.TextMatrix(iLinha, iGrid_ObservacaoMaq_Col) = objProjetoInsumosMaquina.sObservacao
                
    Next
    
    'fixa as linhas do grid
    objGridMaquinas.iLinhasExistentes = colProjetoInsumosMaquinas.Count
    
    Preenche_GridInsumosMaquinas = SUCESSO
    
    Exit Function
            
Erro_Preenche_GridInsumosMaquinas:

    Preenche_GridInsumosMaquinas = gErr
    
    Select Case gErr
    
        Case 134200
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158474)

    End Select

    Exit Function

End Function

Function Preenche_GridMaoDeObra(ByVal colProjetoMaoDeObra As Collection) As Long

Dim lErro As Long
Dim objProjetoMaoDeObra As New ClassProjetoMaoDeObra
Dim objTiposDeMaodeObra As ClassTiposDeMaodeObra
Dim iLinha As Integer
Dim dVariacao As Double

On Error GoTo Erro_Preenche_GridMaoDeObra

    For Each objProjetoMaoDeObra In colProjetoMaoDeObra
    
        Set objTiposDeMaodeObra = New ClassTiposDeMaodeObra
        
        objTiposDeMaodeObra.iCodigo = objProjetoMaoDeObra.iCodMO
        
        'Lê o TiposDeMaodeObra que está sendo Passado
        lErro = CF("TiposDeMaodeObra_Le", objTiposDeMaodeObra)
        If lErro <> SUCESSO And lErro <> 135004 Then gError 135029
       
        'inicializa linha do grid
        iLinha = objProjetoMaoDeObra.iSeq
        
        'inclui linha no grid
        GridMaoDeObra.TextMatrix(iLinha, iGrid_TipoMO_Col) = objProjetoMaoDeObra.iCodMO
        GridMaoDeObra.TextMatrix(iLinha, iGrid_DescricaoMo_Col) = objTiposDeMaodeObra.sDescricao
        GridMaoDeObra.TextMatrix(iLinha, iGrid_UMMO_Col) = objProjetoMaoDeObra.sUMedida
        GridMaoDeObra.TextMatrix(iLinha, iGrid_QuantidadeMO_Col) = Formata_Estoque(objProjetoMaoDeObra.dQuantidade)
        
        'se tem custo unitário calculado ...
        If objProjetoMaoDeObra.dCustoUnitarioCalculado > 0 Then
            
            GridMaoDeObra.TextMatrix(iLinha, iGrid_CustoUnitCalculadoMO_Col) = Format(objProjetoMaoDeObra.dCustoUnitarioCalculado, gobjFAT.sFormatoPrecoUnitario)
        
        End If
        
        'se tem custo unitario informado
        If objProjetoMaoDeObra.dCustoUnitarioInformado > 0 Then
            
            GridMaoDeObra.TextMatrix(iLinha, iGrid_CustoUnitMO_Col) = Format(objProjetoMaoDeObra.dCustoUnitarioInformado, gobjFAT.sFormatoPrecoUnitario)
            
            'inicializa variável
            dVariacao = 0
                    
            'calcula a variação e põe no grid
            Call Calcula_Variacao(objProjetoMaoDeObra.dCustoUnitarioCalculado, objProjetoMaoDeObra.dCustoUnitarioInformado, dVariacao, iGrid_VariacaoMO_Col)
            If dVariacao <> 0 Then
                GridMaoDeObra.TextMatrix(iLinha, iGrid_VariacaoMO_Col) = Format(dVariacao, "Percent")
            End If
            
            'exibe o total do custo em função da quantidade
            GridMaoDeObra.TextMatrix(iLinha, iGrid_CustoTotalMO_Col) = Format(objProjetoMaoDeObra.dQuantidade * objProjetoMaoDeObra.dCustoUnitarioInformado, "Standard")
        
        End If
        
        'exibe a observacao
        GridMaoDeObra.TextMatrix(iLinha, iGrid_ObservacaoMO_Col) = objProjetoMaoDeObra.sObservacao
                
    Next
    
    'fixa as linhas do grid
    objGridMaoDeObra.iLinhasExistentes = colProjetoMaoDeObra.Count
    
    Preenche_GridMaoDeObra = SUCESSO
    
    Exit Function
            
Erro_Preenche_GridMaoDeObra:

    Preenche_GridMaoDeObra = gErr
    
    Select Case gErr
    
        Case 134200
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158475)

    End Select

    Exit Function

End Function



Function Preenche_GridItens(ByVal objProjeto As ClassProjeto) As Long

Dim lErro As Long
Dim objProjetoCusteioItens As New ClassProjetoCusteioItens
Dim objProjetoItens As New ClassProjetoItens
Dim objProdutos As ClassProduto
Dim sProdutoMascarado As String
Dim dCustoTotalItem As Double
Dim dCustoTotalProjeto As Double
Dim dPrecoTotalProjeto As Double

On Error GoTo Erro_Preenche_GridItens

    'Exibe os dados da coleção de ProjetoCusteioItens na tela (GridItens)
    For Each objProjetoItens In objProjeto.colProjetoItens
    
        Set objProdutos = New ClassProduto
        
        objProdutos.sCodigo = objProjetoItens.sProduto
        
        'Le o Produto do Item do Projeto
        lErro = CF("Produto_Le", objProdutos)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 134099
        
        'Mascara o Código do Produto para por na Tela
        lErro = Mascara_RetornaProdutoTela(objProdutos.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 134100
                                
        'Insere no Grid Itens
        GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_ProdutoItens_Col) = sProdutoMascarado
        GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_VersaoProdItens_Col) = objProjetoItens.sVersao
        GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_DescricaoProdItens_Col) = objProdutos.sDescricao
        GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_UMProdItens_Col) = objProjetoItens.sUMedida
        If objProjetoItens.dQuantidade > 0 Then GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_QuantidadeProdItens_Col) = Formata_Estoque(objProjetoItens.dQuantidade)
        
        'Para cada item da coleção ...
        For Each objProjetoCusteioItens In colProjetoCusteioItens
        
            'Verifica se é referente a este Item de Projeto
            If objProjetoCusteioItens.lNumIntDocProjetoItem = objProjetoItens.lNumIntDoc Then
            
                'Continua preenchendo o grid com os dados do Custeio
                If objProjetoCusteioItens.dCustoTotalInsumosKit > 0 Then GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_CustoMPItens_Col) = Format(objProjetoCusteioItens.dCustoTotalInsumosKit, "Standard")
                If objProjetoCusteioItens.dCustoTotalInsumosMaq > 0 Then GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_CustoInsumosMaqItens_Col) = Format(objProjetoCusteioItens.dCustoTotalInsumosMaq, "Standard")
                If objProjetoCusteioItens.dCustoTotalMaoDeObra > 0 Then GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_CustoMOItens_Col) = Format(objProjetoCusteioItens.dCustoTotalMaoDeObra, "Standard")
                
                If objProjetoCusteioItens.dPrecoTotalItem > 0 Then
                    
                    'Se tem Quantidade para dividir ...
                    If objProjetoItens.dQuantidade > 0 Then
                        GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_PrecoUnitItens_Col) = Format(objProjetoCusteioItens.dPrecoTotalItem / objProjetoItens.dQuantidade, gobjFAT.sFormatoPrecoUnitario)
                    End If
                    
                    GridItens.TextMatrix(objProjetoItens.iSeq, iGrid_PrecoTotalItens_Col) = Format(objProjetoCusteioItens.dPrecoTotalItem, "Standard")
                
                End If
                
                'Acumula o Custo Total e o Preço Total
                dCustoTotalItem = objProjetoCusteioItens.dCustoTotalInsumosKit + objProjetoCusteioItens.dCustoTotalInsumosMaq + objProjetoCusteioItens.dCustoTotalMaoDeObra
                dCustoTotalProjeto = dCustoTotalProjeto + dCustoTotalItem
                dPrecoTotalProjeto = dPrecoTotalProjeto + objProjetoCusteioItens.dPrecoTotalItem
                
                Exit For
                
            End If
        
        Next

    Next
    
    'Atualiza as linhas do Grid
    objGridItens.iLinhasExistentes = objProjeto.colProjetoItens.Count
    
    'Lança os valores Totais na Tela
    CustoTotalProjeto.Caption = Format(dCustoTotalProjeto, "Standard")
    PrecoTotalProjeto.Caption = Format(dPrecoTotalProjeto, "Standard")

    Preenche_GridItens = SUCESSO
    
    Exit Function
            
Erro_Preenche_GridItens:

    Preenche_GridItens = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158476)

    End Select

    Exit Function

End Function

Function CalculaNecessidadesProducao(ByVal objProjetoItens As ClassProjetoItens, ByVal objProjetoCusteioItens As ClassProjetoCusteioItens) As Long
    
Dim lErro As Long
Dim dCustoTotalInsumosKit As Double
Dim dCustoTotalInsumosMaq As Double
Dim dCustoTotalMaoDeObra As Double
Dim objProduto As ClassProduto
Dim dCustoTotalItem As Double
Dim dPrecoTotalItem As Double

On Error GoTo Erro_CalculaNecessidadesProducao

    'calcula as Necessidade de Produção de Insumos do Kit
    lErro = CalculaNecessidade_InsumosKit(objProjetoItens, objProjetoCusteioItens, 0, dCustoTotalInsumosKit)
    If lErro <> SUCESSO Then gError 134200
    
    objProjetoCusteioItens.dCustoTotalInsumosKit = dCustoTotalInsumosKit
    
    'calcula as Necessidade de Produção de Insumos da Maquina
    lErro = CalculaNecessidade_InsumosMaquina(objProjetoItens, objProjetoCusteioItens, 0, dCustoTotalInsumosMaq)
    If lErro <> SUCESSO Then gError 134201
    
    objProjetoCusteioItens.dCustoTotalInsumosMaq = dCustoTotalInsumosMaq
    
    'calcula as Necessidade de Produção de Mão-de-Obra da Maquina
    lErro = CalculaNecessidade_MaoDeObra(objProjetoItens, objProjetoCusteioItens, 0, dCustoTotalMaoDeObra)
    If lErro <> SUCESSO Then gError 134202
    
    objProjetoCusteioItens.dCustoTotalMaoDeObra = dCustoTotalMaoDeObra
    
    'Calcula o Custo Total do Item e o Preço Total do Item
    dCustoTotalItem = dCustoTotalInsumosKit + dCustoTotalInsumosMaq + dCustoTotalMaoDeObra
    
    Set objProduto = New ClassProduto
    
    objProduto.sCodigo = objProjetoItens.sProduto
    
    lErro = CF("Produto_Le_PrecoProduto", objProduto, dCustoTotalItem, dPrecoTotalItem)
    If lErro <> SUCESSO Then gError 134202
    
    objProjetoCusteioItens.dPrecoTotalItem = dPrecoTotalItem

    CalculaNecessidadesProducao = SUCESSO
    
    Exit Function
    
Erro_CalculaNecessidadesProducao:

    CalculaNecessidadesProducao = gErr

    Select Case gErr
    
        Case 134200, 134201, 134202
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158477)

    End Select

    Exit Function

End Function

Function Calcula_Variacao(ByVal dCustoUnitarioCalculado As Double, dCustoUnitarioInformado As Double, dVariacao As Double, iColunaGrid As Integer) As Long

Dim lErro As Long
Dim dDiferenca As Double

On Error GoTo Erro_Calcula_Variacao

    Select Case iColunaGrid
    
        'se quero calcular o percentual
        Case Is = iGrid_VariacaoKit_Col, _
                  iGrid_VariacaoMaq_Col, _
                  iGrid_VariacaoMO_Col
        
            'se o calculado não for zero
            If dCustoUnitarioCalculado <> 0 Then
                
                'se o informado não for zero
                If dCustoUnitarioInformado <> 0 Then
                
                    'acho a diferença
                    dDiferenca = dCustoUnitarioInformado - dCustoUnitarioCalculado
                    'e calculo a variacao
                    dVariacao = (dDiferenca / dCustoUnitarioCalculado)
                
                Else
                
                    'a variacao é negativa
                    dVariacao = -1

                End If
                
            Else
                
                'se o informado não for zero
                If dCustoUnitarioInformado <> 0 Then
                
                    'a variacao é total
                    dVariacao = 1
                    
                Else
                
                    'nao tem variacao
                    dVariacao = 0

                End If
                
            End If
        
        'se quero calcular o custo unitario
        Case Is = iGrid_CustoUnitKit_Col, _
                  iGrid_CustoUnitMaq_Col, _
                  iGrid_CustoUnitMO_Col
            
            'independentemente de qual membro da equação for zerado ou negativo ...
            'calcula o informado = (calculado + (calculado * variacao))
            dCustoUnitarioInformado = dCustoUnitarioCalculado + ((dCustoUnitarioCalculado * dVariacao) / 100)
            
    End Select
    
    Calcula_Variacao = SUCESSO
    
    Exit Function
    
Erro_Calcula_Variacao:

    Calcula_Variacao = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158478)
        
    End Select
    
    Exit Function

End Function

Function Localiza_ProjetoCusteioItens(ByVal sProduto As String, ByVal sVersao As String, objProjetoCusteioItens As ClassProjetoCusteioItens) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProjeto As ClassProjeto
Dim objProjetoItens As New ClassProjetoItens
Dim lNumIntDocProjetoItem As Integer
Dim objPCItens As New ClassProjetoCusteioItens

On Error GoTo Erro_Localiza_ProjetoCusteioItens

    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134702

    'se o produto existe cadastrado ...
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
        Set objProjeto = New ClassProjeto
        
        'Le o Projeto para pegar seu NumIntDoc
        lErro = CF("TP_Projeto_Le", Projeto, objProjeto)
        If lErro <> SUCESSO Then gError 134467
        
        'Ajusta a variável alterarada indevidamente pela TP_Projeto_Le
        iProjetoAlterado = 0
        
        'Le os Itens do Projeto
        lErro = CF("Projeto_Le_Itens", objProjeto)
        If lErro <> SUCESSO And lErro <> 134453 Then gError 134096
        
        'Percorre os Itens do Projeto para achar o Produto e Versão
        For Each objProjetoItens In objProjeto.colProjetoItens
        
            If objProjetoItens.sProduto = sProdutoFormatado And objProjetoItens.sVersao = Trim(sVersao) Then
            
                'encontrei ...
                lNumIntDocProjetoItem = objProjetoItens.lNumIntDoc
                Exit For
                
            End If
        
        Next
        
        'verifica na coleção de itens o item selecionado
        For Each objPCItens In colProjetoCusteioItens
        
            'se encontrou ...
            If objPCItens.lNumIntDocProjetoItem = lNumIntDocProjetoItem Then
            
                'Seta o obj a retornar
                Set objProjetoCusteioItens = objPCItens
                    
                'e abandona a pesquisa por itens
                Exit For
                
            End If
    
        Next
        
    End If
    
    Localiza_ProjetoCusteioItens = SUCESSO
    
    Exit Function

Erro_Localiza_ProjetoCusteioItens:

    Localiza_ProjetoCusteioItens = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158479)
    
    End Select
    
    Exit Function

End Function



Function AjustaCustoNecessidade_InsumosKit(ByVal sConteudoAAjustar As String, ByVal iColunaAAjustar As Integer, Optional ByVal dAjusteCustoKit As Double) As Long

Dim lErro As Long
Dim sProduto As String
Dim sVersao As String
Dim objProjetoCusteioItens As ClassProjetoCusteioItens
Dim objProjetoInsumosKit As New ClassProjetoInsumosKit

On Error GoTo Erro_AjustaCustoNecessidade_InsumosKit

    sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
    sVersao = GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col)

    Set objProjetoCusteioItens = New ClassProjetoCusteioItens

    'Localiza o item de custeio na coleção
    lErro = Localiza_ProjetoCusteioItens(sProduto, sVersao, objProjetoCusteioItens)
    If lErro <> SUCESSO Then gError 134702
    
    'Percorre a coleção no objProjetoCusteioItens localizado
    For Each objProjetoInsumosKit In objProjetoCusteioItens.colProjetoInsumosKit
    
        'Verifica se é o item que está alterando ...
        If objProjetoInsumosKit.iSeq = GridKit.Row Then
            
            'encontrou... altera o valor da coluna passada
            Select Case iColunaAAjustar
            
                Case Is = iGrid_CustoUnitKit_Col
                
                    objProjetoInsumosKit.dCustoUnitarioInformado = StrParaDbl(sConteudoAAjustar)
                
                Case Is = iGrid_ObservacaoKit_Col
                    
                    objProjetoInsumosKit.sObservacao = sConteudoAAjustar
                
            End Select
            
            'e encerra a busca
            Exit For
            
        End If
    
    Next
    
    'se a coluna a ajustar é a do Custo
    If iColunaAAjustar = iGrid_CustoUnitKit_Col Then
    
        'Ajusta o valor no obj
        objProjetoCusteioItens.dCustoTotalInsumosKit = objProjetoCusteioItens.dCustoTotalInsumosKit + dAjusteCustoKit
        
        'Ajusta o valor no Total do Tab
        CustoTotalInsumosKit.Caption = Format(objProjetoCusteioItens.dCustoTotalInsumosKit, "Standard")
        
        'Ajusta o valor no Grid de Itens do Projeto
        If objProjetoCusteioItens.dCustoTotalInsumosKit > 0 Then
            GridItens.TextMatrix(GridItens.Row, iGrid_CustoMPItens_Col) = Format(objProjetoCusteioItens.dCustoTotalInsumosKit, "Standard")
        Else
            GridItens.TextMatrix(GridItens.Row, iGrid_CustoMPItens_Col) = ""
        End If
        
        'Recalcula o Custo Total do Item e o Preço Total do Item
        lErro = Recalcula_Totais(objProjetoCusteioItens)
        If lErro <> SUCESSO Then gError 137999

    End If

    AjustaCustoNecessidade_InsumosKit = SUCESSO
    
    Exit Function
    
Erro_AjustaCustoNecessidade_InsumosKit:

    AjustaCustoNecessidade_InsumosKit = gErr
    
    Select Case gErr
    
        Case 137999
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158480)
            
    End Select
    
    Exit Function

End Function

Function AjustaCustoNecessidade_InsumosMaquina(ByVal sConteudoAAjustar As String, ByVal iColunaAAjustar As Integer, Optional ByVal dAjusteCustoMaq As Double) As Long

Dim lErro As Long
Dim sProduto As String
Dim sVersao As String
Dim objProjetoCusteioItens As ClassProjetoCusteioItens
Dim objProjetoInsumosMaquina As New ClassProjetoInsumosMaquina

On Error GoTo Erro_AjustaCustoNecessidade_InsumosMaquina

    sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
    sVersao = GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col)

    Set objProjetoCusteioItens = New ClassProjetoCusteioItens

    'Localiza o item de custeio na coleção
    lErro = Localiza_ProjetoCusteioItens(sProduto, sVersao, objProjetoCusteioItens)
    If lErro <> SUCESSO Then gError 134702
    
    'Percorre a coleção no objProjetoCusteioItens localizado
    For Each objProjetoInsumosMaquina In objProjetoCusteioItens.colProjetoInsumosMaquina
    
        'Verifica se é o item que está alterando ...
        If objProjetoInsumosMaquina.iSeq = GridMaquinas.Row Then
            
            'encontrou... altera o valor da coluna passada
            Select Case iColunaAAjustar
            
                Case Is = iGrid_CustoUnitMaq_Col
                
                    objProjetoInsumosMaquina.dCustoUnitarioInformado = StrParaDbl(sConteudoAAjustar)
                
                Case Is = iGrid_ObservacaoMaq_Col
                    
                    objProjetoInsumosMaquina.sObservacao = sConteudoAAjustar
                
            End Select
            
            'e encerra a busca
            Exit For
            
        End If
    
    Next
    
    'se a coluna a ajustar é a do Custo
    If iColunaAAjustar = iGrid_CustoUnitMaq_Col Then
    
        'Ajusta o valor no obj
        objProjetoCusteioItens.dCustoTotalInsumosMaq = objProjetoCusteioItens.dCustoTotalInsumosMaq + dAjusteCustoMaq
        
        'Ajusta o valor no Total do Tab
        CustoTotalInsumosMaquina.Caption = Format(objProjetoCusteioItens.dCustoTotalInsumosMaq, "Standard")
        
        'Ajusta o valor no Grid de Itens do Projeto
        If objProjetoCusteioItens.dCustoTotalInsumosMaq > 0 Then
            GridItens.TextMatrix(GridItens.Row, iGrid_CustoInsumosMaqItens_Col) = Format(objProjetoCusteioItens.dCustoTotalInsumosMaq, "Standard")
        Else
            GridItens.TextMatrix(GridItens.Row, iGrid_CustoInsumosMaqItens_Col) = ""
        End If
        
        'Recalcula o Custo Total do Item e o Preço Total do Item
        lErro = Recalcula_Totais(objProjetoCusteioItens)
        If lErro <> SUCESSO Then gError 137999

    End If

    AjustaCustoNecessidade_InsumosMaquina = SUCESSO
    
    Exit Function
    
Erro_AjustaCustoNecessidade_InsumosMaquina:

    AjustaCustoNecessidade_InsumosMaquina = gErr
    
    Select Case gErr
        
        Case 137999
            'erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158481)
            
    End Select
    
    Exit Function

End Function

Function AjustaCustoNecessidade_MaoDeObra(ByVal sConteudoAAjustar As String, ByVal iColunaAAjustar As Integer, Optional ByVal dAjusteCustoMO As Double) As Long

Dim lErro As Long
Dim sProduto As String
Dim sVersao As String
Dim objProjetoCusteioItens As ClassProjetoCusteioItens
Dim objProjetoMaoDeObra As New ClassProjetoMaoDeObra

On Error GoTo Erro_AjustaCustoNecessidade_MaoDeObra

    sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
    sVersao = GridItens.TextMatrix(GridItens.Row, iGrid_VersaoProdItens_Col)

    Set objProjetoCusteioItens = New ClassProjetoCusteioItens

    'Localiza o item de custeio na coleção
    lErro = Localiza_ProjetoCusteioItens(sProduto, sVersao, objProjetoCusteioItens)
    If lErro <> SUCESSO Then gError 134702
    
    'Percorre a coleção no objProjetoCusteioItens localizado
    For Each objProjetoMaoDeObra In objProjetoCusteioItens.colProjetoMaoDeObra
    
        'Verifica se é o item que está alterando ...
        If objProjetoMaoDeObra.iSeq = GridMaoDeObra.Row Then
            
            'encontrou... altera o valor da coluna passada
            Select Case iColunaAAjustar
            
                Case Is = iGrid_CustoUnitMO_Col
                
                    objProjetoMaoDeObra.dCustoUnitarioInformado = StrParaDbl(sConteudoAAjustar)
                
                Case Is = iGrid_ObservacaoMO_Col
                    
                    objProjetoMaoDeObra.sObservacao = sConteudoAAjustar
                
            End Select
            
            'e encerra a busca
            Exit For
            
        End If
    
    Next
    
    'se a coluna a ajustar é a do Custo
    If iColunaAAjustar = iGrid_CustoUnitMO_Col Then
    
        'Ajusta o valor no obj
        objProjetoCusteioItens.dCustoTotalMaoDeObra = objProjetoCusteioItens.dCustoTotalMaoDeObra + dAjusteCustoMO
        
        'Ajusta o valor no Total do Tab
        CustoTotalMaoDeObra.Caption = Format(objProjetoCusteioItens.dCustoTotalMaoDeObra, "Standard")
        
        'Ajusta o valor no Grid de Itens do Projeto
        If objProjetoCusteioItens.dCustoTotalMaoDeObra > 0 Then
            GridItens.TextMatrix(GridItens.Row, iGrid_CustoMOItens_Col) = Format(objProjetoCusteioItens.dCustoTotalMaoDeObra, "Standard")
        Else
            GridItens.TextMatrix(GridItens.Row, iGrid_CustoMOItens_Col) = ""
        End If
        
        'Recalcula o Custo Total do Item e o Preço Total do Item
        lErro = Recalcula_Totais(objProjetoCusteioItens)
        If lErro <> SUCESSO Then gError 137999

    End If

    AjustaCustoNecessidade_MaoDeObra = SUCESSO
    
    Exit Function
    
Erro_AjustaCustoNecessidade_MaoDeObra:

    AjustaCustoNecessidade_MaoDeObra = gErr
    
    Select Case gErr
        
        Case 137999
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158482)
            
    End Select
    
    Exit Function

End Function

Function Recalcula_Totais(ByVal objProjetoCusteioItens As ClassProjetoCusteioItens) As Long

Dim lErro As Long
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As ClassProduto
Dim dCustoTotalItem As Double
Dim dPrecoTotalItem As Double
Dim dCustoTotalProjeto As Double
Dim dPrecoTotalProjeto As Double
Dim dQuantidade As Double
Dim iLinha As Integer

On Error GoTo Erro_Recalcula_Totais

    'Recalcula o Custo Total do Item e o Preço Total do Item
    dCustoTotalItem = objProjetoCusteioItens.dCustoTotalInsumosKit + objProjetoCusteioItens.dCustoTotalInsumosMaq + objProjetoCusteioItens.dCustoTotalMaoDeObra
    
    sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_ProdutoItens_Col)
    
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 134702
    
    Set objProduto = New ClassProduto
    
    objProduto.sCodigo = sProdutoFormatado
    
    lErro = CF("Produto_Le_PrecoProduto", objProduto, dCustoTotalItem, dPrecoTotalItem)
    If lErro <> SUCESSO Then gError 134202
        
    'Ajusta o valor do PrecoTotal no obj
    objProjetoCusteioItens.dPrecoTotalItem = dPrecoTotalItem
    
    'Ajusta o valor do PrecoTotal e o Preco Unitario no grid de Itens do Projeto
    If objProjetoCusteioItens.dPrecoTotalItem > 0 Then
    
        dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantidadeProdItens_Col))
        
        'Se tem Quantidade para dividir ...
        If dQuantidade > 0 Then
            GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitItens_Col) = Format(objProjetoCusteioItens.dPrecoTotalItem / dQuantidade, gobjFAT.sFormatoPrecoUnitario)
        End If
        
        GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotalItens_Col) = Format(objProjetoCusteioItens.dPrecoTotalItem, "Standard")
    
    Else
            
        GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitItens_Col) = ""
        GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotalItens_Col) = ""
    
    End If
    
    'Calcula os novos totais do Projeto
    For iLinha = 1 To objGridItens.iLinhasExistentes
    
        dCustoTotalProjeto = dCustoTotalProjeto + (StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_CustoMPItens_Col)) + StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_CustoInsumosMaqItens_Col)) + StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_CustoMOItens_Col)))
        dPrecoTotalProjeto = dPrecoTotalProjeto + StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_PrecoTotalItens_Col))
    
    Next

    'Ajusta o valor dos Totais do Custo e do Preço do Projeto na tela
    CustoTotalProjeto.Caption = Format(dCustoTotalProjeto, "Standard")
    PrecoTotalProjeto.Caption = Format(dPrecoTotalProjeto, "Standard")
    
    Recalcula_Totais = SUCESSO
    
    Exit Function
        
Erro_Recalcula_Totais:

    Recalcula_Totais = gErr
    
    Select Case gErr

        Case 134202, 134702
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158483)

    End Select
    
    Exit Function
        
End Function
Function Gravar_Registro() As Long

Dim lErro As Long
Dim objProjetoCusteio As New ClassProjetoCusteio
Dim iIndice As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o Projeto está preenchido
    If Len(Trim(Projeto.Text)) = 0 Then gError 134084
    
    'Verifica se a Data do Custeio está preenchida
    If Len(Trim(DataCusteio.ClipText)) = 0 Then gError 137127
    
    'Preenche o objProjetoCusteio
    lErro = Move_Tela_Memoria(objProjetoCusteio)
    If lErro <> SUCESSO Then gError 134091

    lErro = Trata_Alteracao(objProjetoCusteio, objProjetoCusteio.lNumIntDocProj)
    If lErro <> SUCESSO Then gError 134092

    'Grava o ProjetoCusteio no Banco de Dados
    lErro = CF("ProjetoCusteio_Grava", objProjetoCusteio)
    If lErro <> SUCESSO Then gError 134093
    
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 134084
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 137127
            Call Rotina_Erro(vbOKOnly, "ERRO_DATACUSTEIO_NAO_PREENCHIDA", gErr)
                
        Case 134091, 134092, 134093
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158484)

    End Select

    Exit Function

End Function

Function Move_Tela_Memoria(objProjetoCusteio As ClassProjetoCusteio) As Long

Dim lErro As Long
Dim objProjeto As ClassProjeto
Dim objProjetoCusteioItens As ClassProjetoCusteioItens

On Error GoTo Erro_Move_Tela_Memoria

    Set objProjeto = New ClassProjeto
    
    'Le o Projeto para pegar seu NumIntDoc
    lErro = CF("TP_Projeto_Le", Projeto, objProjeto)
    If lErro <> SUCESSO Then gError 134467
    
    'Ajusta a variável alterarada indevidamente pela TP_Projeto_Le
    iProjetoAlterado = 0

    objProjetoCusteio.lNumIntDocProj = objProjeto.lNumIntDoc
    
    If StrParaDbl(CustoTotalProjeto.Caption) <> 0 Then
        objProjetoCusteio.dCustoTotalProjeto = StrParaDbl(CustoTotalProjeto.Caption)
    End If
    
    If StrParaDbl(PrecoTotalProjeto.Caption) <> 0 Then
        objProjetoCusteio.dPrecoTotalProjeto = StrParaDbl(PrecoTotalProjeto.Caption)
    End If

    If Len(Trim(DataCusteio.ClipText)) > 0 Then
        objProjetoCusteio.dtDataCusteio = CDate(DataCusteio.Text)
    Else
        objProjetoCusteio.dtDataCusteio = DATA_NULA
    End If
    
    'preenche a coleção dos itens do Projeto
    For Each objProjetoCusteioItens In colProjetoCusteioItens
    
       objProjetoCusteio.colProjetoCusteioItens.Add objProjetoCusteioItens
       
    Next
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 134467
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158485)

    End Select

    Exit Function

End Function

Function IncluiNecessidade_InsumosMaquina(ByVal objMaquinasInsumos As ClassMaquinasInsumos, ByVal objPOMaquinas As ClassPOMaquinas, ByVal objProjetoCusteioItens As ClassProjetoCusteioItens, iSeq As Integer, dCustoTotalInsumosMaq As Double) As Long

Dim lErro As Long
Dim bProdutoNovo As Boolean
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim dQuantidade As Double
Dim dCustoProduto As Double
Dim objProjetoInsumosMaquina As ClassProjetoInsumosMaquina
Dim objAuxProjInsumosMaquina As New ClassProjetoInsumosMaquina

On Error GoTo Erro_IncluiNecessidade_InsumosMaquina

    'Descobre o fator de conversao da UM de Tempo utilizada p/UM de tempo padrão
    lErro = CF("UM_Conversao_Trans", gobjEST.iClasseUMTempo, objMaquinasInsumos.sUMTempo, TAXA_CONSUMO_TEMPO_PADRAO, dFator)
    If lErro <> SUCESSO Then gError 134976
    
    'Converte a quantidade
    dQuantidade = (objMaquinasInsumos.dQuantidade * dFator) * objPOMaquinas.dHorasMaquina
    
    'Vamos incluir na coleção...
    bProdutoNovo = True
    
    'Verifica se já há algum produto insumos de Maquinas na coleção
    For Each objAuxProjInsumosMaquina In objProjetoCusteioItens.colProjetoInsumosMaquina

        'se encontrou ...
        If objAuxProjInsumosMaquina.sProduto = objMaquinasInsumos.sProduto Then
                        
            'subtrai o valor anterior do Total do Custo no acumulador
            dCustoTotalInsumosMaq = dCustoTotalInsumosMaq - (objAuxProjInsumosMaquina.dQuantidade * objAuxProjInsumosMaquina.dCustoUnitarioInformado)
            
            'acumula a quantidade
            objAuxProjInsumosMaquina.dQuantidade = objAuxProjInsumosMaquina.dQuantidade + dQuantidade
            
            'lança o novo valor Total do Custo no acumulador
            dCustoTotalInsumosMaq = dCustoTotalInsumosMaq + (objAuxProjInsumosMaquina.dQuantidade * objAuxProjInsumosMaquina.dCustoUnitarioInformado)
            
            'avisa que acumulou ...
            bProdutoNovo = False
            Exit For
            
        End If
        
    Next
    
    'se não tem o produto ...
    If bProdutoNovo Then
    
        Set objProduto = New ClassProduto
        
        objProduto.sCodigo = objMaquinasInsumos.sProduto

        'Descobre qual seu custo
        lErro = CF("Produto_Le_CustoProduto", objProduto, dCustoProduto)
        If lErro <> SUCESSO Then gError 134202
                                    
        'incrementa o sequencial
        iSeq = iSeq + 1
        
        'reCria o obj
        Set objProjetoInsumosMaquina = New ClassProjetoInsumosMaquina
        
        objProjetoInsumosMaquina.iSeq = iSeq
        objProjetoInsumosMaquina.sProduto = objMaquinasInsumos.sProduto
        objProjetoInsumosMaquina.sUMedida = objMaquinasInsumos.sUMProduto
        objProjetoInsumosMaquina.dQuantidade = dQuantidade
        objProjetoInsumosMaquina.dCustoUnitarioCalculado = dCustoProduto
        objProjetoInsumosMaquina.dCustoUnitarioInformado = dCustoProduto
        
        'lança o valor Total do Custo no acumulador
        dCustoTotalInsumosMaq = dCustoTotalInsumosMaq + (dQuantidade * dCustoProduto)
        
        'e inclui na coleção.
        objProjetoCusteioItens.colProjetoInsumosMaquina.Add objProjetoInsumosMaquina, "X" & Right$(CStr(100000 + iSeq), 5)
        
    End If
        
    IncluiNecessidade_InsumosMaquina = SUCESSO
    
    Exit Function
    
Erro_IncluiNecessidade_InsumosMaquina:

    IncluiNecessidade_InsumosMaquina = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158486)
    
    End Select
    
    Exit Function

End Function

Function IncluiNecessidade_MaoDeObra(ByVal objMaquinaOperadores As ClassMaquinaOperadores, ByVal objPOMaquinas As ClassPOMaquinas, ByVal objProjetoCusteioItens As ClassProjetoCusteioItens, iSeq As Integer, dCustoTotalMaoDeObra As Double) As Long

Dim lErro As Long
Dim bTipoNovo As Boolean
Dim objTipoMO As New ClassTiposDeMaodeObra
Dim dQuantidade As Double
Dim dCustoMO As Double
Dim objProjetoMaoDeObra As ClassProjetoMaoDeObra
Dim objAuxProjMaoDeObra As New ClassProjetoMaoDeObra

On Error GoTo Erro_IncluiNecessidade_MaoDeObra

    'Converte a quantidade
    dQuantidade = (objMaquinaOperadores.iQuantidade * objMaquinaOperadores.dPercentualUso) * objPOMaquinas.dHorasMaquina
    
    'Vamos incluir na coleção...
    bTipoNovo = True
    
    'Verifica se já há algum Tipo de MaoDeObra na coleção
    For Each objAuxProjMaoDeObra In objProjetoCusteioItens.colProjetoMaoDeObra

        'se encontrou ...
        If objAuxProjMaoDeObra.iCodMO = objMaquinaOperadores.iTipoMaoDeObra Then
                        
            'subtrai o valor anterior do Total do Custo no acumulador
            dCustoTotalMaoDeObra = dCustoTotalMaoDeObra - (objAuxProjMaoDeObra.dQuantidade * objAuxProjMaoDeObra.dCustoUnitarioInformado)
            
            'acumula a quantidade
            objAuxProjMaoDeObra.dQuantidade = objAuxProjMaoDeObra.dQuantidade + dQuantidade
            
            'lança o novo valor Total do Custo no acumulador
            dCustoTotalMaoDeObra = dCustoTotalMaoDeObra + (objAuxProjMaoDeObra.dQuantidade * objAuxProjMaoDeObra.dCustoUnitarioInformado)
            
            'avisa que acumulou ...
            bTipoNovo = False
            Exit For
            
        End If
        
    Next
    
    'se não tem o tipo ...
    If bTipoNovo Then
    
        Set objTipoMO = New ClassTiposDeMaodeObra
        
        objTipoMO.iCodigo = objMaquinaOperadores.iTipoMaoDeObra

        'Descobre qual seu custo
        lErro = CF("TiposDeMaodeObra_Le", objTipoMO)
        If lErro <> SUCESSO And lErro <> 135004 Then gError 134202
                                    
        'incrementa o sequencial
        iSeq = iSeq + 1
        
        'reCria o obj
        Set objProjetoMaoDeObra = New ClassProjetoMaoDeObra
        
        objProjetoMaoDeObra.iSeq = iSeq
        objProjetoMaoDeObra.iCodMO = objMaquinaOperadores.iTipoMaoDeObra
        objProjetoMaoDeObra.sUMedida = TAXA_CONSUMO_TEMPO_PADRAO
        objProjetoMaoDeObra.dQuantidade = dQuantidade
        objProjetoMaoDeObra.dCustoUnitarioCalculado = objTipoMO.dCustoHora
        objProjetoMaoDeObra.dCustoUnitarioInformado = objTipoMO.dCustoHora
        
        'lança o valor Total do Custo no acumulador
        dCustoTotalMaoDeObra = dCustoTotalMaoDeObra + (dQuantidade * objTipoMO.dCustoHora)
        
        'e inclui na coleção.
        objProjetoCusteioItens.colProjetoMaoDeObra.Add objProjetoMaoDeObra, "X" & Right$(CStr(100000 + iSeq), 5)
        
    End If
        
    IncluiNecessidade_MaoDeObra = SUCESSO
    
    Exit Function
    
Erro_IncluiNecessidade_MaoDeObra:

    IncluiNecessidade_MaoDeObra = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158487)
    
    End Select
    
    Exit Function

End Function


