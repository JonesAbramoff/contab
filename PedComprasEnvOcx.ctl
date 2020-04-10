VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl PedComprasEnvOcx 
   ClientHeight    =   9195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   KeyPreview      =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   16995
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   8145
      Index           =   4
      Left            =   165
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   16620
      Begin VB.Frame Frame5 
         Caption         =   "Distribuição dos Produtos"
         Height          =   7890
         Left            =   180
         TabIndex        =   29
         Top             =   210
         Width           =   16260
         Begin VB.TextBox ContaContabil 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   7695
            MaxLength       =   50
            TabIndex        =   31
            Top             =   1170
            Width           =   1575
         End
         Begin VB.TextBox DescProduto 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1785
            MaxLength       =   50
            TabIndex        =   30
            Top             =   765
            Width           =   4000
         End
         Begin MSMask.MaskEdBox Quant 
            Height          =   225
            Left            =   6240
            TabIndex        =   32
            Top             =   360
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
         Begin MSMask.MaskEdBox UM 
            Height          =   225
            Left            =   5325
            TabIndex        =   33
            Top             =   1620
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   5
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CentroCusto 
            Height          =   225
            Left            =   2880
            TabIndex        =   34
            Top             =   360
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Prod 
            Height          =   225
            Left            =   270
            TabIndex        =   35
            Top             =   360
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Almoxarifado 
            Height          =   225
            Left            =   5895
            TabIndex        =   36
            Top             =   4035
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridDistribuicao 
            Height          =   2715
            Left            =   255
            TabIndex        =   37
            Top             =   495
            Width           =   15735
            _ExtentX        =   27755
            _ExtentY        =   4789
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
      Height          =   8235
      Index           =   5
      Left            =   165
      TabIndex        =   105
      Top             =   840
      Visible         =   0   'False
      Width           =   16575
      Begin VB.Frame SSFrame1 
         Caption         =   "Bloqueios"
         Height          =   7590
         Left            =   105
         TabIndex        =   107
         Top             =   90
         Width           =   16395
         Begin VB.ComboBox TipoBloqueio 
            Height          =   315
            ItemData        =   "PedComprasEnvOcx.ctx":0000
            Left            =   180
            List            =   "PedComprasEnvOcx.ctx":0002
            TabIndex        =   108
            Top             =   570
            Width           =   3000
         End
         Begin MSMask.MaskEdBox ResponsavelLib 
            Height          =   270
            Left            =   7440
            TabIndex        =   109
            Top             =   4215
            Width           =   3200
            _ExtentX        =   5636
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataLiberacao 
            Height          =   270
            Left            =   5880
            TabIndex        =   110
            Top             =   780
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
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
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodUsuario 
            Height          =   270
            Left            =   3990
            TabIndex        =   111
            Top             =   4230
            Width           =   2500
            _ExtentX        =   4419
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ResponsavelBL 
            Height          =   270
            Left            =   5040
            TabIndex        =   112
            Top             =   5505
            Width           =   3200
            _ExtentX        =   5636
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataBloqueio 
            Height          =   270
            Left            =   2205
            TabIndex        =   113
            Top             =   585
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   476
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
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
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridBloqueios 
            Height          =   2805
            Left            =   120
            TabIndex        =   114
            Top             =   375
            Width           =   16020
            _ExtentX        =   28258
            _ExtentY        =   4948
            _Version        =   393216
            Rows            =   7
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.CommandButton BotaoLiberaBloqueio 
         Caption         =   "Liberação de Bloqueios"
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
         Left            =   150
         TabIndex        =   106
         Top             =   7755
         Width           =   2355
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8175
      Index           =   3
      Left            =   165
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   16545
      Begin VB.Frame Frame2 
         Caption         =   "Local de Entrega"
         Height          =   2910
         Left            =   195
         TabIndex        =   5
         Top             =   180
         Width           =   8595
         Begin VB.Frame FrameTipo 
            BorderStyle     =   0  'None
            Height          =   675
            Index           =   1
            Left            =   4785
            TabIndex        =   100
            Top             =   360
            Visible         =   0   'False
            Width           =   3495
            Begin VB.Label FilialFornec 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   1200
               TabIndex        =   104
               Top             =   360
               Width           =   2145
            End
            Begin VB.Label Label32 
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
               Left            =   645
               TabIndex        =   103
               Top             =   405
               Width           =   465
            End
            Begin VB.Label Fornec 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   1200
               TabIndex        =   102
               Top             =   0
               Width           =   2145
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
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   101
               Top             =   60
               Width           =   1035
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Tipo"
            Height          =   555
            Left            =   270
            TabIndex        =   9
            Top             =   450
            Width           =   3855
            Begin VB.OptionButton TipoDestino 
               Caption         =   "Fornecedor"
               Enabled         =   0   'False
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
               Left            =   2220
               TabIndex        =   11
               Top             =   225
               Width           =   1335
            End
            Begin VB.OptionButton TipoDestino 
               Caption         =   "Filial Empresa"
               Enabled         =   0   'False
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
               Left            =   420
               TabIndex        =   10
               Top             =   225
               Width           =   1515
            End
         End
         Begin VB.Frame FrameTipo 
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            Height          =   795
            Index           =   0
            Left            =   4800
            TabIndex        =   6
            Top             =   330
            Width           =   3495
            Begin VB.Label Label37 
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
               Left            =   390
               TabIndex        =   8
               Top             =   195
               Width           =   465
            End
            Begin VB.Label FilialEmpresa 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   960
               TabIndex        =   7
               Top             =   120
               Width           =   2145
            End
         End
         Begin VB.Label Pais 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4020
            TabIndex        =   23
            Top             =   2355
            Width           =   1995
         End
         Begin VB.Label Estado 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1260
            TabIndex        =   22
            Top             =   2400
            Width           =   495
         End
         Begin VB.Label CEP 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6675
            TabIndex        =   21
            Top             =   1920
            Width           =   930
         End
         Begin VB.Label Cidade 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4020
            TabIndex        =   20
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label Bairro 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1260
            TabIndex        =   19
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label Endereco 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1260
            TabIndex        =   18
            Top             =   1500
            Width           =   6345
         End
         Begin VB.Label Label63 
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
            Height          =   195
            Left            =   3465
            TabIndex        =   17
            Top             =   2400
            Width           =   495
         End
         Begin VB.Label Label65 
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
            Height          =   195
            Left            =   6150
            TabIndex        =   16
            Top             =   1995
            Width           =   465
         End
         Begin VB.Label Label70 
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
            Height          =   195
            Left            =   600
            TabIndex        =   15
            Top             =   1995
            Width           =   585
         End
         Begin VB.Label Label71 
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
            Height          =   195
            Left            =   510
            TabIndex        =   14
            Top             =   2400
            Width           =   675
         End
         Begin VB.Label Label72 
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
            Height          =   195
            Left            =   3285
            TabIndex        =   13
            Top             =   1995
            Width           =   675
         End
         Begin VB.Label Label73 
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
            Height          =   195
            Left            =   300
            TabIndex        =   12
            Top             =   1515
            Width           =   915
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frete"
         Height          =   885
         Left            =   180
         TabIndex        =   24
         Top             =   3165
         Width           =   8625
         Begin VB.Label TransportadoraLabel 
            Caption         =   "Transportadora:"
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
            Left            =   4020
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   28
            Top             =   435
            Width           =   1410
         End
         Begin VB.Label TipoFrete 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1590
            TabIndex        =   27
            Top             =   390
            Width           =   825
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Frete:"
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
            Left            =   555
            TabIndex        =   26
            Top             =   450
            Width           =   945
         End
         Begin VB.Label Transportadora 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5490
            TabIndex        =   25
            Top             =   390
            Width           =   2190
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   8115
      Index           =   2
      Left            =   165
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   16560
      Begin VB.Frame Frame9 
         Caption         =   "Valores"
         Height          =   900
         Index           =   1
         Left            =   135
         TabIndex        =   45
         Top             =   7200
         Width           =   8625
         Begin VB.Label DescontoValor 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4958
            TabIndex        =   59
            Top             =   435
            Width           =   1125
         End
         Begin VB.Label ValorProdutos 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   150
            TabIndex        =   58
            Top             =   435
            Width           =   1125
         End
         Begin VB.Label Label41 
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
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   330
            TabIndex        =   57
            Top             =   225
            Width           =   765
         End
         Begin VB.Label ValorFrete 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1352
            TabIndex        =   56
            Top             =   435
            Width           =   1125
         End
         Begin VB.Label ValorSeguro 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2554
            TabIndex        =   55
            Top             =   435
            Width           =   1125
         End
         Begin VB.Label OutrasDespesas 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3756
            TabIndex        =   54
            Top             =   435
            Width           =   1125
         End
         Begin VB.Label IPIValor 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6160
            TabIndex        =   53
            Top             =   435
            Width           =   1125
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Despesas"
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
            Left            =   3855
            TabIndex        =   52
            Top             =   240
            Width           =   840
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Frete"
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
            Left            =   1695
            TabIndex        =   51
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Desconto"
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
            Left            =   5115
            TabIndex        =   50
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Seguro"
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
            Left            =   2760
            TabIndex        =   49
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "IPI"
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
            Height          =   165
            Index           =   1
            Left            =   6600
            TabIndex        =   48
            Top             =   240
            Width           =   255
         End
         Begin VB.Label ValorTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7365
            TabIndex        =   47
            Top             =   435
            Width           =   1125
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
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
            Left            =   7695
            TabIndex        =   46
            Top             =   240
            Width           =   450
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Itens"
         Height          =   6960
         Left            =   150
         TabIndex        =   43
         Top             =   150
         Width           =   16305
         Begin VB.TextBox DescCompleta 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1140
            MaxLength       =   50
            TabIndex        =   130
            Top             =   915
            Width           =   5460
         End
         Begin MSMask.MaskEdBox TotalMoedaReal 
            Height          =   228
            Left            =   6228
            TabIndex        =   126
            Top             =   1260
            Width           =   1080
            _ExtentX        =   1905
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
         Begin MSMask.MaskEdBox PrecoUnitarioMoedaReal 
            Height          =   228
            Left            =   5004
            TabIndex        =   127
            Top             =   1260
            Width           =   1080
            _ExtentX        =   1905
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
            Format          =   "#,##0.00###"
            PromptChar      =   " "
         End
         Begin VB.TextBox DescricaoProduto 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   1335
            MaxLength       =   50
            TabIndex        =   70
            Top             =   330
            Width           =   4000
         End
         Begin MSMask.MaskEdBox UnidadeMed 
            Height          =   225
            Left            =   2925
            TabIndex        =   71
            Top             =   240
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   5
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PrecoTotal 
            Height          =   225
            Left            =   7560
            TabIndex        =   72
            Top             =   270
            Width           =   1080
            _ExtentX        =   1905
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
            Left            =   6465
            TabIndex        =   73
            Top             =   270
            Width           =   1080
            _ExtentX        =   1905
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
            Format          =   "#,##0.00###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox QuantRecebida 
            Height          =   225
            Left            =   5115
            TabIndex        =   74
            Top             =   270
            Width           =   1260
            _ExtentX        =   2223
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
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   225
            Left            =   4095
            TabIndex        =   75
            Top             =   240
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
            Left            =   15
            TabIndex        =   76
            Top             =   330
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.TextBox Observacao 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   4950
            MaxLength       =   255
            TabIndex        =   62
            Top             =   2400
            Width           =   2445
         End
         Begin VB.ComboBox RecebForaFaixa 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "PedComprasEnvOcx.ctx":0004
            Left            =   2670
            List            =   "PedComprasEnvOcx.ctx":0006
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   2325
            Width           =   2235
         End
         Begin MSMask.MaskEdBox DataLimite 
            Height          =   225
            Left            =   2760
            TabIndex        =   63
            Top             =   1950
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox AliquotaICM 
            Height          =   225
            Left            =   5910
            TabIndex        =   64
            Top             =   1950
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   5
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
         Begin MSMask.MaskEdBox ValorIPI 
            Height          =   225
            Left            =   4890
            TabIndex        =   65
            Top             =   1950
            Width           =   960
            _ExtentX        =   1693
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
         Begin MSMask.MaskEdBox AliquotaIPI 
            Height          =   225
            Left            =   3930
            TabIndex        =   66
            Top             =   1950
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   5
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
         Begin MSMask.MaskEdBox ValorDesconto 
            Height          =   225
            Left            =   1695
            TabIndex        =   67
            Top             =   1950
            Width           =   1035
            _ExtentX        =   1826
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
            Left            =   735
            TabIndex        =   68
            Top             =   1950
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   6
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
         Begin MSMask.MaskEdBox PercentMaisReceb 
            Height          =   225
            Left            =   540
            TabIndex        =   69
            Top             =   2400
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   5
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
         Begin MSFlexGridLib.MSFlexGrid GridItens 
            Height          =   2340
            Left            =   180
            TabIndex        =   44
            Top             =   255
            Width           =   15915
            _ExtentX        =   28072
            _ExtentY        =   4128
            _Version        =   393216
            Rows            =   6
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
      Caption         =   "Frame6"
      Height          =   8205
      Index           =   1
      Left            =   165
      TabIndex        =   1
      Top             =   840
      Width           =   16665
      Begin VB.Frame Frame8 
         Caption         =   "Datas"
         Height          =   1035
         Left            =   225
         TabIndex        =   78
         Top             =   4620
         Width           =   9495
         Begin VB.Label Data 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1770
            TabIndex        =   95
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label DataAlteracao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1770
            TabIndex        =   94
            Top             =   645
            Width           =   1095
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Alteração:"
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
            Left            =   780
            TabIndex        =   93
            Top             =   705
            Width           =   885
         End
         Begin VB.Label DataEmissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6015
            TabIndex        =   92
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label29 
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
            Left            =   5160
            TabIndex        =   91
            Top             =   300
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Envio:"
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
            Left            =   5370
            TabIndex        =   90
            Top             =   705
            Width           =   555
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Registro:"
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
            Left            =   885
            TabIndex        =   89
            Top             =   300
            Width           =   780
         End
         Begin VB.Label DataEnvio 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6015
            TabIndex        =   88
            Top             =   645
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Cabeçalho"
         Height          =   4035
         Left            =   225
         TabIndex        =   77
         Top             =   435
         Width           =   9495
         Begin MSMask.MaskEdBox Codigo 
            Height          =   300
            Left            =   1755
            TabIndex        =   115
            Top             =   255
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin VB.Label LabelObsEmbalagem 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1740
            TabIndex        =   129
            Top             =   1965
            Width           =   2145
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Embalagem:"
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
            Left            =   645
            TabIndex        =   128
            Top             =   2025
            Width           =   1035
         End
         Begin VB.Label EmbalagemLabel 
            AutoSize        =   -1  'True
            Caption         =   "Embalagem:"
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
            Left            =   -10000
            TabIndex        =   125
            Top             =   2025
            Width           =   1035
         End
         Begin VB.Label Taxa 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6000
            TabIndex        =   124
            Top             =   1575
            Width           =   2175
         End
         Begin VB.Label Moeda 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1755
            TabIndex        =   123
            Top             =   1560
            Width           =   2145
         End
         Begin VB.Label TaxaLabel 
            AutoSize        =   -1  'True
            Caption         =   "Taxa:"
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
            Left            =   5445
            TabIndex        =   122
            Top             =   1620
            Width           =   495
         End
         Begin VB.Label MoedaLabel 
            AutoSize        =   -1  'True
            Caption         =   "Moeda:"
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
            Left            =   1050
            TabIndex        =   121
            Top             =   1620
            Width           =   645
         End
         Begin VB.Label Embalagem 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   -10000
            TabIndex        =   120
            Top             =   1965
            Width           =   2145
         End
         Begin VB.Label Label4 
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
            Height          =   195
            Index           =   0
            Left            =   555
            TabIndex        =   99
            Top             =   2415
            Width           =   1095
         End
         Begin VB.Label Observ 
            BorderStyle     =   1  'Fixed Single
            Height          =   1545
            Left            =   1755
            TabIndex        =   98
            Top             =   2370
            Width           =   6390
         End
         Begin VB.Label Contato 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6000
            TabIndex        =   97
            Top             =   1125
            Width           =   2145
         End
         Begin VB.Label Label21 
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
            Left            =   5205
            TabIndex        =   96
            Top             =   1185
            Width           =   735
         End
         Begin VB.Label CondPagto 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1755
            TabIndex        =   87
            Top             =   1125
            Width           =   2145
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
            Height          =   195
            Index           =   0
            Left            =   675
            TabIndex        =   86
            Top             =   750
            Width           =   1035
         End
         Begin VB.Label Label6 
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
            Index           =   0
            Left            =   5475
            TabIndex        =   85
            Top             =   750
            Width           =   465
         End
         Begin VB.Label Fornecedor 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1755
            TabIndex        =   84
            Top             =   690
            Width           =   2145
         End
         Begin VB.Label Filial 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6000
            TabIndex        =   83
            Top             =   690
            Width           =   2145
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Comprador:"
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
            Left            =   4965
            TabIndex        =   82
            Top             =   315
            Width           =   975
         End
         Begin VB.Label Comprador 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6000
            TabIndex        =   81
            Top             =   255
            Width           =   2145
         End
         Begin VB.Label CondPagtoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Cond Pagto:"
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
            Left            =   570
            TabIndex        =   80
            Top             =   1185
            Width           =   1065
         End
         Begin VB.Label CodigoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Nº Pedido:"
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
            Left            =   765
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   79
            Top             =   315
            Width           =   930
         End
      End
      Begin VB.CommandButton BotaoPedidosEnviados 
         Caption         =   "Pedidos Enviados"
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
         Left            =   315
         TabIndex        =   60
         Top             =   6045
         Width           =   2205
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   8205
      Index           =   6
      Left            =   270
      TabIndex        =   116
      Top             =   810
      Visible         =   0   'False
      Width           =   16455
      Begin VB.Frame Frame11 
         Caption         =   "Notas"
         Height          =   7920
         Left            =   30
         TabIndex        =   117
         Top             =   165
         Width           =   16305
         Begin VB.TextBox NotaPC 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   1260
            MaxLength       =   150
            TabIndex        =   118
            Top             =   570
            Width           =   14430
         End
         Begin MSFlexGridLib.MSFlexGrid GridNotas 
            Height          =   3915
            Left            =   210
            TabIndex        =   119
            Top             =   255
            Width           =   15855
            _ExtentX        =   27966
            _ExtentY        =   6906
            _Version        =   393216
            Rows            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   14250
      ScaleHeight     =   495
      ScaleWidth      =   2580
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   75
      Width           =   2640
      Begin VB.CommandButton BotaoEmail 
         Height          =   360
         Left            =   75
         Picture         =   "PedComprasEnvOcx.ctx":0008
         Style           =   1  'Graphical
         TabIndex        =   131
         ToolTipText     =   "Enviar email"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoImprimir 
         Height          =   360
         Left            =   577
         Picture         =   "PedComprasEnvOcx.ctx":09AA
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Imprimir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   2085
         Picture         =   "PedComprasEnvOcx.ctx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1581
         Picture         =   "PedComprasEnvOcx.ctx":0C2A
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoBaixar 
         Height          =   360
         Left            =   1079
         Picture         =   "PedComprasEnvOcx.ctx":115C
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Baixar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   8640
      Left            =   135
      TabIndex        =   0
      Top             =   480
      Width           =   16740
      _ExtentX        =   29528
      _ExtentY        =   15240
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Pedido"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Entrega"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Distribuição"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Bloqueios"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Notas"
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
Attribute VB_Name = "PedComprasEnvOcx"
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
Dim objGridItens As AdmGrid
Dim objGridDistribuicao As AdmGrid
Dim gcolItemPedido As Collection
Dim iFrameTipoDestinoAtual As Integer

Dim iGrid_Produto_Col As Integer
Dim iGrid_DescricaoProduto_Col As Integer
Dim iGrid_UnidadeMed_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_QuantRecebida_Col As Integer
Dim iGrid_PrecoUnitario_Col As Integer
Dim iGrid_PercentDesc_Col As Integer
Dim iGrid_Desconto_Col As Integer
Dim iGrid_PrecoTotal_Col As Integer
Dim iGrid_PrecoUnitarioMoedaReal_Col As Integer
Dim iGrid_TotalMoedaReal_Col As Integer
Dim iGrid_DataLimite_Col As Integer
Dim iGrid_AliquotaIPI_Col As Integer
Dim iGrid_ValorIPIItem_Col As Integer
Dim iGrid_AliquotaICMS_Col As Integer
Dim iGrid_PercentMaisReceb_Col As Integer
Dim iGrid_RecebForaFaixa_Col As Integer
Dim iGrid_Observacao_Col As Integer
Dim iGrid_DescCompleta_Col As Integer 'leo

Dim iGrid_Prod_Col As Integer
Dim iGrid_DescProduto_Col As Integer
Dim iGrid_CentroCusto_Col As Integer
Dim iGrid_Almoxarifado_Col As Integer
Dim iGrid_UM_Col As Integer
Dim iGrid_Quant_Col As Integer
Dim iGrid_ContaContabil_Col As Integer

Dim objGridBloqueio As AdmGrid
Dim iGrid_TipoBloqueio_Col As Integer
Dim iGrid_DataBloqueio_Col As Integer
Dim iGrid_CodUsuario_Col As Integer
Dim iGrid_ResponsavelBL_Col As Integer
Dim iGrid_DataLiberacao_Col As Integer
Dim iGrid_ResponsavelLib_Col As Integer

Dim objGridNotas As AdmGrid
Dim iGrid_NotaPC_Col As Integer


Dim bExibirColReal As Boolean

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoBotaoPedidosEnviados As AdmEvento
Attribute objEventoBotaoPedidosEnviados.VB_VarHelpID = -1

Private Sub BotaoBaixar_Click()

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoBaixar

    GL_objMDIForm.MousePointer = vbHourglass
    
    'se o numero do pedido nao estiver preenchido ---> erro
    If Len(Trim(Codigo.Text)) = 0 Then Error 53145

    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then Error 53146

    've se o Pedido de Compra foi baixado
    lErro = CF("PedidoCompraBaixado_Le_Numero", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56137 Then Error 53147
    'se foi ---> erro
    If lErro = SUCESSO Then Error 53148

    'procura na tabela de Pedido de Compras
    lErro = CF("PedidoCompra_Le_Numero", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56142 Then Error 53149

    'se nao encontrar ---> erro
    If lErro = 56142 Then Error 53150

    'Pede a confirmação da baixa do pedido
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_BAIXA_PEDIDOCOMPRAS", objPedidoCompra.lCodigo)

    If vbMsgRes = vbNo Then Error 62649

    'baixa o Pedido de Compras
    lErro = CF("PedidoCompra_Baixar", objPedidoCompra)
    If lErro <> SUCESSO Then Error 53151

    ' ok Você deve limpar todos os campos da tela não só alguns
    Call Limpa_Tela_PedidoCompras

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoBaixar:

    Select Case Err

        Case 53145
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PEDIDO_NAO_PREENCHIDO", Err)

        Case 53146, 53147, 53149, 53151, 62649

        Case 53148
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDCOMPRA_BAIXADO", Err, objPedidoCompra.lCodigo)

        Case 53150
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", Err, objPedidoCompra.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164373)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim objComprador As New ClassComprador

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    
    bExibirColReal = True
    
    '################################
    'Inserido por Wagner
    Call Formata_Controles
    '################################

    Set objEventoCodigo = New AdmEvento
    Set objEventoBotaoPedidosEnviados = New AdmEvento
    
    objComprador.sCodUsuario = gsUsuario
    
    'Verifica se gsUsuario e comprador
    lErro = CF("Comprador_Le_Usuario", objComprador)
    If lErro <> SUCESSO And lErro <> 50059 Then Error 53095
    If lErro = 50059 Then Error 53096

    'Inicializa mascara do produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then Error 53087

    ' carrega a combo RecebForaFaixa
    lErro = Carrega_RecebForaFaixa()
    If lErro <> SUCESSO Then Error 53092

    Set gcolItemPedido = New Collection

    Set objGridItens = New AdmGrid
    Set objGridDistribuicao = New AdmGrid
    Set objGridBloqueio = New AdmGrid
    Set objGridNotas = New AdmGrid

    'Faz a inicializacao do grid itens
    lErro = Inicializa_Grid_Itens(objGridItens)
    If lErro <> SUCESSO Then Error 53099

    'Faz a inicializacao do grid distribuicao
    lErro = Inicializa_Grid_Distribuicao(objGridDistribuicao)
    If lErro <> SUCESSO Then Error 53100
    
    'Faz a inicializacao do grid bloqueio
    lErro = Inicializa_Grid_Bloqueios(objGridBloqueio)
    If lErro <> SUCESSO Then Error 53187
    
    lErro = Inicializa_Grid_Notas(objGridNotas)
    If lErro <> SUCESSO Then Error 53187
    
    'Carrega a combo de Tipos de Bloqueio
    lErro = Carrega_TipoBloqueio()
    If lErro <> SUCESSO Then Error 53178
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 53087, 53092, 53095, 53099, 53100, 53187, 53178, 53097
        
        Case 53096
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_COMPRADOR", Err, objComprador.sCodUsuario)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164374)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Itens
    
    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Quant Recebida")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Total")
    'Se a moeda for Diferente de Real => Exibe as Colunas de Comparacao
    If bExibirColReal = True Then
        objGridInt.colColuna.Add ("Preço (R$)")
        objGridInt.colColuna.Add ("Total (R$)")
    End If
    objGridInt.colColuna.Add ("Data Limite")
    objGridInt.colColuna.Add ("Alíquota IPI")
    objGridInt.colColuna.Add ("Valor IPI ")
    objGridInt.colColuna.Add ("Alíquota ICMS")
    objGridInt.colColuna.Add ("% a Mais Receb")
    objGridInt.colColuna.Add ("Ação Receb Fora Faixa")
    objGridInt.colColuna.Add ("Observação")
    objGridInt.colColuna.Add ("Desc. Completa") 'leo

    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoProduto.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (QuantRecebida.Name)
    objGridInt.colCampo.Add (PrecoUnitario.Name)
    objGridInt.colCampo.Add (PercentDesc.Name)
    objGridInt.colCampo.Add (ValorDesconto.Name)
    objGridInt.colCampo.Add (PrecoTotal.Name)
    'Se a moeda for Diferente de Real => Exibe as Colunas de Comparacao
    If bExibirColReal = True Then
        objGridInt.colCampo.Add (PrecoUnitarioMoedaReal.Name)
        objGridInt.colCampo.Add (TotalMoedaReal.Name)
    End If
    objGridInt.colCampo.Add (DataLimite.Name)
    objGridInt.colCampo.Add (AliquotaIPI.Name)
    objGridInt.colCampo.Add (ValorIPI.Name)
    objGridInt.colCampo.Add (AliquotaICM.Name)
    objGridInt.colCampo.Add (PercentMaisReceb.Name)
    objGridInt.colCampo.Add (RecebForaFaixa.Name)
    objGridInt.colCampo.Add (Observacao.Name)
    objGridInt.colCampo.Add (DescCompleta.Name) 'leo

   'indica onde estao situadas as colunas do grid
    iGrid_Produto_Col = 1
    iGrid_DescProduto_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_Quantidade_Col = 4
    iGrid_QuantRecebida_Col = 5
    iGrid_PrecoUnitario_Col = 6
    iGrid_PercentDesc_Col = 7
    iGrid_Desconto_Col = 8
    iGrid_PrecoTotal_Col = 9
    
    If bExibirColReal = True Then
        
        iGrid_PrecoUnitarioMoedaReal_Col = 10
        iGrid_TotalMoedaReal_Col = 11
        iGrid_DataLimite_Col = 12
        iGrid_AliquotaIPI_Col = 13
        iGrid_ValorIPIItem_Col = 14
        iGrid_AliquotaICMS_Col = 15
        iGrid_PercentMaisReceb_Col = 16
        iGrid_RecebForaFaixa_Col = 17
        iGrid_Observacao_Col = 18
        iGrid_DescCompleta_Col = 19
        
    Else
    
        iGrid_DataLimite_Col = 10
        iGrid_AliquotaIPI_Col = 11
        iGrid_ValorIPIItem_Col = 12
        iGrid_AliquotaICMS_Col = 13
        iGrid_PercentMaisReceb_Col = 14
        iGrid_RecebForaFaixa_Col = 15
        iGrid_Observacao_Col = 16
        iGrid_DescCompleta_Col = 17
        
    End If

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridItens

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_PEDIDO_COMPRAS + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 18

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'proibido incluir e excluir linhas
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function
 Private Function Inicializa_Grid_Distribuicao(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Distribuicao

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Centro de Custo")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Unidade Medida")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Conta Contábil")

    objGridInt.colCampo.Add (Prod.Name)
    objGridInt.colCampo.Add (DescProduto.Name)
    objGridInt.colCampo.Add (CentroCusto.Name)
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (UM.Name)
    objGridInt.colCampo.Add (Quant.Name)
    objGridInt.colCampo.Add (ContaContabil.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_Prod_Col = 1
    iGrid_DescProduto_Col = 2
    iGrid_CentroCusto_Col = 3
    iGrid_Almoxarifado_Col = 4
    iGrid_UM_Col = 5
    iGrid_Quant_Col = 6
    iGrid_ContaContabil_Col = 7

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridDistribuicao

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_DISTRIBUICAO + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 28

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'proibido incluir e excluir linhas
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Distribuicao = SUCESSO

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_Tela_Extrai

    sTabela = "PedidoCompra_Fornecedor"
    
    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then Error 53101

    'Preenche a coleção colCampoValor
    colCampoValor.Add "Codigo", objPedidoCompra.lCodigo, 0, "Codigo"
    colCampoValor.Add "OutrasDespesas", objPedidoCompra.dOutrasDespesas, 0, "OutrasDespesas"
    colCampoValor.Add "Data", objPedidoCompra.dtData, 0, "Data"
    colCampoValor.Add "DataAlteracao", objPedidoCompra.dtDataAlteracao, 0, "DataAlteracao"
    colCampoValor.Add "DataEnvio", objPedidoCompra.dtDataEnvio, 0, "DataEnvio"
    colCampoValor.Add "DataEmissao", objPedidoCompra.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "ValorDesconto", objPedidoCompra.dValorDesconto, 0, "ValorDesconto"
    colCampoValor.Add "ValorFrete", objPedidoCompra.dValorFrete, 0, "ValorFrete"
    colCampoValor.Add "ValorIPI", objPedidoCompra.dValorIPI, 0, "ValorIPI"
    colCampoValor.Add "ValorSeguro", objPedidoCompra.dValorSeguro, 0, "ValorSeguro"
    colCampoValor.Add "ValorTotal", objPedidoCompra.dValorTotal, 0, "ValorTotal"
    colCampoValor.Add "CondicaoPagto", objPedidoCompra.iCondicaoPagto, 0, "CondicaoPagto"
    colCampoValor.Add "Filial", objPedidoCompra.iFilial, 0, "Filial"
    colCampoValor.Add "FilialDestino", objPedidoCompra.iFilialDestino, 0, "FilialDestino"
    colCampoValor.Add "FilialEmpresa", objPedidoCompra.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "ProxSeqBloqueio", objPedidoCompra.iProxSeqBloqueio, 0, "ProxSeqBloqueio"
    colCampoValor.Add "TipoBaixa", objPedidoCompra.iTipoBaixa, 0, "TipoBaixa"
    colCampoValor.Add "TipoDestino", objPedidoCompra.iTipoDestino, 0, "TipoDestino"
    colCampoValor.Add "FornCliDestino", objPedidoCompra.lFornCliDestino, 0, "FornCliDestino"
    colCampoValor.Add "Fornecedor", objPedidoCompra.lFornecedor, 0, "Fornecedor"
    colCampoValor.Add "NumIntDoc", objPedidoCompra.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "Transportadora", objPedidoCompra.iTransportadora, 0, "Transportadora"
    colCampoValor.Add "Alcada", objPedidoCompra.sAlcada, STRING_BUFFER_MAX_TEXTO, "Alcada"
    colCampoValor.Add "Contato", objPedidoCompra.sContato, STRING_BUFFER_MAX_TEXTO, "Contato"
    colCampoValor.Add "MotivoBaixa", objPedidoCompra.sMotivoBaixa, STRING_BUFFER_MAX_TEXTO, "MotivoBaixa"
    colCampoValor.Add "Observacao", objPedidoCompra.lObservacao, 0, "Observacao"
    colCampoValor.Add "TipoFrete", objPedidoCompra.sTipoFrete, STRING_BUFFER_MAX_TEXTO, "TipoFrete"
    colCampoValor.Add "Embalagem", objPedidoCompra.iEmbalagem, 0, "Embalagem"
    colCampoValor.Add "Taxa", objPedidoCompra.dTaxa, 0, "Taxa"
    colCampoValor.Add "Moeda", objPedidoCompra.iMoeda, 0, "Moeda"
    colCampoValor.Add "Comprador", objPedidoCompra.iComprador, 0, "Comprador"
    colCampoValor.Add "ObsEmbalagem", objPedidoCompra.sObsEmbalagem, STRING_BUFFER_MAX_TEXTO, "ObsEmbalagem"

    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "DataEnvio", OP_DIFERENTE, DATA_NULA

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 53101

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164375)

    End Select

    Exit Sub

End Sub

Private Sub BotaoEmail_Click()

Dim lErro As Long, objBloqueioPC As ClassBloqueioPC
Dim objPedidoCompra As New ClassPedidoCompras
Dim objRelatorio As New AdmRelatorio
Dim sMailTo As String, sFiltro As String
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objEndereco As New ClassEndereco, sInfoEmail As String

On Error GoTo Erro_BotaoEmail_Click

    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then gError 53141

    If objPedidoCompra.lCodigo = 0 Then gError 76053

    'Verifica se o Pedido de Compra informado existe
    lErro = CF("PedidoCompra_Le_Numero", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56142 Then gError 76030
    
    'Se o Pedido não existe ==> erro
    If lErro = 56142 Then gError 76031
        
    lErro = CF("BloqueiosPC_Le", objPedidoCompra)
    If lErro <> SUCESSO Then gError 76057

    For Each objBloqueioPC In objPedidoCompra.colBloqueiosPC
            
        If objBloqueioPC.dtDataLib = DATA_NULA Then gError 76051
    
    Next
    
    If objPedidoCompra.lFornecedor <> 0 And objPedidoCompra.iFilial <> 0 Then

        objFilialFornecedor.lCodFornecedor = objPedidoCompra.lFornecedor
        objFilialFornecedor.iCodFilial = objPedidoCompra.iFilial

        lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 12929 Then gError 129314
         
        If lErro = SUCESSO Then
        
            objEndereco.lCodigo = objFilialFornecedor.lEndereco
            
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO Then gError 129315
        
            sMailTo = objEndereco.sEmail
            
        End If
        
        sInfoEmail = "Fornecedor: " & CStr(objFilialFornecedor.lCodFornecedor) & " - " & Fornecedor.Caption & " . Filial: " & Filial.Caption
        
    End If
    
    If Len(Trim(sMailTo)) = 0 Then gError 129316
    
    'Preenche a Data de Entrada com a Data Atual
    DataEmissao.Caption = Format(gdtDataHoje, "dd/mm/yyyy")

    objPedidoCompra.dtDataEmissao = gdtDataHoje

    'Atualiza data de emissao no BD para a data atual
    lErro = CF("PedidoCompra_Atualiza_DataEmissao", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56348 Then gError 53142

    'se nao encontrar ---> erro
    If lErro = 56348 Then gError 53143

    sFiltro = "REL_PCOM.PC_NumIntDoc = @NPEDCOM"
    lErro = CF("Relatorio_ObterFiltro", "Pedido de Compra Enviado", sFiltro)
    If lErro <> SUCESSO Then gError 76032
    
    'Executa o relatório
    lErro = objRelatorio.ExecutarDiretoEmail("Pedido de Compra Enviado", sFiltro, 0, "PEDCOM", "NPEDCOM", objPedidoCompra.lNumIntDoc, "TTO_EMAIL", sMailTo, "TSUBJECT", "Pedido de Compra " & CStr(objPedidoCompra.lCodigo), "TALIASATTACH", "PedCompra" & CStr(objPedidoCompra.lCodigo), "TINFO_EMAIL", sInfoEmail)
    If lErro <> SUCESSO Then gError 76032
     
    Exit Sub

Erro_BotaoEmail_Click:

    Select Case gErr

        Case 53141, 53142, 76057, 129314, 129315

        Case 53143
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", gErr, objPedidoCompra.lCodigo)

        Case 76030, 76032
        
        Case 76031
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", gErr, objPedidoCompra.lCodigo)
            
        Case 76051
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_BLOQUEADO", gErr, objPedidoCompra.lCodigo)
            
        Case 76053
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDCOMPRA_IMPRESSAO", gErr)
            
        Case 129316
            Call Rotina_Erro(vbOKOnly, "ERRO_EMAIL_NAO_ENCONTRADO", gErr, objPedidoCompra.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164376)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Activate()
    
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_Tela_Preenche

    'Carrega objPedidoCompra com os dados passados em colCampoValor
    objPedidoCompra.dOutrasDespesas = colCampoValor.Item("OutrasDespesas").vValor
    objPedidoCompra.dtData = colCampoValor.Item("Data").vValor
    objPedidoCompra.dtDataAlteracao = colCampoValor.Item("DataAlteracao").vValor
    objPedidoCompra.dtDataEmissao = colCampoValor.Item("DataEmissao").vValor
    objPedidoCompra.dtDataEnvio = colCampoValor.Item("DataEnvio").vValor
    objPedidoCompra.dValorDesconto = colCampoValor.Item("ValorDesconto").vValor
    objPedidoCompra.dValorFrete = colCampoValor.Item("ValorFrete").vValor
    objPedidoCompra.dValorIPI = colCampoValor.Item("ValorIPI").vValor
    objPedidoCompra.dValorSeguro = colCampoValor.Item("ValorSeguro").vValor
    objPedidoCompra.dValorTotal = colCampoValor.Item("ValorTotal").vValor
    objPedidoCompra.iComprador = colCampoValor.Item("Comprador").vValor
    objPedidoCompra.iCondicaoPagto = colCampoValor.Item("CondicaoPagto").vValor
    objPedidoCompra.iFilial = colCampoValor.Item("Filial").vValor
    objPedidoCompra.iFilialDestino = colCampoValor.Item("FilialDestino").vValor
    objPedidoCompra.iProxSeqBloqueio = colCampoValor.Item("ProxSeqBloqueio").vValor
    objPedidoCompra.iTipoBaixa = colCampoValor.Item("TipoBaixa").vValor
    objPedidoCompra.iTipoDestino = colCampoValor.Item("TipoDestino").vValor
    objPedidoCompra.lCodigo = colCampoValor.Item("Codigo").vValor
    objPedidoCompra.lFornCliDestino = colCampoValor.Item("FornCliDestino").vValor
    objPedidoCompra.lFornecedor = colCampoValor.Item("Fornecedor").vValor
    objPedidoCompra.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor
    objPedidoCompra.iTransportadora = colCampoValor.Item("Transportadora").vValor
    objPedidoCompra.sAlcada = colCampoValor.Item("Alcada").vValor
    objPedidoCompra.sContato = colCampoValor.Item("Contato").vValor
    objPedidoCompra.sMotivoBaixa = colCampoValor.Item("MotivoBaixa").vValor
    objPedidoCompra.lObservacao = colCampoValor.Item("Observacao").vValor
    objPedidoCompra.sTipoFrete = colCampoValor.Item("TipoFrete").vValor
    objPedidoCompra.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
    objPedidoCompra.iEmbalagem = colCampoValor.Item("Embalagem").vValor
    objPedidoCompra.iMoeda = colCampoValor.Item("Moeda").vValor
    objPedidoCompra.dTaxa = colCampoValor.Item("Taxa").vValor
    objPedidoCompra.sObsEmbalagem = colCampoValor.Item("ObsEmbalagem").vValor

    ' preenche a tela com os elementos do objPedidoCompra
    lErro = Traz_PedidoCompra_Tela(objPedidoCompra)
    If lErro <> SUCESSO Then Error 53102

    iAlterado = 0

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 53102

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164377)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Function Traz_PedidoCompra_Tela(objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim objObservacao As New ClassObservacao
Dim objFornecedor As New ClassFornecedor
Dim objEndereco As New ClassEndereco
Dim objTransportadora As New ClassTransportadora
Dim objEmbalagem As New ClassEmbalagem
Dim objMoeda As New ClassMoedas
Dim objUsuarios As New ClassUsuarios
Dim objComprador As New ClassComprador

On Error GoTo Erro_Traz_PedidoCompra_Tela

    'Le o Pedido de Compra
    lErro = CF("PedidoCompras_Le", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56118 Then gError 53136
    If lErro = 56118 Then gError 53137
    
    If objPedidoCompra.dtDataEnvio = DATA_NULA Then gError 53103

    ' lê os itens do Pedido de compra
    lErro = CF("ItensPC_Le", objPedidoCompra)
    If lErro <> SUCESSO Then gError 53104

    'Le os Bloqueios do Pedido de Compra
    lErro = CF("BloqueiosPC_Le", objPedidoCompra)
    If lErro <> SUCESSO Then gError 49489
    
    lErro = CF("NotasPedCompras_Le", objPedidoCompra)
    If lErro <> SUCESSO Then gError 103355
    
    Call Limpa_Tela_PedidoCompras

    objComprador.iCodigo = objPedidoCompra.iComprador
    objComprador.iFilialEmpresa = objPedidoCompra.iFilialEmpresa
    
    lErro = CF("Comprador_Le", objComprador)
    If lErro <> SUCESSO And lErro <> 50064 Then gError 11111
    If lErro <> SUCESSO Then gError 22222
    
    objPedidoCompra.sComprador = objComprador.sCodUsuario
        
    objUsuarios.sCodUsuario = objPedidoCompra.sComprador

    'le  o usuário contido na tabela de Usuarios
    lErro = CF("Usuarios_Le", objUsuarios, False)
    If lErro <> SUCESSO And lErro <> 40832 Then Error 53097
    If lErro = 40832 Then Error 53098

    Comprador.Caption = objUsuarios.sNomeReduzido

    If objPedidoCompra.dTaxa > 0 Then Taxa.Caption = Format(objPedidoCompra.dTaxa, FORMATO_TAXA_CONVERSAO_MOEDA)
    
    If objPedidoCompra.iEmbalagem > 0 Then
        
        objEmbalagem.iCodigo = objPedidoCompra.iEmbalagem
        
        lErro = CF("Embalagem_Le", objEmbalagem)
        If lErro <> SUCESSO And lErro <> 82763 Then gError 103392
        
        If lErro = SUCESSO Then Embalagem.Caption = objEmbalagem.sSigla
               
    End If
        
    objMoeda.iCodigo = objPedidoCompra.iMoeda
    
    lErro = CF("Moedas_Le", objMoeda)
    If lErro <> SUCESSO And lErro <> 108821 Then gError 103393
    If lErro = SUCESSO Then Moeda.Caption = objMoeda.iCodigo & SEPARADOR & objMoeda.sNome


    'Se a moeda selecionada for = REAL
    If objMoeda.iCodigo = MOEDA_REAL Then
    
        'Limpa a cotacao
        Taxa.Caption = ""
        
        bExibirColReal = False
        
    Else
            
        bExibirColReal = True
    
    End If
    
    'Coloca os dados na tela
    Codigo.Text = objPedidoCompra.lCodigo
    Contato.Caption = objPedidoCompra.sContato

    Data.Caption = Format(objPedidoCompra.dtData, "dd/mm/yyyy")

    objFornecedor.lCodigo = objPedidoCompra.lFornecedor

    'Lê o Fornecedor
    lErro = CF("Fornecedor_Le", objFornecedor)
    If lErro <> SUCESSO And lErro <> 12729 Then gError 53105
    If lErro = 12729 Then gError 53106

    'Coloca o NomeReduzido do Fornecedor na tela
    Fornecedor.Caption = objFornecedor.sNomeReduzido

    'Passa o CodFornecedor e o CodFilial para o objfilialfornecedor
    objFilialFornecedor.lCodFornecedor = objPedidoCompra.lFornecedor
    objFilialFornecedor.iCodFilial = objPedidoCompra.iFilial

    'Lê o filialforncedor
    lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", objFornecedor.sNomeReduzido, objFilialFornecedor)
    If lErro <> SUCESSO And lErro <> 18272 Then gError 53107
    If lErro = 18272 Then gError 53108 'Se nao encontrou

    'Coloca a filial na tela
    Filial.Caption = objPedidoCompra.iFilial & SEPARADOR & objFilialFornecedor.sNome

   If objPedidoCompra.dtDataAlteracao <> DATA_NULA Then DataAlteracao.Caption = Format(objPedidoCompra.dtDataAlteracao, "dd/mm/yyyy")
   If objPedidoCompra.dtDataEmissao <> DATA_NULA Then DataEmissao.Caption = Format(objPedidoCompra.dtDataEmissao, "dd/mm/yyyy")

    'Preenche o TipoDestino
    TipoDestino(objPedidoCompra.iTipoDestino).Value = True

    If iFrameTipoDestinoAtual = TIPO_DESTINO_EMPRESA Then

        objFilialEmpresa.iCodFilial = objPedidoCompra.iFilialDestino

        'Lê a FilialEmpresa
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 53109
        If lErro = 27378 Then gError 53110

        'Coloca a FilialEmpresa na tela
        FilialEmpresa.Caption = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
        
        ' preencher endereço
        If objFilialEmpresa.objEnderecoEntrega.lCodigo <> 0 Then
            Call Preenche_Endereco(objFilialEmpresa.objEnderecoEntrega)
        Else
            Call Preenche_Endereco(objFilialEmpresa.objEndereco)
        End If

    ElseIf iFrameTipoDestinoAtual = TIPO_DESTINO_FORNECEDOR Then

        objFornecedor.lCodigo = objPedidoCompra.lFornCliDestino

        'Lê o Fornecedor
        lErro = CF("Fornecedor_Le", objFornecedor)
        If lErro <> SUCESSO And lErro <> 12729 Then gError 53111

        If lErro = 12729 Then gError 53112

        'Coloca o Fornecedor na tela.
        Fornec.Caption = objFornecedor.sNomeReduzido

        objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
        objFilialFornecedor.iCodFilial = objPedidoCompra.iFilialDestino

        'le a FilialFornecedor
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", objFornecedor.sNomeReduzido, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 53113
        If lErro = 18272 Then gError 53114 'Não encontrou

        'Coloca a Filial na tela
        FilialFornec.Caption = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome

        If objFornecedor.lEndereco > 0 Then

            objEndereco.lCodigo = objFornecedor.lEndereco

            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO And lErro <> 12309 Then gError 53166
            If lErro = 12309 Then gError 53167 'Se nao encontrou ---> erro

            ' preenche endereço
            Call Preenche_Endereco(objEndereco)
        End If

    End If

    If StrParaInt(objPedidoCompra.sTipoFrete) = TIPO_CIF Then
        TipoFrete.Caption = "CIF"
    Else
        TipoFrete.Caption = "FOB"
    End If

    If objPedidoCompra.iTransportadora <> 0 Then

        objTransportadora.iCodigo = objPedidoCompra.iTransportadora
        'le a transportadora
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 53170
        If lErro = 19250 Then gError 53171 'se nao encontrou ---> gErro

        Transportadora.Caption = objTransportadora.sNomeReduzido
    End If

    If objPedidoCompra.iCondicaoPagto <> 0 Then

        objCondicaoPagto.iCodigo = objPedidoCompra.iCondicaoPagto
        'lê a cond. de pagto
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 53119
        If lErro = 19205 Then gError 53120 'se nao encontrou --->erro

        CondPagto.Caption = objPedidoCompra.iCondicaoPagto & SEPARADOR & objCondicaoPagto.sDescReduzida
    End If

    DataEnvio.Caption = Format(objPedidoCompra.dtDataEnvio, "dd/mm/yyyy")

    'lê a observacao
    If objPedidoCompra.lObservacao > 0 Then

        objObservacao.lNumInt = objPedidoCompra.lObservacao

        lErro = CF("Observacao_Le", objObservacao)
        If lErro <> SUCESSO And lErro <> 53827 Then gError 53128
        If lErro = 53827 Then gError 53129

        Observ.Caption = objObservacao.sObservacao
    End If
    LabelObsEmbalagem.Caption = objPedidoCompra.sObsEmbalagem

    'If objPedidoCompra.dValorProdutos > 0 Then ValorProdutos.Caption = Format(objPedidoCompra.dValorProdutos, "standard")
    If objPedidoCompra.dValorFrete > 0 Then ValorFrete.Caption = Format(objPedidoCompra.dValorFrete, "standard")
    If objPedidoCompra.dValorSeguro > 0 Then ValorSeguro.Caption = Format(objPedidoCompra.dValorSeguro, "standard")
    If objPedidoCompra.dOutrasDespesas > 0 Then OutrasDespesas.Caption = Format(objPedidoCompra.dOutrasDespesas, "standard")
    If objPedidoCompra.dValorDesconto > 0 Then DescontoValor.Caption = Format(objPedidoCompra.dValorDesconto, "standard")
    If objPedidoCompra.dValorIPI > 0 Then IPIValor.Caption = Format(objPedidoCompra.dValorIPI, "standard")

    'preenche o Grid com os Ítens do Pedido Compra
    lErro = Preenche_Grid_Itens(objPedidoCompra)
    If lErro <> SUCESSO Then gError 53121

    ' preenche o Grid de distribuicao atraves do objPedidoCompra
    lErro = Preenche_Grid_Distribuicao(objPedidoCompra)
    If lErro <> SUCESSO Then gError 53122

    'Preenche o GridBloqueio
    lErro = Preenche_Grid_Bloqueio(objPedidoCompra)
    If lErro <> SUCESSO Then gError 56022
            
    'preenche o Grid com as Notas do Pedido Compra
    lErro = Preenche_Grid_Notas(objPedidoCompra)
    If lErro <> SUCESSO Then gError 103346
    
    'preenche o campo ValorTotal e ValorProdutos
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 53124

    iAlterado = 0

    Traz_PedidoCompra_Tela = SUCESSO

    Exit Function

Erro_Traz_PedidoCompra_Tela:

    Traz_PedidoCompra_Tela = gErr

    Select Case gErr

        Case 49489, 53097, 53104, 53105, 53107, 53109, 53111, 53113, 53115, 53117, _
             53119, 53121, 53122, 53124, 53128, 53166, 53168, 53170, 56022, 53136, _
             103346, 103392, 103393, 103355
        
        Case 53098
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", gErr, objUsuarios.sCodUsuario)

        Case 53103
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_ENVIADO", gErr, objPedidoCompra.lCodigo)

        Case 53106
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objFornecedor.lCodigo)

        Case 53110
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)

        Case 53112
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO_2", gErr)

        Case 53108, 53114
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FORNECEDOR_INEXISTENTE", gErr, objFilialFornecedor.lCodFornecedor, objFilialFornecedor.iCodFilial)

        Case 53120
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)

        Case 53129
            Call Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_NAO_CADASTRADA", gErr, objPedidoCompra.lObservacao)
            
        Case 53137
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", gErr, objPedidoCompra.lCodigo)

        Case 53167, 53169
            Call Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO", gErr)

        Case 53171
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, objTransportadora.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164378)

    End Select

    Exit Function

End Function

Private Function Preenche_Grid_Itens(objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim dPercDesc As Double
Dim iItem As Integer
Dim objItemPC As New ClassItemPedCompra
Dim dPrecoTotal As Double, objProduto As New ClassProduto
Dim objObservacao As New ClassObservacao

On Error GoTo Erro_Preenche_Grid_Itens

    Set gcolItemPedido = New Collection

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridItens)

    iIndice = 0

    For Each objItemPC In objPedidoCompra.colItens

        iIndice = iIndice + 1

        lErro = Mascara_RetornaProdutoEnxuto(objItemPC.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then Error 53123

        'Mascara o produto enxuto
        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True

        'Calcula o percentual de desconto
        dPercDesc = objItemPC.dValorDesconto / (objItemPC.dPrecoUnitario * objItemPC.dQuantidade)
        dPrecoTotal = objItemPC.dPrecoUnitario * objItemPC.dQuantidade - objItemPC.dValorDesconto

        'Coloca os dados dos itens na tela
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text

        GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col) = objItemPC.sDescProduto
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemPC.sUM
        If objItemPC.dQuantRecebida > 0 Then GridItens.TextMatrix(iIndice, iGrid_QuantRecebida_Col) = Formata_Estoque(objItemPC.dQuantRecebida)
        If objItemPC.dQuantidade > 0 Then GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemPC.dQuantidade)
        If objItemPC.dPrecoUnitario > 0 Then GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col) = Format(objItemPC.dPrecoUnitario, gobjCOM.sFormatoPrecoUnitario) ' "STANDARD") 'Alterado por Wagner
        If dPercDesc > 0 Then GridItens.TextMatrix(iIndice, iGrid_PercentDesc_Col) = Format(dPercDesc, "Percent")
        If objItemPC.dValorDesconto > 0 Then GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(objItemPC.dValorDesconto, "Standard")
        If dPrecoTotal > 0 Then GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col) = Format(dPrecoTotal, PrecoTotal.Format) 'Alterado por Wagner
        
        If objItemPC.dtDataLimite <> DATA_NULA Then GridItens.TextMatrix(iIndice, iGrid_DataLimite_Col) = Format(objItemPC.dtDataLimite, "dd/mm/yyyy")

        If objItemPC.dPercentMaisReceb > 0 Then GridItens.TextMatrix(iIndice, iGrid_PercentMaisReceb_Col) = Format(objItemPC.dPercentMaisReceb, "Percent")
        If objItemPC.dAliquotaIPI > 0 Then GridItens.TextMatrix(iIndice, iGrid_AliquotaIPI_Col) = Format(objItemPC.dAliquotaIPI, "Percent")
        If objItemPC.dAliquotaICMS > 0 Then GridItens.TextMatrix(iIndice, iGrid_AliquotaICMS_Col) = Format(objItemPC.dAliquotaICMS, "Percent")

        'lê a observacao
        If objItemPC.lObservacao > 0 Then

            objObservacao.lNumInt = objItemPC.lObservacao

            lErro = CF("Observacao_Le", objObservacao)
            If lErro <> SUCESSO And lErro <> 53827 Then Error 53090
            If lErro = 53827 Then Error 53091

            GridItens.TextMatrix(iIndice, iGrid_Observacao_Col) = objObservacao.sObservacao

        End If

        For iItem = 0 To RecebForaFaixa.ListCount - 1
            If objItemPC.iRebebForaFaixa = RecebForaFaixa.ItemData(iItem) Then
                'coloca no Grid Itens RecebForaFaixa
                GridItens.TextMatrix(iIndice, iGrid_RecebForaFaixa_Col) = RecebForaFaixa.List(iItem)
            End If
        Next

        If objItemPC.dValorIPI > 0 Then GridItens.TextMatrix(iIndice, iGrid_ValorIPIItem_Col) = objItemPC.dValorIPI

        'Le o produto
        objProduto.sCodigo = objItemPC.sProduto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 56381
        'Se nao encontrou => erro
        If lErro = 28030 Then Error 56437
        
        'Preenche a descrição completa do produto com a ObsFisica do produto na tabela de produtos
        GridItens.TextMatrix(iIndice, iGrid_DescCompleta_Col) = objProduto.sObsFisica
        
        'Armazena os números internos dos itens
        gcolItemPedido.Add objItemPC.lNumIntDoc

    Next

    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = gcolItemPedido.Count
    
    If objPedidoCompra.iMoeda = MOEDA_REAL Then
        Call ComparativoMoedaReal_Calcula(1)
    ElseIf Len(Trim(Taxa.Caption)) > 0 Then
        Call ComparativoMoedaReal_Calcula(CDbl(Taxa.Caption))
    End If

    Exit Function

Erro_Preenche_Grid_Itens:

    Preenche_Grid_Itens = Err

    Select Case Err

        Case 56437
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)
        
        Case 53090, 53123, 56381

        Case 53091
            Call Rotina_Erro(vbOKOnly, "ERRO_OBSERVACAO_NAO_CADASTRADA", Err, objPedidoCompra.lObservacao)


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164379)

    End Select

    Exit Function

End Function

Private Function Preenche_Grid_Distribuicao(objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iItem As Integer
Dim objItemPC As New ClassItemPedCompra
Dim objLocalizacao As New ClassLocalizacaoItemPC
Dim sCclMascarado As String
Dim sContaMascarada As String
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Preenche_Grid_Distribuicao

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridDistribuicao)

    iIndice = 0
    iItem = 0

    For Each objItemPC In objPedidoCompra.colItens

        iItem = iItem + 1

        For Each objLocalizacao In objItemPC.colLocalizacao

            iIndice = iIndice + 1

            'Coloca os dados de distribuicao na tela
            GridDistribuicao.TextMatrix(iIndice, iGrid_Prod_Col) = GridItens.TextMatrix(iItem, iGrid_Produto_Col)
            GridDistribuicao.TextMatrix(iIndice, iGrid_Quant_Col) = Formata_Estoque(objLocalizacao.dQuantidade)
            GridDistribuicao.TextMatrix(iIndice, iGrid_DescProduto_Col) = GridItens.TextMatrix(iItem, iGrid_DescProduto_Col)
            GridDistribuicao.TextMatrix(iIndice, iGrid_UM_Col) = GridItens.TextMatrix(iItem, iGrid_UnidadeMed_Col)

            If Len(Trim(objLocalizacao.sCcl)) > 0 Then
                lErro = Mascara_MascararCcl(objLocalizacao.sCcl, sCclMascarado)
                If lErro <> SUCESSO Then Error 53125

                GridDistribuicao.TextMatrix(iIndice, iGrid_CentroCusto_Col) = sCclMascarado

            End If

            If objLocalizacao.sContaContabil <> "" Then

                lErro = Mascara_MascararConta(objLocalizacao.sContaContabil, sContaMascarada)
                If lErro <> SUCESSO Then Error 53126

                GridDistribuicao.TextMatrix(iIndice, iGrid_ContaContabil_Col) = sContaMascarada

            End If

            If (objLocalizacao.iAlmoxarifado) > 0 Then
                objAlmoxarifado.iCodigo = objLocalizacao.iAlmoxarifado

                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                If lErro <> SUCESSO Then Error 53127
                GridDistribuicao.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
            End If
        Next

    Next

    Preenche_Grid_Distribuicao = SUCESSO

    Exit Function

Erro_Preenche_Grid_Distribuicao:

    Preenche_Grid_Distribuicao = Err

    Select Case Err

        Case 53125, 53126, 53127

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164380)

    End Select

    Exit Function

End Function

Private Sub CodigoLabel_Click()

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras
Dim colSelecao As New Collection

On Error GoTo Erro_CodigoLabel_Click

    'Move os dados da tela
    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then Error 53130

    'Chama a Tela de browse
    Call Chama_Tela("PedComprasEnvLista", colSelecao, objPedidoCompra, objEventoCodigo)

    Exit Sub

Erro_CodigoLabel_Click:

    Select Case Err

        Case 53130

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164381)

    End Select

    Exit Sub

End Sub

Private Sub objEventoBotaoPedidosEnviados_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_objEventoBotaoPedidosEnviados_evSelecao

    Set objPedidoCompra = obj1

    lErro = Traz_PedidoCompra_Tela(objPedidoCompra)
    If lErro <> SUCESSO Then Error 53133

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoBotaoPedidosEnviados_evSelecao:

    Select Case Err

        Case 53133

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164382)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objPedidoCompra = obj1
    
    lErro = Traz_PedidoCompra_Tela(objPedidoCompra)
    If lErro <> SUCESSO Then Error 53131

    iAlterado = 0

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case Err

        Case 53131

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164383)

    End Select

    Exit Sub

End Sub

Private Sub BotaoPedidosEnviados_Click()

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoPedidosEnviados_Click

    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then Error 53132

    'Chama a tela PedComprasEnvLista
    Call Chama_Tela("PedComprasEnvLista", colSelecao, objPedidoCompra, objEventoBotaoPedidosEnviados)

    Exit Sub

Erro_BotaoPedidosEnviados_Click:

    Select Case Err

        Case 53132

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164384)

    End Select

    Exit Sub

End Sub

Private Function Move_Tela_Memoria(objPedidoCompra As ClassPedidoCompras) As Long

    'guarda a FilialEmpresa e o Codigo em objPedidoCompra
    objPedidoCompra.lCodigo = StrParaLong(Codigo.Text)
    objPedidoCompra.iFilialEmpresa = giFilialEmpresa

    Move_Tela_Memoria = SUCESSO

    Exit Function

End Function

Public Function Trata_Parametros(Optional objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma nota foi passada por parametro
    If Not (objPedidoCompra Is Nothing) Then

        If objPedidoCompra.lNumIntDoc > 0 Then

            lErro = Traz_PedidoCompra_Tela(objPedidoCompra)
            If lErro <> SUCESSO Then Error 53138

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 53138

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164385)

    End Select

    Exit Function

End Function

Private Sub TabStrip1_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

       If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual invisivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index

    End If

End Sub

Private Sub TipoDestino_Click(Index As Integer)

    If Index = iFrameTipoDestinoAtual Then Exit Sub

    'Torna Frame atual invisivel
    FrameTipo(iFrameTipoDestinoAtual).Visible = False

    'Torna Frame correspondente a Index visivel
    FrameTipo(Index).Visible = True

    'Armazena novo valor de iFrameTipoDestinoAtual
    iFrameTipoDestinoAtual = Index

    Call Limpa_Frame_Endereco

End Sub


Private Sub BotaoImprimir_Click()

Dim lErro As Long, objBloqueioPC As ClassBloqueioPC
Dim objPedidoCompra As New ClassPedidoCompras
Dim objRelatorio As New AdmRelatorio, sFiltro As String

On Error GoTo Erro_BotaoImprimir_Click

    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then gError 53141

    If objPedidoCompra.lCodigo = 0 Then gError 76053

    'Verifica se o Pedido de Compra informado existe
    lErro = CF("PedidoCompra_Le_Numero", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56142 Then gError 76030
    
    'Se o Pedido não existe ==> erro
    If lErro = 56142 Then gError 76031
        
    lErro = CF("BloqueiosPC_Le", objPedidoCompra)
    If lErro <> SUCESSO Then gError 76057

    For Each objBloqueioPC In objPedidoCompra.colBloqueiosPC
            
        If objBloqueioPC.dtDataLib = DATA_NULA Then gError 76051
    
    Next
    
    'Preenche a Data de Entrada com a Data Atual
    DataEmissao.Caption = Format(gdtDataHoje, "dd/mm/yyyy")

    objPedidoCompra.dtDataEmissao = gdtDataHoje

    'Atualiza data de emissao no BD para a data atual
    lErro = CF("PedidoCompra_Atualiza_DataEmissao", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56348 Then gError 53142

    'se nao encontrar ---> erro
    If lErro = 56348 Then gError 53143

    sFiltro = "REL_PCOM.PC_NumIntDoc = @NPEDCOM"
    lErro = CF("Relatorio_ObterFiltro", "Pedido de Compra Enviado", sFiltro)
    If lErro <> SUCESSO Then gError 76032
    
    'Executa o relatório
    lErro = objRelatorio.ExecutarDireto("Pedido de Compra Enviado", sFiltro, 0, "PEDCOM", "NPEDCOM", objPedidoCompra.lNumIntDoc)
    If lErro <> SUCESSO Then gError 76032
     
    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 53141, 53142, 76057

        Case 53143
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", gErr, objPedidoCompra.lCodigo)

        Case 76030, 76032
        
        Case 76031
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", gErr, objPedidoCompra.lCodigo)
            
        Case 76051
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_BLOQUEADO", gErr, objPedidoCompra.lCodigo)
            
        Case 76053
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDCOMPRA_IMPRESSAO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164386)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 53144

    'Limpa a tela
    Call Limpa_Tela_PedidoCompras

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 53144

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164387)

    End Select

    Exit Sub

End Sub

Sub Limpa_Tela_PedidoCompras()

    'Limpa a tela
    Call Limpa_Tela(Me)

    'Limpa os outros campos da tela
    Codigo.Text = ""
    Contato.Caption = ""
    Data.Caption = ""
    Fornecedor.Caption = ""
    Observ.Caption = ""
    DataEnvio = ""
    DataAlteracao.Caption = ""
    DataEmissao.Caption = ""
    Filial.Caption = ""
    CondPagto.Caption = ""
    ValorTotal.Caption = ""
    ValorFrete.Caption = ""
    ValorSeguro.Caption = ""
    ValorProdutos.Caption = ""
    OutrasDespesas.Caption = ""
    DescontoValor.Caption = ""
    IPIValor.Caption = ""
    FilialEmpresa.Caption = ""
    TipoFrete.Caption = ""
    Transportadora.Caption = ""
    Fornec.Caption = ""
    FilialFornec.Caption = ""
    Endereco.Caption = ""
    Taxa.Caption = ""
    Moeda.Caption = ""
    Embalagem.Caption = ""
    Comprador.Caption = ""

    Call Limpa_Frame_Endereco

    'Limpa os grids
    Call Grid_Limpa(objGridItens)
    Call Grid_Limpa(objGridDistribuicao)
    Call Grid_Limpa(objGridBloqueio)
    Call Grid_Limpa(objGridNotas)

    Set gcolItemPedido = New Collection

    Exit Sub

End Sub
'ja existe em PedidoCompra
Private Sub Limpa_Frame_Endereco()

    Endereco.Caption = ""
    Bairro.Caption = ""
    Cidade.Caption = ""
    CEP.Caption = ""
    Estado.Caption = ""
    Pais.Caption = ""

    Exit Sub

End Sub

Private Function Carrega_RecebForaFaixa() As Long

    'Limpa a combo
    RecebForaFaixa.Clear

    RecebForaFaixa.AddItem MENSAGEM_NAO_AVISA_ACEITA_RECEBIMENTO
    RecebForaFaixa.ItemData(RecebForaFaixa.NewIndex) = NAO_AVISA_E_ACEITA_RECEBIMENTO

    RecebForaFaixa.AddItem MENSAGEM_REJEITA_RECEBIMENTO
    RecebForaFaixa.ItemData(RecebForaFaixa.NewIndex) = ERRO_E_REJEITA_RECEBIMENTO

    RecebForaFaixa.AddItem MENSAGEM_ACEITA_RECEBIMENTO
    RecebForaFaixa.ItemData(RecebForaFaixa.NewIndex) = AVISA_E_ACEITA_RECEBIMENTO

    Exit Function

End Function

'ja existe em PedidoCompras
Private Function Inicializa_MascaraCcl() As Long
'Inicializa a mascara do centro de custo

Dim sMascaraCcl As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_mascaraccl

    sMascaraCcl = String(STRING_CCL, 0)

    'le a mascara dos centros de custo/lucro
    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then Error 49460

    CentroCusto.Mask = sMascaraCcl

    Inicializa_MascaraCcl = SUCESSO

    Exit Function

Erro_Inicializa_mascaraccl:

    Inicializa_MascaraCcl = Err

    Select Case Err

        Case 49460

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164388)

    End Select

    Exit Function

End Function

Private Sub Preenche_Endereco(objEndereco As ClassEndereco)

Dim objPais As New ClassPais
Dim lErro As Long

On Error GoTo Erro_Preenche_Endereco

    objPais.iCodigo = objEndereco.iCodigoPais

    lErro = CF("Paises_Le", objPais)
    If lErro <> SUCESSO And lErro <> 47876 Then Error 53088
    If lErro = 47876 Then Error 53089

    Endereco.Caption = objEndereco.sEndereco
    Bairro.Caption = objEndereco.sBairro
    Estado.Caption = objEndereco.sSiglaEstado
    Cidade.Caption = objEndereco.sCidade
    Pais.Caption = objPais.sNome
    CEP.Caption = objEndereco.sCEP

    Exit Sub

Erro_Preenche_Endereco:

    Select Case Err

        Case 53088

        Case 53089
            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO", Err, objEndereco.iCodigoPais)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164389)

    End Select

    Exit Sub

End Sub

Function ValorTotal_Calcula() As Long

Dim dPrecoTotal As Double
Dim dValorTotal As Double
Dim iIndice As Integer

On Error GoTo Erro_ValorTotal_Calcula

    For iIndice = 1 To objGridItens.iLinhasExistentes
        
        'Calcula a soma dos valores de produtos
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then

            If StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then
                dValorTotal = dValorTotal + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
            End If
            
        End If
        'Calcula Preco Total das linhas do GridItens
        dPrecoTotal = dPrecoTotal + (StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)))

    Next

    'Coloca na tela o valor dos produtos
    ValorProdutos.Caption = Format(dPrecoTotal, PrecoTotal.Format) 'Alterado por Wagner
    dValorTotal = (dPrecoTotal + StrParaDbl(ValorFrete.Caption) + StrParaDbl(ValorSeguro.Caption) + StrParaDbl(OutrasDespesas.Caption) + StrParaDbl(IPIValor.Caption)) - StrParaDbl(DescontoValor.Caption)

    'Coloca na tela o valor total
    ValorTotal.Caption = Format(dValorTotal, PrecoTotal.Format) 'Alterado por Wagner

    ValorTotal_Calcula = SUCESSO

    Exit Function

Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 164390)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    'libera as variaveis globais
    Set objEventoBotaoPedidosEnviados = Nothing
    Set objEventoCodigo = Nothing

    Set objGridItens = Nothing
    Set objGridDistribuicao = Nothing
    Set objGridBloqueio = Nothing
    Set objGridNotas = Nothing

    Set gcolItemPedido = Nothing

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Pedido de Compra Enviado - Consulta"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PedComprasCons"
    
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
    m_Caption = New_Caption
End Property

'**** fim do trecho a ser copiado *****
Private Sub FornecedorLabel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(FornecedorLabel(Index), Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel(Index), Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label6(Index), Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6(Index), Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label4(Index), Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4(Index), Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label15(Index), Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15(Index), Button, Shift, X, Y)
End Sub
Private Sub Data_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Data, Source, X, Y)
End Sub

Private Sub Data_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Data, Button, Shift, X, Y)
End Sub

Private Sub Codigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Codigo, Source, X, Y)
End Sub

Private Sub Codigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Codigo, Button, Shift, X, Y)
End Sub

Private Sub CondPagto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondPagto, Source, X, Y)
End Sub

Private Sub CondPagto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondPagto, Button, Shift, X, Y)
End Sub

Private Sub Contato_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Contato, Source, X, Y)
End Sub

Private Sub Contato_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Contato, Button, Shift, X, Y)
End Sub

Private Sub DataAlteracao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataAlteracao, Source, X, Y)
End Sub

Private Sub DataAlteracao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataAlteracao, Button, Shift, X, Y)
End Sub

Private Sub Label30_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label30, Source, X, Y)
End Sub

Private Sub Label30_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label30, Button, Shift, X, Y)
End Sub

Private Sub DataEmissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataEmissao, Source, X, Y)
End Sub

Private Sub DataEmissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataEmissao, Button, Shift, X, Y)
End Sub

Private Sub Label29_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label29, Source, X, Y)
End Sub

Private Sub Label29_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label29, Button, Shift, X, Y)
End Sub

Private Sub Fornecedor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Fornecedor, Source, X, Y)
End Sub

Private Sub Fornecedor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Fornecedor, Button, Shift, X, Y)
End Sub

Private Sub Filial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Filial, Source, X, Y)
End Sub

Private Sub Filial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Filial, Button, Shift, X, Y)
End Sub

Private Sub Label28_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label28, Source, X, Y)
End Sub

Private Sub Label28_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label28, Button, Shift, X, Y)
End Sub

Private Sub Comprador_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Comprador, Source, X, Y)
End Sub

Private Sub Comprador_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Comprador, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub CondPagtoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CondPagtoLabel, Source, X, Y)
End Sub

Private Sub CondPagtoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CondPagtoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub

Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub

Private Sub CodigoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodigoLabel, Source, X, Y)
End Sub

Private Sub CodigoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodigoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub DataEnvio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataEnvio, Source, X, Y)
End Sub

Private Sub DataEnvio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataEnvio, Button, Shift, X, Y)
End Sub

Private Sub Observ_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Observ, Source, X, Y)
End Sub

Private Sub Observ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Observ, Button, Shift, X, Y)
End Sub

Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label25, Source, X, Y)
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label25, Button, Shift, X, Y)
End Sub

Private Sub ValorTotal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorTotal, Source, X, Y)
End Sub

Private Sub ValorTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorTotal, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub

Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub

Private Sub IPIValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(IPIValor, Source, X, Y)
End Sub

Private Sub IPIValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(IPIValor, Button, Shift, X, Y)
End Sub

Private Sub OutrasDespesas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(OutrasDespesas, Source, X, Y)
End Sub

Private Sub OutrasDespesas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(OutrasDespesas, Button, Shift, X, Y)
End Sub

Private Sub ValorSeguro_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorSeguro, Source, X, Y)
End Sub

Private Sub ValorSeguro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorSeguro, Button, Shift, X, Y)
End Sub

Private Sub ValorFrete_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorFrete, Source, X, Y)
End Sub

Private Sub ValorFrete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorFrete, Button, Shift, X, Y)
End Sub

Private Sub Label41_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label41, Source, X, Y)
End Sub

Private Sub Label41_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label41, Button, Shift, X, Y)
End Sub

Private Sub ValorProdutos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorProdutos, Source, X, Y)
End Sub

Private Sub ValorProdutos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorProdutos, Button, Shift, X, Y)
End Sub

Private Sub DescontoValor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescontoValor, Source, X, Y)
End Sub

Private Sub DescontoValor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescontoValor, Button, Shift, X, Y)
End Sub

Private Sub TransportadoraLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TransportadoraLabel, Source, X, Y)
End Sub

Private Sub TransportadoraLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TransportadoraLabel, Button, Shift, X, Y)
End Sub

Private Sub TipoFrete_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoFrete, Source, X, Y)
End Sub

Private Sub TipoFrete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoFrete, Button, Shift, X, Y)
End Sub

Private Sub Label31_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label31, Source, X, Y)
End Sub

Private Sub Label31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label31, Button, Shift, X, Y)
End Sub

Private Sub Transportadora_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Transportadora, Source, X, Y)
End Sub

Private Sub Transportadora_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Transportadora, Button, Shift, X, Y)
End Sub

Private Sub Fornec_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Fornec, Source, X, Y)
End Sub

Private Sub Fornec_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Fornec, Button, Shift, X, Y)
End Sub

Private Sub Label32_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label32, Source, X, Y)
End Sub

Private Sub Label32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label32, Button, Shift, X, Y)
End Sub

Private Sub FilialFornec_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialFornec, Source, X, Y)
End Sub

Private Sub FilialFornec_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialFornec, Button, Shift, X, Y)
End Sub

Private Sub Label37_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label37, Source, X, Y)
End Sub

Private Sub Label37_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label37, Button, Shift, X, Y)
End Sub

Private Sub FilialEmpresa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialEmpresa, Source, X, Y)
End Sub

Private Sub FilialEmpresa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialEmpresa, Button, Shift, X, Y)
End Sub

Private Sub Pais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Pais, Source, X, Y)
End Sub

Private Sub Pais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Pais, Button, Shift, X, Y)
End Sub

Private Sub Estado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Estado, Source, X, Y)
End Sub

Private Sub Estado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Estado, Button, Shift, X, Y)
End Sub

Private Sub CEP_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CEP, Source, X, Y)
End Sub

Private Sub CEP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CEP, Button, Shift, X, Y)
End Sub

Private Sub Cidade_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Cidade, Source, X, Y)
End Sub

Private Sub Cidade_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Cidade, Button, Shift, X, Y)
End Sub

Private Sub Bairro_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Bairro, Source, X, Y)
End Sub

Private Sub Bairro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Bairro, Button, Shift, X, Y)
End Sub

Private Sub Endereco_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Endereco, Source, X, Y)
End Sub

Private Sub Endereco_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Endereco, Button, Shift, X, Y)
End Sub

Private Sub Label63_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label63, Source, X, Y)
End Sub

Private Sub Label63_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label63, Button, Shift, X, Y)
End Sub

Private Sub Label65_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label65, Source, X, Y)
End Sub

Private Sub Label65_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label65, Button, Shift, X, Y)
End Sub

Private Sub Label70_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label70, Source, X, Y)
End Sub

Private Sub Label70_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label70, Button, Shift, X, Y)
End Sub

Private Sub Label71_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label71, Source, X, Y)
End Sub

Private Sub Label71_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label71, Button, Shift, X, Y)
End Sub

Private Sub Label72_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label72, Source, X, Y)
End Sub

Private Sub Label72_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label72, Button, Shift, X, Y)
End Sub

Private Sub Label73_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label73, Source, X, Y)
End Sub

Private Sub Label73_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label73, Button, Shift, X, Y)
End Sub

Private Function Carrega_TipoBloqueio() As Long

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_TipoBloqueio

    'Lê o Código e o NOme de Todas os Tipos de Bloqueio do BD
    lErro = CF("Cod_Nomes_Le", "TiposDeBloqueioPC", "Codigo", "NomeReduzido", STRING_NOME_TABELA, colCodigoNome)
    If lErro <> SUCESSO Then Error 53181

    'Carrega a combo de Tipo de Bloqueio
    For Each objCodigoNome In colCodigoNome
        If objCodigoNome.iCodigo <> BLOQUEIO_ALCADA Then

            TipoBloqueio.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            TipoBloqueio.ItemData(TipoBloqueio.NewIndex) = objCodigoNome.iCodigo

        End If
    Next

    Carrega_TipoBloqueio = SUCESSO

    Exit Function

Erro_Carrega_TipoBloqueio:

    Carrega_TipoBloqueio = Err

    Select Case Err

        Case 53181

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164391)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Bloqueios(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Distribuicao

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Bloqueios

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Tipo Bloqueio")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Usuário")
    objGridInt.colColuna.Add ("Responsável")
    objGridInt.colColuna.Add ("Data Liberação")
    objGridInt.colColuna.Add ("Resp. Liberação")

    ' campos de edição do grid
    objGridInt.colCampo.Add (TipoBloqueio.Name)
    objGridInt.colCampo.Add (DataBloqueio.Name)
    objGridInt.colCampo.Add (CodUsuario.Name)
    objGridInt.colCampo.Add (ResponsavelBL.Name)
    objGridInt.colCampo.Add (DataLiberacao.Name)
    objGridInt.colCampo.Add (ResponsavelLib.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_TipoBloqueio_Col = 1
    iGrid_DataBloqueio_Col = 2
    iGrid_CodUsuario_Col = 3
    iGrid_ResponsavelBL_Col = 4
    iGrid_DataLiberacao_Col = 5
    iGrid_ResponsavelLib_Col = 6

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridBloqueios

    'Linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_BLOQUEIOS + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 20

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Bloqueios = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Bloqueios:

    Inicializa_Grid_Bloqueios = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 164392)

    End Select

    Exit Function

End Function

Private Function Preenche_Grid_Bloqueio(objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objBloqueioPC As New ClassBloqueioPC
Dim objTipoDeBloqueioPC As New ClassTipoBloqueioPC

On Error GoTo Erro_Preenche_Grid_Bloqueio

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridBloqueio)

    iIndice = 0

    For Each objBloqueioPC In objPedidoCompra.colBloqueiosPC

        iIndice = iIndice + 1

        GridBloqueios.TextMatrix(iIndice, iGrid_CodUsuario_Col) = objBloqueioPC.sCodUsuario
        GridBloqueios.TextMatrix(iIndice, iGrid_ResponsavelBL_Col) = objBloqueioPC.sResponsavel
        GridBloqueios.TextMatrix(iIndice, iGrid_ResponsavelLib_Col) = objBloqueioPC.sCodUsuarioLib

        objTipoDeBloqueioPC.iCodigo = objBloqueioPC.iTipoBloqueio

        lErro = CF("TipoDeBloqueioPC_Le", objTipoDeBloqueioPC)
        If lErro <> SUCESSO And lErro <> 49143 Then Error 57250
        If lErro = 49143 Then Error 57251

        GridBloqueios.TextMatrix(iIndice, iGrid_TipoBloqueio_Col) = objBloqueioPC.iTipoBloqueio & SEPARADOR & objTipoDeBloqueioPC.sNomeReduzido

        If objBloqueioPC.dtDataLib <> DATA_NULA Then GridBloqueios.TextMatrix(iIndice, iGrid_DataLiberacao_Col) = Format(objBloqueioPC.dtDataLib, "dd/mm/yyyy")
        If (objBloqueioPC.dtData <> DATA_NULA) Then GridBloqueios.TextMatrix(iIndice, iGrid_DataBloqueio_Col) = Format(objBloqueioPC.dtData, "dd/mm/yyyy")
    Next

    objGridBloqueio.iLinhasExistentes = iIndice

    Preenche_Grid_Bloqueio = SUCESSO

    Exit Function

Erro_Preenche_Grid_Bloqueio:

    Preenche_Grid_Bloqueio = Err

    Select Case Err

        Case 57251
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODEBLOQUEIOPC_NAO_CADASTRADO", Err, objTipoDeBloqueioPC.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164393)

    End Select

    Exit Function

End Function

Private Sub BotaoLiberaBloqueio_Click()

Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras

On Error GoTo Erro_BotaoLiberaBloqueio_Click

    'Verifica se o número do Pedido de Compra está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub

    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objPedidoCompra)
    If lErro <> SUCESSO Then Error 49499

    'Chama tela LiberaBloqueioPC
    Call Chama_Tela("LiberaBloqueioPC", objPedidoCompra)

    Exit Sub

Erro_BotaoLiberaBloqueio_Click:

    Select Case Err

        Case 49499

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164394)

    End Select

    Exit Sub

End Sub



Private Sub ComparativoMoedaReal_Calcula(ByVal dTaxa As Double)
'Preenche as colunas INFORMATIVAS de proporção da moeda R$.

Dim iIndice As Integer

On Error GoTo Erro_ComparativoMoedaReal_Calcula

    'Para cada linha do grid de Itens será claculado o correspondente em R$
    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        'Preço Unitário em R$ = Preço Unitário na Moeda selecionada dividido pela taxa de conversão
        GridItens.TextMatrix(iIndice, iGrid_PrecoUnitarioMoedaReal_Col) = Format(StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col)) * dTaxa, gobjCOM.sFormatoPrecoUnitario) 'Alterado por Wagner
        
        'Preço Total em R$ = Preço Unitário em R$ x Quantidade do produto
        GridItens.TextMatrix(iIndice, iGrid_TotalMoedaReal_Col) = Format(StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) * dTaxa, TotalMoedaReal.Format) 'Alterado por Wagner
        
    Next

    Exit Sub
    
Erro_ComparativoMoedaReal_Calcula:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164395)

    End Select
    
    Exit Sub

End Sub

Private Function Preenche_Grid_Notas(objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Preenche_Grid_Notas

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridNotas)

    iIndice = 0

    For iIndice = 1 To objPedidoCompra.colNotasPedCompras.Count

        GridNotas.TextMatrix(iIndice, iGrid_NotaPC_Col) = objPedidoCompra.colNotasPedCompras.Item(iIndice)
        
    Next

    objGridNotas.iLinhasExistentes = iIndice - 1

    Preenche_Grid_Notas = SUCESSO

    Exit Function

Erro_Preenche_Grid_Notas:

    Preenche_Grid_Notas = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164396)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Notas(objGridInt As AdmGrid) As Long
'Executa a Inicialização do gridNotas

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Nota")
    
    ' campos de edição do grid
    objGridInt.colCampo.Add (NotaPC.Name)
    
    'indica onde estao situadas as colunas do grid
    iGrid_NotaPC_Col = 1

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridNotas

    'Linhas do grid
    objGridInt.objGrid.Rows = 25

    GridBloqueios.ColWidth(0) = 300

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 24

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Notas = SUCESSO

    Exit Function

End Function

'??? Já existe na tela de moedas
Public Function Moedas_Le(objMoedas As ClassMoedas) As Long

Dim lComando As Long
Dim lErro As Long
Dim sNome As String
Dim sSimbolo As String

On Error GoTo Erro_Moedas_Le

    'Abre Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 108818

    'Inicializa as strings
    sNome = String(STRING_NOME_MOEDA, 0)
    sSimbolo = String(STRING_SIMBOLO_MOEDA, 0)
    
    'Verifica se existe moeda com o codigo passado
    lErro = Comando_Executar(lComando, "SELECT Nome, Simbolo FROM Moedas WHERE Codigo = ?", sNome, sSimbolo, objMoedas.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 108819

    'Busca o registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 108820

    'Se nao encontrou => Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 108821

    'Transfere os dados
    objMoedas.sNome = sNome
    objMoedas.sSimbolo = sSimbolo
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    Moedas_Le = SUCESSO
    
    Exit Function

Exit Function

Erro_Moedas_Le:

    Moedas_Le = gErr

    Select Case gErr

        Case 108818
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 108819, 108820
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOEDAS", gErr)

        Case 108821

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164397)

    End Select

    Call Comando_Fechar(lComando)

End Function

'##############################################
'Inserido por Wagner
Private Sub Formata_Controles()

    PrecoUnitario.Format = gobjCOM.sFormatoPrecoUnitario
    PrecoUnitarioMoedaReal.Format = gobjCOM.sFormatoPrecoUnitario

End Sub
'##############################################

